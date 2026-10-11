"""Cálculos derivados. São a fonte única para o JSON e a referência para validar as fórmulas do Excel."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, timedelta
from statistics import fmean

from .config import Config
from .leituras import Leitura, Selecao
from .medidas import Semana
from .metricas import INDICADORES, MEDIAS_MOVEIS, POR_CAMPO, RESUMO

JANELA = timedelta(days=7)


@dataclass
class Periodo:
    """Uma semana ou um mês de resumo."""
    rotulo: str
    inicio: date
    fim: date
    leituras: int
    dias_com_medicao: int
    parcial: bool | None                  # None quando não há leituras
    medias: dict[str, float | None]       # por campo, na ordem de metricas.RESUMO
    variacao_peso_kg: float | None = None
    variacao_gordura_pp: float | None = None
    manuais: dict[str, float | None] = field(default_factory=dict)   # cintura/pescoço (chave JSON)


@dataclass
class Indicador:
    inicio_media_7d: float
    atual_media_7d: float
    variacao: float
    ultima_leitura: float
    ritmo_semanal_4s: float | None


@dataclass
class IndicadorManual:
    primeira: float
    data_primeira: date
    ultima: float
    data_ultima: date
    variacao: float
    medicoes: int


@dataclass
class Calculos:
    """Tudo o que o Excel e o JSON precisam, calculado uma única vez."""
    cfg: Config
    selecao: Selecao
    semanas: list[Semana]
    semanais: list[Periodo]
    mensais: list[Periodo]
    moveis: list[dict[str, float]]                       # por leitura: campo -> média de 7 dias
    indicadores: dict[str, Indicador | IndicadorManual | None]
    variacoes_cintura: dict[date, float | None]          # por início de semana

    @property
    def leituras(self) -> list[Leitura]:
        return self.selecao.aceites


def medias_moveis(leituras: list[Leitura]) -> list[dict[str, float]]:
    """Média das leituras em ]t − 7 dias, t], em hora local (igual à fórmula do Excel)."""
    resultado = []
    for atual in leituras:
        janela = [lt for lt in leituras if atual.local - JANELA < lt.local <= atual.local]
        resultado.append({m.campo: fmean(lt[m.campo] for lt in janela) for m in MEDIAS_MOVEIS})
    return resultado


def _periodo(rotulo: str, inicio: date, fim: date, leituras: list[Leitura], cfg: Config, ultimo_dia: date) -> Periodo:
    dentro = [lt for lt in leituras if inicio <= lt.local.date() <= fim]
    return Periodo(
        rotulo=rotulo, inicio=inicio, fim=fim, leituras=len(dentro),
        dias_com_medicao=len({lt.local.date() for lt in dentro}),
        parcial=(inicio < cfg.inicio or fim > ultimo_dia) if dentro else None,
        medias={c: fmean(lt[c] for lt in dentro) if dentro else None for c in RESUMO})


def _variacoes(periodos: list[Periodo]) -> None:
    """Média do período menos a do período anterior (peso em kg; gordura em pontos percentuais)."""
    for anterior, p in zip([None, *periodos], periodos, strict=False):
        for atributo, campo in (("variacao_peso_kg", "Weight"), ("variacao_gordura_pp", "BodyFat")):
            a, b = (anterior.medias[campo] if anterior else None), p.medias[campo]
            setattr(p, atributo, b - a if a is not None and b is not None else None)


def variacoes_cintura(semanas: list[Semana]) -> dict[date, float | None]:
    """Cintura menos a medição anterior disponível (por data de medição)."""
    medidas = [(s.data, s.cintura) for s in semanas if s.cintura is not None and s.cintura > 0]
    resultado = {}
    for s in semanas:
        anteriores = [m for m in medidas if s.data is not None and m[0] < s.data] if s.cintura is not None else []
        resultado[s.inicio] = s.cintura - max(anteriores)[1] if anteriores else None
    return resultado


def _ultimo_dia(leituras: list[Leitura]) -> date:
    return max(lt.local for lt in leituras).date()


def semanais(leituras: list[Leitura], semanas: list[Semana], cfg: Config) -> list[Periodo]:
    ultimo_dia = _ultimo_dia(leituras)
    variacao = variacoes_cintura(semanas)
    periodos = []
    for s in semanas:
        p = _periodo(s.rotulo, s.inicio, s.fim, leituras, cfg, ultimo_dia)
        p.manuais = {"cintura_cm": s.cintura, "variacao_cintura_cm": variacao[s.inicio], "pescoco_cm": s.pescoco}
        periodos.append(p)
    _variacoes(periodos)
    return periodos


def meses_entre(primeiro: date, ultimo: date) -> list[tuple[date, date]]:
    meses, atual = [], primeiro.replace(day=1)
    while atual <= ultimo:
        seguinte = (atual.replace(day=28) + timedelta(days=4)).replace(day=1)
        meses.append((atual, seguinte - timedelta(days=1)))
        atual = seguinte
    return meses


def mensais(leituras: list[Leitura], semanas: list[Semana], cfg: Config, ate: date) -> list[Periodo]:
    ultimo_dia = _ultimo_dia(leituras)
    periodos = []
    for inicio, fim in meses_entre(cfg.inicio, max(ate, ultimo_dia)):
        p = _periodo(f"{inicio:%Y-%m}", inicio, fim, leituras, cfg, ultimo_dia)
        for chave, campo in (("cintura_media_cm", "cintura"), ("pescoco_medio_cm", "pescoco")):
            valores = [getattr(s, campo) for s in semanas
                       if getattr(s, campo) is not None and s.data is not None and inicio <= s.data <= fim]
            p.manuais[chave] = fmean(valores) if valores else None
        periodos.append(p)
    _variacoes(periodos)
    return periodos


def indicadores(leituras: list[Leitura], semanas: list[Semana]) -> dict[str, Indicador | IndicadorManual | None]:
    """Chaves: chave JSON da métrica (peso_kg, …) e cintura_cm/pescoco_cm."""
    # Janelas pela hora local, como no Excel (MIN/MAX da coluna Data/hora); a "última leitura"
    # é a de maior instante UTC (leituras[-1]), que pode não ter a maior hora local na mudança de hora.
    primeira, ultima = min(lt.local for lt in leituras), max(lt.local for lt in leituras)
    inicio = [lt for lt in leituras if lt.local < primeira + JANELA]
    atual = [lt for lt in leituras if lt.local > ultima - JANELA]
    base = [lt for lt in leituras if ultima - 4 * JANELA < lt.local <= ultima - 3 * JANELA]
    resultado: dict[str, Indicador | IndicadorManual | None] = {}
    for campo in INDICADORES:
        a = fmean(lt[campo] for lt in inicio)
        b = fmean(lt[campo] for lt in atual)
        ritmo = (b - fmean(lt[campo] for lt in base)) / 3 if base else None
        resultado[POR_CAMPO[campo].chave] = Indicador(a, b, b - a, leituras[-1][campo], ritmo)
    for nome, campo in (("cintura_cm", "cintura"), ("pescoco_cm", "pescoco")):
        medidas = sorted((s.data, getattr(s, campo)) for s in semanas
                         if getattr(s, campo) is not None and getattr(s, campo) > 0 and s.data is not None)
        resultado[nome] = None if not medidas else IndicadorManual(
            medidas[0][1], medidas[0][0], medidas[-1][1], medidas[-1][0], medidas[-1][1] - medidas[0][1], len(medidas))
    return resultado


def calcular(cfg: Config, selecao: Selecao, semanas: list[Semana], ate: date) -> Calculos:
    leituras = selecao.aceites
    return Calculos(
        cfg=cfg, selecao=selecao, semanas=semanas,
        semanais=semanais(leituras, semanas, cfg),
        mensais=mensais(leituras, semanas, cfg, ate),
        moveis=medias_moveis(leituras),
        indicadores=indicadores(leituras, semanas),
        variacoes_cintura=variacoes_cintura(semanas))
