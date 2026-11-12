"""Cálculos derivados. São a fonte única para o JSON e a referência para validar as fórmulas do Excel."""
from __future__ import annotations

import math
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from statistics import fmean, median

from . import metricas
from .config import Config
from .leituras import Leitura, Selecao
from .medidas import Semana
from .metricas import INDICADORES, MEDIAS_MOVEIS, POR_CAMPO, RESUMO

JANELA = timedelta(days=7)
DIA = timedelta(days=1)
# Os extremos da média de 7 dias só contam a partir do 7.º dia (antes disso a "média" tem poucos dias).
EXTREMOS_DESDE = timedelta(days=6)
TOLERANCIA_EXTREMO = 1e-6          # empates na média móvel: fica a leitura mais antiga (igual no Excel)
VARIACAO_MINIMA_FRACAO = 0.5       # kg: abaixo disto não se calcula "% da variação que foi gordura"
FASE_CAMPOS = ["Weight", "BodyFat", "massa_gorda", "massa_magra"]
FASE_DIAS_MINIMOS_RITMO = 21


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
class Extremo:
    valor: float
    quando: datetime                      # hora local da leitura


@dataclass
class Indicador:
    inicio_media_7d: float
    atual_media_7d: float
    variacao: float
    ultima_leitura: float
    ritmo_semanal_4s: float | None        # declive da regressão linear dos últimos 28 dias, por semana
    # Média de 7 dias desta métrica na leitura em que o PESO teve a média de 7 dias mais baixa.
    no_minimo_do_peso: float | None = None
    variacao_desde_minimo_do_peso: float | None = None
    minimo: Extremo | None = None         # da própria média de 7 dias (só métricas com média móvel)
    maximo: Extremo | None = None


@dataclass
class IndicadorManual:
    primeira: float
    data_primeira: date
    ultima: float
    data_ultima: date
    variacao: float
    medicoes: int
    minima: float | None = None
    data_minima: date | None = None       # a mais antiga, em caso de empate
    variacao_desde_minima: float | None = None


@dataclass
class Composicao:
    """De quanto variou o peso, quanto foi massa gorda e quanto foi massa magra (médias de 7 dias)."""
    peso_kg: float
    massa_gorda_kg: float
    massa_magra_kg: float
    fracao_gordura: float | None          # massa gorda / peso; None se o peso variou menos de 0,5 kg


@dataclass
class ResumoFase:
    nome: str
    objetivo: str | None
    inicio: date
    fim: date
    dias: int
    leituras: int
    ultima_leitura: datetime | None
    inicial: dict[str, float | None]      # por campo de FASE_CAMPOS: média dos primeiros 7 dias da fase
    final: dict[str, float | None]        # média dos 7 dias até à última leitura da fase
    ritmo_peso_semanal: float | None      # variação do peso / semanas entre as duas médias
    cintura_inicial: float | None
    cintura_final: float | None

    def variacao(self, campo: str) -> float | None:
        a, b = self.inicial[campo], self.final[campo]
        return b - a if a is not None and b is not None else None

    @property
    def variacao_cintura(self) -> float | None:
        a, b = self.cintura_inicial, self.cintura_final
        return b - a if a is not None and b is not None else None


@dataclass
class Atipica:
    leitura: Leitura
    motivos: list[str]


@dataclass
class Qualidade:
    atipicas: list[Atipica]
    dias_do_periodo: int                  # da primeira à última leitura, inclusive
    dias_com_leitura: int
    dias_com_varias_leituras: int
    maior_intervalo_sem_leituras_dias: int
    dias_desde_ultima_leitura: int        # em relação à data de referência
    data_referencia: date                 # dia da atualização (nunca anterior à última leitura)


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
    altura_cm: float | None = None                       # altura do perfil (última leitura)
    minimo_peso: Extremo | None = None                   # mínimo da média de 7 dias do peso
    dias_ritmo: list[float | None] = field(default_factory=list)   # por leitura: dias até à última (≤ 0) ou None
    composicao_desde_inicio: Composicao | None = None
    composicao_desde_minimo: Composicao | None = None
    cintura_altura: float | None = None                  # última cintura / altura
    fases: list[ResumoFase] = field(default_factory=list)
    qualidade: Qualidade | None = None

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


def altura_do_perfil(leituras: list[Leitura], cfg: Config) -> float | None:
    """Altura (cm) usada nos indicadores com fita métrica: a da última leitura."""
    altura = leituras[-1]["Height"]
    if isinstance(altura, (int, float)) and not isinstance(altura, bool) and altura > 0:
        return float(altura)
    return cfg.altura_perfil_sem_id_cm


def cintura_altura(cintura: float | None, altura: float | None) -> float | None:
    return cintura / altura if cintura is not None and altura else None


def gordura_fita(cintura: float | None, pescoco: float | None, altura: float | None, sexo: str | None) -> float | None:
    """Estimativa de gordura (%) pelo método da Marinha dos EUA (Hodgdon e Beckett), em centímetros.
    Só para sexo masculino: a fórmula feminina exige a anca, que não é registada."""
    if sexo != "masculino" or cintura is None or pescoco is None or not altura or cintura <= pescoco:
        return None
    return 495 / (1.0324 - 0.19077 * math.log10(cintura - pescoco) + 0.15456 * math.log10(altura)) - 450


def semanais(leituras: list[Leitura], semanas: list[Semana], cfg: Config) -> list[Periodo]:
    ultimo_dia = _ultimo_dia(leituras)
    variacao = variacoes_cintura(semanas)
    altura = altura_do_perfil(leituras, cfg)
    periodos = []
    for s in semanas:
        p = _periodo(s.rotulo, s.inicio, s.fim, leituras, cfg, ultimo_dia)
        p.manuais = {"cintura_cm": s.cintura, "variacao_cintura_cm": variacao[s.inicio], "pescoco_cm": s.pescoco,
                     "cintura_altura": cintura_altura(s.cintura, altura),
                     "gordura_fita_percentagem": gordura_fita(s.cintura, s.pescoco, altura, cfg.sexo)}
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


def dias_para_o_ritmo(leituras: list[Leitura]) -> list[float | None]:
    """Por leitura: dias até à última leitura (≤ 0) se estiver nos últimos 28 dias; senão None.
    É a coluna auxiliar do Excel sobre a qual se calcula o declive (SLOPE)."""
    ultima = max(lt.local for lt in leituras)
    limite = ultima - timedelta(days=metricas.RITMO_DIAS)
    return [(lt.local - ultima).total_seconds() / 86400 if lt.local > limite else None for lt in leituras]


def ritmo_semanal(leituras: list[Leitura], campo: str) -> float | None:
    """Declive da regressão linear das leituras dos últimos 28 dias, em unidades por semana.
    None com menos de 8 leituras ou se estas cobrirem menos de 14 dias."""
    pontos = [(x, lt[campo]) for x, lt in zip(dias_para_o_ritmo(leituras), leituras, strict=True) if x is not None]
    amplitude = -min((x for x, _ in pontos), default=0)
    # Meio segundo de margem, como nas fórmulas do Excel (ver excel/estilos.py).
    if len(pontos) < metricas.RITMO_MINIMO_LEITURAS or amplitude < metricas.RITMO_MINIMO_AMPLITUDE_DIAS - 0.5 / 86400:
        return None
    mx, my = fmean(x for x, _ in pontos), fmean(y for _, y in pontos)
    return 7 * sum((x - mx) * (y - my) for x, y in pontos) / sum((x - mx) ** 2 for x, _ in pontos)


def _media_ate(leituras: list[Leitura], campo: str, quando: datetime) -> float:
    return fmean(lt[campo] for lt in leituras if quando - JANELA < lt.local <= quando)


def extremos(leituras: list[Leitura], moveis: list[dict[str, float]], campo: str) -> tuple[Extremo, Extremo] | None:
    """Mínimo e máximo da média de 7 dias (a partir do 7.º dia de leituras); em empate, a leitura mais antiga."""
    desde = min(lt.local for lt in leituras) + EXTREMOS_DESDE
    validos = [(mm[campo], lt.local) for lt, mm in zip(leituras, moveis, strict=True) if lt.local >= desde]
    if not validos:
        return None
    vmin, vmax = min(v for v, _ in validos), max(v for v, _ in validos)
    return (Extremo(vmin, min(q for v, q in validos if v <= vmin + TOLERANCIA_EXTREMO)),
            Extremo(vmax, min(q for v, q in validos if v >= vmax - TOLERANCIA_EXTREMO)))


def _indicador_manual(medidas: list[tuple[date, float]]) -> IndicadorManual | None:
    if not medidas:
        return None
    medidas = sorted(medidas)
    (d0, v0), (d1, v1) = medidas[0], medidas[-1]
    minima = min(v for _, v in medidas)
    return IndicadorManual(v0, d0, v1, d1, v1 - v0, len(medidas), minima,
                           min(d for d, v in medidas if v == minima), v1 - minima)


def indicadores(leituras: list[Leitura], semanas: list[Semana], moveis: list[dict[str, float]] | None = None,
                cfg: Config | None = None) -> dict[str, Indicador | IndicadorManual | None]:
    """Chaves: chave JSON da métrica (peso_kg, …), cintura_cm, pescoco_cm e gordura_fita_percentagem."""
    # Janelas pela hora local, como no Excel (MIN/MAX da coluna Data/hora); a "última leitura"
    # é a de maior instante UTC (leituras[-1]), que pode não ter a maior hora local na mudança de hora.
    moveis = moveis if moveis is not None else medias_moveis(leituras)
    primeira, ultima = min(lt.local for lt in leituras), max(lt.local for lt in leituras)
    inicio = [lt for lt in leituras if lt.local < primeira + JANELA]
    atual = [lt for lt in leituras if lt.local > ultima - JANELA]
    extremo_peso = extremos(leituras, moveis, "Weight")
    com_media_movel = {m.campo for m in MEDIAS_MOVEIS}
    resultado: dict[str, Indicador | IndicadorManual | None] = {}
    for campo in INDICADORES:
        a = fmean(lt[campo] for lt in inicio)
        b = fmean(lt[campo] for lt in atual)
        ind = Indicador(a, b, b - a, leituras[-1][campo], ritmo_semanal(leituras, campo))
        if extremo_peso:
            ind.no_minimo_do_peso = _media_ate(leituras, campo, extremo_peso[0].quando)
            ind.variacao_desde_minimo_do_peso = b - ind.no_minimo_do_peso
        if campo in com_media_movel and extremo_peso:
            ind.minimo, ind.maximo = extremos(leituras, moveis, campo)
        resultado[POR_CAMPO[campo].chave] = ind
    altura = altura_do_perfil(leituras, cfg) if cfg else None
    sexo = cfg.sexo if cfg else None
    medidas = [s for s in semanas if s.data is not None]
    series = {
        "cintura_cm": [(s.data, s.cintura) for s in medidas if s.cintura is not None and s.cintura > 0],
        "pescoco_cm": [(s.data, s.pescoco) for s in medidas if s.pescoco is not None and s.pescoco > 0],
        "gordura_fita_percentagem": [(s.data, g) for s in medidas
                                     if (g := gordura_fita(s.cintura, s.pescoco, altura, sexo)) is not None],
    }
    resultado.update({nome: _indicador_manual(valores) for nome, valores in series.items()})
    return resultado


def _composicao(ind: dict, atributo: str) -> Composicao | None:
    peso, gorda, magra = (getattr(ind[POR_CAMPO[c].chave], atributo) for c in ("Weight", "massa_gorda", "massa_magra"))
    if peso is None:
        return None
    return Composicao(peso, gorda, magra, gorda / peso if abs(peso) >= VARIACAO_MINIMA_FRACAO else None)


def fases(leituras: list[Leitura], semanas: list[Semana], cfg: Config) -> list[ResumoFase]:
    """Resumo de cada fase configurada: médias de 7 dias no início e no fim, e a variação entre elas."""
    ultimo_dia = _ultimo_dia(leituras)
    resultado = []
    for f, seguinte in zip(cfg.fases, [*cfg.fases[1:], None], strict=False):
        # Uma fase em curso termina na última leitura (não numa data futura), para os dias e o ritmo serem reais.
        fim = max(f.inicio, ultimo_dia)
        if seguinte:
            fim = min(fim, seguinte.inicio - DIA)
        dentro = [lt for lt in leituras if f.inicio <= lt.local.date() <= fim]
        ultima = max((lt.local for lt in dentro), default=None)
        primeiros = [lt for lt in dentro if lt.local.date() < f.inicio + JANELA]
        ultimos = [lt for lt in dentro if lt.local > ultima - JANELA] if dentro else []
        media = lambda grupo, campo: fmean(lt[campo] for lt in grupo) if grupo else None
        inicial = {c: media(primeiros, c) for c in FASE_CAMPOS}
        final = {c: media(ultimos, c) for c in FASE_CAMPOS}
        dias = (fim - f.inicio).days + 1
        cinturas = sorted((s.data, s.cintura) for s in semanas
                          if s.cintura is not None and s.cintura > 0 and s.data is not None and f.inicio <= s.data <= fim)
        fase = ResumoFase(f.nome, f.objetivo, f.inicio, fim, dias, len(dentro), ultima, inicial, final, None,
                          cinturas[0][1] if cinturas else None, cinturas[-1][1] if cinturas else None)
        variacao = fase.variacao("Weight")
        if variacao is not None and dias >= FASE_DIAS_MINIMOS_RITMO:
            fase.ritmo_peso_semanal = variacao / ((dias - 7) / 7)
        resultado.append(fase)
    return resultado


def _numero(valor) -> bool:
    return isinstance(valor, (int, float)) and not isinstance(valor, bool)


def atipicas(leituras: list[Leitura]) -> list[Atipica]:
    """Leituras que se afastam da mediana das outras leituras a ±7 dias. São só assinaladas."""
    vizinhanca = timedelta(days=metricas.ATIPICA_VIZINHANCA_DIAS)
    limites = ([(c, v, False) for c, v in metricas.ATIPICA_ABSOLUTA.items()]
               + [(c, v, True) for c, v in metricas.ATIPICA_RELATIVA.items()])
    resultado = []
    for lt in leituras:
        vizinhas = [o for o in leituras if o is not lt and abs(o.local - lt.local) <= vizinhanca]
        motivos = []
        for campo, limite, relativo in limites:
            valores = [o[campo] for o in vizinhas if _numero(o[campo])]
            if not _numero(lt[campo]) or len(valores) < metricas.ATIPICA_MINIMO_VIZINHAS:
                continue
            centro = median(valores)
            desvio = lt[campo] - centro
            m = POR_CAMPO[campo]
            referencia = f"face à mediana das leituras a ±{metricas.ATIPICA_VIZINHANCA_DIAS} dias ({centro:g})"
            if relativo and centro and abs(desvio) / centro > limite:
                motivos.append(f"{m.nome}: {lt[campo]:g} {m.unidade}, {desvio / centro:+.0%} {referencia}")
            elif not relativo and abs(desvio) > limite:
                motivos.append(f"{m.nome}: {lt[campo]:g} {m.unidade}, {desvio:+.2f} {referencia}")
        if motivos:
            resultado.append(Atipica(lt, motivos))
    return resultado


def qualidade(leituras: list[Leitura], hoje: date) -> Qualidade:
    por_dia: dict[date, int] = {}
    for lt in leituras:
        por_dia[lt.local.date()] = por_dia.get(lt.local.date(), 0) + 1
    dias = sorted(por_dia)
    intervalos = [(b - a).days - 1 for a, b in zip(dias, dias[1:], strict=False)]
    return Qualidade(
        atipicas=atipicas(leituras),
        dias_do_periodo=(dias[-1] - dias[0]).days + 1,
        dias_com_leitura=len(dias),
        dias_com_varias_leituras=sum(1 for n in por_dia.values() if n > 1),
        maior_intervalo_sem_leituras_dias=max(intervalos, default=0),
        dias_desde_ultima_leitura=max((hoje - dias[-1]).days, 0),
        data_referencia=max(hoje, dias[-1]))


def calcular(cfg: Config, selecao: Selecao, semanas: list[Semana], ate: date, hoje: date | None = None) -> Calculos:
    leituras = selecao.aceites
    moveis = medias_moveis(leituras)
    ind = indicadores(leituras, semanas, moveis, cfg)
    altura = altura_do_perfil(leituras, cfg)
    extremo_peso = extremos(leituras, moveis, "Weight")
    cintura = ind["cintura_cm"]
    return Calculos(
        cfg=cfg, selecao=selecao, semanas=semanas,
        semanais=semanais(leituras, semanas, cfg),
        mensais=mensais(leituras, semanas, cfg, ate),
        moveis=moveis,
        indicadores=ind,
        variacoes_cintura=variacoes_cintura(semanas),
        altura_cm=altura,
        minimo_peso=extremo_peso[0] if extremo_peso else None,
        dias_ritmo=dias_para_o_ritmo(leituras),
        composicao_desde_inicio=_composicao(ind, "variacao"),
        composicao_desde_minimo=_composicao(ind, "variacao_desde_minimo_do_peso"),
        cintura_altura=cintura_altura(cintura.ultima if cintura else None, altura),
        fases=fases(leituras, semanas, cfg),
        qualidade=qualidade(leituras, hoje or ate))
