"""Cálculos derivados. São a fonte única para o JSON e a referência para validar as fórmulas do Excel."""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, timedelta
from statistics import fmean

from .config import Config
from .leituras import Leitura
from .medidas import Semana

# (chave JSON, cabeçalho Excel, campo da leitura)
METRICAS_RESUMO = [
    ("peso_medio_kg", "Peso médio (kg)", "Weight"),
    ("gordura_media_percentagem", "Gordura média (%)", "BodyFat"),
    ("musculo_medio_kg", "Músculo médio (kg)", "MuscleMass"),
    ("massa_gorda_media_kg", "Massa gorda média (kg)", "massa_gorda"),
    ("agua_media_percentagem", "Água média (%)", "BodyWater"),
    ("proteina_media_percentagem", "Proteína média (%)", "ProteinRatio"),
    ("massa_ossea_media_kg", "Massa óssea média (kg)", "BoneMass"),
    ("imc_medio", "IMC médio", "BMI"),
    ("gordura_visceral_media", "Gordura visceral média (índice)", "VisceralFat"),
    ("idade_corporal_media", "Idade corporal média (anos)", "MetabolicAge"),
    ("pontuacao_corporal_media", "Pontuação corporal média", "BodyScore"),
]
JANELA = timedelta(days=7)


def valor(leitura: Leitura, campo: str) -> float:
    return leitura.massa_gorda if campo == "massa_gorda" else leitura[campo]


def medias_moveis(leituras: list[Leitura], campos=("Weight", "BodyFat", "MuscleMass")) -> list[dict]:
    """Média das leituras em ]t − 7 dias, t], em hora local (igual à fórmula do Excel)."""
    resultado = []
    for atual in leituras:
        janela = [l for l in leituras if atual.local - JANELA < l.local <= atual.local]
        resultado.append({c: fmean(valor(l, c) for l in janela) for c in campos})
    return resultado


def _periodo(inicio: date, fim: date, leituras: list[Leitura], cfg: Config, ultimo_dia: date) -> dict:
    dentro = [l for l in leituras if inicio <= l.local.date() <= fim]
    p = {"inicio": inicio, "fim": fim, "leituras": len(dentro),
         "dias_com_medicao": len({l.local.date() for l in dentro}),
         "parcial": (inicio < cfg.inicio or fim > ultimo_dia) if dentro else None}
    for chave, _, campo in METRICAS_RESUMO:
        p[chave] = fmean(valor(l, campo) for l in dentro) if dentro else None
    return p


def _variacoes(periodos: list[dict]) -> None:
    anterior = None
    for p in periodos:
        for chave, base in (("variacao_peso_kg", "peso_medio_kg"), ("variacao_gordura_pp", "gordura_media_percentagem")):
            p[chave] = (p[base] - anterior[base]) if anterior and p[base] is not None and anterior[base] is not None else None
        anterior = p


def variacoes_cintura(semanas: list[Semana]) -> dict[date, float | None]:
    """Cintura menos a medição anterior disponível (por data de medição)."""
    medidas = [(s.data, s.cintura) for s in semanas if s.cintura is not None and s.cintura > 0]
    resultado = {}
    for s in semanas:
        if s.cintura is None:
            resultado[s.inicio] = None
            continue
        anteriores = [m for m in medidas if s.data is not None and m[0] < s.data]
        resultado[s.inicio] = s.cintura - max(anteriores)[1] if anteriores else None
    return resultado


def semanais(leituras: list[Leitura], semanas: list[Semana], cfg: Config) -> list[dict]:
    ultimo_dia = max(l.local for l in leituras).date()
    variacao = variacoes_cintura(semanas)
    periodos = []
    for s in semanas:
        p = _periodo(s.inicio, s.fim, leituras, cfg, ultimo_dia)
        p.update(rotulo=s.rotulo, cintura_cm=s.cintura, variacao_cintura_cm=variacao[s.inicio], pescoco_cm=s.pescoco)
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


def mensais(leituras: list[Leitura], semanas: list[Semana], cfg: Config, ate: date) -> list[dict]:
    ultimo_dia = max(l.local for l in leituras).date()
    periodos = []
    for inicio, fim in meses_entre(cfg.inicio, max(ate, ultimo_dia)):
        p = _periodo(inicio, fim, leituras, cfg, ultimo_dia)
        p["rotulo"] = f"{inicio:%Y-%m}"
        for chave, campo in (("cintura_media_cm", "cintura"), ("pescoco_medio_cm", "pescoco")):
            valores = [getattr(s, campo) for s in semanas
                       if getattr(s, campo) is not None and s.data is not None and inicio <= s.data <= fim]
            p[chave] = fmean(valores) if valores else None
        periodos.append(p)
    _variacoes(periodos)
    return periodos


@dataclass
class Indicador:
    inicio_media_7d: float
    atual_media_7d: float
    variacao: float
    ultima_leitura: float
    ritmo_semanal_4s: float | None


def indicadores(leituras: list[Leitura], semanas: list[Semana]) -> dict:
    # Janelas pela hora local, como no Excel (MIN/MAX da coluna Data/hora); a "última leitura"
    # é a de maior instante UTC (leituras[-1]), que pode não ter a maior hora local na mudança de hora.
    primeira, ultima = min(l.local for l in leituras), max(l.local for l in leituras)
    inicio = [l for l in leituras if l.local < primeira + JANELA]
    atual = [l for l in leituras if l.local > ultima - JANELA]
    base = [l for l in leituras if ultima - 4 * JANELA < l.local <= ultima - 3 * JANELA]
    resultado = {}
    for nome, campo in (("peso_kg", "Weight"), ("gordura_corporal_percentagem", "BodyFat"),
                        ("massa_muscular_kg", "MuscleMass")):
        a = fmean(valor(l, campo) for l in inicio)
        b = fmean(valor(l, campo) for l in atual)
        ritmo = (b - fmean(valor(l, campo) for l in base)) / 3 if base else None
        resultado[nome] = Indicador(a, b, b - a, valor(leituras[-1], campo), ritmo)
    for nome, campo in (("cintura_cm", "cintura"), ("pescoco_cm", "pescoco")):
        medidas = sorted((s.data, getattr(s, campo)) for s in semanas
                         if getattr(s, campo) is not None and getattr(s, campo) > 0 and s.data is not None)
        resultado[nome] = None if not medidas else {
            "primeira": medidas[0][1], "data_primeira": medidas[0][0],
            "ultima": medidas[-1][1], "data_ultima": medidas[-1][0],
            "variacao": medidas[-1][1] - medidas[0][1], "medicoes": len(medidas)}
    return resultado


def para_serie_excel(momento: datetime | date) -> float:
    base = datetime(1899, 12, 30)
    if not isinstance(momento, datetime):
        momento = datetime(momento.year, momento.month, momento.day)
    return (momento - base).total_seconds() / 86400
