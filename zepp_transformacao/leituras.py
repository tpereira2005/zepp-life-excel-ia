"""Leitura da exportação Zepp e seleção das pesagens aceites."""
from __future__ import annotations

import json
import math
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path

from .config import Config


@dataclass(frozen=True)
class Metrica:
    campo: str          # nome na exportação do leitor
    chave: str          # chave no JSON para IA
    nome: str           # nome apresentado no Excel
    unidade: str
    significado: str
    origem: str         # balança | aplicação | derivada


# Lista explícita: nada fora desta lista é exportado (nem token, conta ou identificadores).
METRICAS = [
    Metrica("Weight", "peso_kg", "Peso", "kg", "Peso medido pela balança", "balança"),
    Metrica("Height", "altura_cm", "Altura", "cm", "Altura do perfil associada à leitura", "aplicação"),
    Metrica("BMI", "imc", "IMC", "kg/m²", "Índice de massa corporal", "aplicação"),
    Metrica("BodyFat", "gordura_corporal_percentagem", "Gordura corporal", "%", "Estimativa por bioimpedância", "aplicação"),
    Metrica("BodyWater", "agua_corporal_percentagem", "Água corporal", "%", "Estimativa de água corporal", "aplicação"),
    Metrica("BoneMass", "massa_ossea_kg", "Massa óssea", "kg", "Estimativa; não é densidade mineral óssea", "aplicação"),
    Metrica("MuscleMass", "massa_muscular_kg", "Massa muscular", "kg", "Estimativa de massa muscular", "aplicação"),
    Metrica("ProteinRatio", "proteina_percentagem", "Proteína", "%", "Proporção estimada de proteína", "aplicação"),
    Metrica("ProteinMass", "massa_proteica_kg", "Massa proteica", "kg", "Derivada pelo leitor: peso × proteína / 100", "derivada"),
    Metrica("VisceralFat", "gordura_visceral_indice", "Gordura visceral", "índice", "Índice da aplicação; não é percentagem (inteiro no leitor)", "aplicação"),
    Metrica("BasalMetabolism", "metabolismo_basal_kcal_dia", "Metabolismo basal", "kcal/dia", "Estimativa de gasto basal (inteiro no leitor)", "aplicação"),
    Metrica("MetabolicAge", "idade_corporal_anos", "Idade corporal", "anos", "Estimativa; não é idade cronológica", "aplicação"),
    Metrica("BodyScore", "pontuacao_corporal", "Pontuação corporal", "pontos", "Pontuação global da aplicação", "aplicação"),
    Metrica("PhysiqueRating", "classificacao_fisica_codigo", "Classificação física", "código", "Código original, sem interpretação documentada", "aplicação"),
    Metrica("StandBodyWeight", "peso_padrao_kg", "Peso padrão", "kg", "Referência da aplicação; não é um objetivo individual", "aplicação"),
    Metrica("Impedance", "impedancia_ohm", "Impedância", "Ω", "Impedância registada pela balança", "balança"),
]

# Uma leitura só é aceite com composição corporal completa.
CAMPOS_COMPOSICAO = ["BMI", "BodyFat", "BodyWater", "BoneMass", "BasalMetabolism", "MuscleMass",
                     "VisceralFat", "ProteinMass", "ProteinRatio", "MetabolicAge", "BodyScore"]

# Limites largos: apanham erros de leitura evidentes, não julgam a composição corporal.
LIMITES = {
    "Weight": (30, 250), "BMI": (10, 60), "BodyFat": (2, 70), "BodyWater": (25, 80),
    "BoneMass": (0.5, 8), "MuscleMass": (15, 150), "ProteinRatio": (5, 35),
    "VisceralFat": (1, 59), "BasalMetabolism": (700, 4500), "MetabolicAge": (10, 100),
    "BodyScore": (0, 100),
}


@dataclass
class Leitura:
    instante_utc: datetime
    local: datetime                    # hora local sem fuso (como no Excel)
    desvio_utc: str                    # ex.: "+01:00"
    valores: dict = field(default_factory=dict)
    perfil_id: int | None = None

    @property
    def id_medicao(self) -> int:
        return int(self.instante_utc.timestamp())

    def __getitem__(self, campo: str):
        return self.valores.get(campo)

    @property
    def massa_gorda(self) -> float:
        return self["Weight"] * self["BodyFat"] / 100


@dataclass
class Selecao:
    aceites: list[Leitura]
    excluidas: list[tuple[Leitura, str]]   # apenas desde o início da transformação
    contagens: dict


def _numero(valor) -> bool:
    return isinstance(valor, (int, float)) and not isinstance(valor, bool) and math.isfinite(valor)


def carregar_exportacao(caminho: Path) -> list[dict]:
    dados = json.loads(Path(caminho).read_text(encoding="utf-8-sig"))
    if not isinstance(dados, list) or not dados:
        raise ValueError("Exportação vazia ou inválida.")
    return dados


def converter(registo: dict, cfg: Config) -> Leitura:
    data = registo.get("Date")
    peso = registo.get("Weight")
    if not isinstance(data, str) or not _numero(peso) or peso <= 0:
        raise ValueError(f"Pesagem inválida na exportação: {data!r}")
    instante = datetime.fromisoformat(data)
    if instante.tzinfo is None:
        raise ValueError(f"Data sem fuso horário na exportação: {data!r}")
    local = instante.astimezone(cfg.fuso)
    desvio = local.strftime("%z")
    perfil = registo.get("ProfileID")
    return Leitura(
        instante_utc=instante.astimezone(timezone.utc),
        local=local.replace(tzinfo=None),
        desvio_utc=f"{desvio[:3]}:{desvio[3:]}",
        valores={m.campo: registo.get(m.campo) for m in METRICAS},
        perfil_id=int(perfil) if _numero(perfil) else None,
    )


def motivo_exclusao(leitura: Leitura, cfg: Config) -> str | None:
    em_falta = [c for c in CAMPOS_COMPOSICAO if not _numero(leitura[c])]
    if em_falta:
        # O IMC é calculado pela aplicação só com o peso; o resto exige bioimpedância.
        if {c for c in CAMPOS_COMPOSICAO if c != "BMI"} <= set(em_falta):
            return "composição incompleta: só peso, sem bioimpedância"
        return "composição incompleta: falta " + ", ".join(em_falta)
    fora = [c for c, (lo, hi) in LIMITES.items() if not lo <= leitura[c] <= hi]
    if fora:
        return "valor implausível: " + ", ".join(fora)
    if leitura.perfil_id is not None:
        if leitura.perfil_id != cfg.perfil_titular_id:
            return "outro perfil"
    elif cfg.altura_perfil_sem_id_cm is None or leitura["Height"] != cfg.altura_perfil_sem_id_cm:
        return "outro perfil"
    return None


def selecionar(registos: list[dict], cfg: Config) -> Selecao:
    leituras = sorted((converter(r, cfg) for r in registos), key=lambda l: l.instante_utc)
    instantes = [l.id_medicao for l in leituras]
    if len(set(instantes)) != len(instantes):
        raise ValueError("Existem pesagens duplicadas no mesmo instante; rever a exportação.")
    contagens = {"origem": len(leituras), "anteriores_ao_inicio": 0, "incompletas": 0,
                 "implausiveis": 0, "outro_perfil": 0, "aceites": 0}
    aceites, excluidas = [], []
    for leitura in leituras:
        if leitura.local.date() < cfg.inicio:
            contagens["anteriores_ao_inicio"] += 1
            continue
        motivo = motivo_exclusao(leitura, cfg)
        if motivo is None:
            aceites.append(leitura)
            continue
        excluidas.append((leitura, motivo))
        chave = ("incompletas" if motivo.startswith("composição") else
                 "implausiveis" if motivo.startswith("valor") else "outro_perfil")
        contagens[chave] += 1
    contagens["aceites"] = len(aceites)
    if not aceites:
        raise ValueError("Não existem medições completas do perfil titular desde o início.")
    if sum(v for k, v in contagens.items() if k != "origem") != contagens["origem"]:
        raise AssertionError("Contagens de seleção não reconciliadas.")
    return Selecao(aceites, excluidas, contagens)
