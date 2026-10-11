"""Leitura da exportação Zepp e seleção das pesagens aceites."""
from __future__ import annotations

import math
from dataclasses import dataclass, field
from datetime import datetime, timezone

from .config import Config
from .metricas import CAMPOS_COMPOSICAO, DA_EXPORTACAO, LIMITES


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
        if campo == "massa_gorda":
            return self.massa_gorda
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
        valores={m.campo: registo.get(m.campo) for m in DA_EXPORTACAO},
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
    leituras = sorted((converter(r, cfg) for r in registos), key=lambda lt: lt.instante_utc)
    instantes = [lt.id_medicao for lt in leituras]
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
