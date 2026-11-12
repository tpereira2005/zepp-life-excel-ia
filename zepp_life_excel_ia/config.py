"""Configuração pessoal: lida de config.json (privado) com valores de config.example.json."""
from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from zoneinfo import ZoneInfo

RAIZ = Path(__file__).resolve().parents[1]
SEXOS = ("masculino", "feminino")


@dataclass(frozen=True)
class Fase:
    """Fase opcional da transformação (ex.: défice, manutenção). Termina na véspera da fase seguinte."""
    nome: str
    inicio: date
    objetivo: str | None = None


@dataclass(frozen=True)
class Config:
    inicio: date
    fuso: ZoneInfo
    perfil_titular_id: int
    altura_perfil_sem_id_cm: float | None
    protocolo_pesagem: str
    csv_medidas_iniciais: Path | None
    # Opcionais (2.2.0): contexto para os indicadores e para a IA.
    sexo: str | None = None                  # "masculino" | "feminino": necessário para a estimativa pela fita
    ano_nascimento: int | None = None
    objetivo: str | None = None              # texto livre
    fases: tuple[Fase, ...] = ()

    @property
    def fuso_nome(self) -> str:
        return self.fuso.key


def _fases(valor, inicio: date) -> tuple[Fase, ...]:
    if not valor:
        return ()
    if not isinstance(valor, list):
        raise ValueError("config.json: 'fases' tem de ser uma lista.")
    fases = []
    for f in valor:
        if not isinstance(f, dict) or not str(f.get("nome", "")).strip() or not f.get("inicio"):
            raise ValueError("config.json: cada fase precisa de 'nome' e 'inicio' (AAAA-MM-DD).")
        objetivo = f.get("objetivo")
        fases.append(Fase(str(f["nome"]).strip(), date.fromisoformat(f["inicio"]), str(objetivo) if objetivo else None))
    datas = [f.inicio for f in fases]
    if datas != sorted(set(datas)):
        raise ValueError("config.json: as fases têm de estar por ordem de data, sem datas repetidas.")
    if datas[0] < inicio:
        raise ValueError("config.json: a primeira fase não pode começar antes de 'inicio_transformacao'.")
    return tuple(fases)


def carregar(caminho: Path | None = None) -> Config:
    caminho = caminho or RAIZ / "config.json"
    if not caminho.exists():
        raise FileNotFoundError(
            f"Falta {caminho.name}. Copie config.example.json para config.json e ajuste os valores."
        )
    dados = json.loads(caminho.read_text(encoding="utf-8"))
    csv = dados.get("csv_medidas_iniciais")
    csv_path = None
    if csv:
        csv_path = Path(csv)
        if not csv_path.is_absolute():
            csv_path = caminho.parent / csv_path
    altura = dados.get("altura_perfil_sem_id_cm")
    sexo = dados.get("sexo")
    if sexo is not None:
        sexo = str(sexo).strip().lower()
        if sexo not in SEXOS:
            raise ValueError("config.json: 'sexo' tem de ser \"masculino\" ou \"feminino\" (ou ser omitido).")
    ano = dados.get("ano_nascimento")
    if ano is not None and (isinstance(ano, bool) or not isinstance(ano, int) or not 1900 <= ano <= 2100):
        raise ValueError("config.json: 'ano_nascimento' tem de ser um ano (ex.: 1990) ou ser omitido.")
    inicio = date.fromisoformat(dados["inicio_transformacao"])
    return Config(
        inicio=inicio,
        fuso=ZoneInfo(dados.get("fuso_horario", "Europe/Lisbon")),
        perfil_titular_id=int(dados.get("perfil_titular_id", -1)),
        altura_perfil_sem_id_cm=float(altura) if altura is not None else None,
        protocolo_pesagem=dados.get("protocolo_pesagem", ""),
        csv_medidas_iniciais=csv_path,
        sexo=sexo,
        ano_nascimento=ano,
        objetivo=str(dados["objetivo"]) if dados.get("objetivo") else None,
        fases=_fases(dados.get("fases"), inicio),
    )
