"""Configuração pessoal: lida de config.json (privado) com valores de config.example.json."""
from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from zoneinfo import ZoneInfo

RAIZ = Path(__file__).resolve().parents[1]


@dataclass(frozen=True)
class Config:
    inicio: date
    fuso: ZoneInfo
    perfil_titular_id: int
    altura_perfil_sem_id_cm: float | None
    protocolo_pesagem: str
    csv_medidas_iniciais: Path | None

    @property
    def fuso_nome(self) -> str:
        return self.fuso.key


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
    return Config(
        inicio=date.fromisoformat(dados["inicio_transformacao"]),
        fuso=ZoneInfo(dados.get("fuso_horario", "Europe/Lisbon")),
        perfil_titular_id=int(dados.get("perfil_titular_id", -1)),
        altura_perfil_sem_id_cm=float(altura) if altura is not None else None,
        protocolo_pesagem=dados.get("protocolo_pesagem", ""),
        csv_medidas_iniciais=csv_path,
    )
