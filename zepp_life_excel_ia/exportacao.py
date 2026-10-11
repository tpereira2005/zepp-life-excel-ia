"""Exportação do leitor Zepp: leitura e validação de uma nova descarga antes de substituir a anterior."""
from __future__ import annotations

import json
import math
from datetime import datetime
from pathlib import Path


class ExportacaoInvalida(ValueError):
    pass


def carregar(caminho: Path) -> list[dict]:
    dados = json.loads(Path(caminho).read_text(encoding="utf-8-sig"))
    if not isinstance(dados, list) or not dados:
        raise ExportacaoInvalida("Exportação vazia ou inválida.")
    return dados


def _instante(registo: dict) -> int | None:
    """Segundos Unix do instante da pesagem; None se a data for inválida ou não tiver fuso."""
    data = registo.get("Date") if isinstance(registo, dict) else None
    if not isinstance(data, str):
        return None
    try:
        momento = datetime.fromisoformat(data)
    except ValueError:
        return None
    return int(momento.timestamp()) if momento.tzinfo is not None else None


def _peso_valido(registo: dict) -> bool:
    peso = registo.get("Weight")
    return isinstance(peso, (int, float)) and not isinstance(peso, bool) and math.isfinite(peso) and peso > 0


def validar(nova: list[dict], anterior: list[dict] | None = None, permitir_remocoes: bool = False) -> list[str]:
    """Recusa (ExportacaoInvalida) uma descarga vazia, inválida, com duplicados ou que perca pesagens
    anteriores (comparação por instante UTC). Devolve avisos não bloqueantes."""
    if not isinstance(nova, list) or not nova:
        raise ExportacaoInvalida("A exportação está vazia.")
    instantes = set()
    for registo in nova:
        instante = _instante(registo)
        if instante is None or not _peso_valido(registo):
            raise ExportacaoInvalida("A exportação contém pesagens inválidas.")
        if instante in instantes:
            raise ExportacaoInvalida("A exportação contém pesagens duplicadas no mesmo instante.")
        instantes.add(instante)
    avisos = []
    perdidas = [r for r in anterior or [] if _instante(r) is not None and _instante(r) not in instantes]
    if perdidas:
        if not permitir_remocoes:
            exemplos = ", ".join(r["Date"] for r in perdidas[:5])
            raise ExportacaoInvalida(
                f"A nova exportação perdeu {len(perdidas)} pesagem(ns) anterior(es) ({exemplos}). "
                "Atualização interrompida. Se as apagou de propósito na Zepp Life, execute o atualizador "
                "com -PermitirRemocoes.")
        avisos.append(f"{len(perdidas)} pesagem(ns) removida(s) na Zepp foram aceites (-PermitirRemocoes).")
    return avisos
