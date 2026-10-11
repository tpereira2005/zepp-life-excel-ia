"""Compara os valores calculados pelo Microsoft Excel com os cálculos em Python.

O gerador do Excel regista, para cada célula com fórmula que importa, o valor que o Python espera
(`Esperado`). A verificação limita-se a abrir o livro no Excel e a comparar essa lista: os endereços
das células existem num único sítio (onde a célula é escrita).
"""
from __future__ import annotations

import json
import math
import os
import re
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path

from .config import RAIZ

SCRIPT = RAIZ / "scripts" / "verificar-excel-real.ps1"
TOLERANCIA = 1e-6


class ExcelIndisponivel(RuntimeError):
    pass


@dataclass(frozen=True)
class Esperado:
    folha: str
    celula: str
    valor: object        # número, texto, bool (Sim/Não) ou None (célula vazia ou "")
    descricao: str


FOLHAS_OBRIGATORIAS = ["Painel", "Resumo Mensal", "Resumo Semanal", "Medidas Manuais", "Medições Corporais"]


def calcular_no_excel(livro: Path, graficos: Path | None = None, timeout: int = 300, tentativas: int = 3) -> dict:
    """Valores calculados pelo Excel. Falhas intermitentes da automação COM (Excel ocupado, leitura
    incompleta) são repetidas até `tentativas` vezes; diferenças de valores nunca são repetidas."""
    if os.name != "nt" or not SCRIPT.exists():
        raise ExcelIndisponivel("Verificação no Excel só existe em Windows.")
    erro = None
    for tentativa in range(1, tentativas + 1):
        with tempfile.TemporaryDirectory() as pasta:
            saida = Path(pasta) / "valores.json"
            comando = ["powershell.exe", "-NoLogo", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(SCRIPT),
                       "-Livro", str(livro), "-Saida", str(saida), "-FolhasEsperadas", "|".join(FOLHAS_OBRIGATORIAS)]
            if graficos:
                comando += ["-Graficos", str(graficos)]
            try:
                processo = subprocess.run(comando, capture_output=True, text=True, timeout=timeout)
            except subprocess.TimeoutExpired:
                erro = RuntimeError(f"A verificação no Excel excedeu {timeout} s.")
                continue
            if processo.returncode == 3:
                raise ExcelIndisponivel("Microsoft Excel não está instalado.")
            if processo.returncode == 0 and saida.exists():
                valores = json.loads(saida.read_text(encoding="utf-8-sig"))
                em_falta = [f for f in FOLHAS_OBRIGATORIAS if f not in valores.get("folhas", {})]
                if not em_falta:
                    return valores
                erro = RuntimeError("O Excel não devolveu as folhas: " + ", ".join(em_falta))
            else:
                erro = RuntimeError("A verificação no Excel falhou: " + (processo.stderr or processo.stdout).strip()[-800:])
        if tentativa < tentativas:
            print(f"Aviso: verificação no Excel interrompida ({erro}); nova tentativa {tentativa + 1} de {tentativas}.")
    raise erro


def _celula(folha: list, endereco: str):
    letras, linha = re.fullmatch(r"([A-Z]+)(\d+)", endereco).groups()
    col = 0
    for letra in letras:
        col = col * 26 + ord(letra) - 64
    linha = int(linha)
    if linha - 1 >= len(folha) or col - 1 >= len(folha[linha - 1] or []):
        return None
    return folha[linha - 1][col - 1]


def _igual(excel, esperado) -> bool:
    if esperado is None:
        return excel in (None, "")
    if isinstance(esperado, bool):
        return excel == ("Sim" if esperado else "Não")
    if isinstance(esperado, str):
        return excel == esperado
    return isinstance(excel, (int, float)) and math.isclose(excel, esperado, rel_tol=1e-9, abs_tol=TOLERANCIA)


def comparar(valores: dict, esperados: list[Esperado]) -> list[str]:
    """Devolve a lista de diferenças (vazia = Excel e Python concordam)."""
    diferencas = [f"Erro de fórmula: {e}" for e in valores.get("erros", [])]
    folhas = valores["folhas"]
    em_falta = set(FOLHAS_OBRIGATORIAS) - set(folhas)
    if em_falta:
        return diferencas + ["Folhas não lidas no Excel: " + ", ".join(sorted(em_falta))]
    for e in esperados:
        obtido = _celula(folhas[e.folha], e.celula)
        if not _igual(obtido, e.valor):
            diferencas.append(f"{e.folha}!{e.celula} ({e.descricao}): Excel={obtido!r} Python={e.valor!r}")
    return diferencas
