"""Compara os valores calculados pelo Microsoft Excel com os cálculos em Python."""
from __future__ import annotations

import json
import math
import os
import subprocess
import tempfile
from pathlib import Path

from .config import RAIZ
from .resumos import METRICAS_RESUMO, Indicador

SCRIPT = RAIZ / "scripts" / "verificar-excel-real.ps1"
TOLERANCIA = 1e-6


class ExcelIndisponivel(RuntimeError):
    pass


def calcular_no_excel(livro: Path, graficos: Path | None = None, timeout: int = 300) -> dict:
    if os.name != "nt" or not SCRIPT.exists():
        raise ExcelIndisponivel("Verificação no Excel só existe em Windows.")
    with tempfile.TemporaryDirectory() as pasta:
        saida = Path(pasta) / "valores.json"
        comando = ["powershell.exe", "-NoLogo", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(SCRIPT),
                   "-Livro", str(livro), "-Saida", str(saida)]
        if graficos:
            comando += ["-Graficos", str(graficos)]
        processo = subprocess.run(comando, capture_output=True, text=True, timeout=timeout)
        if processo.returncode == 3:
            raise ExcelIndisponivel("Microsoft Excel não está instalado.")
        if processo.returncode != 0 or not saida.exists():
            raise RuntimeError("A verificação no Excel falhou: " + (processo.stderr or processo.stdout).strip()[-800:])
        return json.loads(saida.read_text(encoding="utf-8-sig"))


def _celula(folha: list, endereco: str):
    col = 0
    letras = "".join(c for c in endereco if c.isalpha())
    for letra in letras:
        col = col * 26 + ord(letra) - 64
    linha = int(endereco[len(letras):])
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


def comparar(valores: dict, semanais: list[dict], mensais: list[dict], moveis: list[dict], massas_gordas: list[float],
             indicadores: dict, variacoes_cintura: list, estados: list[str]) -> list[str]:
    """Devolve a lista de diferenças (vazia = Excel e Python concordam)."""
    diferencas = [f"Erro de fórmula: {e}" for e in valores.get("erros", [])]
    folhas = valores["folhas"]

    def verificar(folha, endereco, esperado, descricao):
        obtido = _celula(folhas[folha], endereco)
        if not _igual(obtido, esperado):
            diferencas.append(f"{folha}!{endereco} ({descricao}): Excel={obtido!r} Python={esperado!r}")

    colunas = ["G", "I"] + [chr(ord("K") + i) for i in range(len(METRICAS_RESUMO) - 2)]
    for folha, periodos, extras in (("Resumo Semanal", semanais, [("T", "cintura_cm"), ("U", "variacao_cintura_cm"), ("V", "pescoco_cm")]),
                                    ("Resumo Mensal", mensais, [("T", "cintura_media_cm"), ("U", "pescoco_medio_cm")])):
        for i, p in enumerate(periodos):
            r = i + 4
            verificar(folha, f"D{r}", p["leituras"], "leituras")
            verificar(folha, f"E{r}", p["dias_com_medicao"], "dias")
            verificar(folha, f"F{r}", p["parcial"], "parcial")
            verificar(folha, f"H{r}", p["variacao_peso_kg"], "variação do peso")
            verificar(folha, f"J{r}", p["variacao_gordura_pp"], "variação da gordura")
            for col, (chave, _, _) in zip(colunas, METRICAS_RESUMO):
                verificar(folha, f"{col}{r}", p[chave], chave)
            for col, chave in extras:
                verificar(folha, f"{col}{r}", p[chave], chave)
    for i, (m, gorda) in enumerate(zip(moveis, massas_gordas)):
        r = i + 4
        verificar("Medições Corporais", f"I{r}", gorda, "massa gorda")
        verificar("Medições Corporais", f"N{r}", m["Weight"], "peso 7 dias")
        verificar("Medições Corporais", f"O{r}", m["BodyFat"], "gordura 7 dias")
        verificar("Medições Corporais", f"P{r}", m["MuscleMass"], "músculo 7 dias")
    nomes = {"Completa": "completa", "Incompleta": "incompleta", "Por preencher": "por_preencher", "Sem medição": "sem_medicao"}
    for i, (variacao, estado) in enumerate(zip(variacoes_cintura, estados)):
        r = i + 4
        verificar("Medidas Manuais", f"D{r}", variacao, "variação da cintura")
        obtido = _celula(folhas["Medidas Manuais"], f"F{r}")
        if nomes.get(obtido) != estado:
            diferencas.append(f"Medidas Manuais!F{r}: Excel={obtido!r} Python={estado!r}")
    peso, gordura, musculo = (indicadores[k] for k in ("peso_kg", "gordura_corporal_percentagem", "massa_muscular_kg"))
    for endereco, esperado, descricao in (
            ("A5", peso.ultima_leitura, "último peso"), ("E5", peso.atual_media_7d, "peso 7 dias"),
            ("I5", peso.variacao, "peso desde o início"), ("M5", peso.ritmo_semanal_4s, "ritmo"),
            ("A11", gordura.atual_media_7d, "gordura 7 dias"), ("E11", gordura.variacao, "gordura desde o início"),
            ("I11", musculo.atual_media_7d, "músculo 7 dias"), ("M11", musculo.variacao, "músculo desde o início")):
        verificar("Painel", endereco, esperado, descricao)
    for col_valor, col_var, nome in (("A", "E", "cintura_cm"), ("I", "M", "pescoco_cm")):
        ind = indicadores[nome]
        verificar("Painel", f"{col_valor}17", ind["ultima"] if ind else None, f"última {nome}")
        verificar("Painel", f"{col_var}17", ind["variacao"] if ind and ind["medicoes"] > 1 else None, f"{nome} desde o início")
    return diferencas
