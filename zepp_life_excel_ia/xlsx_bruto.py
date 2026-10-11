"""Lê os valores guardados no OOXML sem adivinhar tipos a partir de formatos personalizados."""
from __future__ import annotations

import posixpath
import re
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from .datas import serie_excel

NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
ERROS = ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!", "#SPILL!", "#CALC!")


def _coluna(endereco: str) -> int:
    n = 0
    for letra in re.match(r"[A-Z]+", endereco).group(0):
        n = n * 26 + ord(letra) - 64
    return n


def ler_livro(caminho: Path, permitir_erros: bool = False) -> dict[str, list[list]]:
    """Devolve {folha: linhas}; datas ficam como número de série Excel (sistema 1900)."""
    folhas = {}
    with zipfile.ZipFile(caminho) as z:
        textos = []
        if "xl/sharedStrings.xml" in z.namelist():
            for si in ET.fromstring(z.read("xl/sharedStrings.xml")):
                # Texto simples (si/t) ou rico (si/r/t); os textos fonéticos (si/rPh) são ignorados.
                textos.append("".join(t.text or "" for t in si.findall("s:t", NS) + si.findall("s:r/s:t", NS)))
        rels = {r.attrib["Id"]: r.attrib["Target"]
                for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
        livro = ET.fromstring(z.read("xl/workbook.xml"))
        props = livro.find("s:workbookPr", NS)
        if props is not None and props.get("date1904") in ("1", "true"):
            raise ValueError("O sistema de datas 1904 não é suportado.")
        for folha in livro.find("s:sheets", NS):
            alvo = rels[folha.attrib[REL_ID]]
            alvo = alvo.lstrip("/") if alvo.startswith("/") else posixpath.normpath(posixpath.join("xl", alvo))
            doc = ET.fromstring(z.read(alvo))
            linhas: list[list] = []
            for linha in doc.findall("s:sheetData/s:row", NS):
                valores: list = []
                for celula in linha:
                    endereco, tipo = celula.get("r"), celula.get("t")
                    no = celula.find("s:v", NS)
                    valor = no.text if no is not None else None
                    if tipo == "s" and valor is not None:
                        valor = textos[int(valor)]
                    elif tipo == "inlineStr":
                        valor = "".join(n.text or "" for n in celula.findall(".//s:t", NS))
                    elif tipo == "d" and valor is not None:
                        # OOXML permite datas ISO 8601 além de números de série.
                        instante = datetime.fromisoformat(valor.replace("Z", "+00:00"))
                        if instante.tzinfo is not None:
                            instante = instante.astimezone(timezone.utc)
                        valor = serie_excel(instante)
                    elif tipo == "b" and valor is not None:
                        valor = valor == "1"
                    elif tipo == "e":
                        if not permitir_erros:
                            raise ValueError(f"Erro de fórmula em {folha.get('name')}!{endereco}: {valor}")
                    elif tipo not in ("str",) and valor is not None:
                        valor = float(valor)
                    col = _coluna(endereco)
                    while len(valores) < col:
                        valores.append(None)
                    valores[col - 1] = valor
                numero = int(linha.get("r"))
                while len(linhas) < numero:
                    linhas.append([])
                linhas[numero - 1] = valores
            folhas[folha.get("name")] = linhas
    return folhas
