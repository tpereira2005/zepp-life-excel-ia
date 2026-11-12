"""Livro Excel nativo (openpyxl): uma folha por módulo.

`construir` devolve a lista de células com fórmulas que importam e o valor que o Python espera em
cada uma; `verificacao` usa-a para confirmar o livro no Microsoft Excel real.
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.workbook.properties import CalcProperties

from ..medidas import FOLHA as FOLHA_MANUAL
from ..resumos import Calculos
from ..verificacao import Esperado
from . import folha_criterios, folha_fases, folha_manuais, folha_medicoes, folha_painel, folha_resumos
from .estilos import COR, ignorar_avisos_de_formula

FOLHAS = ["Painel", "Resumo Mensal", "Resumo Semanal", FOLHA_MANUAL, folha_medicoes.NOME, "Critérios e Qualidade"]


def folhas(calc: Calculos) -> list[str]:
    """Folhas do livro, por ordem. "Fases" só existe quando config.json define fases."""
    if not calc.fases:
        return list(FOLHAS)
    posicao = FOLHAS.index("Resumo Semanal") + 1
    return [*FOLHAS[:posicao], folha_fases.NOME, *FOLHAS[posicao:]]


def construir(caminho: Path, calc: Calculos, gerado_em: datetime) -> list[Esperado]:
    wb = Workbook()
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    wb.properties.creator = "zepp-life-excel-ia"
    wb.properties.title = "Transformação corporal"
    nomes = folhas(calc)
    wb.active.title = nomes[0]
    for nome in nomes[1:]:
        wb.create_sheet(nome)
    painel, mensal, semanal, manual, medicoes, criterios = (wb[n] for n in FOLHAS)
    inicio_txt = f"{calc.cfg.inicio:%d/%m/%Y}"
    atualizado = f"Atualizado em {gerado_em:%d/%m/%Y %H:%M}"
    esperados: list[Esperado] = []

    # As folhas de dados primeiro: as outras referem as colunas delas.
    folha_m = folha_medicoes.construir(medicoes, calc, inicio_txt, atualizado, esperados)
    folha_mm = folha_manuais.construir(manual, calc, esperados)
    colunas_semanal = folha_resumos.construir(semanal, calc.semanais, True, calc.cfg, folha_m, folha_mm, atualizado,
                                              esperados, calc.altura_cm)
    folha_resumos.construir(mensal, calc.mensais, False, calc.cfg, folha_m, folha_mm, atualizado, esperados)
    if calc.fases:
        folha_fases.construir(wb[folha_fases.NOME], calc.fases, folha_m, folha_mm, atualizado, esperados)
    folha_painel.construir(painel, calc, folha_m, folha_mm, colunas_semanal, inicio_txt, atualizado, esperados)
    folha_criterios.construir(criterios, calc, atualizado)

    painel.sheet_properties.tabColor = COR["azul"]
    manual.sheet_properties.tabColor = "F2C744"
    wb.active = 0
    wb.save(caminho)
    ignorar_avisos_de_formula(caminho)
    return esperados
