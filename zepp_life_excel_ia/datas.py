"""Datas no formato do Excel (sistema de datas 1900: número de dias desde 30/12/1899)."""
from __future__ import annotations

import math
from datetime import date, datetime, timedelta

EPOCA = datetime(1899, 12, 30)


def serie_excel(momento: datetime | date) -> float:
    """Número de série Excel de uma data ou data/hora (sem fuso)."""
    if not isinstance(momento, datetime):
        momento = datetime(momento.year, momento.month, momento.day)
    return (momento.replace(tzinfo=None) - EPOCA).total_seconds() / 86400


def data_de_serie(valor: float) -> date:
    """Data (sem hora) correspondente a um número de série Excel."""
    return (EPOCA + timedelta(days=math.floor(valor))).date()
