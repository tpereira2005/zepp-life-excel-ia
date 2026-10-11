"""Gera o exemplo público (dados fictícios) em docs/exemplo/. Uso: .venv\\Scripts\\python scripts\\gerar_demo.py"""
import shutil
import sys
import tempfile
from datetime import date
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

from zepp_life_excel_ia import sintetico  # noqa: E402
from zepp_life_excel_ia.gerar import gerar  # noqa: E402

destino = RAIZ / "docs" / "exemplo"
destino.mkdir(parents=True, exist_ok=True)
with tempfile.TemporaryDirectory() as pasta:
    caminhos = sintetico.criar_projeto(Path(pasta), dias=150, semanas_csv=20)
    gerar(caminhos["origem"], destino / "exemplo-transformacao.xlsx", destino / "exemplo-dados-ia.json",
          config=caminhos["config"], hoje=date(2026, 6, 2))
    shutil.copy(caminhos["origem"], destino / "exemplo-exportacao-zepp.json")
    shutil.copy(caminhos["csv"], destino / "exemplo-medidas-manuais.csv")
print(f"Exemplo fictício criado em {destino}")
