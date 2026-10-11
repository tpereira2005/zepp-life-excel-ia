"""Gera o exemplo público (dados fictícios) em docs/exemplo/. Uso: .venv\\Scripts\\python scripts\\gerar_demo.py"""
import shutil
import sys
import tempfile
from argparse import Namespace
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

from zepp_transformacao import gerar, sintetico  # noqa: E402

destino = RAIZ / "docs" / "exemplo"
destino.mkdir(parents=True, exist_ok=True)
with tempfile.TemporaryDirectory() as pasta:
    caminhos = sintetico.criar_projeto(Path(pasta), dias=150, semanas_csv=20)
    gerar.executar(Namespace(origem=str(caminhos["origem"]), excel=str(destino / "exemplo-transformacao.xlsx"),
                             json=str(destino / "exemplo-dados-ia.json"), existente=None,
                             config=str(caminhos["config"]), hoje="2026-06-02", sem_excel_real=False,
                             graficos=None))
    shutil.copy(caminhos["origem"], destino / "exemplo-exportacao-zepp.json")
    shutil.copy(caminhos["csv"], destino / "exemplo-medidas-manuais.csv")
print(f"Exemplo fictício criado em {destino}")
