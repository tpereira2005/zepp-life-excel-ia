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
    # Com as opções de config.json que a versão 2.2.0 acrescentou: sexo (estimativa pela fita), objetivo e fases.
    caminhos = sintetico.criar_projeto(Path(pasta), dias=150, semanas_csv=20, viragem=0.7, extra={
        "sexo": "masculino", "ano_nascimento": 1990, "objetivo": "Perder gordura e depois manter o peso",
        "fases": [{"nome": "Défice", "inicio": "2026-01-04", "objetivo": "Perder cerca de 0,5 kg por semana"},
                  {"nome": "Manutenção", "inicio": "2026-04-19", "objetivo": "Estabilizar o peso"}]})
    gerar(caminhos["origem"], destino / "exemplo-transformacao.xlsx", destino / "exemplo-dados-ia.json",
          config=caminhos["config"], hoje=date(2026, 6, 2), json_resumo=destino / "exemplo-dados-ia-resumo.json")
    shutil.copy(caminhos["config"], destino / "exemplo-config.json")
    shutil.copy(caminhos["origem"], destino / "exemplo-exportacao-zepp.json")
    shutil.copy(caminhos["csv"], destino / "exemplo-medidas-manuais.csv")
print(f"Exemplo fictício criado em {destino}")
