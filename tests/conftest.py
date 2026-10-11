import sys
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(RAIZ))

from zepp_transformacao import config as configuracao, sintetico  # noqa: E402


@pytest.fixture
def projeto(tmp_path):
    caminhos = sintetico.criar_projeto(tmp_path)
    caminhos["cfg"] = configuracao.carregar(caminhos["config"])
    caminhos["pasta"] = tmp_path
    return caminhos
