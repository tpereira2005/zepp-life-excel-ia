import pytest

from zepp_life_excel_ia import config as configuracao
from zepp_life_excel_ia import sintetico


@pytest.fixture
def projeto(tmp_path):
    caminhos = sintetico.criar_projeto(tmp_path)
    caminhos["cfg"] = configuracao.carregar(caminhos["config"])
    caminhos["pasta"] = tmp_path
    return caminhos
