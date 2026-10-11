from datetime import date, timedelta
from itertools import pairwise
from statistics import fmean

import pytest

from zepp_life_excel_ia import resumos, sintetico
from zepp_life_excel_ia.leituras import selecionar
from zepp_life_excel_ia.medidas import Semana, combinar, ler_csv


@pytest.fixture
def dados(projeto):
    selecao = selecionar(sintetico.exportacao(), projeto["cfg"])
    semanas = combinar({}, ler_csv(projeto["csv"]), projeto["cfg"].inicio, selecao.aceites[-1].local.date())
    return projeto["cfg"], selecao.aceites, semanas


def test_media_semanal_por_leitura_e_dias(dados):
    cfg, leituras, semanas = dados
    for p in resumos.semanais(leituras, semanas, cfg):
        dentro = [lt for lt in leituras if p.inicio <= lt.local.date() <= p.fim]
        assert p.leituras == len(dentro)
        assert p.dias_com_medicao == len({lt.local.date() for lt in dentro})
        if dentro:
            assert p.medias["Weight"] == pytest.approx(fmean(lt["Weight"] for lt in dentro))
            assert p.medias["massa_gorda"] == pytest.approx(fmean(lt["Weight"] * lt["BodyFat"] / 100 for lt in dentro))
    assert any(p.leituras > p.dias_com_medicao for p in resumos.semanais(leituras, semanas, cfg))


def test_variacoes_entre_periodos_consecutivos(dados):
    cfg, leituras, semanas = dados
    periodos = resumos.semanais(leituras, semanas, cfg)
    assert periodos[0].variacao_peso_kg is None
    for a, b in pairwise(periodos):
        if a.medias["Weight"] is not None and b.medias["Weight"] is not None:
            assert b.variacao_peso_kg == pytest.approx(b.medias["Weight"] - a.medias["Weight"])
            assert b.variacao_gordura_pp == pytest.approx(b.medias["BodyFat"] - a.medias["BodyFat"])


def test_periodo_parcial(dados):
    cfg, leituras, semanas = dados
    mensais = resumos.mensais(leituras, semanas, cfg, leituras[-1].local.date())
    assert mensais[0].rotulo == "2026-01" and mensais[0].parcial is True     # começa a 01/01, antes do início
    assert mensais[1].parcial is False
    assert mensais[-1].parcial is True                                         # termina depois da última leitura


def test_variacao_cintura_salta_semanas_sem_medicao():
    d = date(2026, 1, 4)
    semanas = [Semana(d, d, 90.0, 39.0, None), Semana(d + timedelta(7), d + timedelta(7), None, None, "férias"),
               Semana(d + timedelta(14), d + timedelta(15), 88.5, 39.0, None)]
    v = resumos.variacoes_cintura(semanas)
    assert v[d] is None and v[d + timedelta(7)] is None
    assert v[d + timedelta(14)] == pytest.approx(-1.5)


def test_media_movel_sete_dias(dados):
    _, leituras, _ = dados
    moveis = resumos.medias_moveis(leituras)
    i = len(leituras) - 1
    janela = [lt["Weight"] for lt in leituras if leituras[i].local - timedelta(days=7) < lt.local <= leituras[i].local]
    assert moveis[i]["Weight"] == pytest.approx(fmean(janela))
    assert moveis[0]["Weight"] == leituras[0]["Weight"]


def test_indicadores(dados):
    _, leituras, semanas = dados
    ind = resumos.indicadores(leituras, semanas)
    peso = ind["peso_kg"]
    assert peso.variacao == pytest.approx(peso.atual_media_7d - peso.inicio_media_7d)
    assert peso.variacao < 0 and peso.ritmo_semanal_4s is not None
    assert ind["cintura_cm"].variacao == pytest.approx(ind["cintura_cm"].ultima - ind["cintura_cm"].primeira)


def test_eixo_legivel():
    from zepp_life_excel_ia.excel.graficos import eixo_legivel
    fundo, topo, passo = eixo_legivel([37.0, 37.1, 37.3], 3)
    assert fundo <= 37.0 and topo >= 37.3 and topo - fundo >= 3
    assert 5 <= round((topo - fundo) / passo) + 1 <= 7
    assert eixo_legivel([], 4) == (0, 4, 0.8)
