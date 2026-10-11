from datetime import date, datetime, timedelta, timezone

import pytest

from zepp_life_excel_ia import sintetico
from zepp_life_excel_ia.leituras import converter, selecionar


def test_contagens_reconciliadas_e_motivos(projeto):
    registos = sintetico.exportacao()
    selecao = selecionar(registos, projeto["cfg"])
    c = selecao.contagens
    assert c["origem"] == len(registos)
    assert c["aceites"] + c["anteriores_ao_inicio"] + c["incompletas"] + c["implausiveis"] + c["outro_perfil"] == c["origem"]
    assert c["anteriores_ao_inicio"] > 0 and c["incompletas"] > 0
    assert all(m == "composição incompleta: só peso, sem bioimpedância" for _, m in selecao.excluidas)
    assert [lt.instante_utc for lt in selecao.aceites] == sorted(lt.instante_utc for lt in selecao.aceites)
    assert all(lt.local.date() >= projeto["cfg"].inicio for lt in selecao.aceites)


def test_outro_perfil_e_implausivel_excluidos(projeto):
    base = datetime(2026, 2, 2, 7, 0, tzinfo=sintetico.LISBOA)
    registos = [sintetico.registo(base, 80, 22, 58),
                sintetico.registo(base + timedelta(days=1), 80, 22, 58, perfil=12345),
                sintetico.registo(base + timedelta(days=2), 80, 95, 58)]
    selecao = selecionar(registos, projeto["cfg"])
    motivos = sorted(m for _, m in selecao.excluidas)
    assert motivos[0] == "outro perfil" and motivos[1].startswith("valor implausível: BodyFat")
    assert selecao.contagens["outro_perfil"] == 1 and selecao.contagens["implausiveis"] == 1


def test_exportacao_antiga_sem_id_usa_altura(projeto):
    base = datetime(2026, 2, 2, 7, 0, tzinfo=sintetico.LISBOA)
    registos = [sintetico.registo(base, 80, 22, 58, perfil=None, altura=178),
                sintetico.registo(base + timedelta(days=1), 60, 22, 45, perfil=None, altura=160)]
    selecao = selecionar(registos, projeto["cfg"])
    assert len(selecao.aceites) == 1 and selecao.excluidas[0][1] == "outro perfil"


def test_mudanca_de_hora_nao_gera_duplicado(projeto):
    # 25/10/2026: 01:30 acontece duas vezes em Lisboa (UTC+1 e depois UTC+0).
    primeira = datetime(2026, 10, 25, 0, 30, tzinfo=timezone.utc)
    segunda = datetime(2026, 10, 25, 1, 30, tzinfo=timezone.utc)
    registos = [sintetico.registo(primeira, 80, 22, 58), sintetico.registo(segunda, 80.1, 22, 58)]
    selecao = selecionar(registos, projeto["cfg"])
    a, b = selecao.aceites
    assert a.local == b.local == datetime(2026, 10, 25, 1, 30)
    assert (a.desvio_utc, b.desvio_utc) == ("+01:00", "+00:00")
    assert a.id_medicao != b.id_medicao


def test_mesmo_instante_com_desvios_diferentes_e_duplicado(projeto):
    instante = datetime(2026, 3, 1, 8, 0, tzinfo=timezone.utc)
    a = sintetico.registo(instante, 80, 22, 58)
    b = dict(a, Date=instante.astimezone(timezone(timedelta(hours=1))).isoformat())
    with pytest.raises(ValueError, match="duplicadas"):
        selecionar([a, b], projeto["cfg"])


def test_converter_rejeita_dados_invalidos(projeto):
    with pytest.raises(ValueError):
        converter({"Date": "2026-03-01T08:00:00", "Weight": 80}, projeto["cfg"])  # sem fuso
    with pytest.raises(ValueError):
        converter({"Date": "2026-03-01T08:00:00Z", "Weight": 0}, projeto["cfg"])


def test_hora_local_lisboa_verao_e_inverno(projeto):
    inverno = converter(sintetico.registo(datetime(2026, 1, 10, 7, 0, tzinfo=timezone.utc), 80, 22, 58), projeto["cfg"])
    verao = converter(sintetico.registo(datetime(2026, 7, 10, 7, 0, tzinfo=timezone.utc), 80, 22, 58), projeto["cfg"])
    assert inverno.local.hour == 7 and verao.local.hour == 8
    assert inverno.local.date() == date(2026, 1, 10)
