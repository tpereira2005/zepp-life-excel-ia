import json
from datetime import datetime, timedelta, timezone

import pytest

from zepp_life_excel_ia import sintetico
from zepp_life_excel_ia.__main__ import main
from zepp_life_excel_ia.exportacao import ExportacaoInvalida, validar


def _registos(n=5):
    base = datetime(2026, 3, 1, 7, 0, tzinfo=timezone.utc)
    return [sintetico.registo(base + timedelta(days=i), 80 - i * 0.1, 22, 58) for i in range(n)]


def test_aceita_exportacao_que_acrescenta_pesagens():
    anterior = _registos(3)
    assert validar(_registos(5), anterior) == []


@pytest.mark.parametrize("nova, mensagem", [
    ([], "vazia"),
    ([{"Date": "2026-03-01T07:00:00", "Weight": 80}], "inválidas"),     # sem fuso
    ([{"Date": "2026-03-01T07:00:00Z", "Weight": 0}], "inválidas"),
    ([{"Date": "não é data", "Weight": 80}], "inválidas"),
])
def test_recusa_exportacao_invalida(nova, mensagem):
    with pytest.raises(ExportacaoInvalida, match=mensagem):
        validar(nova)


def test_duplicado_pelo_instante_mesmo_com_desvios_diferentes():
    a = _registos(1)[0]
    b = dict(a, Date=datetime.fromisoformat(a["Date"]).astimezone(timezone(timedelta(hours=1))).isoformat())
    with pytest.raises(ExportacaoInvalida, match="duplicadas"):
        validar([a, b])


def test_perda_de_pesagens_recusada_ou_aceite_com_permissao():
    anterior, nova = _registos(5), _registos(5)[:3]
    with pytest.raises(ExportacaoInvalida, match="perdeu 2 pesagem"):
        validar(nova, anterior)
    assert validar(nova, anterior, permitir_remocoes=True) == [
        "2 pesagem(ns) removida(s) na Zepp foram aceites (-PermitirRemocoes)."]


def test_linha_de_comandos(tmp_path, capsys):
    nova, anterior = tmp_path / "nova.json", tmp_path / "anterior.json"
    nova.write_text(json.dumps(_registos(3)), encoding="utf-8")
    anterior.write_text(json.dumps(_registos(4)), encoding="utf-8")
    assert main(["validar-exportacao", "--nova", str(nova), "--anterior", str(anterior)]) == 1
    assert "ERRO: A nova exportação perdeu 1 pesagem" in capsys.readouterr().err
    assert main(["validar-exportacao", "--nova", str(nova), "--anterior", str(anterior), "--permitir-remocoes"]) == 0
    assert main(["validar-exportacao", "--nova", str(nova), "--anterior", str(tmp_path / "nao-existe.json")]) == 0
