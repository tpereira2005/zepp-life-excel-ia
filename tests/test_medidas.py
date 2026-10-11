from datetime import date
from itertools import pairwise

import pytest
from openpyxl import Workbook

from zepp_life_excel_ia import medidas
from zepp_life_excel_ia.medidas import Semana, combinar, domingo, leituras_das_notas, ler_csv, ler_existentes, rotulo


def test_semana_domingo_a_sabado():
    assert domingo(date(2026, 10, 11)) == date(2026, 10, 11)   # domingo
    assert domingo(date(2026, 10, 17)) == date(2026, 10, 11)   # sábado
    assert domingo(date(2026, 10, 12)) == date(2026, 10, 11)   # segunda
    assert rotulo(date(2026, 10, 11)) == "11/10/2026 - 17/10/2026"
    assert medidas.inicio_do_rotulo("11/10/2026 - 17/10/2026") == date(2026, 10, 11)
    with pytest.raises(ValueError):
        medidas.inicio_do_rotulo("12/10/2026 - 18/10/2026")      # não começa ao domingo


def test_csv_inicial(projeto):
    semanas = ler_csv(projeto["csv"])
    s = semanas[date(2026, 1, 4)]
    assert s.cintura is not None and s.notas.startswith("Leituras — cintura:")
    esquecida = semanas[date(2026, 2, 1)]
    assert (esquecida.cintura, esquecida.pescoco, esquecida.notas) == (None, None, "Sem medição: esqueci-me")


def test_leituras_das_notas():
    c, p = leituras_das_notas("Leituras — cintura: 92,4 - 92,5 - 92,3; pescoço: 38 - 38,1 - 38")
    assert c == [92.4, 92.5, 92.3] and p == [38.0, 38.1, 38.0]
    assert leituras_das_notas("Sem medição: férias") == (None, None)
    assert leituras_das_notas(None) == (None, None)


def _livro(caminho, linhas, cabecalhos=medidas.CABECALHOS):
    wb = Workbook()
    ws = wb.active
    ws.title = medidas.FOLHA
    ws.append(["Medidas manuais semanais"])
    ws.append([])
    ws.append(cabecalhos)
    for linha in linhas:
        ws.append(linha)
    wb.save(caminho)


def test_le_por_cabecalho_mesmo_com_coluna_inserida(tmp_path):
    cab = ["Semana", "Nova coluna", "Data da medição", "Cintura (cm)", "Variação da cintura (cm)", "Pescoço (cm)",
           "Estado", "Observações"]
    caminho = tmp_path / "livro.xlsx"
    _livro(caminho, [["04/01/2026 - 10/01/2026", "x", date(2026, 1, 5), 90.5, None, 39.0, "Completa", "nota"]], cab)
    s = ler_existentes(caminho)[date(2026, 1, 4)]
    assert (s.data, s.cintura, s.pescoco, s.notas) == (date(2026, 1, 5), 90.5, 39.0, "nota")


def test_aceita_data_e_numero_escritos_como_texto(tmp_path):
    caminho = tmp_path / "livro.xlsx"
    _livro(caminho, [["04/01/2026 - 10/01/2026", "05/01/2026", "90,5", None, "39.2", None, None]])
    s = ler_existentes(caminho)[date(2026, 1, 4)]
    assert s.data == date(2026, 1, 5)          # dd/mm/aaaa, nunca formato americano
    assert (s.cintura, s.pescoco) == (90.5, 39.2)


def test_rejeita_texto_nao_numerico(tmp_path):
    caminho = tmp_path / "livro.xlsx"
    _livro(caminho, [["04/01/2026 - 10/01/2026", date(2026, 1, 4), "noventa", None, None, None, None]])
    with pytest.raises(ValueError, match="não numérica"):
        ler_existentes(caminho)


def test_semana_duplicada_e_coluna_em_falta(tmp_path):
    caminho = tmp_path / "livro.xlsx"
    linha = ["04/01/2026 - 10/01/2026", date(2026, 1, 4), 90, None, 39, None, None]
    _livro(caminho, [linha, linha])
    with pytest.raises(ValueError, match="duplicada"):
        ler_existentes(caminho)
    _livro(caminho, [linha], ["Semana", "Data da medição", "Cintura (cm)"])
    with pytest.raises(ValueError, match="Colunas em falta"):
        ler_existentes(caminho)


def test_prioridade_excel_sobre_csv_incluindo_apagados(projeto):
    csv = ler_csv(projeto["csv"])
    inicio = date(2026, 1, 4)
    existentes = {inicio: Semana(inicio, date(2026, 1, 6), 91.0, None, None, "excel")}  # pescoço e nota apagados
    semanas = combinar(existentes, csv, inicio, date(2026, 4, 30))
    primeira = semanas[0]
    assert (primeira.data, primeira.cintura, primeira.pescoco, primeira.notas) == (date(2026, 1, 6), 91.0, None, None)
    assert semanas[1].origem == "csv"
    assert semanas[-1].inicio == domingo(date(2026, 4, 30)) and semanas[-1].origem == "nova"
    assert all((b.inicio - a.inicio).days == 7 for a, b in pairwise(semanas))


def test_semanas_existentes_fora_do_intervalo_sao_mantidas():
    antiga = date(2025, 12, 28)
    existentes = {antiga: Semana(antiga, antiga, 95.0, 40.0, None, "excel")}
    semanas = combinar(existentes, {}, date(2026, 1, 4), date(2026, 1, 20))
    assert semanas[0].inicio == antiga


def test_data_fora_da_semana_e_valor_fora_dos_limites():
    inicio = date(2026, 1, 4)
    with pytest.raises(ValueError, match="não pertence"):
        combinar({inicio: Semana(inicio, date(2026, 1, 11), 90.0, None, None, "excel")}, {}, inicio, inicio)
    with pytest.raises(ValueError, match="Falta a data"):
        combinar({inicio: Semana(inicio, None, 90.0, None, None, "excel")}, {}, inicio, inicio)
    with pytest.raises(ValueError, match="fora dos limites"):
        combinar({inicio: Semana(inicio, inicio, 900.0, None, None, "excel")}, {}, inicio, inicio)


def test_estado():
    d = date(2026, 1, 4)
    assert Semana(d, d, 90, 39, None).estado == "completa"
    assert Semana(d, d, 90, None, None).estado == "incompleta"
    assert Semana(d, d, None, None, "Sem medição: férias").estado == "sem_medicao"
    assert Semana(d, d, None, None, None).estado == "por_preencher"
