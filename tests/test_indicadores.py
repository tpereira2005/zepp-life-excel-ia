"""Indicadores da versão 2.2.0: massa magra, mínimo do peso, ritmo por regressão, fita métrica, fases,
leituras atípicas, qualidade dos dados, JSON compacto e cópia das medidas manuais."""
import json
from datetime import date, datetime, timedelta
from pathlib import Path

import jsonschema
import pytest
from openpyxl import load_workbook

from zepp_life_excel_ia import config as configuracao
from zepp_life_excel_ia import gerar, medidas, resumos, sintetico
from zepp_life_excel_ia.leituras import selecionar
from zepp_life_excel_ia.medidas import Semana

ESQUEMAS = Path(__file__).resolve().parents[1] / "schema"
INICIO = sintetico.INICIO


def leituras_de(cfg, pesos, gordura=20.0, hora=8):
    """Uma leitura por dia a partir do início, com os pesos dados."""
    registos = [sintetico.registo(datetime(INICIO.year, INICIO.month, INICIO.day, hora, tzinfo=sintetico.LISBOA)
                                  + timedelta(days=d), peso, gordura, 58) for d, peso in enumerate(pesos)]
    return selecionar(registos, cfg).aceites


def configurar(projeto, **extra):
    dados = json.loads(projeto["config"].read_text(encoding="utf-8"))
    dados.update(extra)
    projeto["config"].write_text(json.dumps(dados, ensure_ascii=False), encoding="utf-8")
    projeto["cfg"] = configuracao.carregar(projeto["config"])
    return projeto


def correr(projeto, nome="livro", hoje="2026-05-02", excel_real=False, **extra):
    pasta = projeto["pasta"]
    resultado = gerar.gerar(projeto["origem"], pasta / f"{nome}.xlsx", pasta / f"{nome}.json", config=projeto["config"],
                            hoje=date.fromisoformat(hoje), verificar_excel=excel_real, **extra)
    return pasta / f"{nome}.xlsx", json.loads((pasta / f"{nome}.json").read_text(encoding="utf-8")), resultado


def test_massa_magra_fecha_a_conta_com_o_peso(projeto):
    leituras = selecionar(sintetico.exportacao(), projeto["cfg"]).aceites
    for lt in leituras:
        assert lt.massa_gorda + lt.massa_magra == pytest.approx(lt["Weight"])
    calc = resumos.calcular(projeto["cfg"], selecionar(sintetico.exportacao(), projeto["cfg"]), [],
                            leituras[-1].local.date())
    c = calc.composicao_desde_inicio
    assert c.massa_gorda_kg + c.massa_magra_kg == pytest.approx(c.peso_kg)
    assert c.fracao_gordura == pytest.approx(c.massa_gorda_kg / c.peso_kg)


def test_minimo_do_peso_ignora_os_primeiros_dias_e_mede_a_recuperacao(projeto):
    # A 1.ª leitura é a mais baixa de todas, mas a média de 7 dias só conta a partir do 7.º dia.
    pesos = [60.0] + [80 - 0.5 * d for d in range(1, 30)] + [65.5 + 0.2 * d for d in range(1, 21)]
    leituras = leituras_de(projeto["cfg"], pesos)
    moveis = resumos.medias_moveis(leituras)
    minimo, maximo = resumos.extremos(leituras, moveis, "Weight")
    # As janelas que acabam nos dias 33 e 34 têm a mesma média (66,0): fica a mais antiga.
    assert minimo.quando.date() == INICIO + timedelta(days=33)
    assert minimo.valor == pytest.approx(66.0) and moveis[34]["Weight"] == pytest.approx(66.0)
    assert maximo.quando.date() == INICIO + timedelta(days=7)
    peso = resumos.indicadores(leituras, [], moveis, projeto["cfg"])["peso_kg"]
    assert peso.no_minimo_do_peso == pytest.approx(minimo.valor)
    assert peso.variacao_desde_minimo_do_peso == pytest.approx(peso.atual_media_7d - minimo.valor)
    assert peso.variacao_desde_minimo_do_peso > 2 and peso.variacao < 0


def test_sem_sete_dias_de_leituras_nao_ha_minimo(projeto):
    leituras = leituras_de(projeto["cfg"], [80, 79.8, 79.9])
    assert resumos.extremos(leituras, resumos.medias_moveis(leituras), "Weight") is None
    peso = resumos.indicadores(leituras, [], None, projeto["cfg"])["peso_kg"]
    assert peso.variacao_desde_minimo_do_peso is None and peso.minimo is None


def test_ritmo_e_o_declive_dos_ultimos_28_dias(projeto):
    cfg = projeto["cfg"]
    # Sobe 2 kg em 30 dias e depois desce 0,1 kg por dia: só a descida recente conta.
    pesos = [78 + d / 15 for d in range(30)] + [80 - 0.1 * d for d in range(40)]
    assert resumos.ritmo_semanal(leituras_de(cfg, pesos), "Weight") == pytest.approx(-0.7)
    assert resumos.ritmo_semanal(leituras_de(cfg, [80 - 0.1 * d for d in range(7)]), "Weight") is None      # < 8 leituras
    assert resumos.ritmo_semanal(leituras_de(cfg, [80 - 0.1 * d for d in range(13)]), "Weight") is None     # < 14 dias
    assert resumos.ritmo_semanal(leituras_de(cfg, [80 - 0.1 * d for d in range(15)]), "Weight") == pytest.approx(-0.7)


def test_gordura_pela_fita_e_cintura_altura():
    assert resumos.gordura_fita(90, 38, 178, "masculino") == pytest.approx(20.15, abs=0.05)
    assert resumos.gordura_fita(90, 38, 178, "feminino") is None       # exigiria a anca
    assert resumos.gordura_fita(90, 38, 178, None) is None
    assert resumos.gordura_fita(38, 38, 178, "masculino") is None      # cintura ≤ pescoço: sem logaritmo
    assert resumos.gordura_fita(90, None, 178, "masculino") is None
    assert resumos.cintura_altura(89, 178) == pytest.approx(0.5)
    assert resumos.cintura_altura(None, 178) is None and resumos.cintura_altura(89, None) is None


def test_indicador_manual_guarda_o_minimo_mais_antigo(projeto):
    d = INICIO
    semanas = [Semana(d + timedelta(7 * i), d + timedelta(7 * i), c, 38.0, None)
               for i, c in enumerate([92.0, 88.0, 86.0, 86.0, 89.5])]
    semanas.insert(2, Semana(d + timedelta(70), d + timedelta(70), None, None, "férias"))
    leituras = leituras_de(projeto["cfg"], [80] * 10)
    cintura = resumos.indicadores(leituras, semanas, None, projeto["cfg"])["cintura_cm"]
    assert (cintura.minima, cintura.data_minima) == (86.0, d + timedelta(14))
    assert cintura.variacao_desde_minima == pytest.approx(3.5) and cintura.variacao == pytest.approx(-2.5)


def test_leituras_atipicas_sao_assinaladas_e_nunca_excluidas(projeto):
    registos = sintetico.exportacao()
    normal = selecionar(registos, projeto["cfg"])
    momento = datetime(2026, 3, 10, 13, 0, tzinfo=sintetico.LISBOA)
    registos.append(sintetico.registo(momento, 84.9, 23.0, 59))        # ~3 kg acima das leituras vizinhas
    selecao = selecionar(registos, projeto["cfg"])
    assert len(selecao.aceites) == len(normal.aceites) + 1              # continua aceite
    novas = [a for a in resumos.atipicas(selecao.aceites) if a.leitura.local == momento.replace(tzinfo=None)]
    assert len(novas) == 1 and novas[0].motivos[0].startswith("Peso: 84.9 kg, +")
    calc = resumos.calcular(projeto["cfg"], selecao, [], date(2026, 5, 2))
    assert calc.qualidade.atipicas and sum(p.leituras for p in calc.mensais) == len(selecao.aceites)


def test_qualidade_dos_dados(projeto):
    momentos = [datetime(2026, 1, dia, 8, minuto, tzinfo=sintetico.LISBOA)
                for dia, minuto in ((4, 0), (5, 0), (6, 0), (10, 0), (10, 5))]
    registos = [sintetico.registo(m, 80, 20, 58) for m in momentos]
    q = resumos.qualidade(selecionar(registos, projeto["cfg"]).aceites, date(2026, 1, 15))
    assert (q.dias_do_periodo, q.dias_com_leitura, q.dias_com_varias_leituras) == (7, 4, 1)
    assert q.maior_intervalo_sem_leituras_dias == 3 and q.dias_desde_ultima_leitura == 5
    assert q.data_referencia == date(2026, 1, 15)


def test_fases_configuradas(projeto):
    configurar(projeto, sexo="masculino", ano_nascimento=1990, objetivo="Chegar aos 78 kg", fases=[
        {"nome": "Défice", "inicio": "2026-01-04", "objetivo": "Perder gordura"},
        {"nome": "Manutenção", "inicio": "2026-03-29"},
        {"nome": "Futura", "inicio": "2027-01-01"}])
    livro, dados, resultado = correr(projeto, excel_real=True)         # compara com o Excel real, se existir
    jsonschema.validate(dados, json.loads((ESQUEMAS / "dados-zepp-life-ia.schema.json").read_text(encoding="utf-8")))
    defice, manutencao, futura = dados["fases"]
    assert (defice["inicio"], defice["fim"], defice["dias"]) == ("2026-01-04", "2026-03-28", 84)
    assert manutencao["fim"] == dados["metadados"]["ultima_medicao_corporal_local"][:10]     # termina na última leitura
    assert defice["peso_kg"]["variacao"] < 0 and defice["ritmo_peso_kg_por_semana"] < 0
    assert defice["massa_gorda_kg"]["variacao"] + defice["massa_magra_kg"]["variacao"] == pytest.approx(
        defice["peso_kg"]["variacao"], abs=0.002)
    assert defice["cintura_cm"]["variacao"] < 0
    assert futura["leituras"] == 0 and futura["peso_kg"]["media_7d_inicial"] is None
    # A data de referência nunca é anterior à última leitura (aqui, 03/05).
    assert dados["contexto"] == {"data_referencia": "2026-05-03", "altura_cm": 178.0, "sexo": "masculino",
                                 "ano_nascimento": 1990, "idade_aproximada_anos": 36, "objetivo": "Chegar aos 78 kg"}
    wb = load_workbook(livro)
    assert wb.sheetnames == ["Painel", "Resumo Mensal", "Resumo Semanal", "Fases", "Medidas Manuais",
                             "Medições Corporais", "Critérios e Qualidade"]
    assert [c.value for c in wb["Fases"]["A"][3:6]] == ["Défice", "Manutenção", "Futura"]
    assert "Fases" in {e.folha for e in resultado.esperados}
    # Com sexo masculino há estimativa pela fita no resumo semanal, nas medidas manuais e nos indicadores.
    cabecalhos = [c.value for c in wb["Resumo Semanal"][3]]
    assert "Gordura pela fita (%)" in cabecalhos and "Cintura ÷ altura" in cabecalhos
    medida = next(m for m in dados["medidas_manuais"] if m["cintura_cm"] is not None)
    assert medida["gordura_fita_percentagem"] == pytest.approx(
        resumos.gordura_fita(medida["cintura_cm"], medida["pescoco_cm"], 178, "masculino"), abs=0.001)
    assert dados["indicadores"]["gordura_fita_percentagem"]["medicoes"] >= 10
    assert wb["Painel"]["N22"].value == "Gordura pela fita"


def test_sem_sexo_nao_ha_estimativa_pela_fita(projeto):
    livro, dados, _ = correr(projeto)
    assert dados["indicadores"]["gordura_fita_percentagem"] is None and dados["fases"] == []
    assert all(m["gordura_fita_percentagem"] is None for m in dados["medidas_manuais"])
    assert dados["indicadores"]["cintura_altura"] == pytest.approx(
        dados["indicadores"]["cintura_cm"]["ultima"] / 178, abs=0.001)
    wb = load_workbook(livro)
    assert "Gordura pela fita (%)" not in [c.value for c in wb["Resumo Semanal"][3]]
    assert wb["Painel"]["N22"].value == "Cintura ÷ altura"


def test_configuracao_invalida(projeto):
    with pytest.raises(ValueError, match="sexo"):
        configurar(projeto, sexo="outro")
    with pytest.raises(ValueError, match="ordem"):
        configurar(projeto, sexo=None, fases=[{"nome": "B", "inicio": "2026-03-01"}, {"nome": "A", "inicio": "2026-02-01"}])
    with pytest.raises(ValueError, match="primeira fase"):
        configurar(projeto, fases=[{"nome": "A", "inicio": "2025-12-01"}])
    with pytest.raises(ValueError, match="ano_nascimento"):
        configurar(projeto, fases=None, ano_nascimento="1990")


def test_json_compacto(projeto):
    pasta = projeto["pasta"]
    _, dados, _ = correr(projeto, json_resumo=pasta / "resumo.json")
    texto = (pasta / "resumo.json").read_text(encoding="utf-8")
    resumo = json.loads(texto)
    jsonschema.validate(resumo, json.loads((ESQUEMAS / "dados-zepp-life-ia-resumo.schema.json").read_text(encoding="utf-8")))
    assert len(texto) < len((pasta / "livro.json").read_text(encoding="utf-8")) / 3
    # Os agregados são exatamente os do ficheiro completo.
    for chave in ("indicadores", "fases", "qualidade_dados", "contexto", "protocolo_pesagem"):
        assert resumo[chave] == dados[chave]
    for chave in ("resumos_semanais", "resumos_mensais"):
        tabela = resumo[chave]
        assert [dict(zip(tabela["colunas"], linha, strict=True)) for linha in tabela["linhas"]] == dados[chave]
    recentes = resumo["medicoes_recentes"]
    ultima = datetime.fromisoformat(dados["metadados"]["ultima_medicao_corporal_local"])
    datas = [datetime.fromisoformat(linha[0]) for linha in recentes["linhas"]]
    esperadas = [m for m in dados["medicoes_corporais"]
                 if datetime.fromisoformat(m["data_hora_local"]) > ultima - timedelta(days=28)]
    assert len(datas) == len(esperadas) > 0 and min(datas) > ultima - timedelta(days=28)
    assert recentes["linhas"][-1][recentes["colunas"].index("peso_kg")] == dados["medicoes_corporais"][-1]["peso_kg"]
    assert "observacoes" not in texto and "AABBCCDDEEFF0011" not in texto
    assert resumo["metadados"]["ficheiro_completo"] == "livro.json"


def test_copia_das_medidas_manuais(projeto, capsys):
    pasta = projeto["pasta"]
    copia = pasta / "medidas-manuais.csv"
    livro, antes, _ = correr(projeto, nome="v1", guardar_copia_manuais=copia)
    semanas = medidas.ler_existentes(livro)
    lidas = medidas.ler_copia(copia)
    assert {i: (s.data, s.cintura, s.pescoco, s.notas) for i, s in semanas.items()} == \
           {i: (s.data, s.cintura, s.pescoco, s.notas) for i, s in lidas.items()}
    # Excel apagado: as medidas vêm da cópia (e não só do CSV inicial).
    linhas = copia.read_text(encoding="utf-8-sig").splitlines()
    assert linhas[7].startswith("2026-02-15;")                          # semana de férias, sem linha no CSV inicial
    copia.write_text("\n".join(linhas[:7]) + '\n2026-02-15;2026-02-17;91.2;39.4;"medida; com ponto e vírgula"\n'
                     + "\n".join(linhas[8:]) + "\n", encoding="utf-8-sig")
    resultado = gerar.gerar(projeto["origem"], pasta / "v2.xlsx", pasta / "v2.json", existente=pasta / "nao-existe.xlsx",
                            config=projeto["config"], hoje=date(2026, 5, 2), verificar_excel=False, copia_manuais=copia)
    assert "repostas a partir de medidas-manuais.csv" in capsys.readouterr().out
    depois = json.loads((pasta / "v2.json").read_text(encoding="utf-8"))
    reposta = next(m for m in depois["medidas_manuais"] if m["semana_inicio"] == "2026-02-15")
    assert (reposta["cintura_cm"], reposta["data_medicao"], reposta["observacoes"]) == \
           (91.2, "2026-02-17", "medida; com ponto e vírgula")
    assert resultado.preservadas == len(lidas)
    # Com o Excel presente, o Excel ganha sempre à cópia.
    resultado = gerar.gerar(projeto["origem"], pasta / "v3.xlsx", pasta / "v3.json", existente=livro,
                            config=projeto["config"], hoje=date(2026, 5, 2), verificar_excel=False, copia_manuais=copia)
    com_excel = json.loads((pasta / "v3.json").read_text(encoding="utf-8"))
    assert next(m for m in com_excel["medidas_manuais"] if m["semana_inicio"] == "2026-02-15")["cintura_cm"] is None
    with pytest.raises(RuntimeError, match="cópia das medidas manuais"):
        copia.write_text("isto não é a cópia\n", encoding="utf-8")
        gerar.gerar(projeto["origem"], pasta / "v4.xlsx", pasta / "v4.json", config=projeto["config"],
                    hoje=date(2026, 5, 2), verificar_excel=False, copia_manuais=copia)


def test_leituras_exatamente_no_limite_das_janelas(projeto):
    """Leituras a exatamente 7, 14, 21 e 28 dias (ao segundo) da última: o Excel real tem de dar o mesmo
    que o Python nas médias móveis, nos cartões e no ritmo (as fórmulas usam meio segundo de margem)."""
    exportacao = json.loads(projeto["origem"].read_text(encoding="utf-8"))
    ultima = max(datetime.fromisoformat(r["Date"]) for r in exportacao)
    primeira = min(datetime.fromisoformat(r["Date"]) for r in exportacao if datetime.fromisoformat(r["Date"]).date() >= INICIO)
    for dias in (7, 14, 21, 28):
        exportacao.append(sintetico.registo(ultima - timedelta(days=dias), 79.3, 21.3, 58.9))
    for dias in (6, 7):
        exportacao.append(sintetico.registo(primeira + timedelta(days=dias), 83.1, 24.6, 59.4))
    projeto["origem"].write_text(json.dumps(exportacao), encoding="utf-8")
    _, dados, resultado = correr(projeto, excel_real=True)             # falharia aqui se o Excel divergisse
    assert resultado.estado_excel in ("indisponivel",) or resultado.estado_excel.startswith("verificado")
    atual = [m["peso_kg"] for m in dados["medicoes_corporais"]
             if datetime.fromisoformat(m["data_hora_local"]) > ultima.replace(tzinfo=None) - timedelta(days=7)]
    assert dados["indicadores"]["peso_kg"]["media_7d_atual"] == pytest.approx(sum(atual) / len(atual), abs=0.001)
