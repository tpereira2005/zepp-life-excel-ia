"""Folha "Painel": cartões com indicadores e gráficos.

Desenhado para o ecrã do utilizador (janela com ~1858 px úteis, 1 carácter de largura ≈ 9 px,
1 ponto de altura ≈ 1,667 px). Um único espaçamento (~27 px) em todo o lado: margens laterais,
espaços entre cartões (horizontais e verticais), antes dos gráficos e entre gráficos.
Colunas A e Q são margens vazias; o conteúdo vai de B a P (4 cartões de 3 colunas + 3 espaços).
As colunas R e S (ocultas) guardam valores auxiliares usados por vários cartões: o mínimo da média de
7 dias do peso, a data desse mínimo e os dias desde a última leitura.
"""
from __future__ import annotations

from dataclasses import dataclass

from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill
from openpyxl.utils import get_column_letter

from ..datas import serie_excel
from ..metricas import RITMO_MINIMO_AMPLITUDE_DIAS, RITMO_MINIMO_LEITURAS
from ..resumos import TOLERANCIA_EXTREMO, VARIACAO_MINIMA_FRACAO, Calculos
from ..verificacao import Esperado
from . import estilos, graficos
from .estilos import BORDA, COR
from .folha_manuais import FolhaManuais
from .folha_medicoes import DIAS_RITMO, FolhaMedicoes
from .folha_resumos import CINTURA_ALTURA, GORDURA_FITA

NOME = "Painel"
ESPACO_LARGURA = 3          # caracteres ≈ 27 px
ESPACO_ALTURA = 16.2        # pontos ≈ 27 px
CARTAO_LARGURA = 15.9       # 3 colunas por cartão: 4 cartões + 5 espaços ≈ 1851 px
LINHA_CARTAO = 17           # pontos
LINHAS_CARTOES = [4, 10, 16, 22]                           # cada fila: rótulo, valor (3 linhas), nota
GRAFICOS_LINHAS = [(28, 45), (47, 64), (66, 83), (85, 102)]  # 18 linhas cada, um espaço depois
LINHA_NOTA = 104
ULTIMA_COLUNA = 16                                         # P
# Valores auxiliares (colunas ocultas): rótulo em R, valor em S.
AUX_ROTULO, AUX_VALOR = "R", "S"
AUX_MINIMO, AUX_DATA_MINIMO, AUX_DIAS_SEM_LEITURAS = "$S$4", "$S$5", "$S$6"
DIAS_AVISO = 3                                             # aviso quando a última leitura tem mais dias do que isto
# Amplitude mínima dos eixos Y (unidade da métrica): evita ampliar ruído de medição.
AMPLITUDE_MINIMA = {"peso": 6, "gordura": 4, "musculo": 3, "cintura": 10, "pescoco": 3}

KG, PERCENTAGEM, CM = '0.00 "kg"', '0.0 "%"', '0.0 "cm"'


def _delta(unidade: str, casas: int = 2) -> str:
    n = "0." + "0" * casas
    return f'+{n} "{unidade}";-{n} "{unidade}";{n} "{unidade}"'


@dataclass
class Cartao:
    rotulo: str
    formula: str
    formato: str
    cor: str
    nota: str               # expressão Excel (texto entre aspas ou fórmula)
    esperado: object


def _cartoes(calc: Calculos, medicoes: FolhaMedicoes, manuais: FolhaManuais, semanal: dict[str, str]) -> list[list[Cartao]]:
    datas, instantes = medicoes.intervalo("Data/hora"), medicoes.intervalo("Instante UTC")
    peso, gordura, musculo, gorda, magra = (medicoes.intervalo_metrica(c) for c in
                                            ("Weight", "BodyFat", "MuscleMass", "massa_gorda", "massa_magra"))
    ultimas7 = lambda v: f"AVERAGEIFS({v},{datas},{estilos.depois_de(f'MAX({datas})-7')})"
    primeiras7 = lambda v: f"AVERAGEIFS({v},{datas},{estilos.antes_de(f'MIN({datas})+7')})"
    # Média de 7 dias na leitura em que a média de 7 dias do peso foi mínima.
    no_minimo = lambda v: (f"AVERAGEIFS({v},{datas},{estilos.depois_de(f'{AUX_DATA_MINIMO}-7')},"
                           f"{datas},{estilos.ate(AUX_DATA_MINIMO)})")
    desde_inicio = lambda v: f"{ultimas7(v)}-{primeiras7(v)}"
    desde_minimo = lambda v: f'IF({AUX_MINIMO}="","",{ultimas7(v)}-{no_minimo(v)})'
    # A última leitura é a de maior instante UTC (não de maior hora local: mudança de hora).
    ultima = lambda v: f"INDEX({v},MATCH(MAX({instantes}),{instantes},0))"
    dias_ritmo = medicoes.intervalo(DIAS_RITMO)
    ritmo = (f'IF(OR(COUNT({dias_ritmo})<{RITMO_MINIMO_LEITURAS},'
             f'-MIN({dias_ritmo})<{RITMO_MINIMO_AMPLITUDE_DIAS}-{estilos.MEIO_SEGUNDO}),"",SLOPE({peso},{dias_ritmo})*7)')
    datas_m = manuais.intervalo("Data da medição")

    def manual_ultima(v):
        return f'IF(COUNT({v})=0,"",SUMIFS({v},{datas_m},_xlfn.MAXIFS({datas_m},{v},">0")))'

    def manual_variacao(v):
        return (f'IF(COUNT({v})<2,"",SUMIFS({v},{datas_m},_xlfn.MAXIFS({datas_m},{v},">0"))'
                f'-SUMIFS({v},{datas_m},_xlfn.MINIFS({datas_m},{v},">0")))')

    def manual_data(v, prefixo):
        data = f'_xlfn.MAXIFS({datas_m},{v},">0")'
        return f'IF(COUNT({v})=0,"","{prefixo} "&{estilos.data_formula(data)})'

    def fracao_gordura(variacao):
        """Nota: que parte da variação do peso foi massa gorda (só quando o peso variou ≥ 0,5 kg no mesmo sentido)."""
        dp, dg = variacao(peso), variacao(gorda)
        return (f'IFERROR(IF(OR(ABS({dp})<{VARIACAO_MINIMA_FRACAO!r},({dg})/({dp})<0),"estimativa da balança",'
                f'ROUND(({dg})/({dp})*100,0)&" % da variação do peso"),"estimativa da balança")')

    com_sinal = lambda expr, casas, unidade: (f'IF({expr}>0,"+","")&{estilos.numero_formula(expr, casas)}&" {unidade}"')
    ind = calc.indicadores
    p, g, m = ind["peso_kg"], ind["gordura_corporal_percentagem"], ind["massa_muscular_kg"]
    mg, mm = ind["massa_gorda_kg"], ind["massa_magra_kg"]
    cintura, fita = ind["cintura_cm"], ind["gordura_fita_percentagem"]
    variacao_manual = lambda i: i.variacao if i and i.medicoes > 1 else None
    media7 = '"média 7 dias: atual − inicial"'
    v_cintura, v_pescoco = manuais.intervalo("Cintura (cm)"), manuais.intervalo("Pescoço (cm)")
    nota_minimo = (f'IF({AUX_MINIMO}="","ainda sem 7 dias de leituras","mínimo "&{estilos.numero_formula(AUX_MINIMO, 2)}'
                   f'&" kg em "&{estilos.data_formula(AUX_DATA_MINIMO)})')
    desde_data_minimo = (f'IF({AUX_MINIMO}="","","desde o mínimo do peso, "&{estilos.data_formula(AUX_DATA_MINIMO)})')
    data_cintura_minima = f"_xlfn.MINIFS({datas_m},{v_cintura},MIN({v_cintura}))"
    nota_cintura_minima = (f'IF(COUNT({v_cintura})=0,"","mínimo "&{estilos.numero_formula(f"MIN({v_cintura})", 1)}'
                           f'&" cm em "&{estilos.data_formula(data_cintura_minima)})')
    altura = calc.altura_cm
    racio = f'IF(COUNT({v_cintura})=0,"",{manual_ultima(v_cintura)}/{altura!r})' if altura else '""'
    if calc.cfg.sexo == "masculino":
        coluna_fita = semanal[GORDURA_FITA]
        coluna_racio = semanal[CINTURA_ALTURA]
        ultimo = lambda v: f'IFERROR(LOOKUP(2,1/({v}<>""),{v}),"")'
        nota_fita = (f'IFERROR("método da Marinha dos EUA  ·  cintura ÷ altura "'
                     f'&{estilos.numero_formula(ultimo(coluna_racio), 2)},"método da Marinha dos EUA")')
        quarto = Cartao("Gordura pela fita", f"={ultimo(coluna_fita)}", PERCENTAGEM, "laranja_claro", nota_fita,
                        fita.ultima if fita else None)
    else:
        nota_racio = (f'IF(COUNT({v_pescoco})=0,"abaixo de 0,5 é a referência habitual",'
                      f'"pescoço: "&{estilos.numero_formula(manual_ultima(v_pescoco), 1)}&" cm")')
        quarto = Cartao("Cintura ÷ altura", f"={racio}", "0.00", "laranja_claro", nota_racio, calc.cintura_altura)
    return [
        [Cartao("Último peso", f"={ultima(peso)}", KG, "azul_claro", '"leitura mais recente"', p.ultima_leitura),
         Cartao("Peso (média 7 dias)", f"={ultimas7(peso)}", KG, "azul_claro", '"últimos 7 dias"', p.atual_media_7d),
         Cartao("Peso desde o início", f"={desde_inicio(peso)}", _delta("kg"), "fundo", media7, p.variacao),
         Cartao("Peso desde o mínimo", f"={desde_minimo(peso)}", _delta("kg"), "fundo", nota_minimo,
                p.variacao_desde_minimo_do_peso)],
        [Cartao("Ritmo recente", f"={ritmo}", _delta("kg/sem."), "fundo", '"tendência dos últimos 28 dias"',
                p.ritmo_semanal_4s),
         Cartao("Gordura (média 7 dias)", f"={ultimas7(gordura)}", PERCENTAGEM, "laranja_claro",
                '"estimativa da balança"', g.atual_media_7d),
         Cartao("Gordura desde o início", f"={desde_inicio(gordura)}", _delta("pp", 1), "fundo",
                '"pontos percentuais"', g.variacao),
         Cartao("Músculo (média 7 dias)", f"={ultimas7(musculo)}", KG, "verde_claro",
                f'"desde o início: "&{com_sinal(f"({desde_inicio(musculo)})", 2, "kg")}', m.atual_media_7d)],
        [Cartao("Massa gorda desde o início", f"={desde_inicio(gorda)}", _delta("kg"), "fundo",
                fracao_gordura(desde_inicio), mg.variacao),
         Cartao("Massa magra desde o início", f"={desde_inicio(magra)}", _delta("kg"), "fundo",
                '"massa magra = peso − massa gorda"', mm.variacao),
         Cartao("Massa gorda desde o mínimo", f"={desde_minimo(gorda)}", _delta("kg"), "fundo",
                f'IF({AUX_MINIMO}="","",{fracao_gordura(lambda v: f"({ultimas7(v)}-{no_minimo(v)})")})',
                mg.variacao_desde_minimo_do_peso),
         Cartao("Massa magra desde o mínimo", f"={desde_minimo(magra)}", _delta("kg"), "fundo", desde_data_minimo,
                mm.variacao_desde_minimo_do_peso)],
        [Cartao("Última cintura", f"={manual_ultima(v_cintura)}", CM, "laranja_claro",
                manual_data(v_cintura, "medida em"), cintura.ultima if cintura else None),
         Cartao("Cintura desde o início", f"={manual_variacao(v_cintura)}", _delta("cm", 1), "fundo",
                '"última − primeira medição"', variacao_manual(cintura)),
         Cartao("Cintura desde o mínimo", f'=IF(COUNT({v_cintura})<2,"",{manual_ultima(v_cintura)}-MIN({v_cintura}))',
                _delta("cm", 1), "fundo", nota_cintura_minima,
                cintura.variacao_desde_minima if cintura and cintura.medicoes > 1 else None),
         quarto],
    ]


def _desenhar_cartao(ws, linha: int, indice: int, c: Cartao) -> str:
    """Desenha um cartão (rótulo, valor em 3 linhas, nota) e devolve a célula do valor."""
    inicio, fim = 2 + 4 * indice, 4 + 4 * indice
    c1, c2 = get_column_letter(inicio), get_column_letter(fim)
    ws.merge_cells(f"{c1}{linha}:{c2}{linha}")
    ws.merge_cells(f"{c1}{linha + 1}:{c2}{linha + 3}")
    ws.merge_cells(f"{c1}{linha + 4}:{c2}{linha + 4}")
    ws[f"{c1}{linha}"] = c.rotulo
    ws[f"{c1}{linha + 1}"] = c.formula
    ws[f"{c1}{linha + 4}"] = f"={c.nota}"
    fundo = PatternFill("solid", fgColor=COR[c.cor])
    for r in range(linha, linha + 5):
        for col in range(inicio, fim + 1):
            celula = ws.cell(r, col)
            celula.fill = fundo
            celula.border = Border(left=BORDA if col == inicio else None, right=BORDA if col == fim else None,
                                   top=BORDA if r == linha else None, bottom=BORDA if r == linha + 4 else None)
    ws[f"{c1}{linha}"].font = estilos.fonte(11, "suave", negrito=True)
    ws[f"{c1}{linha + 1}"].font = estilos.fonte(24, negrito=True)
    ws[f"{c1}{linha + 1}"].number_format = c.formato
    ws[f"{c1}{linha + 4}"].font = estilos.fonte(9, "suave", italico=True)
    for r in (linha, linha + 1, linha + 4):
        ws[f"{c1}{r}"].alignment = Alignment(horizontal="center", vertical="center")
    return f"{c1}{linha + 1}"


def _graficos(calc: Calculos, medicoes: FolhaMedicoes, manuais: FolhaManuais) -> list:
    leituras, moveis, semanas = calc.leituras, calc.moveis, calc.semanas
    inicio, fim = min(lt.local for lt in leituras), max(lt.local for lt in leituras)
    x = medicoes.intervalo("Data/hora")

    def diario(campo, titulo, cor, amplitude, unidade, media):
        return graficos.diario(x, medicoes.intervalo_metrica(campo), medicoes.intervalo(media), titulo, cor,
                               [lt[campo] for lt in leituras], [m[campo] for m in moveis], inicio, fim, amplitude, unidade)

    datas = [s.data for s in semanas if s.data is not None] or [s.inicio for s in semanas]

    def manual(cabecalho, campo, titulo, cor, amplitude):
        return graficos.manual(manuais.intervalo(cabecalho), manuais.intervalo("Data da medição"),
                               [getattr(s, campo) for s in semanas], datas, titulo, cor, amplitude)

    (l1, f1), (l2, f2), (l3, f3), (l4, f4) = GRAFICOS_LINHAS
    # (gráfico, (coluna inicial, linha inicial, coluna final, linha final)), índices a começar em 0
    return [
        (diario("Weight", "Peso (kg)", COR["azul"], AMPLITUDE_MINIMA["peso"], "kg", "Peso, média 7 dias (kg)"),
         (1, l1 - 1, ULTIMA_COLUNA, f1)),
        (diario("BodyFat", "Gordura corporal (%)", COR["laranja"], AMPLITUDE_MINIMA["gordura"], "%",
                "Gordura, média 7 dias (%)"), (1, l2 - 1, ULTIMA_COLUNA, f2)),
        (diario("MuscleMass", "Massa muscular (kg)", COR["verde"], AMPLITUDE_MINIMA["musculo"], "kg",
                "Músculo, média 7 dias (kg)"), (1, l3 - 1, ULTIMA_COLUNA, f3)),
        (manual("Cintura (cm)", "cintura", "Cintura (cm)", COR["laranja"], AMPLITUDE_MINIMA["cintura"]),
         (1, l4 - 1, 8, f4)),
        (manual("Pescoço (cm)", "pescoco", "Pescoço (cm)", COR["azul"], AMPLITUDE_MINIMA["pescoco"]),
         (9, l4 - 1, ULTIMA_COLUNA, f4)),
    ]


def _auxiliares(ws, calc: Calculos, medicoes: FolhaMedicoes, esperados: list[Esperado]) -> None:
    datas = medicoes.intervalo("Data/hora")
    media_peso = medicoes.intervalo("Peso, média 7 dias (kg)")
    completas = f"{datas},{estilos.desde(f'MIN({datas})+6')}"      # leituras a partir do 7.º dia
    minimo, data_minimo = AUX_MINIMO.replace("$", ""), AUX_DATA_MINIMO.replace("$", "")
    dias = AUX_DIAS_SEM_LEITURAS.replace("$", "")
    ws[f"{AUX_ROTULO}{minimo[1:]}"] = "Mínimo da média de 7 dias do peso (kg)"
    ws[minimo] = f'=IF(COUNTIFS({completas})=0,"",_xlfn.MINIFS({media_peso},{completas}))'
    ws[f"{AUX_ROTULO}{data_minimo[1:]}"] = "Data/hora desse mínimo (a mais antiga, em caso de empate)"
    ws[data_minimo] = (f'=IF({AUX_MINIMO}="","",_xlfn.MINIFS({datas},{media_peso},'
                       f'"<="&({AUX_MINIMO}+{TOLERANCIA_EXTREMO!r}),{completas}))')
    ws[data_minimo].number_format = "dd/mm/yyyy hh:mm"
    ws[f"{AUX_ROTULO}{dias[1:]}"] = "Dias desde a última leitura"
    ws[dias] = f"=TODAY()-INT(MAX({datas}))"
    m = calc.minimo_peso
    esperados.append(Esperado(NOME, minimo, m.valor if m else None, "mínimo da média de 7 dias do peso"))
    esperados.append(Esperado(NOME, data_minimo, serie_excel(m.quando) if m else None, "data do mínimo do peso"))
    for coluna in (AUX_ROTULO, AUX_VALOR):
        ws.column_dimensions[coluna].hidden = True


def construir(ws, calc: Calculos, medicoes: FolhaMedicoes, manuais: FolhaManuais, semanal: dict[str, str],
              inicio_txt: str, atualizado: str, esperados: list[Esperado]) -> None:
    datas, instantes = medicoes.intervalo("Data/hora"), medicoes.intervalo("Instante UTC")
    estilos.titulo(ws, "", ULTIMA_COLUNA, coluna="B")
    _auxiliares(ws, calc, medicoes, esperados)
    ws["B1"] = f'="Transformação corporal  ·  {inicio_txt} a "&{estilos.data_formula(f"MAX({datas})")}'
    hora = f"INDEX({datas},MATCH(MAX({instantes}),{instantes},0))"
    atraso = (f'IF({AUX_DIAS_SEM_LEITURAS}>{DIAS_AVISO},"  ·  ATENÇÃO: sem leituras há "&{AUX_DIAS_SEM_LEITURAS}'
              f'&" dias (os valores atuais referem-se à última leitura)","")')
    ws["B2"] = (f'="Última leitura: "&{estilos.data_formula(hora)}&" às "&RIGHT("0"&HOUR({hora}),2)&":"'
                f'&RIGHT("0"&MINUTE({hora}),2)&"  ·  "&COUNT({datas})&" leituras aceites  ·  {atualizado}"&{atraso}')
    ws["B2"].font = estilos.fonte(9, "suave", italico=True)
    ws.conditional_formatting.add("B2", FormulaRule(
        formula=[f"{AUX_DIAS_SEM_LEITURAS}>{DIAS_AVISO}"], font=Font(color=COR["vermelho"], bold=True)))
    for linha, fila in zip(LINHAS_CARTOES, _cartoes(calc, medicoes, manuais, semanal), strict=True):
        for indice, cartao in enumerate(fila):
            celula = _desenhar_cartao(ws, linha, indice, cartao)
            esperados.append(Esperado(NOME, celula, cartao.esperado, cartao.rotulo))
    estilos.larguras(ws, [ESPACO_LARGURA] + [CARTAO_LARGURA] * 3 + ([ESPACO_LARGURA] + [CARTAO_LARGURA] * 3) * 3
                     + [ESPACO_LARGURA])
    espacos = {linha - 1 for linha in LINHAS_CARTOES} | {LINHAS_CARTOES[-1] + 5}
    for r in range(LINHAS_CARTOES[0] - 1, LINHAS_CARTOES[-1] + 6):
        ws.row_dimensions[r].height = ESPACO_ALTURA if r in espacos else LINHA_CARTAO
    for _, fim in GRAFICOS_LINHAS:
        ws.row_dimensions[fim + 1].height = ESPACO_ALTURA
    for grafico, (c0, r0, c1, r1) in _graficos(calc, medicoes, manuais):
        # Âncora a duas células: o gráfico ocupa exatamente as colunas/linhas indicadas.
        ws.add_chart(grafico, TwoCellAnchor(editAs="oneCell", _from=AnchorMarker(col=c0, row=r0),
                                            to=AnchorMarker(col=c1, row=r1)))
    estilos.nota(ws, f"B{LINHA_NOTA}", f"{get_column_letter(ULTIMA_COLUNA)}{LINHA_NOTA}",
                 "\"Desde o mínimo\" = desde a média de 7 dias mais baixa do peso. Massa magra = peso − massa "
                 "gorda. Ritmo = tendência (regressão linear) das leituras dos últimos 28 dias. "
                 "Gráficos de peso, gordura e músculo: pontos = leituras; linha = média móvel de 7 dias. "
                 "Cintura e pescoço: pontos = medições reais; tracejado = semanas sem medição (só visual, sem "
                 "valores inventados). As escalas e os tracejados são recalculados em cada atualização.")
    ws.freeze_panes = "A3"
