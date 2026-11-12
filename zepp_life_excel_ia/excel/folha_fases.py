"""Folha "Fases" (opcional): só existe quando config.json define `fases`.

Uma linha por fase, com as médias de 7 dias no início e no fim da fase e a variação entre elas.
Cada coluna é definida num único sítio (`Coluna`), como nos resumos.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Callable

from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

from ..datas import serie_excel
from ..resumos import FASE_DIAS_MINIMOS_RITMO, ResumoFase
from ..verificacao import Esperado
from . import estilos
from .estilos import COR
from .folha_manuais import FolhaManuais
from .folha_medicoes import FolhaMedicoes

NOME = "Fases"
PRIMEIRA = 4
VARIACAO_1 = '+0.0;-0.0;0.0'
VARIACAO_2 = '+0.00;-0.00;0.00'
NOTA = ("Cada fase vai da sua data de início até à véspera da fase seguinte (a última, até à última leitura). "
        "Inicial = média das leituras dos primeiros 7 dias da fase; final = média dos 7 dias até à última leitura "
        "da fase. kg/semana = variação do peso a dividir pelas semanas entre as duas médias (fases com 21 dias ou "
        "mais). Cintura: primeira e última medição manual dentro da fase. As fases definem-se em config.json.")


@dataclass
class Coluna:
    cabecalho: str
    valor: Callable[[int, ResumoFase], object]          # (linha, fase) -> valor ou fórmula
    formato: str | None = None
    esperado: Callable[[ResumoFase], object] | None = None
    grupo: str | None = None
    largura: float = 13


def _colunas(medicoes: FolhaMedicoes, manuais: FolhaManuais) -> list[Coluna]:
    datas = medicoes.intervalo("Data/hora")
    peso, gordura, gorda, magra = (medicoes.intervalo_metrica(c) for c in
                                   ("Weight", "BodyFat", "massa_gorda", "massa_magra"))
    cinturas, datas_m = manuais.intervalo("Cintura (cm)"), manuais.intervalo("Data da medição")
    # Colunas fixas: B = Início, C = Fim, D = Dias, E = Leituras, F = Última leitura.
    na_fase = lambda r, d=datas: f'{d},">="&$B{r},{d},"<"&$C{r}+1'
    inicial = lambda v, r: f'AVERAGEIFS({v},{datas},">="&$B{r},{datas},"<"&MIN($B{r}+7,$C{r}+1))'
    final = lambda v, r: (f"AVERAGEIFS({v},{datas},{estilos.depois_de(f'$F{r}-7')},{datas},{estilos.ate(f'$F{r}')},"
                          f'{datas},">="&$B{r})')
    se_ha = lambda r, expr: f'=IF($E{r}=0,"",IFERROR({expr},""))'

    def par(nome, v, formato, formato_variacao, campo, grupo):
        """Três colunas: inicial, final e variação de uma métrica."""
        return [
            Coluna(f"{nome} inicial", lambda r, f: se_ha(r, inicial(v, r)), formato, lambda f: f.inicial[campo], grupo),
            Coluna(f"{nome} final", lambda r, f: se_ha(r, final(v, r)), formato, lambda f: f.final[campo], grupo),
            Coluna(f"Variação: {nome.lower()}", lambda r, f: se_ha(r, f"{final(v, r)}-{inicial(v, r)}"),
                   formato_variacao, lambda f: f.variacao(campo), grupo),
        ]

    def cintura(extremo):
        criterios = lambda r: f'{cinturas},">0",{na_fase(r, datas_m)}'
        return lambda r: (f'IF(COUNTIFS({criterios(r)})=0,"",SUMIFS({cinturas},{datas_m},'
                          f"_xlfn.{extremo}({datas_m},{criterios(r)})))")

    primeira_cintura, ultima_cintura = cintura("MINIFS"), cintura("MAXIFS")
    azul, verde, laranja = COR["azul_claro"], COR["verde_claro"], COR["laranja_claro"]
    colunas = [
        Coluna("Fase", lambda r, f: f.nome, "@", largura=24),
        Coluna("Início", lambda r, f: f.inicio, "dd/mm/yyyy", largura=11),
        Coluna("Fim", lambda r, f: f.fim, "dd/mm/yyyy", largura=11),
        Coluna("Dias", lambda r, f: f"=$C{r}-$B{r}+1", "0", lambda f: f.dias, largura=8),
        Coluna("Leituras", lambda r, f: f"=COUNTIFS({na_fase(r)})", "0", lambda f: f.leituras, largura=10),
        Coluna("Última leitura", lambda r, f: f'=IF($E{r}=0,"",_xlfn.MAXIFS({datas},{na_fase(r)}))', "dd/mm/yyyy hh:mm",
               lambda f: serie_excel(f.ultima_leitura) if f.ultima_leitura else None, largura=17),
        *par("Peso (kg)", peso, "0.00", VARIACAO_2, "Weight", azul),
    ]
    variacao_peso = get_column_letter(len(colunas))
    colunas += [
        Coluna("Peso: kg/semana",
               lambda r, f: (f'=IF(OR(${variacao_peso}{r}="",$D{r}<{FASE_DIAS_MINIMOS_RITMO}),"",'
                             f"${variacao_peso}{r}/(($D{r}-7)/7))"),
               VARIACAO_2, lambda f: f.ritmo_peso_semanal, azul),
        *par("Gordura (%)", gordura, "0.0", VARIACAO_1, "BodyFat", azul),
        Coluna("Variação: massa gorda (kg)", lambda r, f: se_ha(r, f"{final(gorda, r)}-{inicial(gorda, r)}"),
               VARIACAO_2, lambda f: f.variacao("massa_gorda"), verde),
        Coluna("Variação: massa magra (kg)", lambda r, f: se_ha(r, f"{final(magra, r)}-{inicial(magra, r)}"),
               VARIACAO_2, lambda f: f.variacao("massa_magra"), verde),
        Coluna("Cintura inicial (cm)", lambda r, f: f"={primeira_cintura(r)}", "0.0", lambda f: f.cintura_inicial, laranja),
        Coluna("Cintura final (cm)", lambda r, f: f"={ultima_cintura(r)}", "0.0", lambda f: f.cintura_final, laranja),
    ]
    c0, c1 = get_column_letter(len(colunas) - 1), get_column_letter(len(colunas))
    colunas += [
        Coluna("Variação: cintura (cm)", lambda r, f: f'=IF(OR({c0}{r}="",{c1}{r}=""),"",{c1}{r}-{c0}{r})',
               VARIACAO_1, lambda f: f.variacao_cintura, laranja),
        Coluna("Objetivo", lambda r, f: f.objetivo, "@", largura=44),
    ]
    assert [c.cabecalho for c in colunas[:6]] == ["Fase", "Início", "Fim", "Dias", "Leituras", "Última leitura"]
    return colunas


def construir(ws, fases: list[ResumoFase], medicoes: FolhaMedicoes, manuais: FolhaManuais, atualizado: str,
              esperados: list[Esperado]) -> None:
    colunas = _colunas(medicoes, manuais)
    ultima = len(fases) + PRIMEIRA - 1
    fim = get_column_letter(len(colunas))
    estilos.titulo(ws, "Fases da transformação", len(colunas), atualizado, coluna="B")
    estilos.cabecalho(ws, PRIMEIRA - 1, [c.cabecalho for c in colunas], {c.cabecalho: c.grupo for c in colunas if c.grupo})
    for r, fase in enumerate(fases, start=PRIMEIRA):
        for j, c in enumerate(colunas, start=1):
            ws.cell(r, j, c.valor(r, fase))
            if c.esperado:
                esperados.append(Esperado(NOME, f"{get_column_letter(j)}{r}", c.esperado(fase), c.cabecalho))
    estilos.corpo_texto(ws, PRIMEIRA, ultima, len(colunas))
    for j, c in enumerate(colunas, start=1):
        for r in range(PRIMEIRA, ultima + 1):
            if c.formato:
                ws.cell(r, j).number_format = c.formato
    for r in range(PRIMEIRA, ultima + 1):
        ws.cell(r, len(colunas)).alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[r].height = 30
    estilos.tabela(ws, "tblFases", f"A{PRIMEIRA - 1}:{fim}{ultima}")
    estilos.larguras(ws, [c.largura for c in colunas])
    ws.freeze_panes = f"B{PRIMEIRA}"
    estilos.nota(ws, f"B{ultima + 2}", f"{get_column_letter(min(len(colunas), 15))}{ultima + 2}", NOTA)
    ws.row_dimensions[ultima + 2].height = 42
