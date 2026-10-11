"""Gráficos nativos do Excel: leituras diárias com média móvel e medidas manuais com tracejados."""
from __future__ import annotations

import math
from datetime import date, datetime
from itertools import pairwise

from openpyxl.chart import LineChart, Reference, ScatterChart, Series
from openpyxl.chart.axis import ChartLines, DateAxis
from openpyxl.chart.data_source import NumFmt
from openpyxl.chart.legend import Legend
from openpyxl.chart.marker import DataPoint, Marker
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText, Text
from openpyxl.chart.title import Title
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties, RegularTextRun
from openpyxl.drawing.text import Font as FonteDesenho

from ..datas import serie_excel
from .estilos import COR, FONTE


def eixo_legivel(valores, amplitude_minima: float) -> tuple[float, float, float]:
    """Mínimo, máximo e intervalo 'redondos' com margem de 12% e cerca de 5–7 marcas."""
    valores = [v for v in valores if isinstance(v, (int, float)) and math.isfinite(v)]
    if not valores:
        return 0, amplitude_minima, amplitude_minima / 5
    lo, hi = min(valores), max(valores)
    amplitude = max(hi - lo, amplitude_minima)
    baixo = max(0, (lo + hi - amplitude) / 2 - amplitude * 0.12)
    alto = (lo + hi + amplitude) / 2 + amplitude * 0.12
    expoente = math.floor(math.log10((alto - baixo) / 5))
    candidatos = []
    for potencia in range(expoente - 1, expoente + 2):
        for fator in (1, 2, 2.5, 5, 10):
            passo = fator * 10 ** potencia
            fundo, topo = math.floor(baixo / passo) * passo, math.ceil(alto / passo) * passo
            marcas = round((topo - fundo) / passo) + 1
            candidatos.append((abs(marcas - 6) + (0 if 5 <= marcas <= 7 else 10), topo - fundo, fundo, topo, passo))
    _, _, fundo, topo, passo = min(candidatos)
    return round(fundo, 10), round(topo, 10), round(passo, 10)


def _texto(tamanho: int = 900, negrito: bool = False, cor: str = COR["suave"]) -> RichText:
    props = CharacterProperties(sz=tamanho, b=negrito, solidFill=cor, latin=FonteDesenho(typeface=FONTE))
    return RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=props), endParaRPr=props)])


def _titulo(texto: str, tamanho: int = 1200, negrito: bool = True, cor: str = COR["titulo"]) -> Title:
    props = CharacterProperties(sz=tamanho, b=negrito, solidFill=cor, latin=FonteDesenho(typeface=FONTE))
    paragrafo = Paragraph(pPr=ParagraphProperties(defRPr=props), r=[RegularTextRun(rPr=props, t=texto)])
    return Title(tx=Text(rich=RichText(p=[paragrafo])), overlay=False)


def _estilo(ch, titulo: str, largura_cm: float, altura_cm: float) -> None:
    ch.title = _titulo(titulo)
    ch.width, ch.height = largura_cm, altura_cm
    ch.legend = None
    for eixo in (ch.x_axis, ch.y_axis):
        eixo.delete = False
        eixo.txPr = _texto()
        eixo.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill=COR["linha"]))
    ch.y_axis.majorGridlines = ChartLines(spPr=GraphicalProperties(ln=LineProperties(solidFill="EEF1F4")))
    ch.x_axis.majorGridlines = None


def _escala_y(ch, valores, amplitude: float, formato: str = "0.0") -> None:
    fundo, topo, passo = eixo_legivel(valores, amplitude)
    ch.y_axis.scaling.min, ch.y_axis.scaling.max, ch.y_axis.majorUnit = fundo, topo, passo
    ch.y_axis.numFmt = NumFmt(formatCode=formato, sourceLinked=False)


def _escala_datas(eixo, inicio: date | datetime, fim: date | datetime) -> None:
    """Eixo de datas dd/mm com uma marca a cada 4 semanas."""
    eixo.numFmt = NumFmt(formatCode="dd/mm", sourceLinked=False)
    eixo.scaling.min, eixo.scaling.max = math.floor(serie_excel(inicio)), math.ceil(serie_excel(fim))
    eixo.majorUnit = 28


def _linha(cor: str, largura_pt: float = 2.0, tracejado: str | None = None) -> GraphicalProperties:
    return GraphicalProperties(ln=LineProperties(solidFill=cor, w=int(largura_pt * 12700), prstDash=tracejado))


def diario(x: str, valores_ref: str, medias_ref: str, titulo: str, cor: str, valores: list[float],
           medias: list[float], inicio: datetime, fim: datetime, amplitude: float, unidade: str) -> ScatterChart:
    """Leituras diárias em pontos e média móvel de 7 dias em linha. *_ref: intervalos absolutos."""
    ch = ScatterChart(scatterStyle="lineMarker")
    _estilo(ch, titulo, 30, 8.5)
    eixo_x = Reference(range_string=x)
    pontos = Series(Reference(range_string=valores_ref), eixo_x, title="Leituras")
    pontos.marker = Marker(symbol="circle", size=4)
    pontos.marker.graphicalProperties = GraphicalProperties(solidFill=COR["cinza_ponto"], ln=LineProperties(noFill=True))
    pontos.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))
    media = Series(Reference(range_string=medias_ref), eixo_x, title="Média móvel de 7 dias")
    media.marker = Marker(symbol="none")
    media.graphicalProperties = _linha(cor, 2.25)
    media.smooth = False
    pontos.smooth = False
    ch.series += [pontos, media]
    ch.legend = Legend(legendPos="t", overlay=False, txPr=_texto())
    _escala_y(ch, list(valores) + list(medias), amplitude, "0.0")
    ch.y_axis.title = _titulo(unidade, 900, False, COR["suave"])
    _escala_datas(ch.x_axis, inicio, fim)
    return ch


def manual(valores_ref: str, datas_ref: str, valores: list[float | None], datas: list[date], titulo: str, cor: str,
           amplitude: float) -> LineChart:
    """Linha contínua entre medições consecutivas, tracejado sobre semanas sem medição.
    valores: um por semana (None = sem medição); datas: datas usadas para a escala do eixo."""
    ch = LineChart()
    _estilo(ch, titulo, 15, 8)
    ch.add_data(Reference(range_string=valores_ref))
    ch.set_categories(Reference(range_string=datas_ref))
    serie = ch.series[0]
    serie.smooth = False
    serie.graphicalProperties = _linha(cor)
    serie.marker = Marker(symbol="circle", size=5)
    serie.marker.graphicalProperties = GraphicalProperties(solidFill=cor, ln=LineProperties(solidFill=cor))
    medidos = [i for i, v in enumerate(valores) if v is not None]
    for anterior, atual in pairwise(medidos):
        if atual - anterior > 1:
            # O formato de um ponto aplica-se ao segmento que termina nele.
            ponto = DataPoint(idx=atual, marker=Marker(symbol="circle", size=5))
            ponto.marker.graphicalProperties = GraphicalProperties(solidFill=cor, ln=LineProperties(solidFill=cor))
            ponto.graphicalProperties = _linha(cor, 2.0, "dash")
            serie.dPt.append(ponto)
    ch.display_blanks = "span"
    ch.x_axis = DateAxis(crossAx=100, delete=False, txPr=_texto(),
                         spPr=GraphicalProperties(ln=LineProperties(solidFill=COR["linha"])))
    ch.x_axis.baseTimeUnit = "days"
    ch.x_axis.majorTimeUnit = "days"
    _escala_datas(ch.x_axis, min(datas), max(datas))
    ch.y_axis.crossAx = 500
    _escala_y(ch, valores, amplitude, "0.0")
    return ch
