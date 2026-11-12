"""Folha "Critérios e Qualidade": dicionário de campos, regras, qualidade dos dados, leituras atípicas e excluídas."""
from __future__ import annotations

from openpyxl.styles import Alignment, Border

from .. import metricas
from ..metricas import DA_EXPORTACAO, DERIVADAS, LIMITES
from ..resumos import Calculos
from . import estilos
from .estilos import BORDA

COLUNAS = 6


def _regras(calc: Calculos) -> list[tuple[str, str]]:
    cfg, c = calc.cfg, calc.selecao.contagens
    perfil = f"Titular: identificador Zepp {cfg.perfil_titular_id} (o leitor pede apenas este membro). "
    if cfg.altura_perfil_sem_id_cm:
        perfil += f"Exportações antigas sem identificador: altura {cfg.altura_perfil_sem_id_cm:g} cm."
    absolutos = "; ".join(f"{metricas.POR_CAMPO[k].nome.lower()} {v:g} ".replace(".", ",")
                          + ("pontos percentuais" if metricas.POR_CAMPO[k].unidade == "%" else metricas.POR_CAMPO[k].unidade)
                          for k, v in metricas.ATIPICA_ABSOLUTA.items())
    relativos = "; ".join(f"{metricas.POR_CAMPO[k].nome.lower()} {v:.0%}" for k, v in metricas.ATIPICA_RELATIVA.items())
    fita = ("Cintura ÷ altura: última cintura a dividir pela altura do perfil. " +
            ("Gordura pela fita: método da Marinha dos EUA (cintura, pescoço e altura); é uma estimativa, "
             "independente da balança." if cfg.sexo == "masculino" else
             "Gordura pela fita (método da Marinha dos EUA): só é calculada com \"sexo\": \"masculino\" em "
             "config.json (a fórmula feminina exige a medida da anca)."))
    return [
        ("Início", f"Leituras com data local a partir de {cfg.inicio:%d/%m/%Y} ({cfg.fuso_nome})."),
        ("Protocolo de pesagem", cfg.protocolo_pesagem or "Não indicado."),
        ("Perfil", perfil),
        ("Composição completa", "IMC, gordura, água, massa óssea, metabolismo, músculo, visceral, proteína, "
                                "massa proteica, idade corporal e pontuação com valores numéricos."),
        ("Plausibilidade", "; ".join(f"{k} {lo}–{hi}" for k, (lo, hi) in LIMITES.items())),
        ("Semanas", "Domingo a sábado, em hora local."),
        ("Médias", "Por leitura (não por dia). 'Dias com medição' mostra quantos dias distintos contribuíram."),
        ("Mínimo do peso", "A média de 7 dias mais baixa do peso, a contar do 7.º dia de leituras. 'Desde o mínimo' "
                           "compara a média dos últimos 7 dias com a média de 7 dias nessa data."),
        ("Ritmo recente", f"Declive da regressão linear das leituras dos últimos {metricas.RITMO_DIAS} dias, por "
                          f"semana. Exige {metricas.RITMO_MINIMO_LEITURAS} leituras em pelo menos "
                          f"{metricas.RITMO_MINIMO_AMPLITUDE_DIAS} dias."),
        ("Massa magra", "Peso menos massa gorda. A soma das variações de massa gorda e massa magra é a variação do peso."),
        ("Fita métrica", fita),
        ("Leituras atípicas", f"Assinaladas (nunca excluídas) quando se afastam da mediana das outras leituras a "
                              f"±{metricas.ATIPICA_VIZINHANCA_DIAS} dias mais do que: {absolutos}; {relativos}."),
        ("Contagens", f"origem {c['origem']} · anteriores ao início {c['anteriores_ao_inicio']} · "
                      f"incompletas {c['incompletas']} · implausíveis {c['implausiveis']} · "
                      f"outro perfil {c['outro_perfil']} · aceites {c['aceites']}"),
        ("Limites", "Composição, idade e pontuação são estimativas da aplicação, não diagnósticos clínicos."),
        ("Atualização", "Sincronizar a Zepp Life, guardar e fechar o Excel e executar ATUALIZAR-EXCEL.cmd. "
                        "Depois de editar só medidas manuais: RECALCULAR-SEM-DESCARREGAR.cmd."),
    ]


def _qualidade(calc: Calculos) -> list[tuple[str, str]]:
    q, semanas = calc.qualidade, calc.semanas
    estados = [s.estado for s in semanas]
    return [
        ("Dias com leitura", f"{q.dias_com_leitura} de {q.dias_do_periodo} dias entre a primeira e a última leitura."),
        ("Maior intervalo", f"{q.maior_intervalo_sem_leituras_dias} dia(s) seguidos sem leituras."),
        ("Repesagens", f"{q.dias_com_varias_leituras} dia(s) com mais de uma leitura (todas contam nas médias)."),
        ("Medidas manuais", f"{estados.count('completa')} semanas completas · {estados.count('incompleta')} incompletas · "
                            f"{estados.count('sem_medicao')} sem medição · {estados.count('por_preencher')} por preencher."),
        ("Leituras atípicas", f"{len(q.atipicas)} de {len(calc.leituras)} leituras (lista abaixo)."),
    ]


def _bloco(ws, r: int, titulo: str, pares: list[tuple[str, str]]) -> int:
    ws.cell(r, 1, titulo).font = estilos.fonte(12, "titulo", negrito=True)
    for nome, texto in pares:
        r += 1
        ws.cell(r, 1, nome).font = estilos.fonte(negrito=True)
        ws.cell(r, 1).alignment = Alignment(vertical="center")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=COLUNAS)
        ws.cell(r, 2, texto).font = estilos.fonte()
        ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[r].height = 30
    return r


def _tabela_leituras(ws, r: int, titulo: str, linhas: list[tuple]) -> int:
    """Tabela simples Data/hora · Peso · Motivo (o motivo ocupa as colunas C a F)."""
    ws.cell(r, 1, titulo).font = estilos.fonte(12, "titulo", negrito=True)
    r += 1
    estilos.cabecalho(ws, r, ["Data/hora", "Peso (kg)", "Motivo"])
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=COLUNAS)
    ws.row_dimensions[r].height = 24
    for leitura, motivo in linhas:
        r += 1
        ws.cell(r, 1, leitura.local).number_format = "dd/mm/yyyy hh:mm"
        ws.cell(r, 2, leitura["Weight"]).number_format = "0.00"
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=COLUNAS)
        ws.cell(r, 3, motivo).alignment = Alignment(wrap_text=True, vertical="center")
        for col in range(1, 4):
            ws.cell(r, col).font = estilos.fonte()
        if len(motivo) > 95:
            ws.row_dimensions[r].height = 30
    return r


def construir(ws, calc: Calculos, atualizado: str) -> None:
    leituras, semanas, selecao = calc.leituras, calc.semanas, calc.selecao
    estilos.titulo(ws, "Critérios e qualidade dos dados", COLUNAS, atualizado)
    estilos.cabecalho(ws, 3, ["Campo", "Unidade", "Significado", "Cobertura", "Origem", "Campo na exportação"])
    linhas = [[m.nome, m.unidade, m.significado,
               f"{sum(isinstance(lt[m.campo], (int, float)) for lt in leituras)} de {len(leituras)} leituras",
               m.origem, m.campo] for m in DA_EXPORTACAO]
    linhas += [[m.nome, m.unidade, m.significado, f"{len(leituras)} leituras", m.origem, "—"] for m in DERIVADAS]
    linhas += [
        ["Médias de 7 dias", "kg / %", "Média das leituras em ]t − 7 dias, t]", f"{len(leituras)} leituras",
         "derivada", "—"],
        ["Cintura", "cm", "Circunferência ao nível do umbigo (média de 3 leituras)",
         f"{sum(s.cintura is not None for s in semanas)} de {len(semanas)} semanas", "manual", "—"],
        ["Pescoço", "cm", "Circunferência do pescoço (média de 3 leituras)",
         f"{sum(s.pescoco is not None for s in semanas)} de {len(semanas)} semanas", "manual", "—"],
    ]
    for i, linha in enumerate(linhas, start=4):
        for j, v in enumerate(linha, start=1):
            ws.cell(i, j, v)
    fim = 3 + len(linhas)
    estilos.corpo_texto(ws, 4, fim, COLUNAS)
    for r in range(4, fim + 1):
        for c in range(1, COLUNAS + 1):
            ws.cell(r, c).alignment = Alignment(wrap_text=True, vertical="center")
            ws.cell(r, c).border = Border(bottom=BORDA)
    r = _bloco(ws, fim + 2, "Regras", _regras(calc))
    r = _bloco(ws, r + 2, "Qualidade dos dados", _qualidade(calc))
    atipicas = [(a.leitura, "; ".join(a.motivos)) for a in calc.qualidade.atipicas]
    r = _tabela_leituras(ws, r + 2, f"Leituras atípicas, assinaladas mas incluídas nas médias ({len(atipicas)})", atipicas)
    _tabela_leituras(ws, r + 2, f"Pesagens excluídas desde o início ({len(selecao.excluidas)})", selecao.excluidas)
    estilos.larguras(ws, [24, 12, 46, 22, 14, 20])
    ws.freeze_panes = "A4"
