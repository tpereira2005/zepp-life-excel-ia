"""Gera o Excel e o JSON para IA a partir da exportação Zepp e das medidas manuais existentes."""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

from . import config as configuracao
from . import excel, exportacao, ia_json, medidas, resumos, verificacao
from .leituras import selecionar
from .resumos import Calculos


@dataclass
class Resultado:
    calculos: Calculos
    dados: dict                  # conteúdo do JSON para IA
    esperados: list[verificacao.Esperado]
    preservadas: int             # semanas lidas do Excel existente
    estado_excel: str            # "verificado no Excel …", "nao_verificado" ou "indisponivel"


def _medidas_manuais(cfg, existente: Path | None, inicio: date, ate: date) -> tuple[list[medidas.Semana], int]:
    existentes = {}
    if existente and existente.exists():
        try:
            existentes = medidas.ler_existentes(existente)
        except Exception as erro:
            raise RuntimeError(f"Não foi possível preservar as medidas manuais: {erro}") from erro
    csv = {}
    if cfg.csv_medidas_iniciais and cfg.csv_medidas_iniciais.exists():
        csv = medidas.ler_csv(cfg.csv_medidas_iniciais)
    elif not existentes and cfg.csv_medidas_iniciais:
        print(f"Aviso: CSV inicial não encontrado ({cfg.csv_medidas_iniciais}); medidas manuais começam vazias.")
    return medidas.combinar(existentes, csv, inicio, ate), len(existentes)


def _verificar_no_excel(caminho: Path, esperados: list[verificacao.Esperado], graficos: Path | None,
                        calc: Calculos) -> str:
    try:
        valores = verificacao.calcular_no_excel(caminho, graficos)
    except verificacao.ExcelIndisponivel as aviso:
        print(f"Aviso: {aviso} Verificação no Excel real ignorada.")
        return "indisponivel"
    diferencas = verificacao.comparar(valores, esperados)
    if diferencas:
        raise RuntimeError("O Excel calculou valores diferentes do esperado:\n  " + "\n  ".join(diferencas[:20]))
    print(f"Excel real: {len(calc.semanais)} semanas, {len(calc.mensais)} meses e {len(calc.leituras)} leituras "
          f"confirmados; sem erros de fórmula.")
    return f"verificado no Excel {valores.get('versao_excel', '')}".strip()


def gerar(origem: Path, excel_saida: Path, json_saida: Path, *, existente: Path | None = None,
          config: Path | None = None, hoje: date | None = None, verificar_excel: bool = True,
          graficos: Path | None = None, agora: datetime | None = None) -> Resultado:
    """Lê a exportação e as medidas manuais, calcula tudo uma vez, escreve o Excel, confirma-o no
    Microsoft Excel (se pedido e disponível) e escreve o JSON com o hash do Excel."""
    cfg = configuracao.carregar(config)
    agora = agora or datetime.now(cfg.fuso)
    hoje = hoje or agora.date()
    selecao = selecionar(exportacao.carregar(origem), cfg)
    ate = max(max(lt.local for lt in selecao.aceites).date(), hoje)
    semanas, preservadas = _medidas_manuais(cfg, existente, cfg.inicio, ate)
    calc = resumos.calcular(cfg, selecao, semanas, ate)

    excel_saida.parent.mkdir(parents=True, exist_ok=True)
    esperados = excel.construir(excel_saida, calc, agora)
    print(f"Excel: {excel_saida}")
    estado = _verificar_no_excel(excel_saida, esperados, graficos, calc) if verificar_excel else "nao_verificado"

    dados = ia_json.construir(calc, excel_saida, agora)
    json_saida.write_text(ia_json.serializar(dados), encoding="utf-8")
    print(f"JSON: {json_saida}")
    c = selecao.contagens
    print(f"LEITURAS={c['aceites']};EXCLUIDAS={len(selecao.excluidas)};SEMANAS={len(semanas)};"
          f"MESES={len(calc.mensais)};PRESERVADAS={preservadas};EXCEL={estado}")
    return Resultado(calc, dados, esperados, preservadas, estado)
