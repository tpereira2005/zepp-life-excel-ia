"""Gera o Excel e o JSON para IA a partir da exportação Zepp e das medidas manuais existentes.

Uso: python -m zepp_transformacao.gerar --origem exportacao.json --excel saida.xlsx --json saida.json
         [--existente atual.xlsx] [--config config.json] [--hoje AAAA-MM-DD] [--sem-excel-real] [--graficos pasta]
"""
from __future__ import annotations

import argparse
import sys
from datetime import date, datetime
from pathlib import Path

from . import config as configuracao
from . import excel, ia_json, medidas, resumos, verificacao
from .leituras import carregar_exportacao, selecionar


def executar(args) -> dict:
    cfg = configuracao.carregar(Path(args.config) if args.config else None)
    agora = datetime.now(cfg.fuso)
    hoje = date.fromisoformat(args.hoje) if args.hoje else agora.date()

    selecao = selecionar(carregar_exportacao(Path(args.origem)), cfg)
    leituras = selecao.aceites

    existentes = {}
    if args.existente and Path(args.existente).exists():
        try:
            existentes = medidas.ler_existentes(Path(args.existente))
        except Exception as erro:
            raise RuntimeError(f"Não foi possível preservar as medidas manuais: {erro}") from erro
    csv = {}
    if cfg.csv_medidas_iniciais and cfg.csv_medidas_iniciais.exists():
        csv = medidas.ler_csv(cfg.csv_medidas_iniciais)
    elif not existentes and cfg.csv_medidas_iniciais:
        print(f"Aviso: CSV inicial não encontrado ({cfg.csv_medidas_iniciais}); medidas manuais começam vazias.")
    ultima_data = max(max(l.local for l in leituras).date(), hoje)
    semanas = medidas.combinar(existentes, csv, cfg.inicio, ultima_data)

    semanais = resumos.semanais(leituras, semanas, cfg)
    mensais = resumos.mensais(leituras, semanas, cfg, ultima_data)
    moveis = resumos.medias_moveis(leituras)
    indicadores = resumos.indicadores(leituras, semanas)
    variacoes = resumos.variacoes_cintura(semanas)

    caminho_excel = Path(args.excel)
    caminho_excel.parent.mkdir(parents=True, exist_ok=True)
    excel.construir(caminho_excel, cfg, selecao, semanas, semanais, mensais, moveis, agora)
    print(f"Excel: {caminho_excel}")

    estado_excel = "nao_verificado"
    if not args.sem_excel_real:
        try:
            valores = verificacao.calcular_no_excel(caminho_excel, Path(args.graficos) if args.graficos else None)
            diferencas = verificacao.comparar(
                valores, semanais, mensais, moveis, [l.massa_gorda for l in leituras], indicadores,
                [variacoes[s.inicio] for s in semanas], [s.estado for s in semanas])
            if diferencas:
                raise RuntimeError("O Excel calculou valores diferentes do esperado:\n  " + "\n  ".join(diferencas[:20]))
            estado_excel = f"verificado no Excel {valores.get('versao_excel', '')}".strip()
            print(f"Excel real: {len(semanais)} semanas, {len(mensais)} meses e {len(leituras)} leituras confirmados; sem erros de fórmula.")
        except verificacao.ExcelIndisponivel as aviso:
            print(f"Aviso: {aviso} Verificação no Excel real ignorada.")

    dados = ia_json.construir(cfg, selecao, semanas, variacoes, semanais, mensais, moveis, indicadores,
                              caminho_excel, agora)
    Path(args.json).write_text(ia_json.serializar(dados), encoding="utf-8")
    print(f"JSON: {args.json}")
    c = selecao.contagens
    print(f"LEITURAS={c['aceites']};EXCLUIDAS={len(selecao.excluidas)};SEMANAS={len(semanas)};MESES={len(mensais)};"
          f"PRESERVADAS={len(existentes)};EXCEL={estado_excel}")
    return {"selecao": selecao, "semanas": semanas, "semanais": semanais, "mensais": mensais, "dados": dados}


def main(argv=None) -> int:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--origem", required=True, help="exportação JSON do leitor Zepp")
    p.add_argument("--excel", required=True, help="livro Excel a criar")
    p.add_argument("--json", required=True, help="JSON para IA a criar")
    p.add_argument("--existente", help="livro atual, de onde se preservam as medidas manuais")
    p.add_argument("--config", help="caminho do config.json")
    p.add_argument("--hoje", help="data de referência (testes)")
    p.add_argument("--sem-excel-real", action="store_true", help="não validar no Microsoft Excel")
    p.add_argument("--graficos", help="exportar os gráficos do Painel em PNG para esta pasta")
    args = p.parse_args(argv)
    try:
        executar(args)
    except Exception as erro:
        print(f"ERRO: {erro}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
