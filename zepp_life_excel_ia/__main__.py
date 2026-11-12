"""Linha de comandos.

  python -m zepp_life_excel_ia gerar --origem exportacao.json --excel saida.xlsx --json saida.json
         [--existente atual.xlsx] [--config config.json] [--hoje AAAA-MM-DD] [--sem-excel-real] [--graficos pasta]
         [--json-resumo resumo.json] [--copia-manuais copia.csv] [--guardar-copia-manuais nova-copia.csv]
  python -m zepp_life_excel_ia validar-exportacao --nova nova.json [--anterior anterior.json] [--permitir-remocoes]

Erros: mensagem numa linha que começa por "ERRO:" e código de saída 1 (2 em validar-exportacao quando
o único problema são pesagens que desapareceram: o atualizador pode perguntar se são aceites).
"""
from __future__ import annotations

import argparse
import json
import sys
from datetime import date
from pathlib import Path

from . import exportacao
from .gerar import gerar


def _gerar(a) -> None:
    gerar(Path(a.origem), Path(a.excel), Path(a.json), existente=Path(a.existente) if a.existente else None,
          config=Path(a.config) if a.config else None, hoje=date.fromisoformat(a.hoje) if a.hoje else None,
          verificar_excel=not a.sem_excel_real, graficos=Path(a.graficos) if a.graficos else None,
          json_resumo=Path(a.json_resumo) if a.json_resumo else None,
          copia_manuais=Path(a.copia_manuais) if a.copia_manuais else None,
          guardar_copia_manuais=Path(a.guardar_copia_manuais) if a.guardar_copia_manuais else None)


def _validar(a) -> None:
    ler = lambda caminho: json.loads(Path(caminho).read_text(encoding="utf-8-sig"))
    anterior = ler(a.anterior) if a.anterior and Path(a.anterior).exists() else None
    for aviso in exportacao.validar(ler(a.nova), anterior, a.permitir_remocoes):
        print(f"Aviso: {aviso}")
    print("Exportação válida.")


def main(argv=None) -> int:
    p = argparse.ArgumentParser(prog="python -m zepp_life_excel_ia", description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    comandos = p.add_subparsers(dest="comando", required=True)

    g = comandos.add_parser("gerar", help="gerar o Excel e o JSON para IA")
    g.add_argument("--origem", required=True, help="exportação JSON do leitor Zepp")
    g.add_argument("--excel", required=True, help="livro Excel a criar")
    g.add_argument("--json", required=True, help="JSON para IA a criar")
    g.add_argument("--existente", help="livro atual, de onde se preservam as medidas manuais")
    g.add_argument("--config", help="caminho do config.json")
    g.add_argument("--hoje", help="data de referência AAAA-MM-DD (testes)")
    g.add_argument("--sem-excel-real", action="store_true", help="não validar no Microsoft Excel")
    g.add_argument("--graficos", help="exportar os gráficos do Painel em PNG para esta pasta")
    g.add_argument("--json-resumo", help="também criar a versão compacta do JSON para IA")
    g.add_argument("--copia-manuais", help="cópia CSV das medidas manuais, usada se o livro atual não existir")
    g.add_argument("--guardar-copia-manuais", help="escrever aqui a cópia CSV atualizada das medidas manuais")
    g.set_defaults(funcao=_gerar)

    v = comandos.add_parser("validar-exportacao", help="validar uma nova descarga antes de a usar")
    v.add_argument("--nova", required=True, help="nova exportação JSON")
    v.add_argument("--anterior", help="exportação anterior (para detetar pesagens perdidas)")
    v.add_argument("--permitir-remocoes", action="store_true", help="aceitar pesagens apagadas na Zepp Life")
    v.set_defaults(funcao=_validar)

    args = p.parse_args(argv)
    try:
        args.funcao(args)
    except exportacao.PesagensRemovidas as erro:
        print(f"ERRO: {erro}", file=sys.stderr)
        return 2
    except Exception as erro:
        print(f"ERRO: {erro}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
