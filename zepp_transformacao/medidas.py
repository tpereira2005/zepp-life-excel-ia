"""Medidas manuais semanais (cintura e pescoço): preservação, importação inicial e validação."""
from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path

from .xlsx_bruto import ler_livro

FOLHA = "Medidas Manuais"
CABECALHOS = ["Semana", "Data da medição", "Cintura (cm)", "Variação da cintura (cm)",
              "Pescoço (cm)", "Estado", "Observações"]
LIMITES = {"cintura": (1, 300), "pescoco": (1, 150)}
EPOCA_EXCEL = date(1899, 12, 30)


def domingo(dia: date) -> date:
    return dia - timedelta(days=dia.isoweekday() % 7)


def rotulo(inicio: date) -> str:
    fim = inicio + timedelta(days=6)
    return f"{inicio:%d/%m/%Y} - {fim:%d/%m/%Y}"


def inicio_do_rotulo(texto: str) -> date:
    m = re.fullmatch(r"(\d{2})/(\d{2})/(\d{4}) - (\d{2})/(\d{2})/(\d{4})", texto.strip())
    if not m:
        raise ValueError(f"Semana com formato inesperado: {texto!r}")
    inicio = date(int(m[3]), int(m[2]), int(m[1]))
    if inicio.isoweekday() != 7 or rotulo(inicio) != texto.strip():
        raise ValueError(f"Semana inválida (deve ir de domingo a sábado): {texto!r}")
    return inicio


@dataclass
class Semana:
    inicio: date
    data: date | None
    cintura: float | None
    pescoco: float | None
    notas: str | None
    origem: str = "nova"          # excel | csv | nova

    @property
    def rotulo(self) -> str:
        return rotulo(self.inicio)

    @property
    def fim(self) -> date:
        return self.inicio + timedelta(days=6)

    @property
    def medida(self) -> bool:
        return self.cintura is not None or self.pescoco is not None

    @property
    def estado(self) -> str:
        if not self.medida:
            return "sem_medicao" if self.notas else "por_preencher"
        return "completa" if self.cintura is not None and self.pescoco is not None else "incompleta"


def _data(valor, semana: str) -> date | None:
    if valor is None or valor == "":
        return None
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        return EPOCA_EXCEL + timedelta(days=math.floor(valor))
    if isinstance(valor, str):
        texto = valor.strip()
        for formato in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y %H:%M", "%Y-%m-%dT%H:%M:%S"):
            try:
                return datetime.strptime(texto, formato).date()
            except ValueError:
                pass
    raise ValueError(f"Data da medição ilegível na semana {semana}: {valor!r}. Use dd/mm/aaaa.")


def _medida(valor, semana: str, nome: str) -> float | None:
    if valor is None or (isinstance(valor, str) and not valor.strip()):
        return None
    if isinstance(valor, str):
        try:
            valor = float(valor.strip().replace(",", "."))
        except ValueError:
            raise ValueError(f"{nome.capitalize()} não numérica na semana {semana}: {valor!r}") from None
    if isinstance(valor, bool) or not isinstance(valor, (int, float)) or not math.isfinite(valor):
        raise ValueError(f"{nome.capitalize()} inválida na semana {semana}: {valor!r}")
    return float(valor)


def ler_existentes(caminho: Path) -> dict[date, Semana]:
    """Lê a folha Medidas Manuais pelos nomes dos cabeçalhos (não pela posição das colunas)."""
    linhas = ler_livro(caminho, permitir_erros=True).get(FOLHA)
    if linhas is None:
        raise ValueError(f"A folha '{FOLHA}' não existe no Excel atual.")
    indice = next((i for i, l in enumerate(linhas) if l and l[0] == "Semana"), None)
    if indice is None:
        raise ValueError("Cabeçalhos das medidas manuais não encontrados.")
    cab = [str(c).strip() if c is not None else None for c in linhas[indice]]
    obrigatorios = {"Semana", "Data da medição", "Cintura (cm)", "Pescoço (cm)", "Observações"}
    em_falta = obrigatorios - set(cab)
    if em_falta:
        raise ValueError("Colunas em falta nas medidas manuais: " + ", ".join(sorted(em_falta)))
    col = {nome: cab.index(nome) for nome in obrigatorios}
    semanas: dict[date, Semana] = {}
    for linha in linhas[indice + 1:]:
        valor = lambda nome: linha[col[nome]] if col[nome] < len(linha) else None
        texto = valor("Semana")
        if texto is None or str(texto).strip() == "":
            continue
        texto = str(texto)
        inicio = inicio_do_rotulo(texto)
        if inicio in semanas:
            raise ValueError(f"Semana duplicada nas medidas manuais: {texto}")
        notas = valor("Observações")
        semanas[inicio] = Semana(
            inicio=inicio,
            data=_data(valor("Data da medição"), texto),
            cintura=_medida(valor("Cintura (cm)"), texto, "cintura"),
            pescoco=_medida(valor("Pescoço (cm)"), texto, "pescoço"),
            notas=None if notas is None or str(notas) == "" else str(notas),
            origem="excel",
        )
    return semanas


def ler_csv(caminho: Path) -> dict[date, Semana]:
    """CSV inicial: Data;leituras cintura;média cintura;leituras pescoço;média pescoço."""
    semanas: dict[date, Semana] = {}
    linhas = caminho.read_text(encoding="utf-8-sig").strip().splitlines()[1:]
    for linha in linhas:
        if not linha.strip():
            continue
        data_txt, leituras_cintura, cintura_txt, leituras_pescoco, pescoco_txt = linha.split(";")
        dia = datetime.strptime(data_txt.strip(), "%d/%m/%Y").date()
        media = lambda t: None if t.strip() in ("-", "") else float(t.strip().replace(",", "."))
        cintura, pescoco = media(cintura_txt), media(pescoco_txt)
        notas = ("Sem medição: esqueci-me" if cintura is None and pescoco is None else
                 f"Leituras — cintura: {leituras_cintura.strip()}; pescoço: {leituras_pescoco.strip()}")
        semanas[domingo(dia)] = Semana(domingo(dia), dia, cintura, pescoco, notas, "csv")
    return semanas


def combinar(existentes: dict[date, Semana], csv: dict[date, Semana], primeira: date, ultima: date) -> list[Semana]:
    """Excel atual > CSV inicial > semana nova. Campos apagados no Excel continuam apagados."""
    primeira = min([domingo(primeira), *existentes])
    ultima = max([domingo(ultima), *existentes])
    semanas, inicio = [], primeira
    while inicio <= ultima:
        if inicio in existentes:
            semanas.append(existentes[inicio])
        elif inicio in csv:
            semanas.append(csv[inicio])
        else:
            semanas.append(Semana(inicio, inicio, None, None, None, "nova"))
        inicio += timedelta(days=7)
    validar(semanas)
    return semanas


def validar(semanas: list[Semana]) -> None:
    for s in semanas:
        for nome, valor in (("cintura", s.cintura), ("pescoco", s.pescoco)):
            lo, hi = LIMITES[nome]
            if valor is not None and not lo <= valor <= hi:
                raise ValueError(f"Medida manual fora dos limites ({nome} = {valor}) na semana {s.rotulo}.")
        if s.medida:
            if s.data is None:
                raise ValueError(f"Falta a data da medição na semana {s.rotulo}.")
            if not s.inicio <= s.data <= s.fim:
                raise ValueError(f"A data {s.data:%d/%m/%Y} não pertence à semana {s.rotulo}.")


def leituras_das_notas(notas: str | None) -> tuple[list[float] | None, list[float] | None]:
    """Extrai as leituras individuais quando a nota segue 'Leituras — cintura: a - b - c; pescoço: x - y - z'."""
    if not notas:
        return None, None
    m = re.search(r"cintura:\s*([^;]+);\s*pesco[çc]o:\s*(.+)$", notas, flags=re.IGNORECASE)
    if not m:
        return None, None

    def numeros(texto: str) -> list[float] | None:
        try:
            valores = [float(p.strip().replace(",", ".")) for p in texto.split(" - ")]
        except ValueError:
            return None
        return valores or None

    return numeros(m[1]), numeros(m[2])
