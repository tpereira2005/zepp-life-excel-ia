"""Dados fictícios e determinísticos para testes e demonstrações (sem dados pessoais)."""
from __future__ import annotations

import json
import random
from datetime import date, datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

LISBOA = ZoneInfo("Europe/Lisbon")
INICIO = date(2026, 1, 4)  # domingo


def registo(momento: datetime, peso: float, gordura: float, musculo: float, completo: bool = True,
            perfil: int | None = -1, altura: float = 178.0) -> dict:
    r = {"Date": momento.isoformat(timespec="seconds"), "Weight": round(peso, 2), "BMI": round(peso / (altura / 100) ** 2, 6),
         "Height": altura, "Source": "AABBCCDDEEFF0011"}
    if completo:
        r.update({"BodyFat": round(gordura, 1), "BodyWater": round(70 - gordura * 0.7, 1), "BoneMass": 3.1,
                  "MetabolicAge": 30, "MuscleMass": round(musculo, 2), "ProteinRatio": 19.5,
                  "ProteinMass": round(peso * 0.195, 6), "VisceralFat": 9 if gordura > 20 else 8,
                  "BasalMetabolism": int(1500 + musculo * 3), "BodyScore": 80, "StandBodyWeight": 70.4,
                  "Impedance": 500, "EncryptImpedance": "500"})
    if perfil is not None:
        r["ProfileID"] = perfil
    return r


def exportacao(dias: int = 120, semente: int = 7, inicio: date = INICIO) -> list[dict]:
    aleatorio = random.Random(semente)
    registos = []
    for d in range(-5, dias):
        dia = inicio + timedelta(days=d)
        if aleatorio.random() < 0.08:          # alguns dias sem pesagem
            continue
        hora = 10 if dia.isoweekday() >= 6 else 7
        momento = datetime(dia.year, dia.month, dia.day, hora, aleatorio.randint(0, 59), tzinfo=LISBOA)
        t = max(d, 0) / max(dias, 1)
        peso = 84 - 7 * t + aleatorio.gauss(0, 0.35)
        gordura = 25 - 5 * t + aleatorio.gauss(0, 0.4)
        musculo = 59.5 - 0.6 * t + aleatorio.gauss(0, 0.25)
        registos.append(registo(momento, peso, gordura, musculo))
        if d % 23 == 11:                       # pesagem só de peso (sem impedância)
            registos.append(registo(momento + timedelta(hours=5), peso + 0.6, 0, 0, completo=False))
        if d % 31 == 17:                       # repesagem minutos depois
            registos.append(registo(momento + timedelta(minutes=3), peso - 0.1, gordura, musculo))
    registos.sort(key=lambda r: r["Date"], reverse=True)   # a API devolve do mais recente para o mais antigo
    return registos


def csv_manual(semanas: int = 14, semente: int = 3, inicio: date = INICIO) -> str:
    aleatorio = random.Random(semente)
    linhas = ["Data;Cintura — medições (cm);Cintura — média (cm);Pescoço — medições (cm);Pescoço — média (cm)"]
    for s in range(semanas):
        dia = inicio + timedelta(days=7 * s)
        if s in (6, 7):
            continue                            # férias: sem linha
        if s == 4:
            linhas.append(f"{dia:%d/%m/%Y};esqueci-me;-;esqueci-me;-")
            continue
        cintura = 96 - 0.6 * s
        pescoco = 40 - 0.08 * s
        c = [round(cintura + aleatorio.uniform(-0.15, 0.15), 1) for _ in range(3)]
        p = [round(pescoco + aleatorio.uniform(-0.1, 0.1), 1) for _ in range(3)]
        fmt = lambda v: f"{v:g}".replace(".", ",")
        media = lambda v: f"{round(sum(v) / 3, 1):g}".replace(".", ",")
        linhas.append(f"{dia:%d/%m/%Y};{' - '.join(map(fmt, c))};{media(c)};{' - '.join(map(fmt, p))};{media(p)}")
    return "\n".join(linhas) + "\n"


def criar_projeto(pasta: Path, dias: int = 120, semanas_csv: int = 14) -> dict:
    """Cria config.json, exportação e CSV fictícios numa pasta; devolve os caminhos."""
    pasta.mkdir(parents=True, exist_ok=True)
    (pasta / "inputs").mkdir(exist_ok=True)
    csv = pasta / "inputs" / "medicoes.csv"
    csv.write_text(csv_manual(semanas_csv), encoding="utf-8")
    config = pasta / "config.json"
    config.write_text(json.dumps({
        "inicio_transformacao": INICIO.isoformat(), "fuso_horario": "Europe/Lisbon", "perfil_titular_id": -1,
        "altura_perfil_sem_id_cm": 178, "protocolo_pesagem": "Ao acordar, depois de ir à casa de banho.",
        "csv_medidas_iniciais": "inputs/medicoes.csv"}, ensure_ascii=False, indent=2), encoding="utf-8")
    origem = pasta / "exportacao.json"
    origem.write_text(json.dumps(exportacao(dias), indent=1), encoding="utf-8")
    return {"config": config, "origem": origem, "csv": csv}
