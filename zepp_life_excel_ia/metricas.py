"""Registo único das métricas: tudo o que define uma métrica (JSON, Excel, resumos) está aqui.

A ordem de `METRICAS` é a ordem no JSON. As folhas do Excel e os resumos escolhem as métricas
que mostram (e por que ordem) através das listas `COLUNAS_MEDICOES` e `RESUMO`.
"""
from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class Metrica:
    campo: str                     # nome na exportação do leitor ("massa_gorda" = derivada aqui)
    chave: str                     # chave no JSON para IA
    nome: str                      # nome legível (folha Critérios)
    unidade: str
    significado: str
    origem: str                    # balança | aplicação | derivada
    da_exportacao: bool = True     # False = calculada por este programa
    # Folha "Medições Corporais"
    coluna: str | None = None
    formato: str = "0.00"
    largura: float = 12
    # Resumos semanais e mensais
    resumo: str | None = None              # chave JSON
    cabecalho_resumo: str | None = None
    formato_resumo: str = "0.00"


# Lista explícita: nada fora desta lista é exportado (nem token, conta ou identificadores).
METRICAS = [
    Metrica("Weight", "peso_kg", "Peso", "kg", "Peso medido pela balança", "balança",
            coluna="Peso (kg)", largura=11, resumo="peso_medio_kg", cabecalho_resumo="Peso médio (kg)"),
    Metrica("Height", "altura_cm", "Altura", "cm", "Altura do perfil associada à leitura", "aplicação"),
    Metrica("BMI", "imc", "IMC", "kg/m²", "Índice de massa corporal", "aplicação",
            coluna="IMC", largura=9, resumo="imc_medio", cabecalho_resumo="IMC médio", formato_resumo="0.0"),
    Metrica("BodyFat", "gordura_corporal_percentagem", "Gordura corporal", "%", "Estimativa por bioimpedância",
            "aplicação", coluna="Gordura (%)", formato="0.0", largura=11, resumo="gordura_media_percentagem",
            cabecalho_resumo="Gordura média (%)", formato_resumo="0.0"),
    Metrica("BodyWater", "agua_corporal_percentagem", "Água corporal", "%", "Estimativa de água corporal", "aplicação",
            coluna="Água corporal (%)", formato="0.0", largura=13, resumo="agua_media_percentagem",
            cabecalho_resumo="Água média (%)"),
    Metrica("BoneMass", "massa_ossea_kg", "Massa óssea", "kg", "Estimativa; não é densidade mineral óssea", "aplicação",
            coluna="Massa óssea (kg)", resumo="massa_ossea_media_kg", cabecalho_resumo="Massa óssea média (kg)"),
    Metrica("MuscleMass", "massa_muscular_kg", "Massa muscular", "kg", "Estimativa de massa muscular", "aplicação",
            coluna="Massa muscular (kg)", largura=13, resumo="musculo_medio_kg", cabecalho_resumo="Músculo médio (kg)"),
    Metrica("ProteinRatio", "proteina_percentagem", "Proteína", "%", "Proporção estimada de proteína", "aplicação",
            coluna="Proteína (%)", formato="0.0", largura=11, resumo="proteina_media_percentagem",
            cabecalho_resumo="Proteína média (%)"),
    Metrica("ProteinMass", "massa_proteica_kg", "Massa proteica", "kg", "Derivada pelo leitor: peso × proteína / 100",
            "derivada"),
    Metrica("VisceralFat", "gordura_visceral_indice", "Gordura visceral", "índice",
            "Índice da aplicação; não é percentagem (inteiro no leitor)", "aplicação",
            coluna="Gordura visceral (índice)", formato="0", largura=13, resumo="gordura_visceral_media",
            cabecalho_resumo="Gordura visceral média (índice)", formato_resumo="0.0"),
    Metrica("BasalMetabolism", "metabolismo_basal_kcal_dia", "Metabolismo basal", "kcal/dia",
            "Estimativa de gasto basal (inteiro no leitor)", "aplicação",
            coluna="Metabolismo basal (kcal/dia)", formato="#,##0", largura=15),
    Metrica("MetabolicAge", "idade_corporal_anos", "Idade corporal", "anos", "Estimativa; não é idade cronológica",
            "aplicação", coluna="Idade corporal (anos)", formato="0", resumo="idade_corporal_media",
            cabecalho_resumo="Idade corporal média (anos)", formato_resumo="0.0"),
    Metrica("BodyScore", "pontuacao_corporal", "Pontuação corporal", "pontos", "Pontuação global da aplicação",
            "aplicação", coluna="Pontuação corporal", formato="0", resumo="pontuacao_corporal_media",
            cabecalho_resumo="Pontuação corporal média", formato_resumo="0.0"),
    Metrica("PhysiqueRating", "classificacao_fisica_codigo", "Classificação física", "código",
            "Código original, sem interpretação documentada", "aplicação"),
    Metrica("StandBodyWeight", "peso_padrao_kg", "Peso padrão", "kg",
            "Referência da aplicação; não é um objetivo individual", "aplicação"),
    Metrica("Impedance", "impedancia_ohm", "Impedância", "Ω", "Impedância registada pela balança", "balança"),
    Metrica("massa_gorda", "massa_gorda_kg", "Massa gorda", "kg", "peso_kg × gordura_corporal_percentagem / 100",
            "derivada", da_exportacao=False, coluna="Massa gorda (kg)", resumo="massa_gorda_media_kg",
            cabecalho_resumo="Massa gorda média (kg)"),
]
POR_CAMPO = {m.campo: m for m in METRICAS}
DA_EXPORTACAO = [m for m in METRICAS if m.da_exportacao]

# Colunas de métricas na folha "Medições Corporais", por ordem (a seguir a "Data/hora").
COLUNAS_MEDICOES = ["Weight", "BMI", "BodyFat", "BodyWater", "BoneMass", "MuscleMass", "ProteinRatio",
                    "massa_gorda", "VisceralFat", "BasalMetabolism", "MetabolicAge", "BodyScore"]

# Métricas dos resumos, por ordem. As duas primeiras têm coluna de variação a seguir.
RESUMO = ["Weight", "BodyFat", "MuscleMass", "massa_gorda", "BodyWater", "ProteinRatio", "BoneMass",
          "BMI", "VisceralFat", "MetabolicAge", "BodyScore"]


@dataclass(frozen=True)
class MediaMovel:
    campo: str
    coluna: str          # cabeçalho em "Medições Corporais"
    chave: str           # chave no JSON
    significado: str


MEDIAS_MOVEIS = [
    MediaMovel("Weight", "Peso, média 7 dias (kg)", "peso_media_7d_kg", "Média das leituras aceites em ]t − 7 dias, t]"),
    MediaMovel("BodyFat", "Gordura, média 7 dias (%)", "gordura_media_7d_percentagem", "Média móvel de 7 dias da gordura"),
    MediaMovel("MuscleMass", "Músculo, média 7 dias (kg)", "musculo_media_7d_kg",
               "Média móvel de 7 dias da massa muscular"),
]

# Uma leitura só é aceite com composição corporal completa.
CAMPOS_COMPOSICAO = ["BMI", "BodyFat", "BodyWater", "BoneMass", "BasalMetabolism", "MuscleMass",
                     "VisceralFat", "ProteinMass", "ProteinRatio", "MetabolicAge", "BodyScore"]

# Limites largos: apanham erros de leitura evidentes, não julgam a composição corporal.
LIMITES = {
    "Weight": (30, 250), "BMI": (10, 60), "BodyFat": (2, 70), "BodyWater": (25, 80),
    "BoneMass": (0.5, 8), "MuscleMass": (15, 150), "ProteinRatio": (5, 35),
    "VisceralFat": (1, 59), "BasalMetabolism": (700, 4500), "MetabolicAge": (10, 100),
    "BodyScore": (0, 100),
}

# Indicadores do Painel e do JSON (média 7 dias inicial/atual, variação, ritmo).
INDICADORES = ["Weight", "BodyFat", "MuscleMass"]
