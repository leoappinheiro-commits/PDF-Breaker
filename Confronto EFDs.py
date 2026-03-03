"""Confronto de registros C170 entre EFD Fiscal e EFD Contribuições.

O script lê arquivos TXT da estrutura:
MAIN_DIR/Resultado/EFD Fiscal
MAIN_DIR/Resultado/EFD Contribuições

E gera um Excel com os itens (C170) presentes no Fiscal e ausentes no
Contribuições.
"""

from __future__ import annotations

import argparse
import logging
import re
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd


# Índices dos campos no layout SPED (split por "|").
# Exemplo de linha: |C100|...|
# Após split("|"), o registro fica na posição 1.
INDICE_REGISTRO = 1

# 0150 - cadastro participante
INDICE_0150_COD_PART = 2
INDICE_0150_NOME = 3
INDICE_0150_CNPJ = 5

# 0200 - cadastro item
INDICE_0200_COD_ITEM = 2
INDICE_0200_DESCR_ITEM = 3
INDICE_0200_TIPO_ITEM = 7
INDICE_0200_NCM = 8

# C100 - cabeçalho da nota
INDICE_C100_COD_PART = 4
INDICE_C100_SERIE = 8
INDICE_C100_NUM_DOC = 9
INDICE_C100_DT_DOC = 10

# C170 - itens da nota
INDICE_C170_COD_ITEM = 3
INDICE_C170_VL_ITEM = 7
INDICE_C170_CFOP = 11

# A100 - cabeçalho de serviços
INDICE_A100_IND_OPER = 2
INDICE_A100_COD_PART = 4
INDICE_A100_SERIE = 6
INDICE_A100_NUM_DOC = 8
INDICE_A100_DT_DOC = 10

# A170 - itens de serviços
INDICE_A170_COD_ITEM = 3

# 0000 - abertura
INDICE_0000_DT_INI = 4

# Camada dinâmica de layout para D100/C500 por tipo e ano do SPED.
# Os índices abaixo consideram o split por "|" da linha do arquivo.
LAYOUTS = {
    "FISCAL": {
        2014: {
            "D100": {"COD_PART": 4, "SERIE": 7, "NUM_DOC": 9, "DT_DOC": 11, "VL_DOC": 15, "CFOP": 14},
            "C500": {"COD_PART": 2, "SERIE": 5, "NUM_DOC": 8, "DT_DOC": 9, "VL_DOC": 11},
        },
        2015: {
            "D100": {"COD_PART": 4, "SERIE": 7, "NUM_DOC": 9, "DT_DOC": 11, "VL_DOC": 15, "CFOP": 14},
            "C500": {"COD_PART": 2, "SERIE": 5, "NUM_DOC": 8, "DT_DOC": 9, "VL_DOC": 11},
        },
        2016: {
            "D100": {"COD_PART": 4, "SERIE": 7, "NUM_DOC": 9, "DT_DOC": 11, "VL_DOC": 15, "CFOP": 14},
            "C500": {"COD_PART": 2, "SERIE": 5, "NUM_DOC": 8, "DT_DOC": 9, "VL_DOC": 11},
        },
        2017: {
            "D100": {"COD_PART": 4, "SERIE": 7, "NUM_DOC": 9, "DT_DOC": 11, "VL_DOC": 15, "CFOP": 14},
            "C500": {"COD_PART": 2, "SERIE": 5, "NUM_DOC": 8, "DT_DOC": 9, "VL_DOC": 11},
        },
        2018: {
            "D100": {"COD_PART": 4, "SERIE": 7, "NUM_DOC": 9, "DT_DOC": 11, "VL_DOC": 15, "CFOP": 14},
            "C500": {"COD_PART": 2, "SERIE": 5, "NUM_DOC": 8, "DT_DOC": 9, "VL_DOC": 11},
        },
        2019: {
            "D100": {"COD_PART": 4, "SERIE": 7, "NUM_DOC": 9, "DT_DOC": 11, "VL_DOC": 15, "CFOP": 14},
            "C500": {"COD_PART": 2, "SERIE": 5, "NUM_DOC": 8, "DT_DOC": 9, "VL_DOC": 11},
        },
        2020: {
            "D100": {"COD_PART": 4, "SERIE": 7, "NUM_DOC": 9, "DT_DOC": 11, "VL_DOC": 15, "CFOP": 14},
            "C500": {"COD_PART": 2, "SERIE": 5, "NUM_DOC": 8, "DT_DOC": 9, "VL_DOC": 11},
        },
        2021: {
            "D100": {"COD_PART": 4, "SERIE": 7, "NUM_DOC": 9, "DT_DOC": 11, "VL_DOC": 15, "CFOP": 14},
            "C500": {"COD_PART": 2, "SERIE": 5, "NUM_DOC": 8, "DT_DOC": 9, "VL_DOC": 11},
        },
        2022: {
            "D100": {"COD_PART": 4, "SERIE": 7, "NUM_DOC": 9, "DT_DOC": 11, "VL_DOC": 15, "CFOP": 14},
            "C500": {"COD_PART": 2, "SERIE": 5, "NUM_DOC": 8, "DT_DOC": 9, "VL_DOC": 11},
        },
        2023: {
            "D100": {"COD_PART": 4, "SERIE": 7, "NUM_DOC": 9, "DT_DOC": 11, "VL_DOC": 15, "CFOP": 14},
            "C500": {"COD_PART": 2, "SERIE": 5, "NUM_DOC": 8, "DT_DOC": 9, "VL_DOC": 11},
        },
        2024: {
            "D100": {"COD_PART": 4, "SERIE": 7, "NUM_DOC": 9, "DT_DOC": 11, "VL_DOC": 15, "CFOP": 14},
            "C500": {"COD_PART": 2, "SERIE": 5, "NUM_DOC": 8, "DT_DOC": 9, "VL_DOC": 11},
        },
        2025: {
            "D100": {"COD_PART": 4, "SERIE": 7, "NUM_DOC": 9, "DT_DOC": 11, "VL_DOC": 15, "CFOP": 14},
            "C500": {"COD_PART": 2, "SERIE": 5, "NUM_DOC": 8, "DT_DOC": 9, "VL_DOC": 11},
        },
    },
    "CONTRIBUICOES": {
        2014: {
            "D100": {"COD_PART": 4, "SERIE": 7, "NUM_DOC": 9, "DT_DOC": 11, "VL_DOC": 15, "CFOP": 14},
            "C500": {"COD_PART": 2, "SERIE": 5, "NUM_DOC": 8, "DT_DOC": 9, "VL_DOC": 11},
        },
        2015: {
            "D100": {"COD_PART": 4, "SERIE": 7, "NUM_DOC": 9, "DT_DOC": 11, "VL_DOC": 15, "CFOP": 14},
            "C500": {"COD_PART": 2, "SERIE": 5, "NUM_DOC": 8, "DT_DOC": 9, "VL_DOC": 11},
        },
        2016: {
            "D100": {"COD_PART": 4, "SERIE": 7, "NUM_DOC": 9, "DT_DOC": 11, "VL_DOC": 15, "CFOP": 14},
            "C500": {"COD_PART": 2, "SERIE": 5, "NUM_DOC": 8, "DT_DOC": 9, "VL_DOC": 11},
        },
        2017: {
            "D100": {"COD_PART": 4, "SERIE": 7, "NUM_DOC": 9, "DT_DOC": 11, "VL_DOC": 15, "CFOP": 14},
            "C500": {"COD_PART": 2, "SERIE": 5, "NUM_DOC": 8, "DT_DOC": 9, "VL_DOC": 11},
        },
        2018: {
            "D100": {"COD_PART": 4, "SERIE": 7, "NUM_DOC": 9, "DT_DOC": 11, "VL_DOC": 15, "CFOP": 14},
            "C500": {"COD_PART": 2, "SERIE": 5, "NUM_DOC": 8, "DT_DOC": 9, "VL_DOC": 11},
        },
        2019: {
            "D100": {"COD_PART": 4, "SERIE": 7, "NUM_DOC": 9, "DT_DOC": 11, "VL_DOC": 15, "CFOP": 14},
            "C500": {"COD_PART": 2, "SERIE": 5, "NUM_DOC": 8, "DT_DOC": 9, "VL_DOC": 11},
        },
        2020: {
            "D100": {"COD_PART": 4, "SERIE": 7, "NUM_DOC": 9, "DT_DOC": 11, "VL_DOC": 15, "CFOP": 14},
            "C500": {"COD_PART": 2, "SERIE": 5, "NUM_DOC": 8, "DT_DOC": 9, "VL_DOC": 11},
        },
        2021: {
            "D100": {"COD_PART": 4, "SERIE": 7, "NUM_DOC": 9, "DT_DOC": 11, "VL_DOC": 15, "CFOP": 14},
            "C500": {"COD_PART": 2, "SERIE": 5, "NUM_DOC": 8, "DT_DOC": 9, "VL_DOC": 11},
        },
        2022: {
            "D100": {"COD_PART": 4, "SERIE": 7, "NUM_DOC": 9, "DT_DOC": 11, "VL_DOC": 15, "CFOP": 14},
            "C500": {"COD_PART": 2, "SERIE": 5, "NUM_DOC": 8, "DT_DOC": 9, "VL_DOC": 11},
        },
        2023: {
            "D100": {"COD_PART": 4, "SERIE": 7, "NUM_DOC": 9, "DT_DOC": 11, "VL_DOC": 15, "CFOP": 14},
            "C500": {"COD_PART": 2, "SERIE": 5, "NUM_DOC": 8, "DT_DOC": 9, "VL_DOC": 11},
        },
        2024: {
            "D100": {"COD_PART": 4, "SERIE": 7, "NUM_DOC": 9, "DT_DOC": 11, "VL_DOC": 15, "CFOP": 14},
            "C500": {"COD_PART": 2, "SERIE": 5, "NUM_DOC": 8, "DT_DOC": 9, "VL_DOC": 11},
        },
        2025: {
            "D100": {"COD_PART": 4, "SERIE": 7, "NUM_DOC": 9, "DT_DOC": 11, "VL_DOC": 15, "CFOP": 14},
            "C500": {"COD_PART": 2, "SERIE": 5, "NUM_DOC": 8, "DT_DOC": 9, "VL_DOC": 11},
        },
    },
}



COLUNAS_CNAE_CONFIG = [
    "CNAE",
    "Descrição",
    "Regime_PIS_COFINS",
    "Setor_Economico",
    "Aplica_Credito_Presumido",
    "Cadeia_Agro",
    "Permite_Credito_Ativo",
    "Permite_Credito_Energia",
    "Permite_Credito_Frete",
    "Permite_Credito_Aluguel",
    "Permite_Credito_Armazenagem",
    "Permite_Credito_Importacao",
    "Permite_Credito_Exportacao",
]

COLUNAS_MATRIZ_CNAE = [
    "CNAE",
    "Tipo_Item",
    "Permite_Credito",
    "Peso_Score",
    "Nivel_Risco",
    "Fundamentacao_Tecnica",
    "Observacao_Estrategica",
]

FUNDAMENTO_REGIME_CUMULATIVO = "Leis 10.637/2002 e 10.833/2003 - regime cumulativo não gera crédito"
FUNDAMENTO_MONOFASICO = "Lei 10.925/2004, Lei 10.147/2000 e Lei 10.485/2002 - incidência concentrada/monofásica"
FUNDAMENTO_ST = "IN RFB 2.121/2022 e Lei 9.718/1998 - vedação para itens sob substituição tributária"
FUNDAMENTO_TEMA_779 = "Tema 779/STJ e IN RFB 2.121/2022 - essencialidade/relevância do insumo"
FUNDAMENTO_AGRO = "Lei 10.925/2004 - hipóteses de crédito presumido para cadeia agro"
FUNDAMENTO_DECRETO_8426 = "Decreto 8.426/2015 - contexto de alíquotas PIS/COFINS"

MATRIZ_PADRAO = {
    "Permite_Credito": "Depende",
    "Peso_Score": 0,
    "Nivel_Risco": "Médio",
    "Fundamentacao_Tecnica": FUNDAMENTO_TEMA_779,
    "Observacao_Estrategica": "Aplicar validação documental e jurídica específica.",
}

TIPO_ITEM_KEYWORDS = [
    ("Energia", ["energia", "eletrica", "elétrica", "kwh", "c500"]),
    ("Insumo_Produtivo", ["insumo", "materia prima", "matéria prima", "produto intermediario", "produto intermediário"]),
    ("Embalagem", ["embalagem", "caixa", "rotulo", "rótulo", "saco", "filme"]),
    ("Imobilizado", ["imobilizado", "ativo", "maquina", "máquina", "equipamento"]),
    ("Manutencao_Industrial", ["manutencao", "manutenção", "industrial", "reparo", "lubrificante"]),
    ("Frete_Aquisicao", ["frete", "transporte carga", "cte", "ct-e"]),
    ("Armazenagem", ["armazenagem", "armazenamento", "estocagem"]),
    ("Aluguel", ["aluguel", "locacao", "locação", "arrendamento"]),
    ("Servico_Tecnico", ["servico tecnico", "serviço técnico", "engenharia", "assistencia tecnica", "assistência técnica"]),
    ("Administrativo", ["administrativo", "contabilidade", "juridico", "jurídico", "rh", "recursos humanos"]),
    ("Marketing_Publicidade", ["marketing", "publicidade", "propaganda", "midia", "mídia"]),
    ("TI_Sistemas", ["software", "sistema", "ti", "tecnologia", "licenca", "licença"]),
    ("Equipamento_Protecao", ["epi", "equipamento protecao", "equipamento proteção", "uniforme", "seguranca", "segurança"]),
    ("Transporte_Funcionarios", ["vale transporte", "fretado", "transporte funcionarios", "transporte funcionários"]),
]

TIPO_ITEM_DESCRICAO = {
    "00": "Mercadoria para Revenda",
    "01": "Matéria-Prima",
    "02": "Embalagem",
    "03": "Produto em Processo",
    "04": "Produto Acabado",
    "05": "Subproduto",
    "06": "Produto Intermediário",
    "07": "Material de Uso e Consumo",
    "08": "Ativo Imobilizado",
    "09": "Serviços",
    "10": "Outros Insumos",
    "99": "Outras",
}


MONOFASICO_NCM = {
    "27101259", "27101921", "27101932", "27101999", "27111910", "27112910", "22030000"
}

ST_NCM = {
    "22011000", "22021000", "24022000", "30049099", "33030010", "34011190"
}

SERVICOS_SEM_CREDITO = {
    "servico administrativo", "serviço administrativo", "contabilidade", "advocaticio", "advocatício",
    "consultoria", "rh", "recursos humanos", "limpeza", "vigilancia", "vigilância", "marketing",
}

CNAE_AGRO_PREFIXOS = ("01", "02", "03")
TIPO_ITEM_INDUSTRIAL = {"01", "02", "03", "04", "06", "10"}


def _normalizar_texto(valor: str) -> str:
    return re.sub(r"\s+", " ", str(valor or "").strip().lower())


def _normalizar_ncm(ncm: str) -> str:
    return re.sub(r"\D", "", str(ncm or ""))[:8]


def _parse_valor_brasileiro(serie: pd.Series) -> pd.Series:
    return pd.to_numeric(
        serie.astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False),
        errors="coerce",
    ).fillna(0.0)


def carregar_configuracao_cnae(main_dir: Path) -> Dict[str, str]:
    """Carrega configuração tributária do CNAE principal em MAIN_DIR/CNAE.xlsx."""
    caminho_cnae = main_dir / "CNAE.xlsx"
    if not caminho_cnae.exists():
        raise ProcessamentoErro(f"Arquivo CNAE não encontrado: {caminho_cnae}")

    try:
        df_cnae = pd.read_excel(caminho_cnae, sheet_name="Sheet1")
    except Exception as exc:
        raise ProcessamentoErro(f"Erro ao ler arquivo CNAE {caminho_cnae}: {exc}") from exc

    colunas_obrigatorias = {"CNAE", "Descrição", "Regime_PIS_COFINS"}
    if not colunas_obrigatorias.issubset(df_cnae.columns):
        raise ProcessamentoErro(
            "Arquivo CNAE inválido. Colunas mínimas esperadas: 'CNAE', 'Descrição' e 'Regime_PIS_COFINS'."
        )

    for coluna in COLUNAS_CNAE_CONFIG:
        if coluna not in df_cnae.columns:
            df_cnae[coluna] = ""

    df_cnae = df_cnae.dropna(subset=["CNAE"]).copy()
    if df_cnae.empty:
        raise ProcessamentoErro("Arquivo CNAE sem registros válidos na coluna 'CNAE'.")

    registro = df_cnae.iloc[0].to_dict()
    registro["CNAE"] = re.sub(r"\D", "", str(registro.get("CNAE", "")))
    if not registro["CNAE"]:
        raise ProcessamentoErro("CNAE principal inválido no arquivo CNAE.xlsx.")

    return {col: str(registro.get(col, "") or "").strip() for col in COLUNAS_CNAE_CONFIG}


def carregar_cnae(main_dir: Path) -> Tuple[str, str]:
    """Compatibilidade retroativa: retorna CNAE e descrição do principal."""
    config = carregar_configuracao_cnae(main_dir)
    return config["CNAE"], config["Descrição"]


def carregar_matriz_cnae(main_dir: Path) -> pd.DataFrame:
    """Carrega matriz CNAE x Tipo_Item; retorna vazio quando arquivo auxiliar não existir."""
    caminho_matriz = main_dir / "Matriz_CNAE_Insumo.xlsx"
    if not caminho_matriz.exists():
        logging.warning("Matriz CNAE não encontrada: %s. Aplicando regras padrão.", caminho_matriz)
        return pd.DataFrame(columns=COLUNAS_MATRIZ_CNAE)

    try:
        matriz_df = pd.read_excel(caminho_matriz, sheet_name="Sheet1")
    except Exception as exc:
        logging.warning("Erro ao ler Matriz_CNAE_Insumo.xlsx (%s). Aplicando regras padrão.", exc)
        return pd.DataFrame(columns=COLUNAS_MATRIZ_CNAE)

    for coluna in COLUNAS_MATRIZ_CNAE:
        if coluna not in matriz_df.columns:
            matriz_df[coluna] = ""

    matriz_df = matriz_df[COLUNAS_MATRIZ_CNAE].copy()
    matriz_df["CNAE"] = matriz_df["CNAE"].astype(str).str.replace(r"\D", "", regex=True).str.strip()
    matriz_df["Tipo_Item"] = matriz_df["Tipo_Item"].astype(str).str.strip()
    matriz_df["Permite_Credito"] = matriz_df["Permite_Credito"].astype(str).str.strip()
    matriz_df["Nivel_Risco"] = matriz_df["Nivel_Risco"].astype(str).str.strip()
    matriz_df["Fundamentacao_Tecnica"] = matriz_df["Fundamentacao_Tecnica"].astype(str).str.strip()
    matriz_df["Observacao_Estrategica"] = matriz_df["Observacao_Estrategica"].astype(str).str.strip()
    matriz_df["Peso_Score"] = pd.to_numeric(matriz_df["Peso_Score"], errors="coerce").fillna(0).astype(int)
    return matriz_df


def classificar_tipo_item(descricao_item: str) -> str:
    """Classifica descrição em Tipo_Item estratégico via palavras-chave."""
    descricao = _normalizar_texto(descricao_item)
    for tipo_item, palavras in TIPO_ITEM_KEYWORDS:
        if any(palavra in descricao for palavra in palavras):
            return tipo_item
    return "Insumo_Produtivo" if descricao else "Administrativo"


def aplicar_matriz_cnae(row: pd.Series, config_cnae: Dict[str, str], matriz_df: pd.DataFrame) -> Dict[str, object]:
    """Aplica regra da matriz CNAE x Tipo_Item com fallback padrão."""
    tipo_item = classificar_tipo_item(row.get("descr_item", ""))
    cnae = re.sub(r"\D", "", str(config_cnae.get("CNAE", "")))

    if matriz_df.empty:
        matriz_item = MATRIZ_PADRAO.copy()
    else:
        filtro = (matriz_df["CNAE"] == cnae) & (matriz_df["Tipo_Item"] == tipo_item)
        matriz_match = matriz_df.loc[filtro].head(1)
        matriz_item = matriz_match.iloc[0].to_dict() if not matriz_match.empty else MATRIZ_PADRAO.copy()

    return {
        "Tipo_Item": tipo_item,
        "Permite_Credito_Matriz": str(matriz_item.get("Permite_Credito", MATRIZ_PADRAO["Permite_Credito"])) or MATRIZ_PADRAO["Permite_Credito"],
        "Peso_Score_Matriz": int(pd.to_numeric(matriz_item.get("Peso_Score", 0), errors="coerce") or 0),
        "Nivel_Risco_Matriz": str(matriz_item.get("Nivel_Risco", MATRIZ_PADRAO["Nivel_Risco"])) or MATRIZ_PADRAO["Nivel_Risco"],
        "Fundamentacao_Tecnica_Matriz": str(matriz_item.get("Fundamentacao_Tecnica", MATRIZ_PADRAO["Fundamentacao_Tecnica"])) or MATRIZ_PADRAO["Fundamentacao_Tecnica"],
        "Observacao_Estrategica": str(matriz_item.get("Observacao_Estrategica", MATRIZ_PADRAO["Observacao_Estrategica"])) or MATRIZ_PADRAO["Observacao_Estrategica"],
    }


def avaliar_credito_objetivo(ncm: str, descricao: str, cnae: str, config_cnae: Dict[str, str]) -> Tuple[str, str, str, str]:
    """Avalia regime, hipóteses legais e vedações objetivas."""
    ncm_norm = _normalizar_ncm(ncm)
    descricao_norm = _normalizar_texto(descricao)
    cnae_norm = re.sub(r"\D", "", str(cnae or ""))
    regime = _normalizar_texto(config_cnae.get("Regime_PIS_COFINS", ""))

    if regime != "nao_cumulativo":
        return "Crédito vedado", FUNDAMENTO_REGIME_CUMULATIVO, "Baixo", "Bloqueado_Regime"

    if ncm_norm in MONOFASICO_NCM:
        return "Crédito vedado", FUNDAMENTO_MONOFASICO, "Baixo", "Vedacao_Monofasico"

    if ncm_norm in ST_NCM:
        return "Crédito vedado", FUNDAMENTO_ST, "Baixo", "Vedacao_ST"

    if any(serv in descricao_norm for serv in SERVICOS_SEM_CREDITO):
        return "Crédito improvável", "Leis 10.637/2002 e 10.833/2003 - serviço sem vínculo de insumo", "Médio", "Hipotese_Fragil"

    if any(cnae_norm.startswith(prefixo) for prefixo in CNAE_AGRO_PREFIXOS):
        return "Crédito possível", FUNDAMENTO_AGRO, "Médio", "Hipotese_Agro"

    return "Necessita análise interpretativa", f"{FUNDAMENTO_TEMA_779}; {FUNDAMENTO_DECRETO_8426}", "Médio", "Analise_Interpretativa"


def calcular_score_credito(row: pd.Series) -> int:
    """Calcula score automático para triagem preliminar de crédito."""
    score = 0
    ncm = _normalizar_ncm(row.get("ncm", ""))
    descricao = _normalizar_texto(row.get("descr_item", ""))
    cnae = re.sub(r"\D", "", str(row.get("CNAE", "")))
    tipo_item = str(row.get("tipo_item", "")).strip()
    cfop = str(row.get("cfop", "")).strip()

    if ncm in MONOFASICO_NCM:
        score -= 100
    if ncm in ST_NCM:
        score -= 100
    if any(serv in descricao for serv in SERVICOS_SEM_CREDITO):
        score -= 40
    if any(cnae.startswith(prefixo) for prefixo in CNAE_AGRO_PREFIXOS):
        score += 30
    if tipo_item in TIPO_ITEM_INDUSTRIAL:
        score += 20
    if any(chave in descricao for chave in ("energia eletrica", "energia elétrica", "c500")):
        score += 40
    if cfop.startswith(("13", "23", "53", "63", "73")) or "frete" in descricao:
        score += 30

    return score


def classificar_score(score: int) -> str:
    if score >= 60:
        return "Crédito Estratégico"
    if 40 <= score <= 59:
        return "Crédito Provável"
    if 15 <= score <= 39:
        return "Crédito Possível"
    if -10 <= score <= 14:
        return "Analisar"
    return "Crédito Improvável"


def gerar_resumo_oportunidades(df_analitico: pd.DataFrame) -> pd.DataFrame:
    """Gera resumo de oportunidades por classificação final."""
    colunas_saida = ["Classificacao_Final", "Soma_Potencial_Credito", "Quantidade_Itens", "%_Sobre_Total_Potencial"]
    if df_analitico.empty:
        return pd.DataFrame(columns=colunas_saida)

    total_potencial = df_analitico["Potencial_Credito"].sum()
    resumo = (
        df_analitico.groupby("Classificacao_Final", dropna=False, as_index=False)
        .agg(
            Soma_Potencial_Credito=("Potencial_Credito", "sum"),
            Quantidade_Itens=("Classificacao_Final", "size"),
        )
    )
    resumo["%_Sobre_Total_Potencial"] = (
        resumo["Soma_Potencial_Credito"] / total_potencial * 100 if total_potencial > 0 else 0.0
    )
    return resumo.sort_values(by="Soma_Potencial_Credito", ascending=False).reset_index(drop=True)


class ProcessamentoErro(Exception):
    """Erro de processamento do confronto C170."""


def detectar_layout_sped(arquivo_txt: Path) -> Tuple[str, int]:
    """Detecta tipo de SPED e ano de layout com busca robusta do registro 0000."""
    tipo_sped = "CONTRIBUICOES" if "contrib" in str(arquivo_txt).lower() else "FISCAL"
    primeiras_linhas: List[str] = []

    def _normalizar_linha(linha: str) -> str:
        linha = linha.replace("\ufeff", "").replace("\x00", "")
        return linha.lstrip()

    def _extrair_ano_da_linha(linha: str) -> int | None:
        linha = _normalizar_linha(linha).rstrip("\n\r")
        if not linha:
            return None

        partes = [p.strip() for p in linha.split("|")]
        if "0000" in partes:
            indice_0000 = partes.index("0000")

            # Layouts variam (EFD Fiscal x Contribuições e versões), então
            # buscamos a primeira data DDMMYYYY após o registro 0000.
            for campo in partes[indice_0000 + 1 :]:
                if len(campo) == 8 and campo.isdigit():
                    return int(campo[-4:])

            # Fallback conservador para layouts esperados no histórico.
            for deslocamento in (3, 4, 5):
                dt_ini_idx = indice_0000 + deslocamento
                if dt_ini_idx < len(partes):
                    dt_ini = partes[dt_ini_idx]
                    if len(dt_ini) == 8 and dt_ini.isdigit():
                        return int(dt_ini[-4:])

        # Fallback regex: captura a primeira data DDMMYYYY após |0000|.
        match = re.search(r"\|\s*0000\s*\|(?:[^|]*\|)*?(\d{8})\|", linha)
        if match:
            dt_ini = match.group(1)
            return int(dt_ini[-4:])

        return None


    for encoding in ("utf-8-sig", "latin-1"):
        try:
            with arquivo_txt.open("r", encoding=encoding, errors="ignore") as f:
                for linha_bruta in f:
                    sublinhas = [linha_bruta]
                    if "\n" in linha_bruta:
                        sublinhas = [trecho for trecho in linha_bruta.split("\n") if trecho]

                    for linha in sublinhas:
                        linha_limpa = _normalizar_linha(linha)
                        if len(primeiras_linhas) < 10:
                            primeiras_linhas.append(repr(linha_limpa.rstrip("\n\r")))

                        if "0000" not in linha_limpa:
                            continue

                        ano_layout = _extrair_ano_da_linha(linha_limpa)
                        if ano_layout is None:
                            continue

                        return tipo_sped, ano_layout
        except OSError as exc:
            raise ProcessamentoErro(f"Erro ao ler arquivo {arquivo_txt}: {exc}") from exc

    logging.debug("Registro 0000 não localizado. Primeiras linhas do arquivo %s:", arquivo_txt)
    for idx, linha_debug in enumerate(primeiras_linhas, start=1):
        logging.debug("Linha %s: %s", idx, linha_debug)

    raise ProcessamentoErro(f"Registro 0000 não encontrado no arquivo: {arquivo_txt}")

def obter_layout_registro(tipo_sped: str, ano_layout: int, registro: str) -> Dict[str, int]:
    """Retorna mapeamento de índices para um registro conforme tipo/ano."""
    if tipo_sped not in LAYOUTS:
        raise ProcessamentoErro(f"Tipo de SPED não mapeado: {tipo_sped}")

    layouts_tipo = LAYOUTS[tipo_sped]
    if ano_layout not in layouts_tipo:
        raise ProcessamentoErro(
            f"Layout do ano {ano_layout} ainda não configurado no dicionário LAYOUTS."
        )

    if registro not in layouts_tipo[ano_layout]:
        raise ProcessamentoErro(
            f"Registro {registro} não configurado para {tipo_sped}/{ano_layout} no LAYOUTS."
        )

    return layouts_tipo[ano_layout][registro]


def carregar_arquivos(pasta_base: Path) -> List[Path]:
    """Carrega todos os arquivos .txt de uma pasta.

    Args:
        pasta_base: Caminho da pasta que contém os arquivos TXT.

    Returns:
        Lista ordenada de arquivos TXT.

    Raises:
        ProcessamentoErro: Se pasta não existir ou não houver arquivos.
    """
    if not pasta_base.exists() or not pasta_base.is_dir():
        raise ProcessamentoErro(f"Pasta não encontrada: {pasta_base}")

    arquivos = sorted(pasta_base.glob("*.txt"))
    if not arquivos:
        raise ProcessamentoErro(f"Nenhum arquivo TXT encontrado em: {pasta_base}")

    return arquivos


def _obter_campo(partes: List[str], indice: int) -> str:
    """Retorna campo por índice com segurança."""
    return partes[indice].strip() if len(partes) > indice else ""


def _parece_data(valor: str) -> bool:
    valor = (valor or "").strip()
    return len(valor) == 8 and valor.isdigit()


def _parece_valor(valor: str) -> bool:
    valor = (valor or "").strip()
    if not valor:
        return False
    return bool(re.fullmatch(r"\d{1,3}(?:\.\d{3})*(?:,\d+)?|\d+(?:,\d+)?", valor))


def _parece_cfop(valor: str) -> bool:
    valor = (valor or "").strip()
    return len(valor) == 4 and valor.isdigit()


def _ajustar_campos_d100(registro: Dict[str, str], partes: List[str], idx_data: int) -> Dict[str, str] | None:
    """Aplica heurísticas para corrigir deslocamentos comuns de layout no D100."""
    vl_item = (registro.get("vl_item") or "").strip()
    cfop = (registro.get("cfop") or "").strip()

    # Alguns layouts trocam posições de VL_DOC/CFOP.
    if (not _parece_valor(vl_item) or _parece_cfop(vl_item)) and _parece_valor(cfop):
        registro["vl_item"], registro["cfop"] = cfop, vl_item
        vl_item, cfop = registro["vl_item"], registro["cfop"]

    # Busca data válida no entorno caso o índice do layout venha deslocado.
    if not _parece_data(registro.get("data", "")):
        for desloc in range(-2, 4):
            candidato = _obter_campo(partes, idx_data + desloc)
            if _parece_data(candidato):
                registro["data"] = candidato
                break

    # CFOP é opcional em alguns layouts de Contribuições.
    if not _parece_cfop(cfop):
        registro["cfop"] = ""

    # Descarta linhas claramente inválidas para confronto documental.
    if not registro.get("num_nota", "").strip() or not _parece_data(registro.get("data", "")):
        return None

    return registro


def _extrair_chave_acesso_linha(linha: str) -> str:
    """Extrai chave de acesso (44 dígitos) de uma linha SPED, quando existir."""
    linha = (linha or "").replace("﻿", "")
    match = re.search(r"(?<!\d)(\d{44})(?!\d)", linha)
    return match.group(1) if match else ""


def _ajustar_campos_c500(registro: Dict[str, str], partes: List[str], idx_data: int, idx_valor: int) -> Dict[str, str] | None:
    """Aplica heurísticas para corrigir deslocamentos comuns de layout no C500."""
    if not _parece_data(registro.get("data", "")):
        for desloc in range(-2, 4):
            candidato = _obter_campo(partes, idx_data + desloc)
            if _parece_data(candidato):
                registro["data"] = candidato
                break

    if not _parece_valor(registro.get("vl_item", "")):
        for desloc in range(-2, 5):
            candidato = _obter_campo(partes, idx_valor + desloc)
            if _parece_valor(candidato):
                registro["vl_item"] = candidato
                break

    if not registro.get("num_nota", "").strip() or not _parece_data(registro.get("data", "")):
        return None

    return registro


def criar_chave_acesso_item(df: pd.DataFrame) -> pd.DataFrame:
    """Cria chave baseada em chave de acesso + código do item (quando houver)."""
    if df.empty:
        df = df.copy()
        df["chave_acesso_item"] = ""
        return df

    df = df.copy()
    df["chave_acesso_item"] = (
        df.get("chave_acesso", "").astype(str).str.strip()
        + "|"
        + df.get("cod_item", "").astype(str).str.strip()
    )
    return df


def criar_chave_acesso_documento(df: pd.DataFrame) -> pd.DataFrame:
    """Cria chave baseada apenas na chave de acesso (quando houver)."""
    if df.empty:
        df = df.copy()
        df["chave_acesso_doc"] = ""
        return df

    df = df.copy()
    df["chave_acesso_doc"] = df.get("chave_acesso", "").astype(str).str.strip()
    return df


def _confrontar_prioridade_chave_acesso(
    df_fiscal: pd.DataFrame,
    df_contrib: pd.DataFrame,
    coluna_chave_acesso: str,
    coluna_chave_fallback: str,
) -> pd.DataFrame:
    """Confronta priorizando chave de acesso para linhas com chave; fallback para demais."""
    if df_fiscal.empty:
        return df_fiscal.copy()

    fiscal = df_fiscal.copy()
    contrib = df_contrib.copy()

    fiscal[coluna_chave_acesso] = fiscal[coluna_chave_acesso].astype(str).str.strip()
    contrib[coluna_chave_acesso] = contrib[coluna_chave_acesso].astype(str).str.strip()

    fiscal_com = fiscal[fiscal[coluna_chave_acesso] != ""]
    fiscal_sem = fiscal[fiscal[coluna_chave_acesso] == ""]

    chaves_acesso_contrib = set(contrib.loc[contrib[coluna_chave_acesso] != "", coluna_chave_acesso].tolist())
    div_com = fiscal_com[~fiscal_com[coluna_chave_acesso].isin(chaves_acesso_contrib)]

    chaves_fallback_contrib = set(contrib[coluna_chave_fallback].dropna().astype(str).tolist())
    div_sem = fiscal_sem[~fiscal_sem[coluna_chave_fallback].isin(chaves_fallback_contrib)]

    return pd.concat([div_com, div_sem], ignore_index=True)


def extrair_0150(partes: List[str]) -> Dict[str, str]:
    """Extrai dados relevantes do registro 0150."""
    return {
        "cod_part": _obter_campo(partes, INDICE_0150_COD_PART),
        "nome_part": _obter_campo(partes, INDICE_0150_NOME),
        "cnpj": _obter_campo(partes, INDICE_0150_CNPJ),
    }


def extrair_0200(partes: List[str]) -> Dict[str, str]:
    """Extrai dados relevantes do registro 0200."""
    tipo_item = _obter_campo(partes, INDICE_0200_TIPO_ITEM)
    return {
        "cod_item": _obter_campo(partes, INDICE_0200_COD_ITEM),
        "descr_item": _obter_campo(partes, INDICE_0200_DESCR_ITEM),
        "ncm": _obter_campo(partes, INDICE_0200_NCM),
        "tipo_item": tipo_item,
        "tipo_item_desc": TIPO_ITEM_DESCRICAO.get(tipo_item, "Tipo não mapeado"),
    }


def extrair_c170(arquivos_txt: List[Path], enriquecer: bool = False) -> pd.DataFrame:
    """Extrai registros C170 vinculando dados do C100 e, quando possível, CNPJ via 0150.

    Args:
        arquivos_txt: Lista de arquivos SPED TXT.

    Returns:
        DataFrame com colunas mínimas para confronto.
    """
    registros: List[Dict[str, str]] = []

    for arquivo in arquivos_txt:
        mapa_participantes: Dict[str, Dict[str, str]] = {}
        mapa_itens: Dict[str, Dict[str, str]] = {}
        contexto_c100: Dict[str, str] = {}

        try:
            with arquivo.open("r", encoding="latin-1", errors="ignore") as f:
                for linha in f:
                    if not linha.startswith("|"):
                        continue

                    partes = linha.rstrip("\n\r").split("|")
                    registro = _obter_campo(partes, INDICE_REGISTRO)

                    if registro == "0150":
                        participante = extrair_0150(partes)
                        cod_part = participante["cod_part"]
                        if cod_part:
                            mapa_participantes[cod_part] = participante

                    elif registro == "0200" and enriquecer:
                        item = extrair_0200(partes)
                        cod_item = item["cod_item"]
                        if cod_item:
                            mapa_itens[cod_item] = item

                    elif registro == "C100":
                        cod_part = _obter_campo(partes, INDICE_C100_COD_PART)
                        participante = mapa_participantes.get(cod_part, {})
                        contexto_c100 = {
                            "cod_part": cod_part,
                            "nome_part": participante.get("nome_part", ""),
                            "cnpj": participante.get("cnpj", ""),
                            "num_nota": _obter_campo(partes, INDICE_C100_NUM_DOC),
                            "serie": _obter_campo(partes, INDICE_C100_SERIE),
                            "data": _obter_campo(partes, INDICE_C100_DT_DOC),
                            "chave_acesso": _extrair_chave_acesso_linha(linha),
                        }

                    elif registro == "C170":
                        cod_item = _obter_campo(partes, INDICE_C170_COD_ITEM)
                        item_info = mapa_itens.get(cod_item, {}) if enriquecer else {}
                        # Se não houver contexto C100, ainda registramos o C170,
                        # mas campos da nota ficam vazios.
                        registros.append(
                            {
                                "arquivo": arquivo.name,
                                "cod_part": contexto_c100.get("cod_part", ""),
                                "nome_part": contexto_c100.get("nome_part", ""),
                                "cnpj": contexto_c100.get("cnpj", ""),
                                "num_nota": contexto_c100.get("num_nota", ""),
                                "serie": contexto_c100.get("serie", ""),
                                "data": contexto_c100.get("data", ""),
                                "chave_acesso": contexto_c100.get("chave_acesso", ""),
                                "cod_item": cod_item,
                                "descr_item": item_info.get("descr_item", ""),
                                "ncm": item_info.get("ncm", ""),
                                "tipo_item": item_info.get("tipo_item", ""),
                                "tipo_item_desc": item_info.get("tipo_item_desc", ""),
                                "cfop": _obter_campo(partes, INDICE_C170_CFOP),
                                "vl_item": _obter_campo(partes, INDICE_C170_VL_ITEM),
                            }
                        )
        except OSError as exc:
            raise ProcessamentoErro(f"Erro ao ler arquivo {arquivo}: {exc}") from exc

    colunas = [
        "arquivo",
        "cod_part",
        "nome_part",
        "cnpj",
        "num_nota",
        "serie",
        "data",
        "chave_acesso",
        "cod_item",
        "descr_item",
        "ncm",
        "tipo_item",
        "tipo_item_desc",
        "cfop",
        "vl_item",
    ]
    return pd.DataFrame(registros, columns=colunas)


def criar_chave(df: pd.DataFrame) -> pd.DataFrame:
    """Cria chave de comparação: (Número da Nota + Série + CNPJ + Código do Item)."""
    if df.empty:
        df = df.copy()
        df["chave"] = ""
        return df

    df = df.copy()
    df["chave"] = (
        df["num_nota"].astype(str).str.strip()
        + "|"
        + df["serie"].astype(str).str.strip()
        + "|"
        + df["cnpj"].astype(str).str.strip()
        + "|"
        + df["cod_item"].astype(str).str.strip()
    )
    return df


def criar_chave_documento(df: pd.DataFrame) -> pd.DataFrame:
    """Cria chave de comparação: (Número Documento + Série + CNPJ + Data)."""
    if df.empty:
        df = df.copy()
        df["chave"] = ""
        return df

    df = df.copy()
    df["chave"] = (
        df["num_nota"].astype(str).str.strip()
        + "|"
        + df["serie"].astype(str).str.strip()
        + "|"
        + df["cnpj"].astype(str).str.strip()
        + "|"
        + df["data"].astype(str).str.strip()
    )
    return df


def confrontar(df_fiscal: pd.DataFrame, df_contribuicoes: pd.DataFrame) -> pd.DataFrame:
    """Retorna C170 do Fiscal ausentes no Contribuições (prioriza chave de acesso+item)."""
    if df_fiscal.empty:
        return df_fiscal.copy()

    fiscal = criar_chave_acesso_item(df_fiscal)
    contrib = criar_chave_acesso_item(df_contribuicoes)
    divergentes = _confrontar_prioridade_chave_acesso(fiscal, contrib, "chave_acesso_item", "chave")

    colunas_saida = [
        "arquivo",
        "cod_part",
        "nome_part",
        "cnpj",
        "num_nota",
        "serie",
        "data",
        "chave_acesso",
        "cod_item",
        "descr_item",
        "ncm",
        "tipo_item",
        "tipo_item_desc",
        "cfop",
        "vl_item",
        "chave",
    ]
    colunas_presentes = [col for col in colunas_saida if col in divergentes.columns]
    return divergentes[colunas_presentes]


def extrair_a170(arquivos_txt: List[Path]) -> pd.DataFrame:
    """Extrai A170 do Contribuições apenas para A100 com IND_OPER = 0."""
    registros: List[Dict[str, str]] = []

    for arquivo in arquivos_txt:
        mapa_participantes: Dict[str, Dict[str, str]] = {}
        contexto_a100: Dict[str, str] = {}

        try:
            with arquivo.open("r", encoding="latin-1", errors="ignore") as f:
                for linha in f:
                    if not linha.startswith("|"):
                        continue

                    partes = linha.rstrip("\n\r").split("|")
                    registro = _obter_campo(partes, INDICE_REGISTRO)

                    if registro == "0150":
                        participante = extrair_0150(partes)
                        cod_part = participante["cod_part"]
                        if cod_part:
                            mapa_participantes[cod_part] = participante

                    elif registro == "A100":
                        cod_part = _obter_campo(partes, INDICE_A100_COD_PART)
                        participante = mapa_participantes.get(cod_part, {})
                        contexto_a100 = {
                            "ind_oper": _obter_campo(partes, INDICE_A100_IND_OPER),
                            "cod_part": cod_part,
                            "nome_part": participante.get("nome_part", ""),
                            "cnpj": participante.get("cnpj", ""),
                            "num_nota": _obter_campo(partes, INDICE_A100_NUM_DOC),
                            "serie": _obter_campo(partes, INDICE_A100_SERIE),
                            "data": _obter_campo(partes, INDICE_A100_DT_DOC),
                            "chave_acesso": _extrair_chave_acesso_linha(linha),
                        }

                    elif registro == "A170" and contexto_a100.get("ind_oper") == "0":
                        registros.append(
                            {
                                "arquivo": arquivo.name,
                                "cod_part": contexto_a100.get("cod_part", ""),
                                "nome_part": contexto_a100.get("nome_part", ""),
                                "cnpj": contexto_a100.get("cnpj", ""),
                                "num_nota": contexto_a100.get("num_nota", ""),
                                "serie": contexto_a100.get("serie", ""),
                                "data": contexto_a100.get("data", ""),
                                "chave_acesso": contexto_a100.get("chave_acesso", ""),
                                "cod_item": _obter_campo(partes, INDICE_A170_COD_ITEM),
                            }
                        )
        except OSError as exc:
            raise ProcessamentoErro(f"Erro ao ler arquivo {arquivo}: {exc}") from exc

    colunas = ["arquivo", "cod_part", "nome_part", "cnpj", "num_nota", "serie", "data", "chave_acesso", "cod_item"]
    return pd.DataFrame(registros, columns=colunas)


def confrontar_c170_a170(df_c170_remanescente: pd.DataFrame, df_a170: pd.DataFrame) -> pd.DataFrame:
    """Retorna C170 remanescente não encontrado em A170 (IND_OPER=0), priorizando chave de acesso+item."""
    if df_c170_remanescente.empty:
        return df_c170_remanescente.copy()

    fiscal = criar_chave_acesso_item(df_c170_remanescente)
    a170 = criar_chave(criar_chave_acesso_item(df_a170))
    return _confrontar_prioridade_chave_acesso(fiscal, a170, "chave_acesso_item", "chave")


def extrair_d100(arquivos_txt: List[Path]) -> pd.DataFrame:
    """Extrai D100 (fretes) com camada dinâmica de layout por tipo/ano."""
    registros: List[Dict[str, str]] = []

    for arquivo in arquivos_txt:
        tipo_sped, ano_layout = detectar_layout_sped(arquivo)
        layout_d100 = obter_layout_registro(tipo_sped, ano_layout, "D100")

        mapa_participantes: Dict[str, Dict[str, str]] = {}
        try:
            with arquivo.open("r", encoding="latin-1", errors="ignore") as f:
                for linha in f:
                    if not linha.startswith("|"):
                        continue
                    partes = linha.rstrip("\n\r").split("|")
                    registro = _obter_campo(partes, INDICE_REGISTRO)

                    if registro == "0150":
                        participante = extrair_0150(partes)
                        cod_part = participante["cod_part"]
                        if cod_part:
                            mapa_participantes[cod_part] = participante
                    elif registro == "D100":
                        cod_part = _obter_campo(partes, layout_d100["COD_PART"])
                        participante = mapa_participantes.get(cod_part, {})
                        registro_d100 = {
                            "arquivo": arquivo.name,
                            "cod_part": cod_part,
                            "nome_part": participante.get("nome_part", ""),
                            "cnpj": participante.get("cnpj", ""),
                            "num_nota": _obter_campo(partes, layout_d100["NUM_DOC"]),
                            "serie": _obter_campo(partes, layout_d100["SERIE"]),
                            "data": _obter_campo(partes, layout_d100["DT_DOC"]),
                            "vl_item": _obter_campo(partes, layout_d100["VL_DOC"]),
                            "cfop": _obter_campo(partes, layout_d100.get("CFOP", -1)),
                            "chave_acesso": _extrair_chave_acesso_linha(linha),
                        }
                        registro_d100 = _ajustar_campos_d100(registro_d100, partes, layout_d100["DT_DOC"])
                        if registro_d100 is not None:
                            registros.append(registro_d100)
        except OSError as exc:
            raise ProcessamentoErro(f"Erro ao ler arquivo {arquivo}: {exc}") from exc

    colunas = ["arquivo", "cod_part", "nome_part", "cnpj", "num_nota", "serie", "data", "vl_item", "cfop", "chave_acesso"]
    return pd.DataFrame(registros, columns=colunas)


def confrontar_d100(df_fiscal: pd.DataFrame, df_contribuicoes: pd.DataFrame) -> pd.DataFrame:
    """Retorna D100 do Fiscal ausentes no Contribuições (prioriza chave de acesso)."""
    if df_fiscal.empty:
        return df_fiscal.copy()

    fiscal = criar_chave_documento(criar_chave_acesso_documento(df_fiscal))
    contrib = criar_chave_documento(criar_chave_acesso_documento(df_contribuicoes))
    return _confrontar_prioridade_chave_acesso(fiscal, contrib, "chave_acesso_doc", "chave")


def extrair_c500(arquivos_txt: List[Path]) -> pd.DataFrame:
    """Extrai C500 (energia/comunicação) com camada dinâmica de layout por tipo/ano."""
    registros: List[Dict[str, str]] = []

    for arquivo in arquivos_txt:
        tipo_sped, ano_layout = detectar_layout_sped(arquivo)
        layout_c500 = obter_layout_registro(tipo_sped, ano_layout, "C500")

        mapa_participantes: Dict[str, Dict[str, str]] = {}
        try:
            with arquivo.open("r", encoding="latin-1", errors="ignore") as f:
                for linha in f:
                    if not linha.startswith("|"):
                        continue
                    partes = linha.rstrip("\n\r").split("|")
                    registro = _obter_campo(partes, INDICE_REGISTRO)

                    if registro == "0150":
                        participante = extrair_0150(partes)
                        cod_part = participante["cod_part"]
                        if cod_part:
                            mapa_participantes[cod_part] = participante
                    elif registro == "C500":
                        cod_part = _obter_campo(partes, layout_c500["COD_PART"])
                        participante = mapa_participantes.get(cod_part, {})
                        registro_c500 = {
                            "arquivo": arquivo.name,
                            "cod_part": cod_part,
                            "nome_part": participante.get("nome_part", ""),
                            "cnpj": participante.get("cnpj", ""),
                            "num_nota": _obter_campo(partes, layout_c500["NUM_DOC"]),
                            "serie": _obter_campo(partes, layout_c500["SERIE"]),
                            "data": _obter_campo(partes, layout_c500["DT_DOC"]),
                            "vl_item": _obter_campo(partes, layout_c500["VL_DOC"]),
                            "cfop": "",
                            "chave_acesso": _extrair_chave_acesso_linha(linha),
                        }
                        registro_c500 = _ajustar_campos_c500(
                            registro_c500,
                            partes,
                            layout_c500["DT_DOC"],
                            layout_c500["VL_DOC"],
                        )
                        if registro_c500 is not None:
                            registros.append(registro_c500)
        except OSError as exc:
            raise ProcessamentoErro(f"Erro ao ler arquivo {arquivo}: {exc}") from exc

    colunas = ["arquivo", "cod_part", "nome_part", "cnpj", "num_nota", "serie", "data", "vl_item", "cfop", "chave_acesso"]
    return pd.DataFrame(registros, columns=colunas)


def confrontar_c500(df_fiscal: pd.DataFrame, df_contribuicoes: pd.DataFrame) -> pd.DataFrame:
    """Retorna C500 do Fiscal ausentes no Contribuições (prioriza chave de acesso)."""
    if df_fiscal.empty:
        return df_fiscal.copy()

    fiscal = criar_chave_documento(criar_chave_acesso_documento(df_fiscal))
    contrib = criar_chave_documento(criar_chave_acesso_documento(df_contribuicoes))
    return _confrontar_prioridade_chave_acesso(fiscal, contrib, "chave_acesso_doc", "chave")


def aplicar_filtro_cfop(df: pd.DataFrame, caminho_cfop: Path) -> pd.DataFrame:
    """Aplica filtro de CFOP com base no arquivo TAX Engine _ CFOP.xlsx."""
    if df.empty:
        return df.copy()

    if not caminho_cfop.exists():
        raise ProcessamentoErro(f"Arquivo de filtro CFOP não encontrado: {caminho_cfop}")

    try:
        df_cfop = pd.read_excel(caminho_cfop)
    except Exception as exc:
        raise ProcessamentoErro(f"Erro ao ler arquivo de CFOP {caminho_cfop}: {exc}") from exc

    colunas_esperadas = {"CFOP", "Seleção CFOP"}
    if not colunas_esperadas.issubset(df_cfop.columns):
        raise ProcessamentoErro(
            "Arquivo de CFOP inválido. Colunas esperadas: 'CFOP' e 'Seleção CFOP'."
        )

    df_cfop = df_cfop.copy()
    df_cfop["CFOP"] = df_cfop["CFOP"].astype(str).str.strip()
    df_cfop["Seleção CFOP"] = df_cfop["Seleção CFOP"].astype(str).str.strip().str.lower()

    cfops_considerar = set(df_cfop.loc[df_cfop["Seleção CFOP"] == "considerar", "CFOP"])

    df_filtrado = df.copy()
    df_filtrado["cfop"] = df_filtrado["cfop"].astype(str).str.strip()
    return df_filtrado[df_filtrado["cfop"].isin(cfops_considerar)].copy()


def gerar_resumo_sintetico(df_divergente: pd.DataFrame) -> pd.DataFrame:
    """Gera resumo sintético por descrição do item."""
    if df_divergente.empty:
        return pd.DataFrame(columns=["descr_item", "valor_item_total", "percentual_total"])

    resumo = df_divergente.copy()
    resumo["vl_item_num"] = _parse_valor_brasileiro(resumo["vl_item"])

    agrupado = (
        resumo.groupby("descr_item", dropna=False, as_index=False)["vl_item_num"]
        .sum()
        .rename(columns={"vl_item_num": "valor_item_total"})
    )

    total_geral = agrupado["valor_item_total"].sum()
    if total_geral > 0:
        agrupado["percentual_total"] = (agrupado["valor_item_total"] / total_geral) * 100
    else:
        agrupado["percentual_total"] = 0.0

    return agrupado.sort_values(by="valor_item_total", ascending=False).reset_index(drop=True)


def gerar_resumo_por_fornecedor(df_divergente: pd.DataFrame) -> pd.DataFrame:
    """Gera resumo por fornecedor com soma e percentual sobre o total."""
    if df_divergente.empty:
        return pd.DataFrame(columns=["nome_part", "cnpj", "valor_item_total", "percentual_total"])

    resumo = df_divergente.copy()
    resumo["vl_item_num"] = _parse_valor_brasileiro(resumo["vl_item"])

    agrupado = (
        resumo.groupby(["nome_part", "cnpj"], dropna=False, as_index=False)["vl_item_num"]
        .sum()
        .rename(columns={"vl_item_num": "valor_item_total"})
    )
    total_geral = agrupado["valor_item_total"].sum()
    agrupado["percentual_total"] = (agrupado["valor_item_total"] / total_geral * 100) if total_geral > 0 else 0.0
    return agrupado.sort_values(by="valor_item_total", ascending=False).reset_index(drop=True)


def gerar_resumo_por_tipo_item(df_analitico: pd.DataFrame) -> pd.DataFrame:
    if df_analitico.empty:
        return pd.DataFrame(columns=["Tipo_Item", "Qtd_Itens", "Potencial_Credito_Total", "Score_Medio"])

    return (
        df_analitico.groupby("Tipo_Item", dropna=False, as_index=False)
        .agg(
            Qtd_Itens=("Tipo_Item", "size"),
            Potencial_Credito_Total=("Potencial_Credito", "sum"),
            Score_Medio=("Score_Credito", "mean"),
        )
        .sort_values(by="Potencial_Credito_Total", ascending=False)
        .reset_index(drop=True)
    )


def gerar_mapa_risco(df_analitico: pd.DataFrame) -> pd.DataFrame:
    if df_analitico.empty:
        return pd.DataFrame(columns=["Nivel_Risco", "Classificacao_Final", "Qtd_Itens", "Potencial_Credito_Total"])

    return (
        df_analitico.groupby(["Nivel_Risco", "Classificacao_Final"], dropna=False, as_index=False)
        .agg(Qtd_Itens=("Classificacao_Final", "size"), Potencial_Credito_Total=("Potencial_Credito", "sum"))
        .sort_values(by=["Nivel_Risco", "Potencial_Credito_Total"], ascending=[True, False])
        .reset_index(drop=True)
    )


def gerar_saida(
    df_divergente: pd.DataFrame,
    df_resumo: pd.DataFrame,
    df_resumo_oportunidades: pd.DataFrame,
    df_resumo_tipo_item: pd.DataFrame,
    df_resumo_fornecedor_analitico: pd.DataFrame,
    df_mapa_risco: pd.DataFrame,
    df_d100_divergente: pd.DataFrame,
    df_resumo_fretes: pd.DataFrame,
    df_c500_divergente: pd.DataFrame,
    df_resumo_c500: pd.DataFrame,
    caminho_saida: Path,
) -> None:
    """Gera arquivo Excel com abas analíticas e resumos dos confrontos."""
    try:
        with pd.ExcelWriter(caminho_saida) as writer:
            df_divergente.to_excel(writer, sheet_name="ANALÍTICO", index=False)
            df_resumo.to_excel(writer, sheet_name="Resumo Sintetico", index=False)
            df_resumo_oportunidades.to_excel(writer, sheet_name="Resumo_Oportunidades", index=False)
            df_resumo_tipo_item.to_excel(writer, sheet_name="Resumo_Por_Tipo_Item", index=False)
            df_resumo_fornecedor_analitico.to_excel(writer, sheet_name="Resumo_Por_Fornecedor", index=False)
            df_mapa_risco.to_excel(writer, sheet_name="Mapa_Risco", index=False)
            df_d100_divergente.to_excel(writer, sheet_name="Analitico - Fretes", index=False)
            df_resumo_fretes.to_excel(writer, sheet_name="Resumo - Fretes", index=False)
            df_c500_divergente.to_excel(writer, sheet_name="Analitico - C500", index=False)
            df_resumo_c500.to_excel(writer, sheet_name="Resumo - C500", index=False)
    except Exception as exc:  # erro de engine/IO
        raise ProcessamentoErro(f"Erro ao gerar arquivo de saída {caminho_saida}: {exc}") from exc


def executar(main_dir: Path) -> Tuple[int, int, int, int, int, int, int, float, Path]:
    """Orquestra execução completa."""
    pasta_resultado = main_dir / "Resultado"
    pasta_contrib = pasta_resultado / "EFD Contribuições"
    pasta_fiscal = pasta_resultado / "EFD Fiscal"

    arquivos_contrib = carregar_arquivos(pasta_contrib)
    arquivos_fiscal = carregar_arquivos(pasta_fiscal)

    df_contrib = criar_chave(extrair_c170(arquivos_contrib, enriquecer=False))
    df_fiscal = criar_chave(extrair_c170(arquivos_fiscal, enriquecer=True))

    df_divergente_c170 = confrontar(df_fiscal, df_contrib)
    df_a170 = extrair_a170(arquivos_contrib)
    df_divergente = confrontar_c170_a170(df_divergente_c170, df_a170)

    caminho_cfop = pasta_resultado / "TAX Engine _ CFOP.xlsx"
    df_divergente_filtrado = aplicar_filtro_cfop(df_divergente, caminho_cfop)

    config_cnae = carregar_configuracao_cnae(main_dir)
    cnae_principal = config_cnae["CNAE"]
    matriz_cnae_df = carregar_matriz_cnae(main_dir)

    df_divergente_filtrado = df_divergente_filtrado.copy()
    df_divergente_filtrado["CNAE"] = cnae_principal
    df_divergente_filtrado["Tipo_Item"] = df_divergente_filtrado["descr_item"].map(classificar_tipo_item)

    if df_divergente_filtrado.empty:
        for coluna in [
            "Classificacao_Objetiva", "Fundamento_Objetivo", "Risco_Objetivo", "Status_Camada", "Fundamento_Legal",
            "Nivel_Risco", "Permite_Credito_Matriz", "Peso_Score_Matriz", "Score_Base", "Score_Credito",
            "Classificacao_Final", "Observacao_Estrategica",
        ]:
            df_divergente_filtrado[coluna] = pd.Series(dtype="object")
    else:
        avaliacao_objetiva = df_divergente_filtrado.apply(
            lambda row: avaliar_credito_objetivo(
                row.get("ncm", ""), row.get("descr_item", ""), cnae_principal, config_cnae
            ),
            axis=1,
            result_type="expand",
        )
        avaliacao_objetiva.columns = ["Classificacao_Objetiva", "Fundamento_Objetivo", "Risco_Objetivo", "Status_Camada"]
        df_divergente_filtrado[["Classificacao_Objetiva", "Fundamento_Objetivo", "Risco_Objetivo", "Status_Camada"]] = avaliacao_objetiva

        if matriz_cnae_df.empty:
            df_divergente_filtrado["Permite_Credito_Matriz"] = MATRIZ_PADRAO["Permite_Credito"]
            df_divergente_filtrado["Peso_Score_Matriz"] = MATRIZ_PADRAO["Peso_Score"]
            df_divergente_filtrado["Nivel_Risco_Matriz"] = MATRIZ_PADRAO["Nivel_Risco"]
            df_divergente_filtrado["Fundamentacao_Tecnica_Matriz"] = MATRIZ_PADRAO["Fundamentacao_Tecnica"]
            df_divergente_filtrado["Observacao_Estrategica"] = MATRIZ_PADRAO["Observacao_Estrategica"]
        else:
            matriz_lookup = matriz_cnae_df.rename(
                columns={
                    "Permite_Credito": "Permite_Credito_Matriz",
                    "Peso_Score": "Peso_Score_Matriz",
                    "Nivel_Risco": "Nivel_Risco_Matriz",
                    "Fundamentacao_Tecnica": "Fundamentacao_Tecnica_Matriz",
                }
            )
            matriz_lookup = matriz_lookup.drop_duplicates(subset=["CNAE", "Tipo_Item"], keep="first")
            df_divergente_filtrado = df_divergente_filtrado.merge(
                matriz_lookup[
                    [
                        "CNAE",
                        "Tipo_Item",
                        "Permite_Credito_Matriz",
                        "Peso_Score_Matriz",
                        "Nivel_Risco_Matriz",
                        "Fundamentacao_Tecnica_Matriz",
                        "Observacao_Estrategica",
                    ]
                ],
                on=["CNAE", "Tipo_Item"],
                how="left",
            )
            df_divergente_filtrado["Permite_Credito_Matriz"] = df_divergente_filtrado["Permite_Credito_Matriz"].fillna(
                MATRIZ_PADRAO["Permite_Credito"]
            )
            df_divergente_filtrado["Peso_Score_Matriz"] = pd.to_numeric(
                df_divergente_filtrado["Peso_Score_Matriz"], errors="coerce"
            ).fillna(MATRIZ_PADRAO["Peso_Score"])
            df_divergente_filtrado["Nivel_Risco_Matriz"] = df_divergente_filtrado["Nivel_Risco_Matriz"].fillna(
                MATRIZ_PADRAO["Nivel_Risco"]
            )
            df_divergente_filtrado["Fundamentacao_Tecnica_Matriz"] = df_divergente_filtrado[
                "Fundamentacao_Tecnica_Matriz"
            ].fillna(MATRIZ_PADRAO["Fundamentacao_Tecnica"])
            df_divergente_filtrado["Observacao_Estrategica"] = df_divergente_filtrado["Observacao_Estrategica"].fillna(
                MATRIZ_PADRAO["Observacao_Estrategica"]
            )

        df_divergente_filtrado["Score_Base"] = df_divergente_filtrado.apply(calcular_score_credito, axis=1)
        ajuste_fornecedor = (
            df_divergente_filtrado["cnpj"].astype(str).str.replace(r"\D", "", regex=True).str.len().eq(14).astype(int) * 5
        )
        ajuste_agro = (
            df_divergente_filtrado["Status_Camada"].eq("Hipotese_Agro").astype(int)
            * 15
            * (_normalizar_texto(config_cnae.get("Aplica_Credito_Presumido", "")).startswith("sim"))
        )
        ajuste_matriz_permissao = df_divergente_filtrado["Permite_Credito_Matriz"].str.lower().map({"sim": 10, "não": -20, "nao": -20, "depende": 0}).fillna(0)

        df_divergente_filtrado["Score_Credito"] = (
            df_divergente_filtrado["Score_Base"].astype(int)
            + df_divergente_filtrado["Peso_Score_Matriz"].astype(int)
            + ajuste_fornecedor.astype(int)
            + ajuste_agro.astype(int)
            + ajuste_matriz_permissao.astype(int)
        )
        df_divergente_filtrado["Classificacao_Final"] = df_divergente_filtrado["Score_Credito"].apply(classificar_score)
        df_divergente_filtrado["Nivel_Risco"] = df_divergente_filtrado["Nivel_Risco_Matriz"].where(
            df_divergente_filtrado["Nivel_Risco_Matriz"].astype(str).str.len() > 0,
            df_divergente_filtrado["Risco_Objetivo"],
        )
        df_divergente_filtrado["Fundamento_Legal"] = (
            df_divergente_filtrado["Fundamento_Objetivo"].astype(str)
            + " | "
            + df_divergente_filtrado["Fundamentacao_Tecnica_Matriz"].astype(str)
        ).str.strip(" |")

    df_divergente_filtrado["vl_item_num"] = _parse_valor_brasileiro(df_divergente_filtrado["vl_item"])
    df_divergente_filtrado["Potencial_Credito"] = df_divergente_filtrado["vl_item_num"] * 0.0925
    df_divergente_filtrado = df_divergente_filtrado.sort_values(
        by=["Score_Credito", "Potencial_Credito"], ascending=[False, False]
    ).reset_index(drop=True)

    colunas_analitico = [
        "arquivo", "cod_part", "nome_part", "cnpj", "num_nota", "serie", "data", "cod_item", "descr_item", "ncm",
        "tipo_item", "tipo_item_desc", "cfop", "vl_item", "CNAE", "Status_Camada", "Tipo_Item", "Score_Credito",
        "Classificacao_Final", "Nivel_Risco", "Fundamento_Legal", "Potencial_Credito", "Observacao_Estrategica",
    ]
    df_divergente_filtrado = df_divergente_filtrado[[c for c in colunas_analitico if c in df_divergente_filtrado.columns]]

    df_resumo = gerar_resumo_sintetico(df_divergente_filtrado)
    df_resumo_oportunidades = gerar_resumo_oportunidades(df_divergente_filtrado)
    df_resumo_tipo_item = gerar_resumo_por_tipo_item(df_divergente_filtrado)
    df_resumo_fornecedor_analitico = gerar_resumo_por_fornecedor(df_divergente_filtrado)
    df_mapa_risco = gerar_mapa_risco(df_divergente_filtrado)

    df_d100_fiscal = extrair_d100(arquivos_fiscal)
    df_d100_contrib = extrair_d100(arquivos_contrib)
    df_d100_divergente = confrontar_d100(df_d100_fiscal, df_d100_contrib)
    df_resumo_fretes = gerar_resumo_por_fornecedor(df_d100_divergente)

    df_c500_fiscal = extrair_c500(arquivos_fiscal)
    df_c500_contrib = extrair_c500(arquivos_contrib)
    df_c500_divergente = confrontar_c500(df_c500_fiscal, df_c500_contrib)
    df_resumo_c500 = gerar_resumo_por_fornecedor(df_c500_divergente)

    arquivo_saida = main_dir / "Notas_C170_nao_escrituradas_no_EFD_Contribuicoes.xlsx"
    gerar_saida(
        df_divergente_filtrado,
        df_resumo,
        df_resumo_oportunidades,
        df_resumo_tipo_item,
        df_resumo_fornecedor_analitico,
        df_mapa_risco,
        df_d100_divergente,
        df_resumo_fretes,
        df_c500_divergente,
        df_resumo_c500,
        arquivo_saida,
    )

    total_vedado_automatico = int((df_divergente_filtrado["Classificacao_Objetiva"] == "Crédito vedado").sum())
    total_credito_provavel = int((df_divergente_filtrado["Classificacao_Final"] == "Crédito Provável").sum())
    potencial_total = float(df_divergente_filtrado["Potencial_Credito"].sum())

    return (
        len(df_fiscal),
        len(df_divergente_c170),
        len(df_divergente_filtrado),
        len(df_d100_divergente),
        len(df_c500_divergente),
        total_vedado_automatico,
        total_credito_provavel,
        potencial_total,
        arquivo_saida,
    )


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Confronta registros C170 do EFD Fiscal vs EFD Contribuições e gera "
            "Excel com itens ausentes no Contribuições."
        )
    )
    parser.add_argument(
        "--main-dir",
        default=".",
        help="Diretório base (MAIN_DIR) que contém a pasta Resultado.",
    )
    return parser


def main() -> int:
    parser = _build_parser()
    args = parser.parse_args()

    main_dir = Path(args.main_dir).resolve()

    try:
        (
            total_fiscal,
            total_pos_c170_c170,
            total_pos_a170,
            total_d100_divergente,
            total_c500_divergente,
            total_vedado_automatico,
            total_credito_provavel,
            potencial_total,
            saida,
        ) = executar(main_dir)
    except ProcessamentoErro as exc:
        print(f"[ERRO] {exc}")
        return 1
    except Exception as exc:  # salvaguarda para erro inesperado
        print(f"[ERRO INESPERADO] {exc}")
        return 1

    print("Processamento concluído com sucesso.")
    print(f"Total C170 inicial: {total_fiscal}")
    print(f"Total após confronto C170 x C170: {total_pos_c170_c170}")
    print(f"Total após confronto com A170: {total_pos_a170}")
    print(f"Total D100 divergente: {total_d100_divergente}")
    print(f"Total C500 divergente: {total_c500_divergente}")
    print(f"Total itens analisados: {total_pos_a170}")
    print(f"Total vedados automaticamente: {total_vedado_automatico}")
    print(f"Total crédito provável: {total_credito_provavel}")
    print(f"Potencial financeiro total estimado: {potencial_total:.2f}")
    print(f"Arquivo gerado: {saida}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
