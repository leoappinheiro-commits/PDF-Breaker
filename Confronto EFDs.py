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
    resumo["vl_item_num"] = (
        resumo["vl_item"].astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    )
    resumo["vl_item_num"] = pd.to_numeric(resumo["vl_item_num"], errors="coerce").fillna(0.0)

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
    resumo["vl_item_num"] = (
        resumo["vl_item"].astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    )
    resumo["vl_item_num"] = pd.to_numeric(resumo["vl_item_num"], errors="coerce").fillna(0.0)

    agrupado = (
        resumo.groupby(["nome_part", "cnpj"], dropna=False, as_index=False)["vl_item_num"]
        .sum()
        .rename(columns={"vl_item_num": "valor_item_total"})
    )
    total_geral = agrupado["valor_item_total"].sum()
    agrupado["percentual_total"] = (agrupado["valor_item_total"] / total_geral * 100) if total_geral > 0 else 0.0
    return agrupado.sort_values(by="valor_item_total", ascending=False).reset_index(drop=True)


def gerar_saida(
    df_divergente: pd.DataFrame,
    df_resumo: pd.DataFrame,
    df_d100_divergente: pd.DataFrame,
    df_resumo_fretes: pd.DataFrame,
    df_c500_divergente: pd.DataFrame,
    df_resumo_c500: pd.DataFrame,
    caminho_saida: Path,
) -> None:
    """Gera arquivo Excel com abas analíticas e resumos dos confrontos."""
    try:
        with pd.ExcelWriter(caminho_saida) as writer:
            df_divergente.to_excel(writer, sheet_name="Analitico", index=False)
            df_resumo.to_excel(writer, sheet_name="Resumo Sintetico", index=False)
            df_d100_divergente.to_excel(writer, sheet_name="Analitico - Fretes", index=False)
            df_resumo_fretes.to_excel(writer, sheet_name="Resumo - Fretes", index=False)
            df_c500_divergente.to_excel(writer, sheet_name="Analitico - C500", index=False)
            df_resumo_c500.to_excel(writer, sheet_name="Resumo - C500", index=False)
    except Exception as exc:  # erro de engine/IO
        raise ProcessamentoErro(f"Erro ao gerar arquivo de saída {caminho_saida}: {exc}") from exc


def executar(main_dir: Path) -> Tuple[int, int, int, int, int, Path]:
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
    df_resumo = gerar_resumo_sintetico(df_divergente_filtrado)

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
        df_d100_divergente,
        df_resumo_fretes,
        df_c500_divergente,
        df_resumo_c500,
        arquivo_saida,
    )

    return (
        len(df_fiscal),
        len(df_divergente_c170),
        len(df_divergente_filtrado),
        len(df_d100_divergente),
        len(df_c500_divergente),
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
    print(f"Arquivo gerado: {saida}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
