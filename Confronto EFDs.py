"""Confronto de registros C170 entre EFD Fiscal e EFD Contribuições.

O script lê arquivos TXT da estrutura:
MAIN_DIR/Resultado/EFD Fiscal
MAIN_DIR/Resultado/EFD Contribuições

E gera um Excel com os itens (C170) presentes no Fiscal e ausentes no
Contribuições.
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd


# Índices dos campos no layout SPED (split por "|").
# Exemplo de linha: |C100|...|
# Após split("|"), o registro fica na posição 1.
INDICE_REGISTRO = 1

# 0150 - cadastro participante
INDICE_0150_COD_PART = 2
INDICE_0150_CNPJ = 5

# C100 - cabeçalho da nota
INDICE_C100_COD_PART = 4
INDICE_C100_SERIE = 8
INDICE_C100_NUM_DOC = 9
INDICE_C100_DT_DOC = 10

# C170 - itens da nota
INDICE_C170_COD_ITEM = 3
INDICE_C170_VL_ITEM = 7
INDICE_C170_CFOP = 11


class ProcessamentoErro(Exception):
    """Erro de processamento do confronto C170."""


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


def extrair_c170(arquivos_txt: List[Path]) -> pd.DataFrame:
    """Extrai registros C170 vinculando dados do C100 e, quando possível, CNPJ via 0150.

    Args:
        arquivos_txt: Lista de arquivos SPED TXT.

    Returns:
        DataFrame com colunas mínimas para confronto.
    """
    registros: List[Dict[str, str]] = []

    for arquivo in arquivos_txt:
        mapa_participantes: Dict[str, str] = {}
        contexto_c100: Dict[str, str] = {}

        try:
            with arquivo.open("r", encoding="latin-1", errors="ignore") as f:
                for linha in f:
                    if not linha.startswith("|"):
                        continue

                    partes = linha.rstrip("\n\r").split("|")
                    registro = _obter_campo(partes, INDICE_REGISTRO)

                    if registro == "0150":
                        cod_part = _obter_campo(partes, INDICE_0150_COD_PART)
                        cnpj = _obter_campo(partes, INDICE_0150_CNPJ)
                        if cod_part:
                            mapa_participantes[cod_part] = cnpj

                    elif registro == "C100":
                        cod_part = _obter_campo(partes, INDICE_C100_COD_PART)
                        contexto_c100 = {
                            "cnpj": mapa_participantes.get(cod_part, ""),
                            "num_nota": _obter_campo(partes, INDICE_C100_NUM_DOC),
                            "serie": _obter_campo(partes, INDICE_C100_SERIE),
                            "data": _obter_campo(partes, INDICE_C100_DT_DOC),
                        }

                    elif registro == "C170":
                        # Se não houver contexto C100, ainda registramos o C170,
                        # mas campos da nota ficam vazios.
                        registros.append(
                            {
                                "arquivo": arquivo.name,
                                "cnpj": contexto_c100.get("cnpj", ""),
                                "num_nota": contexto_c100.get("num_nota", ""),
                                "serie": contexto_c100.get("serie", ""),
                                "data": contexto_c100.get("data", ""),
                                "cod_item": _obter_campo(partes, INDICE_C170_COD_ITEM),
                                "cfop": _obter_campo(partes, INDICE_C170_CFOP),
                                "vl_item": _obter_campo(partes, INDICE_C170_VL_ITEM),
                            }
                        )
        except OSError as exc:
            raise ProcessamentoErro(f"Erro ao ler arquivo {arquivo}: {exc}") from exc

    colunas = ["arquivo", "cnpj", "num_nota", "serie", "data", "cod_item", "cfop", "vl_item"]
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


def confrontar(df_fiscal: pd.DataFrame, df_contribuicoes: pd.DataFrame) -> pd.DataFrame:
    """Retorna registros C170 presentes no Fiscal e ausentes no Contribuições."""
    if df_fiscal.empty:
        return df_fiscal.copy()

    chaves_contrib = set(df_contribuicoes["chave"].dropna().astype(str).tolist())
    divergentes = df_fiscal[~df_fiscal["chave"].isin(chaves_contrib)].copy()

    colunas_saida = [
        "arquivo",
        "cnpj",
        "num_nota",
        "serie",
        "data",
        "cod_item",
        "cfop",
        "vl_item",
        "chave",
    ]
    return divergentes[colunas_saida]


def gerar_saida(df_divergente: pd.DataFrame, caminho_saida: Path) -> None:
    """Gera arquivo Excel de saída."""
    try:
        df_divergente.to_excel(caminho_saida, index=False)
    except Exception as exc:  # erro de engine/IO
        raise ProcessamentoErro(f"Erro ao gerar arquivo de saída {caminho_saida}: {exc}") from exc


def executar(main_dir: Path) -> Tuple[int, int, int, Path]:
    """Orquestra execução completa."""
    pasta_resultado = main_dir / "Resultado"
    pasta_contrib = pasta_resultado / "EFD Contribuições"
    pasta_fiscal = pasta_resultado / "EFD Fiscal"

    arquivos_contrib = carregar_arquivos(pasta_contrib)
    arquivos_fiscal = carregar_arquivos(pasta_fiscal)

    df_contrib = criar_chave(extrair_c170(arquivos_contrib))
    df_fiscal = criar_chave(extrair_c170(arquivos_fiscal))

    df_divergente = confrontar(df_fiscal, df_contrib)

    arquivo_saida = main_dir / "Notas_C170_nao_escrituradas_no_EFD_Contribuicoes.xlsx"
    gerar_saida(df_divergente, arquivo_saida)

    return len(df_fiscal), len(df_contrib), len(df_divergente), arquivo_saida


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
        total_fiscal, total_contrib, total_divergente, saida = executar(main_dir)
    except ProcessamentoErro as exc:
        print(f"[ERRO] {exc}")
        return 1
    except Exception as exc:  # salvaguarda para erro inesperado
        print(f"[ERRO INESPERADO] {exc}")
        return 1

    print("Processamento concluído com sucesso.")
    print(f"Total de C170 no EFD Fiscal: {total_fiscal}")
    print(f"Total de C170 no EFD Contribuições: {total_contrib}")
    print(f"Total divergente (no Fiscal e não no Contribuições): {total_divergente}")
    print(f"Arquivo gerado: {saida}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())