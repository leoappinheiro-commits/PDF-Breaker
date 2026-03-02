import os
import argparse
import datetime
import logging
import shutil
import sys
import time
from dataclasses import dataclass
from pathlib import Path

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, WebDriverException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import pandas as pd

# Deixe como None para baixar automaticamente uma versão compatível.
# Só preencha se você realmente precisar fixar manualmente uma versão específica.
driver_version = None


def obter_caminho_chromedriver(versao=None):
    """Garante que o caminho retornado pelo webdriver-manager seja o executável do chromedriver."""
    caminho = ChromeDriverManager(driver_version=versao).install() if versao else ChromeDriverManager().install()
    caminho_path = Path(caminho)

    if caminho_path.is_file() and caminho_path.suffix.lower() == ".exe":
        return str(caminho_path)

    # Alguns ambientes retornam o diretório de cache ou um arquivo que não é executável.
    pasta = caminho_path if caminho_path.is_dir() else caminho_path.parent
    candidatos = list(pasta.rglob("chromedriver.exe"))
    if candidatos:
        return str(candidatos[0])

    raise RuntimeError(
        f"Não foi possível localizar o executável 'chromedriver.exe' em: {pasta}. "
        "Defina o caminho manualmente para evitar o erro WinError 193."
    )


i = 0
Chaves = pd.read_excel(r'C:\Users\leonardopinheiro\Documents\Lista de Chaves.xlsx')

# Configurar as opções do Chrome
chrome_options = Options()
chrome_options.add_argument('--disable-popup-blocking')
chrome_options.add_argument('--safebrowsing-disable-download-protection')
chrome_options.add_argument('--no-sandbox')

# Inicializar o driver do Chrome
# 1) Tenta Selenium Manager (resolve automaticamente o driver correto para o Chrome instalado)
# 2) Em caso de falha, usa webdriver-manager com validação de caminho
try:
    navegador = webdriver.Chrome(options=chrome_options)
except Exception:
    driver_path = obter_caminho_chromedriver(driver_version)
    service = Service(executable_path=driver_path)
    navegador = webdriver.Chrome(service=service, options=chrome_options)

# Acessar a página desejada
navegador.get("https://consultadanfe.com/#")

# Diretório de downloads
downloads_dir = r'C:\Users\leonardopinheiro\Downloads'


# Realizar qualquer interação adicional no navegador, se necessário
for i, chave in Chaves.iterrows():
    campo_nome = navegador.find_element(By.ID, 'chave').send_keys(chave['Chaves_Coluna'])
    time.sleep(5)
    campo_nome = navegador.find_element(By.CLASS_NAME, 'g-recaptcha').click()
    time.sleep(5)
    campo_nome = navegador.find_element(By.XPATH, '//*[@id="modalNFe"]/div/div/div[2]/div/div[3]/p[2]/a').click()
    time.sleep(5)
    campo_nome = navegador.find_element(By.XPATH, '//*[@id="modalNFe"]/div/div/div[3]/a').click()
    time.sleep(5)
    campo_nome = navegador.find_element(By.ID, 'chave').clear()
    time.sleep(5)

    # Mover o arquivo baixado para a pasta XML
    arquivo_baixado = max([os.path.join(downloads_dir, f) for f in os.listdir(downloads_dir)], key=os.path.getctime)
    data_atual = datetime.datetime.now().strftime("%d-%m-%Y")
    xml_dir = os.path.join(downloads_dir, "XML " + data_atual)
    os.makedirs(xml_dir, exist_ok=True)
    shutil.move(arquivo_baixado, xml_dir)

# Fechar o navegador
navegador.quit()
from webdriver_manager.chrome import ChromeDriverManager


@dataclass(frozen=True)
class DownloadConfig:
    input_file: Path
    output_dir: Path
    download_dir: Path
    url: str = "https://consultadanfe.com/#"
    key_column: str = "Chaves_Coluna"
    wait_seconds: int = 5


def setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    )


def parse_args() -> DownloadConfig:
    parser = argparse.ArgumentParser(
        description="Baixa XMLs do Consulta DANFE com base em uma planilha de chaves."
    )
    parser.add_argument(
        "--input",
        default=r"C:\Users\leonardopinheiro\Documents\Lista de Chaves.xlsx",
        help="Caminho da planilha de entrada.",
    )
    parser.add_argument(
        "--output",
        default=None,
        help="Diretório de saída para os XMLs processados. Padrão: pasta de download.",
    )
    parser.add_argument(
        "--download-dir",
        default=r"C:\Users\leonardopinheiro\Downloads",
        help="Diretório padrão de downloads do navegador.",
    )

    args = parser.parse_args()
    download_dir = Path(args.download_dir)
    output_dir = Path(args.output) if args.output else download_dir

    return DownloadConfig(
        input_file=Path(args.input),
        output_dir=output_dir,
        download_dir=download_dir,
    )


def initialize_selenium_driver() -> webdriver.Chrome:
    logger = logging.getLogger("selenium_init")
    try:
        chrome_options = Options()
        chrome_options.add_argument("--disable-popup-blocking")
        chrome_options.add_argument("--safebrowsing-disable-download-protection")
        chrome_options.add_argument("--no-sandbox")

        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        logger.info("driver_initialized=true")
        return driver
    except WebDriverException as exc:
        logger.exception("driver_initialized=false error=%s", exc)
        raise


def latest_downloaded_file(download_dir: Path) -> Path:
    files = [p for p in download_dir.iterdir() if p.is_file()]
    if not files:
        raise FileNotFoundError(f"Nenhum arquivo encontrado em: {download_dir}")
    return max(files, key=lambda path: path.stat().st_ctime)


def postprocess_downloaded_file(download_dir: Path, output_dir: Path) -> Path:
    logger = logging.getLogger("postprocess")
    downloaded_file = latest_downloaded_file(download_dir)
    date_suffix = datetime.datetime.now().strftime("%d-%m-%Y")
    destination_dir = output_dir / f"XML {date_suffix}"
    destination_dir.mkdir(parents=True, exist_ok=True)
    destination_file = destination_dir / downloaded_file.name

    try:
        moved_file = Path(shutil.move(str(downloaded_file), str(destination_file)))
        logger.info(
            "file_moved=true source=%s destination=%s",
            downloaded_file,
            moved_file,
        )
        return moved_file
    except OSError as exc:
        logger.exception(
            "file_moved=false source=%s destination=%s error=%s",
            downloaded_file,
            destination_file,
            exc,
        )
        raise


def process_downloads(config: DownloadConfig) -> None:
    logger = logging.getLogger("download_logic")

    if not config.input_file.exists():
        raise FileNotFoundError(f"Arquivo de entrada não encontrado: {config.input_file}")
    if not config.download_dir.exists():
        raise FileNotFoundError(f"Diretório de download não encontrado: {config.download_dir}")

    keys_df = pd.read_excel(config.input_file)
    if config.key_column not in keys_df.columns:
        raise KeyError(f"Coluna '{config.key_column}' não encontrada na planilha.")

    browser = initialize_selenium_driver()
    try:
        browser.get(config.url)
        logger.info("page_opened=true url=%s", config.url)

        for index, row in keys_df.iterrows():
            key_value = row[config.key_column]
            logger.info("processing_key=true index=%s key=%s", index, key_value)

            try:
                browser.find_element(By.ID, "chave").send_keys(key_value)
                time.sleep(config.wait_seconds)
                browser.find_element(By.CLASS_NAME, "g-recaptcha").click()
                time.sleep(config.wait_seconds)
                browser.find_element(
                    By.XPATH, '//*[@id="modalNFe"]/div/div/div[2]/div/div[3]/p[2]/a'
                ).click()
                time.sleep(config.wait_seconds)
                browser.find_element(By.XPATH, '//*[@id="modalNFe"]/div/div/div[3]/a').click()
                time.sleep(config.wait_seconds)
                browser.find_element(By.ID, "chave").clear()
                time.sleep(config.wait_seconds)

                postprocess_downloaded_file(config.download_dir, config.output_dir)
            except (NoSuchElementException, WebDriverException, OSError, FileNotFoundError) as exc:
                logger.exception(
                    "processing_key=false index=%s key=%s error=%s",
                    index,
                    key_value,
                    exc,
                )
                continue
    finally:
        browser.quit()
        logger.info("browser_closed=true")


def main() -> int:
    setup_logging()
    logger = logging.getLogger("main")

    try:
        config = parse_args()
        logger.info(
            "config_loaded=true input=%s output=%s download_dir=%s",
            config.input_file,
            config.output_dir,
            config.download_dir,
        )
        process_downloads(config)
        return 0
    except Exception as exc:  # captura final para retorno com código de erro
        logger.exception("execution_failed=true error=%s", exc)
        return 1


if __name__ == "__main__":
    sys.exit(main())
