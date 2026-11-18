# M:\Business Intelligence\SCRIPTS PYTHON\eztools\executoreztools.py

import os
import time
import shutil
import logging
import subprocess
import pandas as pd
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
# executado em 03/11/2026


def configurar_logger(log_dir: Path) -> tuple[logging.Logger, Path]:
    """Configura log em arquivo e console."""

    log_dir.mkdir(parents=True, exist_ok=True)

    caminho_log = log_dir / f"{datetime.now():%Y-%m-%d_%H-%M-%S}.log"

    logger = logging.getLogger()
    logger.setLevel(logging.INFO)

    if logger.hasHandlers():
        logger.handlers.clear()

    formatter = logging.Formatter(
        "[%(asctime)s] [%(levelname)s] [%(filename)s:%(lineno)d] %(message)s"
    )

    file_handler = logging.FileHandler(caminho_log, encoding="utf-8")
    file_handler.setFormatter(formatter)

    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(stream_handler)

    return logger, caminho_log


def criar_driver(download_dir: Path) -> webdriver.Chrome:
    """Abre Chrome com download automático na pasta configurada."""

    options = Options()

    prefs = {
        "download.default_directory": str(download_dir),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }

    options.add_experimental_option("prefs", prefs)
    options.add_argument("--start-maximized")

    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=options)


def fazer_login(driver, wait, url, usuario, senha):
    """Acessa Eztools e faz login."""

    logging.info("Acessando Eztools.")
    driver.get(url)

    wait.until(EC.element_to_be_clickable((By.ID, "username"))).send_keys(usuario)
    wait.until(EC.element_to_be_clickable((By.ID, "password"))).send_keys(senha)

    wait.until(
        EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Entrar')]"))
    ).click()

    logging.info("Login enviado.")
    time.sleep(6)


def limpar_filtro_status(driver, wait=None):
    """
    Limpa especificamente o filtro Status quando ele estiver com valor selecionado.
    Ex.: '1 item selecionado(s)' / valor hidden = OPEN.
    """

    logging.info("Tentando limpar especificamente o filtro Status.")

    try:
        # Localiza o componente que contém o valor OPEN / Itens em aberto
        wrappers = driver.find_elements(
            By.XPATH,
            "//div[contains(@class,'vscomp-wrapper')][.//input[contains(@class,'vscomp-hidden-input') and @value='OPEN']]"
        )

        if not wrappers:
            logging.warning("Nenhum filtro Status com valor OPEN foi localizado.")
            return False

        wrapper = wrappers[0]

        clear_button = wrapper.find_element(By.CLASS_NAME, "vscomp-clear-button")
        driver.execute_script("arguments[0].click();", clear_button)

        time.sleep(1.5)

        logging.info("Filtro Status limpo com sucesso.")
        return True

    except Exception as erro:
        logging.warning(f"Não foi possível limpar especificamente o filtro Status: {erro}")
        return False


def limpar_downloads_eztools(download_dir: Path):
    """Remove arquivos residuais da Eztools antes de iniciar nova coleta."""
    padroes = ["Relatorio.xlsx", "Relatorio*.xlsx", "*.crdownload"]

    for padrao in padroes:
        for arquivo in download_dir.glob(padrao):
            try:
                arquivo.unlink()
                logging.info(f"Arquivo residual removido da pasta Downloads: {arquivo}")
            except Exception as erro:
                logging.warning(f"Não foi possível remover arquivo residual {arquivo}: {erro}")


def aguardar_relatorio_finalizado(driver, timeout=360):
    """
    Aguarda a mensagem 'Relatório Finalizado!'.
    Aguardar isso é importante para não fechar o navegador antes da Eztools gerar o arquivo.
    """

    logging.info("Aguardando finalização do relatório na Eztools.")

    limite = time.time() + timeout

    while time.time() < limite:
        if driver.find_elements(By.XPATH, "//*[contains(text(),'Relatório Finalizado')]"):
            logging.info("Mensagem 'Relatório Finalizado!' localizada.")
            return

        time.sleep(5)

    logging.warning("Mensagem de relatório finalizado não apareceu dentro do tempo limite.")


def validar_excel_com_dados(caminho_arquivo: Path, nome_fluxo: str) -> bool:
    """Valida se o Excel baixado possui dados reais."""
    try:
        df_teste = pd.read_excel(caminho_arquivo, nrows=20)

        if df_teste.empty:
            logging.warning(f"{nome_fluxo}: arquivo baixado está vazio: {caminho_arquivo}")
            return False

        colunas_validas = [col for col in df_teste.columns if not str(col).startswith("Unnamed")]
        if not colunas_validas:
            logging.warning(f"{nome_fluxo}: arquivo baixado não possui colunas válidas: {caminho_arquivo}")
            return False

        logging.info(f"{nome_fluxo}: validação OK. Linhas amostradas: {len(df_teste)} | Colunas: {len(df_teste.columns)}")
        return True

    except Exception as erro:
        logging.warning(f"{nome_fluxo}: erro ao validar arquivo Excel baixado: {erro}")
        return False


def aguardar_download_valido(download_dir: Path, inicio: float, nome_fluxo: str, timeout=360) -> Path:
    """Aguarda um .xlsx novo e valida se possui dados."""
    logging.info("Aguardando arquivo baixado válido na pasta Downloads.")

    limite = time.time() + timeout

    while time.time() < limite:
        arquivos_temp = list(download_dir.glob("*.crdownload"))

        arquivos = [
            arquivo
            for arquivo in download_dir.glob("*.xlsx")
            if arquivo.stat().st_mtime >= inicio
        ]

        if arquivos and not arquivos_temp:
            arquivo = max(arquivos, key=lambda x: x.stat().st_mtime)

            # espera estabilizar o tamanho do arquivo
            tamanho_1 = arquivo.stat().st_size
            time.sleep(3)
            tamanho_2 = arquivo.stat().st_size

            if tamanho_1 != tamanho_2:
                logging.info(f"Arquivo ainda está estabilizando: {arquivo}")
                time.sleep(5)
                continue

            logging.info(f"Arquivo baixado localizado: {arquivo} | Tamanho: {tamanho_2} bytes")

            if validar_excel_com_dados(arquivo, nome_fluxo):
                return arquivo

            raise Exception(f"{nome_fluxo}: o arquivo baixado está vazio ou sem dados válidos: {arquivo}")

        time.sleep(5)

    raise TimeoutError(f"{nome_fluxo}: download não foi concluído dentro do tempo limite.")


def aguardar_download(download_dir: Path, inicio: float, timeout=360) -> Path:
    """Aguarda arquivo .xlsx novo aparecer na pasta de downloads."""

    logging.info("Aguardando arquivo baixado na pasta Downloads.")

    limite = time.time() + timeout

    while time.time() < limite:
        arquivos = [
            arquivo
            for arquivo in download_dir.glob("*.xlsx")
            if arquivo.stat().st_mtime >= inicio
            and not arquivo.name.endswith(".crdownload")
        ]

        arquivos_temp = list(download_dir.glob("*.crdownload"))

        if arquivos and not arquivos_temp:
            arquivo = max(arquivos, key=lambda x: x.stat().st_mtime)
            logging.info(f"Arquivo baixado localizado: {arquivo}")
            return arquivo

        time.sleep(5)

    raise TimeoutError("Download não foi concluído dentro do tempo limite.")


def substituir_arquivo(arquivo_origem: Path, arquivo_destino: Path):
    """Substitui o arquivo de destino pelo arquivo baixado."""

    arquivo_destino.parent.mkdir(parents=True, exist_ok=True)

    if arquivo_destino.exists():
        logging.info(f"Removendo arquivo existente: {arquivo_destino}")
        arquivo_destino.unlink()

    shutil.move(str(arquivo_origem), str(arquivo_destino))
    logging.info(f"Arquivo movido para: {arquivo_destino}")


def mover_tratado_para_bkp(arquivo_tratado: Path, pasta_bkp: Path):
    """Move o arquivo tratado anterior para backup com data/hora no nome."""

    pasta_bkp.mkdir(parents=True, exist_ok=True)

    if not arquivo_tratado.exists():
        logging.warning(f"Arquivo tratado anterior não encontrado: {arquivo_tratado}")
        return

    data_ref = datetime.now().strftime("%Y%m%d_%H%M%S")
    novo_nome = f"{arquivo_tratado.stem}_{data_ref}{arquivo_tratado.suffix}"
    destino = pasta_bkp / novo_nome

    shutil.move(str(arquivo_tratado), str(destino))
    logging.info(f"Arquivo tratado anterior movido para backup: {destino}")


def executar_script_python(caminho_script: Path):
    """Executa o script de tratamento e aguarda finalizar."""

    logging.info(f"Executando script de tratamento: {caminho_script}")

    processo = subprocess.run(
        ["python", str(caminho_script)],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        cwd=str(caminho_script.parent)
    )

    if processo.stdout:
        logging.info(processo.stdout)

    if processo.stderr:
        logging.warning(processo.stderr)

    if processo.returncode != 0:
        raise Exception(f"Erro ao executar script: {caminho_script}")

    logging.info(f"Script executado com sucesso: {caminho_script}")


def coletar_relatorio_eztools(
    nome_fluxo: str,
    menu_principal: str,
    submenu: str,
    id_refresh: str,
    id_excel: str,
    arquivo_destino: Path,
    arquivo_tratado: Path,
    pasta_bkp: Path,
    script_tratamento: Path,
    config: dict
):
    """Executa uma coleta completa da Eztools."""

    logging.info(f"Iniciando coleta: {nome_fluxo}")
    
    limpar_downloads_eztools(config["download_dir"])

    driver = None
    inicio_download = time.time()

    try:
        driver = criar_driver(config["download_dir"])
        wait = WebDriverWait(driver, 60)

        fazer_login(
            driver,
            wait,
            config["eztools_url"],
            config["eztools_user"],
            config["eztools_password"]
        )

        logging.info(f"Clicando no menu principal: {menu_principal}")
        wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, f"//span[normalize-space()='{menu_principal}']")
            )
        ).click()

        time.sleep(2)

        logging.info(f"Clicando no submenu: {submenu}")
        wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, f"//a[normalize-space()='{submenu}']")
            )
        ).click()

        time.sleep(6)

        filtro_limpo = limpar_filtro_status(driver, wait)

        logging.info("Atualizando relatório após limpeza do filtro Status.")
        botao_refresh = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, f"//button[@data-id='{id_refresh}' and @data-acttype='REFRESH']")
            )
        )
        driver.execute_script("arguments[0].click();", botao_refresh)

        time.sleep(10)

        logging.info("Solicitando exportação Excel.")
        inicio_download = time.time()

        wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, f"//button[@data-id='{id_excel}' and @data-key='EXCELSOLICS1']")
            )
        ).click()

        aguardar_relatorio_finalizado(driver, timeout=360)

        arquivo_baixado = aguardar_download_valido(
            config["download_dir"],
            inicio_download,
            nome_fluxo,
            timeout=360
        )

        substituir_arquivo(arquivo_baixado, arquivo_destino)

    finally:
        if driver:
            time.sleep(5)
            driver.quit()
            logging.info("Navegador encerrado.")

    mover_tratado_para_bkp(arquivo_tratado, pasta_bkp)
    executar_script_python(script_tratamento)

    logging.info(f"Coleta finalizada com sucesso: {nome_fluxo}")


def main():
    caminho_env = Path(r"M:\Business Intelligence\SCRIPTS PYTHON\.env")
    load_dotenv(caminho_env)

    log_dir = Path(r"M:\Business Intelligence\SCRIPTS PYTHON\eztools\logs")
    logger, caminho_log = configurar_logger(log_dir)

    logger.info("Início da execução do robô Eztools.")
    logger.info(f"Arquivo de log: {caminho_log}")

    config = {
        "eztools_url": os.getenv("EZTOOLS_URL"),
        "eztools_user": os.getenv("EZTOOLS_USER"),
        "eztools_password": os.getenv("EZTOOLS_PASSWORD"),
        "download_dir": Path(os.getenv("DOWNLOAD_DIR")),
    }

    if not all(config.values()):
        raise ValueError("Parâmetros obrigatórios da Eztools não encontrados no .env.")
    
    coletar_relatorio_eztools(
        nome_fluxo="Eztools DP",
        menu_principal="Depto Pessoal",
        submenu="Atendimento",
        id_refresh="78",
        id_excel="167",
        arquivo_destino=Path(r"M:\Business Intelligence\FONTES DE DADOS\eztools DP\eztools_dp_2025.xlsx"),
        arquivo_tratado=Path(r"M:\Business Intelligence\FONTES DE DADOS\eztools DP\Arquivo tratado\eztools_dp.xlsx"),
        pasta_bkp=Path(r"M:\Business Intelligence\FONTES DE DADOS\eztools DP\bkp"),
        script_tratamento=Path(r"M:\Business Intelligence\SCRIPTS PYTHON\eztools\eztools_dp\eztools_dp_corrigido.py"),
        config=config
    )
    
    coletar_relatorio_eztools(
        nome_fluxo="Eztools Financeiro",
        menu_principal="Financeiro",
        submenu="Rateio Contas",
        id_refresh="9",
        id_excel="166",
        arquivo_destino=Path(r"M:\Business Intelligence\FONTES DE DADOS\eztools financeiro\eztools_financeiro.xlsx"),
        arquivo_tratado=Path(r"M:\Business Intelligence\FONTES DE DADOS\eztools financeiro\Arquivo tratado\eztools_fc.xlsx"),
        pasta_bkp=Path(r"M:\Business Intelligence\FONTES DE DADOS\eztools financeiro\antigos"),
        script_tratamento=Path(r"M:\Business Intelligence\SCRIPTS PYTHON\eztools\eztools_financeiro\eztools_financeiro_corrigido.py"),
        config=config
    )

    logger.info("Execução do robô Eztools concluída com sucesso.")


if __name__ == "__main__":
    main()