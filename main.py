import os
import time
import traceback
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

DOWNLOAD_DIR = os.path.join(os.getcwd(), "CAPES_RIS")
BASE_URL = (
    "https://www.periodicos.capes.gov.br/index.php/acervo/buscador.html?"
    "q=all%3Acontains%28%26quot%3Be-waste%26quot%3B%29+OR+all%3Acontains%28"
    "%26quot%3Belectronic+waste%26quot%3B%29&source=all&open_access%5B%5D="
    "open_access%3D%3D1&type%5B%5D=type%3D%3DArtigo&type%5B%5D=type%3D%3DCap"
    "%C3%ADtulo+de+livro&type%5B%5D=type%3D%3DLivro&publishyear_min%5B%5D="
    "1994&publishyear_max%5B%5D=2025&mode=advanced&source=all"
)

def setup_driver():
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options = webdriver.ChromeOptions()
    options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(options=options)
    # habilita download via CDP
    driver.command_executor._commands["send_command"] = (
        "POST", '/session/$sessionId/chromium/send_command'
    )
    driver.execute("send_command", {
        'cmd': 'Page.setDownloadBehavior',
        'params': {'behavior': 'allow', 'downloadPath': DOWNLOAD_DIR}
    })
    return driver

def wait_for_new_ris(initial_files, timeout=60):
    """
    Aguarda que apareça um novo arquivo .ris no diretório de download.
    Retorna o nome do arquivo, ou None se não aparecer em timeout.
    """
    end = time.time() + timeout
    while time.time() < end:
        current = set(os.listdir(DOWNLOAD_DIR))
        added = current - initial_files
        for f in added:
            if f.lower().endswith('.ris'):
                return f
        time.sleep(1)
    return None

def accept_cookies(driver):
    try:
        btn = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Aceitar')]"))
        )
        btn.click()
        time.sleep(1)
    except Exception:
        pass  # nada a fazer se não aparecer o popup

def select_all_items(driver):
    try:
        chk = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input#checkbox-all"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", chk)
        driver.execute_script("document.querySelectorAll('div.blockUI').forEach(e=>e.remove());")
        driver.execute_script("arguments[0].click();", chk)
        return True
    except Exception as e:
        print("Erro ao marcar checkbox:", e)
        return False

def export_ris(driver, page_number):
    try:
        # fecha abas extras
        for handle in driver.window_handles[1:]:
            driver.switch_to.window(handle)
            driver.close()
        driver.switch_to.window(driver.window_handles[0])

        # abre menu Exportar
        export_btn = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH,
                "//a[@role='button' and contains(., 'Exportar')]"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", export_btn)
        driver.execute_script("arguments[0].click();", export_btn)

        # aguarda dropdown
        WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "div.dropdown-menu.show"))
        )

        # snapshot dos arquivos atuais
        initial = set(os.listdir(DOWNLOAD_DIR))

        # clica em RIS
        ris_opt = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "exportRIS"))
        )
        driver.execute_script("arguments[0].click();", ris_opt)

        # aguarda novo .ris
        new_file = wait_for_new_ris(initial, timeout=60)
        if not new_file:
            print("Nenhum novo .ris foi baixado dentro do timeout.")
            return False

        # renomeia para incluir página e timestamp
        src = os.path.join(DOWNLOAD_DIR, new_file)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        dst = os.path.join(DOWNLOAD_DIR, f"pagina_{page_number}_{ts}.ris")

        # garante que o arquivo não esteja travado
        for _ in range(5):
            try:
                os.rename(src, dst)
                print(f"Salvo: {dst}")
                return True
            except PermissionError:
                time.sleep(1)
        print(f"Falha ao renomear {src}")
        return False

    except Exception:
        print("Erro na exportação:", traceback.format_exc())
        return False

def main():
    driver = setup_driver()
    driver.get(BASE_URL)
    driver.maximize_window()
    accept_cookies(driver)

    page = 1
    while True:
        print(f"\n=== Processando página {page} ===")
        if not select_all_items(driver):
            print("Abortando: não conseguiu selecionar itens.")
            break
        if not export_ris(driver, page):
            print("Abortando: falha na exportação.")
            break

        # navega para próxima
        try:
            next_btn = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH,
                    "//button[@aria-label='Página seguinte']"))
            )
            if "disabled" in next_btn.get_attribute("class"):
                print("Última página alcançada.")
                break
            driver.execute_script("arguments[0].click();", next_btn)
            page += 1
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input#checkbox-all"))
            )
            time.sleep(2)
        except Exception as e:
            print("Erro de navegação:", e)
            break

    driver.quit()
    print("Processo finalizado!")

if __name__ == "__main__":
    main()