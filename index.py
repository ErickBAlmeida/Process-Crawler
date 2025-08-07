import os
import re
import subprocess
import time

from dotenv import load_dotenv
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from win10toast import ToastNotifier


class App:
    def __init__(self):
        
        load_dotenv()

        self.notifier = ToastNotifier()

        self.wb = load_workbook("SP.xlsx")
        self.sheet = self.wb.active

        subprocess.Popen([
            os.getenv("CHROME_PATH"),
            f"--remote-debugging-port={os.getenv('DEBUG_PORT')}",
            f"--user-data-dir={os.getenv('USER_DATA_DIR')}"
        ])

        #   Configurações do Chrome para se conectar via DevTools
        options = Options()
        options.debugger_address = f"127.0.0.1:{os.getenv('DEBUG_PORT')}"
        options.add_argument("--start-maximized")
        
        self.navegador = webdriver.Chrome(service=Service(), options=options)
        self.navegador.get(os.getenv("LINK"))

        self.run()

    def logar(self):
        self.navegador.find_element(By.ID, "identificacao").click()

        btnCertificado = WebDriverWait(self.navegador, 10).until(
            EC.presence_of_element_located((By.ID, "linkAbaCertificado"))
        )
        btnCertificado.click()

        time.sleep(3)
        self.navegador.find_element(By.ID, "submitCertificado").click()

    def navegar(self):
        burguerBtn = WebDriverWait(self.navegador, 10).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "header__navbar__menu-hamburger"))
        )
        burguerBtn.click()

        time.sleep(1)
        self.navegador.find_element(By.XPATH, '//*[@id="root"]/div/header/nav/aside[1]/div[1]/nav/ul/li[2]/button').click()
        time.sleep(1)
        self.navegador.find_element(By.XPATH, '//*[@id="root"]/div/header/nav/aside[1]/div[1]/nav/ul/li[2]/ul/li[1]/a').click()

    def pesquisar(self, val):

        num_processo = re.sub(r'[^0-9]','', val)

        if num_processo.isdigit() != True:
            print("\n\n❌ Processo não é valida, indo para a proxima.\n\n")
            return False
        
        self.navegador.find_element(By.ID, 'numeroDigitoAnoUnificado').send_keys(num_processo[:13])
        self.navegador.find_element(By.XPATH, "//*[@id='foroNumeroUnificado']").send_keys(num_processo[-4:])
        self.navegador.find_element(By.ID, 'botaoConsultarProcessos').click()
        
    def ponteiro(self):
        for row in self.sheet.iter_rows(min_row=2, max_col=1):
            cell_a = row[0]
            num_processo = str(cell_a.value).strip()
            self.linha = cell_a.row

            yield num_processo
        
    def main(self):
        time.sleep(5)
        self.navegador.find_element(By.ID, 'setaVoltar').click()
        
    def run(self):
        self.logar()
        self.navegar()
        
        for num_processo in self.ponteiro():
            if self.pesquisar(num_processo) != False:
                self.main()
            
            else:
                print("Seguindo para o próximo processo...")
                continue

try:
    app = App()

except Exception as e:
    print("ERROR!!")