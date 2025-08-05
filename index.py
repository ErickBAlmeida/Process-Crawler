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

        # self.wb = load_workbook("RJ.xlsx")
        # self.sheet = self.wb.active

        #   Configurar Navegador

        subprocess.Popen([
            os.getenv("CHROME_PATH"),
            f"--remote-debugging-port={os.getenv("DEBUG_PORT")}",
            f"--user-data-dir={os.getenv("USER_DATA_DIR")}"
        ])

        time.sleep(3)

        #   Configurações do Chrome para se conectar via DevTools
        options = Options()
        options.debugger_address = f"127.0.0.1:{os.getenv("DEBUG_PORT")}"
        options.add_argument("--start-maximized")
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        
        self.navegador = webdriver.Chrome(service=Service(), options=options)
        self.navegador.get(os.getenv("LINK"))

        self.run()
    
    def run(self):
        pass

try:
    app = App()

except Exception as e:
    print("ERROR!!")