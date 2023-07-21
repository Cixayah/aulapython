from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.common.by import By
import os
import time
import pyautogui
import pandas as pd

excel = "pBaseBot.xlsx"
df = pd.read_excel(excel)


# Configuração do WebDriver do Firefox
service = Service(GeckoDriverManager().install())
driver = webdriver.Firefox(service=service)

driver.maximize_window()  # Abre o navegador maximizado

# Constantes
caminhoLocal = os.getcwd()  # diretório atual
linkExterno = "https://forms.office.com/Pages/ResponsePage.aspx?id=DQSIkWdsW0yxEjajBLZtrQAAAAAAAAAAAAZ__hB_CkpUMlVOSVQzTFJLSUhPVDM1RFoyRlhMNVhZRC4u"
linkArquivo = r"C:\caminho\para\o\geckodriver"  # diretório do arquivo (geckodriver)
driver.get(linkExterno)  # executar o navegador
driver.find_element(
    By.XPATH,
    "/html/body/div[1]/div/div[1]/div/div/div[3]/div/div/div[2]/div[2]/div[1]/div[2]/div/span/input",
).send_keys("")
driver.find_element(
    By.XPATH,
    "/html/body/div[1]/div/div[1]/div/div/div[3]/div/div/div[2]/div[2]/div[2]/div[2]/div/span/input",
).send_keys("")
driver.find_element(
    By.XPATH,
    "/html/body/div[1]/div/div[1]/div/div/div[3]/div/div/div[2]/div[2]/div[3]/div[2]/div/span/input",
).send_keys("")
driver.find_element(
    By.XPATH,
    "/html/body/div[1]/div/div[1]/div/div/div[3]/div/div/div[2]/div[2]/div[4]/div[2]/div/span/input",
).send_keys("")
driver.find_element(
    By.XPATH,
    "/html/body/div/div/div[1]/div/div/div[3]/div/div/div[2]/div[2]/div[5]/div[2]/div/span/input",
).send_keys("")
driver.find_element(
    By.XPATH,
    "/html/body/div/div/div[1]/div/div/div[3]/div/div/div[2]/div[2]/div[6]/div[2]/div/span/input",
).send_keys("")
driver.find_element(
    By.XPATH,
    "/html/body/div/div/div[1]/div/div/div[3]/div/div/div[2]/div[2]/div[7]/div[2]/div/span/input",
).send_keys("")
driver.find_element(
    By.XPATH,
    "/html/body/div/div/div[1]/div/div/div[3]/div/div/div[2]/div[2]/div[8]/div[2]/div/span/input",
).send_keys("")