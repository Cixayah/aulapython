from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import os
import time
import pyautogui
import pandas as pd

########################################################
excel = pd.read_excel("pBaseBot.xlsx")

# Configuração do WebDriver do Firefox
service = Service(GeckoDriverManager().install())
driver = webdriver.Firefox(service=service)


driver.maximize_window()  # Abre o navegador maximizado

# Constantes
caminhoLocal = os.getcwd()  # diretório atual
linkExterno = "https://forms.office.com/Pages/ResponsePage.aspx?id=DQSIkWdsW0yxEjajBLZtrQAAAAAAAAAAAAZ__hB_CkpUMlVOSVQzTFJLSUhPVDM1RFoyRlhMNVhZRC4u"
linkArquivo = r"C:\caminho\para\o\geckodriver"  # diretório do arquivo (geckodriver)
driver.get(linkExterno)  # executar o navegador

# LOOP FOR
for index, row in excel.iterrows():
    nProcesso = row["Nº DO PROCESSO"]
    cDaAcao = row["CLASSE DA AÇÃO"]
    competencia = row["COMPETÊNCIA"]
    dDeAtuacao = pd.to_datetime(row["DATA DE ATUAÇÃO"]).strftime("%d/%m/%Y")  # Convert date to the correct format
    sDeOrigem = row["SUBSEÇÃO DE ORIGEM"]
    situacao = row["SITUAÇÃO"]
    oJulgador = row["ÓRGÃO JULGADOR"]
    juiz = row["JUIZ"]

    driver.find_element(
        By.XPATH,
        "/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[3]/div[1]/div[1]/div[2]/div[2]/div[1]/div[2]/div[1]/span[1]/input[1]",
    ).send_keys(str(nProcesso))
    driver.find_element(
        By.XPATH,
        "/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[3]/div[1]/div[1]/div[2]/div[2]/div[2]/div[2]/div[1]/span[1]/input[1]",
    ).send_keys(cDaAcao)
    driver.find_element(
        By.XPATH,
        "/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[3]/div[1]/div[1]/div[2]/div[2]/div[3]/div[2]/div[1]/span[1]/input[1]",
    ).send_keys(competencia)
    driver.find_element(
        By.XPATH,
        "/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[3]/div[1]/div[1]/div[2]/div[2]/div[4]/div[2]/div[1]/span[1]/input[1]",
    ).send_keys(dDeAtuacao)
    driver.find_element(
        By.XPATH,
        "/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[3]/div[1]/div[1]/div[2]/div[2]/div[5]/div[2]/div[1]/span[1]/input[1]",
    ).send_keys(sDeOrigem)
    driver.find_element(
        By.XPATH,
        "/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[3]/div[1]/div[1]/div[2]/div[2]/div[6]/div[2]/div[1]/span[1]/input[1]",
    ).send_keys(situacao)
    driver.find_element(
        By.XPATH,
        "/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[3]/div[1]/div[1]/div[2]/div[2]/div[7]/div[2]/div[1]/span[1]/input[1]",
    ).send_keys(oJulgador)
    driver.find_element(
        By.XPATH,
        "/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[1]/div[3]/div[1]/div[1]/div[2]/div[2]/div[8]/div[2]/div[1]/span[1]/input[1]",
    ).send_keys(juiz)
    driver.find_element(By.XPATH,'/html/body/div/div/div[1]/div/div/div[3]/div/div/div[2]/div[3]/div/button').click()
    time.sleep(3)
    pyautogui.click(x=380, y=374)


    time.sleep(1)

#
