
import pandas as pd

import ctypes
import os

from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import ElementClickInterceptedException, NoSuchElementException


from time import sleep

ctypes.windll.kernel32.SetThreadExecutionState(0x80000002)
user_name = os.getenv('USERNAME')

# ##########################################################################

def main():

    tabela = pd.read_excel(r"C:\Users\{0}\Desktop\Bot_Ticket\main\base.xlsx".format(user_name))
    options = Options()
    s = Service(ChromeDriverManager().install())
    options.add_argument("--start-maximized")
    options.add_argument("--force-device-scale-factor=0.99")
    options.add_argument("--profile-directory=Default")
    options.add_argument(r"--user-data-dir=C:\Users\{0}\AppData\Local\Google\Chrome\User Data".format(user_name))
   
    for linha in tabela.index:
        print(f"Ticket - N°{linha}:\n{tabela.loc[linha]}")
        driver = webdriver.Chrome(service=s, options=options)
        driver.get("https://ifood.atlassian.net/issues/?filter=40153")
        sleep(10)
        try:
            WebDriverWait(driver, 10).until(ec.element_to_be_clickable((By.CSS_SELECTOR, 'div[data-testid="create-button-wrapper"]'))).click()
            sleep(10)
        except ElementClickInterceptedException:
            WebDriverWait(driver, 10).until(ec.element_to_be_clickable((By.CSS_SELECTOR, "'#createGlobalItemIconButton'"))).click()

            
        status = tabela.loc[linha,'Status'].upper()
        
        if status == 'DISPONIVEL' or status == 'DESCARTE':
            #Campo Resumo
            camp_resumo = driver.find_element(By.CSS_SELECTOR, "#summary-field")
            ActionChains(driver) \
            .scroll_to_element(camp_resumo) \
            .perform()
            sleep(5)

            camp_resumo.click()
            st = tabela.loc[linha, 'Service_Tag']
            camp_resumo.send_keys()
            camp_resumo.send_keys(f"Analise Tecnica - {str(st)} - {status}")
            sleep(5)
        
        elif status == 'BATERIA' or status == 'DEFEITO':
            #Campo Resumo
            camp_resumo = driver.find_element(By.CSS_SELECTOR, "#summary-field")
            ActionChains(driver) \
            .scroll_to_element(camp_resumo) \
            .perform()
            sleep(5)

            camp_resumo.click()
            st = tabela.loc[linha, 'Service_Tag']
            camp_resumo.send_keys()
            camp_resumo.send_keys(f"Analise Tecnica - {str(st)} - {status}")
            sleep(5)

        elif status == 'FABRICANTE':
            #Campo Resumo
            camp_resumo = driver.find_element(By.CSS_SELECTOR, "#summary-field")
            ActionChains(driver) \
                .scroll_to_element(camp_resumo) \
                .perform()

            sleep(5)

            camp_resumo.click()
            camp_resumo.send_keys("Fabricante - ")
            st = tabela.loc[linha, 'Service_Tag']
            camp_resumo.send_keys(str(st))

        else:
            print("Erro Status informado nao corresponde a um opçao valida!!!")

        #Campo Descriçao
        camp_Descricao = driver.find_element(By.CSS_SELECTOR, '#ak-editor-textarea')
        camp_Descricao.click()
        camp_Descricao.send_keys(f"Patrimonio: {tabela.loc[linha,'Patrimonio']}\n"
                                 f"Modelo: {tabela.loc[linha,'Modelo']}\n"
                                 f"Service Tag: {tabela.loc[linha,'Service_Tag']}\n"
                                 f"Garantia: {tabela.loc[linha,'Garantia'].strftime('%d/%m/%y')}\n")       
        sleep(1)

        #Campo ATRIBUIR ao Tecnico
        botao_atribuir = driver.find_element(By.CSS_SELECTOR, '#assignee-field')
        ActionChains(driver) \
            .scroll_to_element(botao_atribuir) \
            .perform()

        
        tecnico = tabela.loc[linha,'Tecnico'].lower()
        for texto in tecnico[:5]:
            botao_atribuir.send_keys(str(texto))
            sleep(0.1)

        sleep(2)
        botao_atribuir.send_keys(Keys.ENTER)
        
        #Atribuir ao lider
        atribuir_Lider = driver.find_element(By.CSS_SELECTOR, '#reporter-field')
        ActionChains(driver) \
            .scroll_to_element(atribuir_Lider) \
            .perform()

        atribuir_Lider.send_keys("cds.jgregorio")
        sleep(1.5)
        atribuir_Lider.send_keys(Keys.ENTER)
        sleep(1)

        #Campo service tag 2
        camp_service = driver.find_element(By.CSS_SELECTOR, '#customfield_15185-field')
        ActionChains(driver) \
            .scroll_to_element(camp_service) \
            .perform()

        camp_service.click()
        camp_service.send_keys(str(st),Keys.ENTER)
        sleep(5)

        #Campo Ver Problema - Visualiza o chamado  
        try:
            driver.find_element(By.CLASS_NAME, "css-1ixg46l").click()
            sleep(10)
        except NoSuchElementException:
            driver.find_element(By.CSS_SELECTOR,'a.css-1ixg46l')
            sleep(10)
            

        #Campo em aberto 
        try:
            driver.find_element(By.CSS_SELECTOR, r'#issue\.fields\.status-view\.status-button > span.css-178ag6o').click()
            sleep(2.5)
        except NoSuchElementException:
            driver.find_element(By.CSS_SELECTOR, r"#issue\.fields\.status-view\.status-button > span.css-5a6fwh")
            sleep(2.5)

        #campo em andamento
        try:
            driver.find_element(By.CSS_SELECTOR, '#react-select-2-option-0 > div > div > span._16jlkb7n').click()
            sleep(2.7)
        except NoSuchElementException:
            driver.find_element(By.CSS_SELECTOR, "#react-select-2-option-0 > div > div > span._1o9zidpf._1e0c1txw > span").click()
            sleep(2.7)

        #Classificaçao Harware
        driver.find_element(By.CSS_SELECTOR, "#customfield_14051").click()
        sleep(1.5)
        driver.find_element(By.CSS_SELECTOR, "#customfield_14051 > option.option-group-15949").click()

        #campo frente de atendimento
        driver.find_element(By.ID,'customfield_19857').click()
        driver.find_element(By.CSS_SELECTOR,'#customfield_19857 > option:nth-child(2)').click()

        sleep(1.5)

        #Campo responder ao cliente 
        camp_Resp_Cliente = driver.find_element(By.ID,'comment')
        camp_Resp_Cliente.click()
        chave_txt ={}
        #dicionario txt
        with open("Res_cliente.txt", "r") as cliente:
            for txt in cliente:
                txt = txt.strip()
                chave, valor = txt.split(": ",1)
                chave_txt[chave] = valor
            
            verificacao = status
            if verificacao in chave_txt:
                novo_valor = chave_txt[verificacao]
            camp_Resp_Cliente.send_keys(f"{verificacao}: {novo_valor}")
            print(f"{verificacao}: {novo_valor}")
        
        #campo classificaçao 
        try:
            driver.find_element(By.CSS_SELECTOR, '#issue-workflow-transition-submit').click()
            sleep(2.5)
        except NoSuchElementException:
            driver.find_element(By.ID,"issue-workflow-transition-submit")
            sleep(2.5)
        
        #Associaçao de prblema 
        camp_Pesquisa = driver.find_element(By.CSS_SELECTOR,'input[data-test-id="search-dialog-input"]')
        camp_Pesquisa.click()
        camp_Pesquisa.send_keys(str(st),Keys.ENTER)

        sleep(10)
        camp_associacao = driver.find_element(By.CLASS_NAME,'css-1kxozrh')
        camp_associacao.click()
        sleep(5)

        camp_Copiar = driver.find_element(By.CSS_SELECTOR,'span[data-testid="issue.common.component.permalink-button.button.link-icon"]')
        camp_Copiar.click()
        sleep(2)

        driver.back()
        sleep(5)
        driver.back()

        #botao associçao de problema 
        botao_ass = driver.find_element(By.CSS_SELECTOR, 'button[data-testid="issue.views.issue-base.foundation.quick-add.link-button.ui.link-dropdown-button"]')
        botao_ass.click()
        sleep(1.5)

        botao_ass2 = driver.find_element(By.CSS_SELECTOR,'button[data-testid="issue.issue-view.views.issue-base.foundation.quick-add.quick-add-item.link-issue"]')
        botao_ass2.click()
        sleep(1.5)

        ctrlV = driver.find_element(By.CSS_SELECTOR,'.issue-links-search__input')
        ctrlV.send_keys(Keys.CONTROL,'v',Keys.ENTER)
        sleep(1.5)

        botao_ass3 = driver.find_element(By.CSS_SELECTOR,'div[data-testid="issue.issue-view.views.issue-base.content.issue-links.add.issue-links-add-view.link-button"]')
        botao_ass3.click()
        sleep(2)

        driver.quit()


if __name__ == "__main__":
    main()
ctypes.windll.kernel32.SetThreadExecutionState(0x00000002)
