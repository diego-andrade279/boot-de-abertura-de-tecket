#BOOT- ADEMIR 
from datetime import datetime
import pandas as pd
import ctypes
import os
import pyautogui
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
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    NoSuchElementException,
)
from time import sleep
import time

ctypes.windll.kernel32.SetThreadExecutionState(0x80000002)
user_name = os.getenv("USERNAME")
numero_Ticket = ""
# ##########################################################################


def main():
    temp = time.time()
    global numero_Ticket
    tabela = pd.read_excel(
        r"C:\Users\{0}\Desktop\Bot_Ticket\main\base.xlsx".format(user_name)
    )
    criar_pasta_pt = r"C:\Users\{0}\Desktop\Bot_Ticket\main\Imagens".format(user_name)
    for p in tabela.index:
        nome_Pasta = str(tabela.loc[p, "Patrimonio"])
        caminho_completo = os.path.join(criar_pasta_pt, nome_Pasta)
        os.makedirs(caminho_completo, exist_ok=True)
    print(tabela)
    pyautogui.alert("Inserir as imagens antes de continuar")

    options = Options()
    s = Service(ChromeDriverManager().install())
    options.add_argument("--start-maximized")
    options.add_argument("--force-device-scale-factor=0.8")
    options.add_argument("--profile-directory=Default")
    options.add_argument(
        r"--user-data-dir=C:\Users\{0}\AppData\Local\Google\Chrome\User Data".format(
            user_name
        )
    )

    for linha in tabela.index:
        print(
            f"Ticket - N°{linha}:\n{tabela.loc[linha,['Patrimonio','Modelo','Service_Tag']]}\n"
        )
        driver = webdriver.Chrome(service=s, options=options)
        driver.get("https://ifood.atlassian.net/issues/?filter=40153")
        sleep(10)
        try:
            WebDriverWait(driver, 10).until(
                ec.element_to_be_clickable(
                    (By.CSS_SELECTOR, 'div[data-testid="create-button-wrapper"]')
                )
            ).click()
            sleep(10)
        except ElementClickInterceptedException:
            WebDriverWait(driver, 10).until(
                ec.element_to_be_clickable(
                    (By.CSS_SELECTOR, "'#createGlobalItemIconButton'")
                )
            ).click()

        status = tabela.loc[linha, "Status"].upper()

        if status == "DISPONIVEL" or status == "DESCARTE":
            # Campo Resumo
            camp_resumo = driver.find_element(By.CSS_SELECTOR, "#summary-field")
            ActionChains(driver).scroll_to_element(camp_resumo).perform()
            sleep(2)

            camp_resumo.click()
            st = tabela.loc[linha, "Service_Tag"]
            camp_resumo.send_keys()
            camp_resumo.send_keys(f"Analise Tecnica - {str(st)} - {status}")
            sleep(2)

        elif status == "BATERIA" or status == "DEFEITO":
            # Campo Resumo
            camp_resumo = driver.find_element(By.CSS_SELECTOR, "#summary-field")
            ActionChains(driver).scroll_to_element(camp_resumo).perform()
            sleep(2)

            camp_resumo.click()
            st = tabela.loc[linha, "Service_Tag"]
            camp_resumo.send_keys()
            camp_resumo.send_keys(f"Analise Tecnica - {str(st)} - {status}")
            sleep(2)

        elif status == "FABRICANTE":
            # Campo Resumo
            camp_resumo = driver.find_element(By.CSS_SELECTOR, "#summary-field")
            ActionChains(driver).scroll_to_element(camp_resumo).perform()

            sleep(2)

            camp_resumo.click()
            camp_resumo.send_keys("Fabricante - ")
            st = tabela.loc[linha, "Service_Tag"]
            camp_resumo.send_keys(str(st))

        else:
            print("Erro Status informado nao corresponde a um opçao valida!!!")

        # Campo Descriçao
        camp_Descricao = driver.find_element(By.CSS_SELECTOR, "#ak-editor-textarea")
        camp_Descricao.click()
        camp_Descricao.send_keys(
            f"Patrimonio: {tabela.loc[linha,'Patrimonio']}\n"
            f"Modelo: {tabela.loc[linha,'Modelo']}\n"
            f"Service Tag: {tabela.loc[linha,'Service_Tag']}\n"
            f"Garantia: {tabela.loc[linha,'Garantia'].strftime('%d/%m/%y')}\n"
        )
        sleep(1)

        # Campo ATRIBUIR ao Tecnico
        botao_atribuir = driver.find_element(By.CSS_SELECTOR, "#assignee-field")
        ActionChains(driver).scroll_to_element(botao_atribuir).perform()

        tecnico = tabela.loc[linha, "Tecnico"].lower()
        for texto in tecnico:
            botao_atribuir.send_keys(str(texto))
            sleep(0.1)

        sleep(3)
        botao_atribuir.send_keys(Keys.ENTER)

        # Atribuir ao lider
        atribuir_Lider = driver.find_element(By.CSS_SELECTOR, "#reporter-field")
        ActionChains(driver).scroll_to_element(atribuir_Lider).perform()

        atribuir_Lider.send_keys("cds.jgregorio")
        sleep(1.5)
        atribuir_Lider.send_keys(Keys.ENTER)
        sleep(1)

        # Campo service tag 2
        inserir_img = driver.find_element(By.CSS_SELECTOR, ".css-z4z5yt")
        ActionChains(driver).scroll_to_element(inserir_img).perform()

        sleep(2)
        # Adicionar imagens ao formulario
        diretorio_imagens = rf"C:\Users\{user_name}\Desktop\Bot_Ticket\main\Imagens\{str(tabela.loc[linha,'Patrimonio'])}"

        # Listar todas as imagens na pasta
        imagens = os.listdir(diretorio_imagens)
        print(imagens)
        for imagem in imagens:
            if imagem.endswith((".png", ".jpg", ".jpeg", ".gif", ".bmp")):
                caminho = os.path.join(diretorio_imagens, imagem)
                print(caminho)
                # Encontrar o campo de upload de imagem
                botao_browse = driver.find_element(By.CSS_SELECTOR, ".css-z4z5yt")
                botao_browse.click()
                sleep(2)
                pyautogui.write(caminho)
                pyautogui.press("enter")
                sleep(1.5)
        sleep(5)

        camp_service = driver.find_element(By.CSS_SELECTOR, "#customfield_15185-field")
        ActionChains(driver).scroll_to_element(camp_service).perform()
        camp_service.click()
        camp_service.send_keys(str(st))

        botao_create = driver.find_element(
            By.CSS_SELECTOR,
            '[data-testid="issue-create.common.ui.footer.create-button"]',
        )
        botao_create.click()
        sleep(10)

        # Campo Ver Problema - Visualiza o chamado
        try:
            driver.find_element(By.CLASS_NAME, "css-1ixg46l").click()
            sleep(10)
        except NoSuchElementException:
            driver.find_element(By.CSS_SELECTOR, "a.css-1ixg46l")
            sleep(10)

        # Campo em aberto
        try:
            driver.find_element(
                By.CSS_SELECTOR,
                r"#issue\.fields\.status-view\.status-button > span.css-178ag6o",
            ).click()
            sleep(2)
        except NoSuchElementException:
            driver.find_element(
                By.CSS_SELECTOR,
                r"#issue\.fields\.status-view\.status-button > span.css-5a6fwh",
            )
            sleep(2)

        # campo em andamento
        try:
            driver.find_element(
                By.CSS_SELECTOR, "#react-select-2-option-0 > div > div > span._16jlkb7n"
            ).click()
            sleep(2)
        except NoSuchElementException:
            driver.find_element(
                By.CSS_SELECTOR,
                "#react-select-2-option-0 > div > div > span._1o9zidpf._1e0c1txw > span",
            ).click()
            sleep(2)
        sleep(2)
        # Classificaçao Harware
        driver.find_element(By.CSS_SELECTOR, "#customfield_14051").click()
        sleep(1)
        driver.find_element(
            By.CSS_SELECTOR, "#customfield_14051 > option.option-group-15949"
        ).click()

        # campo frente de atendimento
        driver.find_element(By.ID, "customfield_19857").click()
        driver.find_element(
            By.CSS_SELECTOR, "#customfield_19857 > option:nth-child(2)"
        ).click()

        sleep(1)

        # Campo responder ao cliente
        camp_Resp_Cliente = driver.find_element(By.ID, "comment")
        camp_Resp_Cliente.click()
        chave_txt = {}

        # dicionario txt
        with open(r"C:\Users\cds.dpaixao_ifood3rd\Desktop\Bot_Ticket\main\Res_cliente.txt", "r") as cliente:
            for txt in cliente:
                txt = txt.strip()
                chave, valor = txt.split(": ", 1)
                chave_txt[chave] = valor

            verificacao = status
            if verificacao in chave_txt:
                novo_valor = chave_txt[verificacao]
            camp_Resp_Cliente.send_keys(f"{verificacao}: {novo_valor}")

        # inserir imagens

        # campo classificaçao
        try:
            driver.find_element(
                By.CSS_SELECTOR, "#issue-workflow-transition-submit"
            ).click()
            sleep(2.5)
        except NoSuchElementException:
            driver.find_element(By.ID, "issue-workflow-transition-submit")
            sleep(2.5)

        sleep(5)
        # Associaçao de prblema
        camp_Pesquisa = driver.find_element(
            By.CSS_SELECTOR, 'input[data-test-id="search-dialog-input"]'
        )
        camp_Pesquisa.click()
        camp_Pesquisa.send_keys(str(st), Keys.ENTER)

        sleep(10)
        try:
            camp_associacao = driver.find_element(By.CSS_SELECTOR, '.css-1uslh63')
            camp_associacao.click()
            sleep(5)

        except NoSuchElementException:
            camp_associacao = driver.find_element(By.CSS_SELECTOR, "#ak-main-content > div > div._1e0c1txw._2lx21bp4._1n261q9c._1h6d1fzn._1dqonqa1._189et94y._2rkopd34._1reo15vq._18m915vq > div._1e0c1txw._kqswh2mm._2lx21bp4._1n261q9c._1reo1wug._18m91wug._ouuo1ylp > table > tbody > tr._4t3i1ylp._kqswpfqs._nt751r31._49pcglyw._1hvw1o36._16qs1ahy._bfhk29zg._irr329zg > td:nth-child(2) > div > div > a")
            camp_associacao.click()
            sleep(5)

        camp_Copiar = driver.find_element(
            By.CSS_SELECTOR,
            'span[data-testid="issue.common.component.permalink-button.button.link-icon"]',
        )
        camp_Copiar.click()
        sleep(2)

        driver.back()
        sleep(10)
        driver.back()
        
        # botao associçao de problema
        botao_ass = driver.find_element(
            By.CSS_SELECTOR,
            'button[data-testid="issue.views.issue-base.foundation.quick-add.link-button.ui.link-dropdown-button"]',
        )
        botao_ass.click()
        sleep(1.5)

        botao_ass2 = driver.find_element(
            By.CSS_SELECTOR,
            'button[data-testid="issue.issue-view.views.issue-base.foundation.quick-add.quick-add-item.link-issue"]',
        )
        botao_ass2.click()
        sleep(1.5)

        ctrlV = driver.find_element(By.CSS_SELECTOR, ".issue-links-search__input")
        ctrlV.send_keys(Keys.CONTROL, "v", Keys.ENTER)
        sleep(1.5)

        botao_ass3 = driver.find_element(
            By.CSS_SELECTOR,
            'div[data-testid="issue.issue-view.views.issue-base.content.issue-links.add.issue-links-add-view.link-button"]',
        )
        botao_ass3.click()
        sleep(2)

        # Copiando o numero do ticket
        n_ticket = driver.find_element(
            By.CSS_SELECTOR,
            'a[data-testid="issue.views.issue-base.foundation.breadcrumbs.current-issue.item"]',
        )
        n_ticket2 = n_ticket.get_attribute("href")
        n_ticket2 = n_ticket2[34:]
        numero_Ticket = n_ticket2

        if (
            (status == "DISPONIVEL")
            or (status == "DEFEITO")
            or (status == "BATERIA")
            or (status == "DESCARTE")
            or (status == "DOAÇAO")
        ):
            # Finalizaçao do chamado
            botao_Em_andamento = driver.find_element(
                By.CSS_SELECTOR,
                '[data-testid="issue-field-status.ui.status-view.status-button.status-button"]',
            )
            botao_Em_andamento.click()
            sleep(5)
            try:
                camp_resolvido = driver.find_element(By.ID, "react-select-6-option-9")
                ActionChains(driver).scroll_to_element(camp_resolvido).perform()
                sleep(5)
                camp_resolvido.click()
                sleep(5)

            except NoSuchElementException:
                try:
                    camp_resolvido = driver.find_element(By.CSS_SELECTOR, "react-select-5-option-9")
                    ActionChains(driver).scroll_to_element(camp_resolvido).perform()
                    sleep(2.5)
                    camp_resolvido.click()
                    sleep(5)
                except NoSuchElementException:
                    id_result = driver.find_element(By.CSS_SELECTOR, ".css-rz73v4-option")
                    id_result = id_result.get_attribute("id")
                    id_result = id_result[:-1]
                    camp_resolvido = f"{id_result}9"
                    print(camp_resolvido)
                    camp_resolvido = driver.find_element(By.ID, camp_resolvido)
                    ActionChains(driver).scroll_to_element(camp_resolvido).perform()
                    sleep(2.5)
                    camp_resolvido.click()
                    sleep(5)


            # validaçao Defeito do equipamento
            camp_Defeito_equip = driver.find_element(By.CSS_SELECTOR, '[name="customfield_15907"]')
            camp_Defeito_equip.click()
            sleep(2)

            estado_equip = tabela.loc[linha, "Estado_Equipamento"].upper().strip()
            if estado_equip == "NOVO":
                camp_Defeito_equip = driver.find_element(By.CSS_SELECTOR, "#customfield_15907 > option:nth-child(4)")
                camp_Defeito_equip.click()
                sleep(2)
            else:
                camp_Defeito_equip = driver.find_element(By.CSS_SELECTOR, "#customfield_15907 > option:nth-child(3)")
                camp_Defeito_equip.click()
                sleep(2)

            # Campo troca
            if estado_equip == "NOVO":
                camp_troca = driver.find_element(By.CSS_SELECTOR, "#customfield_18105 > option:nth-child(3)")
                camp_troca.click()
                sleep(2)
            else:
                camp_troca = driver.find_element(By.CSS_SELECTOR, "#customfield_18105 > option:nth-child(2)")
                camp_troca.click()
                sleep(2)

            # Patrimonio encerrando chamado
            pt = tabela.loc[linha, "Patrimonio"]
            input_patrimonio = driver.find_element(By.CSS_SELECTOR, '[name="customfield_14879"]')
            input_patrimonio.send_keys(str(pt))
            input_patrimonio.send_keys(Keys.ENTER)

        elif status == "FABRICANTE":
            botao_pen_fornecedor = driver.find_element(By.CSS_SELECTOR, '[id="react-select-18-option-6"]')
        driver.quit()

        # INTEGRAÇAO GLPI
        options = Options()
        s = Service(ChromeDriverManager().install())
        options.add_argument("--start-maximized")
        options.add_argument("--force-device-scale-factor=0.7")
        options.add_argument("--profile-directory=Default")
        options.add_argument(
            r"--user-data-dir=C:\Users\{0}\AppData\Local\Google\Chrome\User Data".format(
                user_name
            )
        )

        st = tabela.loc[linha, "Service_Tag"]

        driver = webdriver.Chrome(service=s, options=options)
        driver.get(
            f"https://inventory.ifoodcorp.com.br/front/computer.php?is_deleted=0&as_map=0&browse=0&criteria%5B0%5D%5Blink%5D=AND&criteria%5B0%5D%5Bfield%5D=view&criteria%5B0%5D%5Bsearchtype%5D=contains&criteria%5B0%5D%5Bvalue%5D={st}&itemtype=Computer&start=0&_glpi_csrf_token=51691ee852929c2669adf0a485cb1a60204963f4b6c129dccca56918a9735867&sort%5B%5D=1&order%5B%5D=ASC"
        )
        sleep(10)

        id_comput = 'a[id^="Computer_"]'
        camp_Patrimonio = driver.find_element(By.CSS_SELECTOR, id_comput)
        id_comput = camp_Patrimonio.get_attribute("id")
        camp_Patrimonio = driver.find_element(By.ID, id_comput)
        camp_Patrimonio.click()
        sleep(3.5)

        driver.find_element(
            By.CSS_SELECTOR, 'a.nav-link[title="Informações Gerais"]'
        ).click()
        sleep(2.5)
        driver.find_element(By.CSS_SELECTOR, '[data-select2-id="31"]').click()
        sleep(2.5)

        tc = tabela.loc[linha, "Tecnico"].split()
        camp_Analista = driver.find_element(
            By.CSS_SELECTOR, "input.select2-search__field"
        )
        camp_Analista.send_keys(str(tc[0]))
        sleep(1.5)
        camp_Analista.send_keys(Keys.ENTER)
        sleep(1.5)
        camp_Tp_Analise = driver.find_element(By.CSS_SELECTOR, '[data-select2-id="29"]')
        sleep(1.5)
        camp_Tp_Analise.click()
        sleep(1)

        # Campo Tipo de analise
        try:
            class_pai = "ul.select2-results__options"  # nome da class do menu suspenso
            camp_Tp_Analise2 = driver.find_element(By.CSS_SELECTOR, class_pai)
            id_pai = camp_Tp_Analise2.get_attribute("id")
            pai = f"#{id_pai} li[title='Análise Técnica - '] > span[title='Análise Técnica - ']"
            camp_Tp_Analise2 = driver.find_element(By.CSS_SELECTOR, pai)
            camp_Tp_Analise2.click()
            sleep(1.5)

        except NoSuchElementException:
            pass
            sleep(1.5)

        # Campo Data da Analise
        class_input_data = "div.input-group.flatpickr"
        camp_Data = driver.find_element(By.CSS_SELECTOR, class_input_data)
        id_input = camp_Data.get_attribute("id")
        camp_Data = driver.find_element(By.ID, id_input)
        camp_Data.click()
        sleep(1.5)

        camp_Data2 = driver.find_element(
            By.CSS_SELECTOR, 'button.ms-2.btn.btn-outline-secondary[btn-id="0"]'
        )
        camp_Data2.click()
        sleep(1.5)

        # Campo Numero do chamado
        id_n_Ticket = 'input[name="ndechamadodeatfield"]'
        camp_N_chamado = driver.find_element(By.CSS_SELECTOR, id_n_Ticket)
        id_n_Ticket = camp_N_chamado.get_attribute("id")
        camp_N_chamado = driver.find_element(By.ID, id_n_Ticket)
        camp_N_chamado.click()
        sleep(1.5)
        camp_N_chamado.clear()
        sleep(1.5)
        camp_N_chamado.send_keys(str(numero_Ticket))
        sleep(1.5)
        camp_estado_EQ = driver.find_element(By.CSS_SELECTOR, '[data-select2-id="33"]')
        camp_estado_EQ.click()
        sleep(1)

        status = tabela.loc[linha, "Status"].upper()
        if status == "DISPONIVEL":
            x = driver.find_element(By.CSS_SELECTOR, ".select2-search__field")
            sleep(2.5)
            x.send_keys(str(status))
            sleep(1)
            x.send_keys(Keys.ENTER)
            sleep(1.5)

        elif status == "FABRICANTE":
            x = driver.find_element(By.CSS_SELECTOR, ".select2-search__field")
            sleep(2.5)
            x.send_keys(str(status))
            sleep(1)
            x.send_keys(Keys.ENTER)
            sleep(1.5)

        elif status == "DESCARTE":
            x = driver.find_element(By.CSS_SELECTOR, ".select2-search__field")
            sleep(2.5)
            x.send_keys(str(status))
            sleep(1)
            x.send_keys(Keys.ENTER)
            sleep(1.5)

        elif status == "BATERIA":
            x = driver.find_element(By.CSS_SELECTOR, ".select2-search__field")
            sleep(2.5)
            x.send_keys(str(status))
            sleep(1)
            x.send_keys(Keys.ENTER)
            sleep(1.5)

        # localizaçao estoque
        camp_loc_estoque = driver.find_element(
            By.CSS_SELECTOR, '[data-select2-id="41"]'
        )
        camp_loc_estoque.click()
        sleep(1.5)
        # campo prateleras
        camp_prateleira = driver.find_element(
            By.CSS_SELECTOR, "input.select2-search__field"
        )
        fabri = tabela.loc[linha, "Fabricante"].upper().strip()

        # Processo logico para alocar os equipamento
        processador = str(tabela.loc[linha, "Processador"]).upper().strip()
        modelo = tabela.loc[linha, "Modelo"].strip()

        if status == "DISPONIVEL":

            if (
                modelo == "Latitude 3440 i7 16GB"
                and processador == "I7"
                and status == "DISPONIVEL"
            ):
                camp_prateleira.send_keys(str("D3C1"))
                sleep(2)
                camp_prateleira.send_keys(Keys.ENTER)
                sleep(2.5)

            elif (
                modelo == "Latitude 3440 i7 16GB"
                and processador == "I5"
                and status == "DISPONIVEL"
            ):
                camp_prateleira.send_keys(str("D3B1"))
                sleep(2)
                camp_prateleira.send_keys(Keys.ENTER)
                sleep(2.5)

            elif processador == "I7" and fabri == "DELL" or fabri == "DELL INC.":
                camp_prateleira.send_keys(str("D2B1"))
                sleep(2)
                camp_prateleira.send_keys(Keys.ENTER)
                sleep(2.5)

            elif processador == "I7" and fabri == "LENOVO":
                camp_prateleira.send_keys(str("D1B1"))
                sleep(2)
                camp_prateleira.send_keys(Keys.ENTER)
                sleep(2.5)

            elif processador == "I5" and fabri == "DELL" or fabri == "DELL INC.":
                camp_prateleira.send_keys(str("D5B1"))
                sleep(2)
                camp_prateleira.send_keys(Keys.ENTER)
                sleep(2.5)

            elif processador == "I5" and fabri == "LENOVO":
                camp_prateleira.send_keys(str("D5C1"))
                sleep(2)
                camp_prateleira.send_keys(Keys.ENTER)
                sleep(2.5)

            elif processador == "I5" and fabri == "APPLE" or fabri == "APPLE INC.":
                camp_prateleira.send_keys(str("D25B1"))
                sleep(2)
                camp_prateleira.send_keys(Keys.ENTER)
                sleep(2.5)

            elif processador == "I7" and fabri == "APPLE" or fabri == "APPLE INC.":
                camp_prateleira.send_keys(str("D25B1"))
                sleep(2)
                camp_prateleira.send_keys(Keys.ENTER)
                sleep(2.5)

            elif (
                (processador == "M1" and fabri == "APPLE" or fabri == "APPLE INC.")
                or (processador == "M2" and fabri == "APPLE" or fabri == "APPLE INC.")
                or (processador == "M3" and fabri == "APPLE" or fabri == "APPLE INC.")
            ):
                camp_prateleira.send_keys(str("D21B1"))
                sleep(2)
                camp_prateleira.send_keys(Keys.ENTER)
                sleep(2.5)

        elif status == "FABRICANTE":

            if processador == "I7" and status == "FABRICANTE":
                camp_prateleira.send_keys(str("AT4B1"))
                sleep(2)
                camp_prateleira.send_keys(Keys.ENTER)
                sleep(2.5)

                try:
                    cam_defeito = driver.find_element(
                        By.CSS_SELECTOR, '[data-select2-id="35"]'
                    )
                    cam_defeito.click()
                    sleep(1.5)

                    cam_input_defeito = driver.find_element(
                        By.CSS_SELECTOR, "input.select2-search__field"
                    )
                    cam_input_defeito.send_keys("Placa Mae")
                    sleep(2)
                    cam_input_defeito.send_keys(Keys.ENTER)
                    sleep(1.5)

                except NoSuchElementException:
                    cam_defeito = driver.find_element(
                        By.CSS_SELECTOR, '[data-select2-id="13"]'
                    )
                    cam_defeito.click()
                    sleep(1.5)

                    cam_input_defeito = driver.find_element(
                        By.CSS_SELECTOR, "input.select2-search__field"
                    )
                    cam_input_defeito.send_keys("Placa Mae")
                    sleep(2)
                    cam_input_defeito.send_keys(Keys.ENTER)
                    sleep(1.5)

            elif processador == "I5" and status == "FABRICANTE":
                camp_prateleira.send_keys(str("AT4B2"))
                sleep(2)
                camp_prateleira.send_keys(Keys.ENTER)
                sleep(2.5)

                try:
                    cam_defeito = driver.find_element(
                        By.CSS_SELECTOR, '[data-select2-id="35"]'
                    )
                    cam_defeito.click()
                    sleep(1.5)

                    cam_input_defeito = driver.find_element(
                        By.CSS_SELECTOR, "input.select2-search__field"
                    )
                    cam_input_defeito.send_keys("Placa Mae")
                    sleep(2)
                    cam_input_defeito.send_keys(Keys.ENTER)
                    sleep(1.5)

                except NoSuchElementException:
                    cam_defeito = driver.find_element(
                        By.CSS_SELECTOR, '[data-select2-id="13"]'
                    )
                    cam_defeito.click()
                    sleep(1.5)

                    cam_input_defeito = driver.find_element(
                        By.CSS_SELECTOR, "input.select2-search__field"
                    )
                    cam_input_defeito.send_keys("Placa Mae")
                    sleep(2)
                    cam_input_defeito.send_keys(Keys.ENTER)
                    sleep(1.5)

        elif status == "DESCARTE":

            if fabri == "DELL":
                camp_prateleira.send_keys("D15D1")
                sleep(1.5)
                camp_prateleira.send_keys(Keys.ENTER)
                sleep(2)
                try:
                    cam_defeito = driver.find_element(
                        By.CSS_SELECTOR, '[data-select2-id="35"]'
                    )
                    cam_defeito.click()
                    sleep(1.5)

                    cam_input_defeito = driver.find_element(
                        By.CSS_SELECTOR, "input.select2-search__field"
                    )
                    cam_input_defeito.send_keys("Placa Mae")
                    sleep(2)
                    cam_input_defeito.send_keys(Keys.ENTER)
                    sleep(1.5)

                except NoSuchElementException:
                    cam_defeito = driver.find_element(
                        By.CSS_SELECTOR, '[data-select2-id="13"]'
                    )
                    cam_defeito.click()
                    sleep(1.5)

                    cam_input_defeito = driver.find_element(
                        By.CSS_SELECTOR, "input.select2-search__field"
                    )
                    cam_input_defeito.send_keys("Placa Mae")
                    sleep(2)
                    cam_input_defeito.send_keys(Keys.ENTER)
                    sleep(1.5)

            elif fabri == "LENOVO":
                camp_prateleira.send_keys("D11B1")
                sleep(1.5)
                camp_prateleira.send_keys(Keys.ENTER)
                sleep(2)
                try:
                    cam_defeito = driver.find_element(
                        By.CSS_SELECTOR, '[data-select2-id="35"]'
                    )
                    cam_defeito.click()
                    sleep(1.5)

                    cam_input_defeito = driver.find_element(
                        By.CSS_SELECTOR, "input.select2-search__field"
                    )
                    cam_input_defeito.send_keys("Placa Mae")
                    sleep(2)
                    cam_input_defeito.send_keys(Keys.ENTER)
                    sleep(1.5)

                except NoSuchElementException:
                    cam_defeito = driver.find_element(
                        By.CSS_SELECTOR, '[data-select2-id="13"]'
                    )
                    cam_defeito.click()
                    sleep(1.5)

                    cam_input_defeito = driver.find_element(
                        By.CSS_SELECTOR, "input.select2-search__field"
                    )
                    cam_input_defeito.send_keys("Placa Mae")
                    sleep(2)
                    cam_input_defeito.send_keys(Keys.ENTER)
                    sleep(1.5)

            elif fabri == "APPLE":
                camp_prateleira.send_keys("D14B1")
                sleep(1.5)
                camp_prateleira.send_keys(Keys.ENTER)
                sleep(2)
                try:
                    cam_defeito = driver.find_element(
                        By.CSS_SELECTOR, '[data-select2-id="35"]'
                    )
                    cam_defeito.click()
                    sleep(1.5)

                    cam_input_defeito = driver.find_element(
                        By.CSS_SELECTOR, "input.select2-search__field"
                    )
                    cam_input_defeito.send_keys("Placa Mae")
                    sleep(2)
                    cam_input_defeito.send_keys(Keys.ENTER)
                    sleep(1.5)

                except NoSuchElementException:
                    cam_defeito = driver.find_element(
                        By.CSS_SELECTOR, '[data-select2-id="13"]'
                    )
                    cam_defeito.click()
                    sleep(1.5)

                    cam_input_defeito = driver.find_element(
                        By.CSS_SELECTOR, "input.select2-search__field"
                    )
                    cam_input_defeito.send_keys("Placa Mae")
                    sleep(2)
                    cam_input_defeito.send_keys(Keys.ENTER)
                    sleep(1.5)
            else:
                print("Opcao nao encontrada!!!")
                pass
        # BATERIA
        elif status == "BATERIA":
            hoje = datetime.now().strftime("%d/%m/%Y")
            garantia = tabela.loc[linha, "Garantia"].strftime("%d/%m/%y")

            if fabri == "DELL":
                if hoje >= garantia:  # maquinas bateria sem garantia
                    camp_prateleira.send_keys("D10C1")
                    sleep(1.5)
                    camp_prateleira.send_keys(Keys.ENTER)
                    sleep(2)
                else:  # maquinas bateria com garantia
                    camp_prateleira.send_keys("D10B1")
                    sleep(1.5)
                    camp_prateleira.send_keys(Keys.ENTER)
                    sleep(2)

                try:
                    cam_defeito = driver.find_element(
                        By.CSS_SELECTOR, '[data-select2-id="35"]'
                    )
                    cam_defeito.click()
                    sleep(1.5)

                    cam_input_defeito = driver.find_element(
                        By.CSS_SELECTOR, "input.select2-search__field"
                    )
                    cam_input_defeito.send_keys("Bateria")
                    sleep(2)
                    cam_input_defeito.send_keys(Keys.ENTER)
                    sleep(1.5)

                except NoSuchElementException:
                    cam_defeito = driver.find_element(
                        By.CSS_SELECTOR, '[data-select2-id="13"]'
                    )
                    cam_defeito.click()
                    sleep(1.5)

                    cam_input_defeito = driver.find_element(
                        By.CSS_SELECTOR, "input.select2-search__field"
                    )
                    cam_input_defeito.send_keys("BATERIA")
                    sleep(2)
                    cam_input_defeito.send_keys(Keys.ENTER)
                    sleep(1.5)

            elif fabri == "LENOVO":
                if hoje >= garantia:  # maquinas bateria sem garantia
                    camp_prateleira.send_keys("D10C1")
                    sleep(1.5)
                    camp_prateleira.send_keys(Keys.ENTER)
                    sleep(2)
                else:  # maquinas bateria com garantia
                    camp_prateleira.send_keys("D10B1")
                    sleep(1.5)
                    camp_prateleira.send_keys(Keys.ENTER)
                    sleep(2)
                try:
                    cam_defeito = driver.find_element(
                        By.CSS_SELECTOR, '[data-select2-id="35"]'
                    )
                    cam_defeito.click()
                    sleep(1.5)

                    cam_input_defeito = driver.find_element(
                        By.CSS_SELECTOR, "input.select2-search__field"
                    )
                    cam_input_defeito.send_keys("BATERIA")
                    sleep(2)
                    cam_input_defeito.send_keys(Keys.ENTER)
                    sleep(1.5)

                except NoSuchElementException:
                    cam_defeito = driver.find_element(
                        By.CSS_SELECTOR, '[data-select2-id="13"]'
                    )
                    cam_defeito.click()
                    sleep(1.5)

                    cam_input_defeito = driver.find_element(
                        By.CSS_SELECTOR, "input.select2-search__field"
                    )
                    cam_input_defeito.send_keys("BATERIA")
                    sleep(2)
                    cam_input_defeito.send_keys(Keys.ENTER)
                    sleep(1.5)

            elif fabri == "APPLE":
                if hoje >= garantia:  # maquinas bateria sem garantia
                    camp_prateleira.send_keys("D14B4")
                    sleep(1.5)
                    camp_prateleira.send_keys(Keys.ENTER)
                    sleep(2)
                else:  # maquinas bateria com garantia
                    camp_prateleira.send_keys("D14B4")
                    sleep(1.5)
                    camp_prateleira.send_keys(Keys.ENTER)
                    sleep(2)
                try:
                    cam_defeito = driver.find_element(
                        By.CSS_SELECTOR, '[data-select2-id="35"]'
                    )
                    cam_defeito.click()
                    sleep(1.5)

                    cam_input_defeito = driver.find_element(
                        By.CSS_SELECTOR, "input.select2-search__field"
                    )
                    cam_input_defeito.send_keys("BATERIA")
                    sleep(2)
                    cam_input_defeito.send_keys(Keys.ENTER)
                    sleep(1.5)

                except NoSuchElementException:
                    cam_defeito = driver.find_element(
                        By.CSS_SELECTOR, '[data-select2-id="13"]'
                    )
                    cam_defeito.click()
                    sleep(1.5)

                    cam_input_defeito = driver.find_element(
                        By.CSS_SELECTOR, "input.select2-search__field"
                    )
                    cam_input_defeito.send_keys("Placa Mae")
                    sleep(2)
                    cam_input_defeito.send_keys(Keys.ENTER)
                    sleep(1.5)
            else:
                print("Opcao nao encontrada!!!")

        else:
            pass
        # CAMPO 0BSERVAÇAO
        camp_obs = driver.find_element(By.CSS_SELECTOR, '[name="observaofieldthree"]')
        camp_obs.click()
        sleep(1)
        camp_obs.clear()
        sleep(1)
        camp_obs.send_keys(
            "Analise tecnica relizada com sucesso: -" + str(numero_Ticket)
        )
        sleep(1.5)

        # botao salve
        bot_salve = driver.find_element(
            By.CSS_SELECTOR, '[name="update_fields_values"]'
        )
        bot_salve.click()
        sleep(2.5)

        driver.quit()
        fim = time.time()
        tem_exe = (fim - temp) / 60
        print(f"Tampo de execuçao: {tem_exe:.2f}")
        tem_exe = -tem_exe


if __name__ == "__main__":
    temp = time.time()
    main()
    fim = time.time()
    tem_exe = (fim - temp) / 60
    print(f"Tampo de execuçao: {tem_exe:.2f}")
ctypes.windll.kernel32.SetThreadExecutionState(0x00000002)
