import tkinter as tk
import webbrowser
import time
import json
import sys
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials
import gspread
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import requests




mensagem_except_selenium = "ELEMENTO NÃO ENCONTRADO NA PAGINA TENTE NOVAMENTE OU ENTRE EM CONTATO COM O DESENVOLVEDOR"
def get_JSON_pastebin():
    # URL do Pastebin com as credenciais
    pastebin_url = "COLQUE AQUI A URL DO PASTEBIN COM AS CREDENCIAIS"

    # Baixar o conteúdo do Pastebin
    response = requests.get(pastebin_url)

    # Verificar se a solicitação foi bem-sucedida
    if response.status_code == 200:
        credentials_json = response.text
        # Carregar o JSON das credenciais
        credentials_dict = json.loads(credentials_json)
        return credentials_dict
    else:
        print(f"Falha ao acessar o Pastebin: {response.status_code}")
        return None

# Chamada da função
credentials_dict = get_JSON_pastebin()   

def autenticar_google_sheets():
    if credentials_dict is None:
        raise Exception("Não foi possível obter as credenciais do Pastebin.")
    
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
             "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
    client = gspread.authorize(creds)
    return client
# VARIÁVEIS PRIMORDIAIS

client = autenticar_google_sheets()
planilha = client.open('NOME DA PLANILHA GOOGLE SHEETS')
folha = planilha.worksheet('MENSAGENS AUTOMATICAS')
tempo_maximo_espera = 15
limite_mensagens= 0
valores = folha.get_all_values()
header = valores[0]

# Função para obter a mensagem personalizada
def obter_mensagem_personalizada(codigo_mensagem, valores, indice_codigo_mensagem_p, indice_mensagem_p):
    for linha in valores[1:]:
        if linha[indice_codigo_mensagem_p] == codigo_mensagem:
            return linha[indice_mensagem_p]
    return None

# Função para exibir a mensagem personalizada no chatbot
def exibir_mensagem_personalizada(mensagem, driver):
    
    js_add_text_to_input = """
    var elm = arguments[0], txt = arguments[1];
    elm.value += txt;
    elm.dispatchEvent(new Event('change'));
    """
    
    textarea = WebDriverWait(driver, tempo_maximo_espera).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, ".MuiInputBase-input.MuiOutlinedInput-input.MuiInputBase-inputMultiline.MuiOutlinedInput-inputMultiline")))
    textarea.click()

    
    driver.execute_script(js_add_text_to_input, textarea, mensagem)
    textarea.click()
    textarea.send_keys(Keys.SPACE)
    textarea.send_keys(Keys.BACKSPACE)
    time.sleep(.2)
    textarea.send_keys(Keys.RETURN)
        
# Função para tratar o alerta do chatbot
def tratar_alerta(driver, folha, indice_linha, indice_status, alert_text):
    global limite_mensagens
       
    if "Já existe um atendimento para esse contato" in alert_text:
        wait = WebDriverWait(driver, 15)
        alert_element = wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, ".Toastify__toast-body")))

        elemento_fechar = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "#add-chat-dialog > div.MuiDialog-container.MuiDialog-scrollPaper > div > div > div:nth-child(1) > div > div:nth-child(2) > button > span.MuiIconButton-label > svg"))
        )
        elemento_fechar.click()

        inicio_nome = alert_text.find("sendo atendido por ") + len("sendo atendido por ")
        fim_nome_protocolo = alert_text.find(" - Protocolo")
        atendente = alert_text[inicio_nome:fim_nome_protocolo]
        protocolo = alert_text[fim_nome_protocolo + len(" - Protocolo "):]

        mensagem_tratativa = f"Atendido por {atendente} - Protocolo {protocolo}"
        folha.update_cell(indice_linha, indice_status+1, mensagem_tratativa)
        data_hoje = datetime.today().strftime("%d/%m/%Y")
        indice_data = valores[0].index("DATA ENVIO") + 1
        folha.update_cell(indice_linha, indice_data, data_hoje)
        print(mensagem_tratativa)

    elif "Esse número não utiliza o WhatsApp" in alert_text:
        wait = WebDriverWait(driver, 15)
        alert_element = wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, ".Toastify__toast-body")))

        elemento_fechar = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "#add-chat-dialog > div.MuiDialog-container.MuiDialog-scrollPaper > div > div > div:nth-child(1) > div > div:nth-child(2) > button > span.MuiIconButton-label > svg"))
        )
        elemento_fechar.click()
            
        mensagem_tratativa = "Número não é WhatsApp"
        folha.update_cell(indice_linha, indice_status + 1, mensagem_tratativa)
        limite_mensagens +1
        data_hoje = datetime.today().strftime("%d/%m/%Y")
        indice_data = valores[0].index("DATA ENVIO") + 1
        folha.update_cell(indice_linha, indice_data, data_hoje)
        print(mensagem_tratativa)

        print()
# Configuração da interface gráfica com Tkinter


def indice_feito_parametros(folha, header):
    # Obter o índice da coluna "FEITO"
    indice_feito = header.index("FEITO")
    
    # Função para contar quantos "NÃO" existem na coluna "FEITO"
    def contar_nao(folha, indice_feito):
        valores = folha.get_all_values()
        restante_nao = 0  # Inicializar a variável de contagem
        
        for linha in valores[1:]:  # Ignorar a primeira linha de cabeçalho
            if linha[indice_feito] == "NÃO":
                restante_nao += 1  # Contar quantos "NÃO" existem
        
        return restante_nao

    # Função para obter o índice da primeira linha que contém "NÃO"
    def obter_indice_proximo_nao(folha, indice_feito):
        valores = folha.get_all_values()
        for indice_linha, linha in enumerate(valores[1:], start=2):
            if linha[indice_feito] == "NÃO":
                return indice_linha
        return None

    # Chamar as funções e armazenar os resultados
    restante_nao = contar_nao(folha, indice_feito)
    indice_linha = obter_indice_proximo_nao(folha, indice_feito)

    return restante_nao, indice_linha

restante_nao, indice_linha = indice_feito_parametros(folha, header)

print(restante_nao,indice_linha)



def run_automation(limite_mensagens):  # Ensure it has a default value
    valores = folha.get_all_values()
    restante_nao, indice_linha = indice_feito_parametros(folha, header)
    if restante_nao <= 0 and (indice_linha == 0 or indice_linha is None):
        print("SEM CLIENTE PARA ENVIO DE MENSAGEM NA PLANILHA")
        return    
    
    
    # Definir os índices das colunas de interesse
    indice_associado = header.index("ASSOCIADO")
    indice_placa = header.index("PLACA")
    indice_telefone = header.index("TELEFONE")
    indice_status = header.index("STATUS")
    indice_codigo_mensagem = header.index("CODIGO_MENSAGEM")
    indice_feito = header.index("FEITO")
    indice_data = header.index("DATA ENVIO")
    indice_atendente = header.index("ATENDENTE")
    indice_login = header.index("LOGIN")
    indice_senha = header.index("SENHA")
    indice_codigo_mensagem_p = header.index("CODIGO_MENSAGEM_P")
    indice_mensagem_p = header.index("MENSAGEM_PERSONALIZADA")
    indice_codigo_boleto = header.index("CODIGO_BOLETO")
    indice_link_boleto = header.index("LINK_BOLETO")
    # Obter o atendente selecionado
    atendente_selecionado = folha.cell(1, 11).value

    login = None
    senha = None
    for linha in valores[1:]:
        if linha[indice_atendente] == atendente_selecionado:
            login = linha[indice_login]
            senha = linha[indice_senha]
            break

    if login is None or senha is None:
        raise ValueError(f"Atendente '{atendente_selecionado}' não encontrado na planilha.")

    chrome_options = Options()
    chrome_options.add_argument("--mute-audio") 
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    
    
    # Certifique-se de que o serviço está sendo passado corretamente
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    tempo_maximo_espera = 15
    driver.maximize_window()
    # Abrir a página do CHATBOT
    driver.get("SITE DO CHATBOT AQUI")
    
    # Preencher os campos de login e senha
    elemento_login = WebDriverWait(driver, tempo_maximo_espera).until(
    EC.visibility_of_element_located((By.XPATH, "//*[@id='root']/div/div[1]/div/div/div/div/div/div/form/div/div[1]/div/div/input"))
    )
    
    try:        
        elemento_senha = WebDriverWait(driver, tempo_maximo_espera).until(
            EC.visibility_of_element_located((By.XPATH, "//*[@id='root']/div/div[1]/div/div/div/div/div/div/form/div/div[2]/div/div/input"))
        )
        elemento_login.send_keys(login)
        elemento_senha.send_keys(senha)

        elemento_entrar = driver.find_element(By.XPATH, "//*[@id='root']/div/div[1]/div/div/div/div/div/div/form/div/div[3]/div/button/span[1]")
        elemento_entrar.click()
    except TimeoutException:
        print(mensagem_except_selenium)
    # Processar cada linha da planilha
    
    
    def loop_principal(limite_mensagens):
        restante_nao, indice_linha = indice_feito_parametros(folha, header)

        while restante_nao > 0 and indice_linha is not None:
            valores = folha.get_all_values()
            linha = valores[indice_linha - 1]  # -1 porque 'valores' é indexado em 0
            associado = linha[indice_associado]
            placa = linha[indice_placa]
            telefone = linha[indice_telefone]
            codigo_mensagem = linha[indice_codigo_mensagem]
            feito = linha[indice_feito]
            boletoB = linha[indice_codigo_boleto]
            boletoL = linha[indice_link_boleto]

            print(f"Associado: {associado}, Placa: {placa}, Telefone: {telefone}, codigo_mensagem: {codigo_mensagem}, Feito? {feito}, Boleto {boletoB}")

            if feito == "NÃO":
                try:
                    elemento_novo = WebDriverWait(driver, tempo_maximo_espera).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "#root > div > div.sc-kEYyzF.iXCCiQ.fullHeight.overflow-y-hidden > div.main-content.flex-column.fullHeight.overflow-y-hidden > div > div.left > div > div:nth-child(2) > div > div:nth-child(1) > div > div > div:nth-child(1) > div > div.MuiGrid-root.p-1.MuiGrid-item > div > div:nth-child(2) > a > span.MuiIconButton-label > svg"))
                    )
                    elemento_novo.click()
                    time.sleep(1)

                    xpath_departamento = WebDriverWait(driver, tempo_maximo_espera).until(
                        EC.visibility_of_element_located((By.XPATH, "//*[@id='mui-component-select-category']"))
                    )
                    action_chains = ActionChains(driver)
                    time.sleep(1.5)
                    action_chains.double_click(xpath_departamento).perform()

                    xpath_canal = WebDriverWait(driver, tempo_maximo_espera).until(
                        EC.visibility_of_element_located((By.XPATH, "//*[@id='mui-component-select-channel']"))
                    )
                    xpath_canal.click()

                    xpath_numero = WebDriverWait(driver, tempo_maximo_espera).until(
                        EC.visibility_of_element_located((By.XPATH, "//*[@id='menu-channel']/div[3]/ul/li[2]/div/div[2]/p[2]"))
                    )
                    xpath_numero.click()

                    xpath_telefone = WebDriverWait(driver, tempo_maximo_espera).until(
                        EC.visibility_of_element_located((By.XPATH, "//*[@id='add-chat-dialog']/div[3]/div/div/div[3]/div/form/div[1]/div[4]/div/div[1]/div/input"))
                    )
                    xpath_telefone.send_keys(telefone)
                    time.sleep(1.5)

                    try:
                        wait = WebDriverWait(driver, 5)
                        element_confirmar = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#react-autowhatever-1--item-0 > div")))
                        element_confirmar.click()
                    except TimeoutException:
                        print()

                    xpath_criar_tendimento = WebDriverWait(driver, tempo_maximo_espera).until(
                        EC.visibility_of_element_located((By.XPATH, "//*[@id='add-chat-dialog']/div[3]/div/div/div[3]/div/form/div[2]/div[2]/button/span[1]"))
                    )
                    xpath_criar_tendimento.click()

                    wait = WebDriverWait(driver, 10)
                    alert_element = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".Toastify__toast-body")))
                    alert_text = alert_element.text
                    tratar_alerta(driver, folha, indice_linha, indice_status, alert_text)

                    mensagem_base = obter_mensagem_personalizada(codigo_mensagem, valores, indice_codigo_mensagem_p, indice_mensagem_p)
                    if mensagem_base:
                        mensagem_formatada = mensagem_base.format(
                            associado=associado,
                            placa=placa,
                            telefone=telefone,
                            atendente=atendente_selecionado,
                            codigoboleto=boletoB,
                            linkboleto=boletoL
                        )
                        time.sleep(2)
                        exibir_mensagem_personalizada(mensagem_formatada, driver)
                        mensagem_tratativa = "Mensagem enviada"
                        folha.update_cell(indice_linha, indice_status + 1, mensagem_tratativa)
                        data_hoje = datetime.today().strftime("%d/%m/%Y")
                        indice_data = header.index("DATA ENVIO") + 1
                        folha.update_cell(indice_linha, indice_data, data_hoje)
                        print(mensagem_tratativa)
                    else:
                        print(f"Nenhuma mensagem personalizada encontrada para a codigo_mensagem: {codigo_mensagem}")

                except TimeoutException:
                    print()

                limite_mensagens -= 1
                if limite_mensagens <= 0:
                    print("Limite de mensagens alcançado.")
                    break
                print(f"Resta: {limite_mensagens}")
            elif feito == "":
                print("sem mais associados para envio na planilha")

            # Atualizar o restante de "NÃO" e o índice da próxima linha a ser processada
            restante_nao, indice_linha = indice_feito_parametros(folha, header)

        time.sleep(10)

    loop_principal(limite_mensagens)

# Função para redirecionar texto para a área de texto do Tkinter
class TextRedirector:
    def __init__(self, text_widget, stream):
        self.text_widget = text_widget
        self.stream = stream

    def write(self, message):
        self.text_widget.config(state='normal')
        self.text_widget.insert(tk.END, message)
        self.text_widget.config(state='disabled')
        self.text_widget.yview(tk.END)
        # Também escrever no stream original (console)
        self.stream.write(message)

    def flush(self):
        pass


def open_google_sheet():
    webbrowser.open("LINK DA PLANILHA AQUI")
    
def parar():
    sys.exit("parando a automação")