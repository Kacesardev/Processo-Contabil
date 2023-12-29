import pymysql
import base64
import openpyxl
import smtplib

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.service import Service

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains

import pandas as pd
import time

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment

import time
import datetime
import calendar
import os, sys


from configuracoes import *
import gravar_log_database

# Defina as variáveis
descricao_log = "Contabilização Manual"
codigo_script_log = 9 # Variavel de identificacao do Script no Banco de Dados - Tabela Catalogo
error_response = ''

regionais_lista = {'itaipu':'Itaipu','itaipunorte':'Itaipu Norte','equipo':'Equipo','quintaroda':'Quinta Roda', 'wlm-sede':'WLM-Sede', 'csc':'CSC'}
opcoes_empresa = {'itaipu':'WLM - REGIONAL MINAS','itaipunorte':'WLM - REGIONAL NORTE','equipo':'WLM - REGIONAL RIO','quintaroda':'WLM -  REGIONAL SAO PAULO', 'wlm-sede':'WLM - MATRIZ', 'csc':'WLM - CSC-MG'}
arquivo = "Robô Contabilização Manual Oficial.xlsx"

start = time.time()
status_script = "Ok"

driver = None

HOST_NAME = '172.10.10.8'
USER_NAME = 'aW1fYWNlc3Nv'
PASWD = 'U2NhbmlhQDIwMTk='

if len(sys.argv) >= 2:
    var_regional = sys.argv[1]
else:
    print("Parâmetros inválidos acesse com: Regional")
    sys.exit()
print(descricao_log+"\n")
print("definição das variáveis.\n")
local_arquivo = r"\\" + pasta_rede + "\itaipu-fs\Interdep\CSC-Contabilidade\Contabilidade - CSC\Contabilização Automática\\" + regionais_lista[var_regional] + "\\"
url_login = "https://acp-siconnet.scania.com.br/sicomnet-ace3/wlm/sicomweb.gen.gen.pag.Login.cls"

def decod(var):
    text_var = base64.b64decode(var.encode('ascii'))
    return text_var.decode('ascii')

def carrega_credencial(tipo_acesso):
    try:
        con = pymysql.connect(host=HOST_NAME, user=decod(USER_NAME), password=decod(PASWD), db='automacao_processos', charset='utf8')
        query_db = con.cursor()
    except Exception as e:
        print("Erro: Impossível conectar ao Banco de Dados, Erro: " + str(e))

    try:
        sql_command = "SELECT login, senha FROM cofre_acessos WHERE descricao=%s"
        query_db.execute(sql_command, (tipo_acesso,))
        row_db = query_db.fetchone()
    except Exception as e:
        print("Erro: Impossível acessar a tabela de acessos, Erro: " + str(e))
    finally:
        query_db.close()
        con.close()
    return row_db[0], row_db[1]

user_sicon, pwd_sicon = carrega_credencial('SiconNet')

#user_sicon = 'WLMCPJ'
#pwd_sicon = 'Nov,0119'



def send_email(subject, message):
    print("Email será enviado.\n")
    # Configurar as credenciais do servidor de email
    #smtp_server = "pod51028.outlook.com"
    smtp_port = 587
    #smtp_username = "sistemas@wlm.com.br"
    #smtp_password = "wlm_2011"
    sender_email = 'sistemas@wlm.com.br'
    receiver_email = 'carlos.junior@wlmequipo.com.br'

    # Criar o objeto de email
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = 'Robo de Contabilização Manual - Vecna'

    # Adicionar o corpo da mensagem
    body = message
    msg.attach(MIMEText(body, 'Olá,'
    'Aqui é o Vecna Robo de Contabilização,'
    'A automação foi inicializada.'))

    # Enviar o email
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_user, smtp_pass)
            server.send_message(msg)
        print("Email enviado com sucesso")
    except Exception as e:
        print(f"Erro ao enviar o email: {str(e)}")

send_email("Inicio da Rotina","Aqui é o Vecna Robo de Contabilização e a automação foi iniciada.")
    

def gravar_excel(valor_celula, posicoes_linha):
    wb = load_workbook(local_arquivo + "\\" + arquivo)
    ws = wb['Lançamentos Contábeis']
    for posicao in posicoes_linha:
        ws['T' + str(int(posicao) + 2)].value = str(valor_celula)
    try:
        wb.save(local_arquivo + "\\" + arquivo)
    except:
        print("Erro ao salvar no arquivo Excel, verifique se o arquivo esta aberto.")

posicao_linha_excel = []
lotes_existentes = []
numero_lancamento = []
valor_total = []
dia_lote = []

conta_debito = []
conta_credito = []
valor = []
digitado = []

centro_custo = []
conta_corrente = []

# Obter a regional
def obtem_regional(arquivo):
    regional = os.path.dirname(arquivo).split("\\")[-1]
    return regional

# Obter os dados do arquivo
print("obtendo dados do arquivo.\n")
df = pd.read_excel(local_arquivo + arquivo, dtype="str")
df["Valor"] = df["Valor"].astype(float)
df["Valor Total"] = df["Valor Total"].astype(float)
df["Valor"] = df["Valor"].apply(lambda x: '{:.2f}'.format(x))
df["Valor Total"] = df["Valor Total"].apply(lambda x: '{:.2f}'.format(x))
df.fillna('', inplace=True)

posicao_linha_excel = df.index.to_list()
lotes_existentes = df["Lote"].to_list()
numero_lancamento = df["Número de Lançamentos"].to_list()
valor_total = df["Valor Total"].to_list()
dia_lote = df["Dia do Lote"].to_list()

conta_debito = df["Conta Débito"].to_list()
conta_credito = df["Conta Crédito"].to_list()
valor = df["Valor"].to_list()
digitado = df["Digitado"].to_list()

centro_custo = df["Centro de Custos"].to_list()
conta_corrente = df["Conta Corrente"].to_list()

# Abrir o navegador
def abre_browser():
    print("browser sendo aberto.\n")
    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-automation")
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-extensions")
    options.add_argument("--proxy-server='direct://'")
    options.add_argument("--proxy-bypass-list=*")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--disable-notifications")  # Desabilitar notificações
    driver = webdriver.Chrome(service=Service(), options=options)
    return driver

# Entrar no login
def entra_login(driver, url_login):
    print("login será realizado.\n")
    driver.get(url_login)
    return driver

# Aceitar alerta
def aceita_alerta():
    print("aceitando alerta.\n")
    alert = driver.switch_to.alert
    alert.accept()

# Fazer login
def faz_login(driver, login, senha):
    print("fazendo login.\n")
    input_usuario = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "control_11")))
    input_senha = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.ID, "control_15")))
    input_usuario.click()
    input_usuario.clear()
    input_usuario.send_keys(login)
    input_senha.click()
    input_senha.clear()
    input_senha.send_keys(senha)
    botao = driver.find_element(By.ID, 'botaoLogin')
    botao.click()
    return driver

# Fazer logoff
def faz_logoff(driver):
    print("fazendo logoff.\n")
    time.sleep(3)
    driver.switch_to.default_content()
    driver.find_element(By.ID, "image_33").click()
    driver.switch_to.window(driver.window_handles[0])
    elemento = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.ID, "zen29")))
    driver.close()
    return driver

# Entrar em lançamento contábil
def entra_lac_contab(driver):
    print("entrando em lançamento contabil.\n")
    driver.switch_to.window(driver.window_handles[1])
    javascript_code = "zenPage.carregaFrame('sicomweb.ctm.mv.pag.LancamentoContabLote.cls', 552, 0);"
    driver.execute_script(javascript_code)
    return driver

# Selecionar o dropdown
def seleciona_dropdown(driver):
    print("selecionando dropdown.\n")
    iframe = driver.find_element(By.ID, 'iframe_34')
    driver.switch_to.frame(iframe)
    driver.find_element(By.ID, "btn_13").click()
    driver.execute_script("zenPage.getComponent(13).showDropdown();")
    return driver

# Capturar opções
def captura_opcoes(driver):
    print("capturando opções de empresa.\n")
    combo = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'table[class="comboboxTable"]')))
    combo = combo.text
    lista_combo = []
    combo_split = combo.split("\n")
    for i, texto in enumerate(combo_split):
        if i % 2 == 0:
            lista_combo.append(texto)
    print("Selecione a empresa desejada:\n")
    selecao_empresa = input(lista_combo).replace("'", "")
    return selecao_empresa

# Selecionar a empresa
def seleciona_empresa(driver, selecao_empresa):
    print(f"selecionando empresa. {selecao_empresa}\n")
    time.sleep(5)
    driver.find_element(By.CSS_SELECTOR, f"tr[zentext='{selecao_empresa}']").click()
    driver.find_element(By.CSS_SELECTOR, 'body[id="zenBody"]').click()
    botao_ok = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, 'control_14'))).click()
    return driver

def cadastra_1a_parte(driver, lotes_Existentes, i):
    print("cadastrando 1a parte.\n")
    lote = lotes_existentes[i]
    numero = numero_lancamento[i]
    valor = valor_total[i]
    dia = dia_lote[i]
    input_lote = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, 'control_60')))
    input_lote.click()
    input_lote.clear()
    input_lote.send_keys(lote)
    input_numero = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, 'control_62')))
    input_numero.click()
    input_numero.clear()
    input_numero.send_keys(numero)
    input_valor = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, 'control_63')))
    input_valor.click()
    input_valor.clear()
    input_valor.send_keys(valor)
    input_dia = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, 'control_64')))
    input_dia.click()
    input_dia.clear()
    input_dia.send_keys(dia)
    botao_gravar = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, 'control_16'))).click()
    aceita_alerta() 

# Clicar em lançamento
def abre_lancamento(driver):
    print("abrindo lançamento.\n")
    driver.find_element(By.CSS_SELECTOR, 'a.link#a_27').click()

# Efetuar lançamento
def efetua_lancamento(driver, i):
    print("efetuando 2a parte do lançamento contabil.\n")
    cd = conta_debito[i]
    cc = conta_credito[i]
    vlr = valor[i]
    digit = digitado[i]
    centro_c = centro_custo[i]
    conta_corr = conta_corrente[i]
    input_conta_debito = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, 'control_60')))
    input_conta_debito.click()
    input_conta_debito.clear()
    input_conta_debito.send_keys(cd)
    input_conta_credito = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, 'control_61')))
    input_conta_credito.click()
    input_conta_credito.clear()
    input_conta_credito.send_keys(cc)
    input_valor = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, 'control_62')))
    input_valor.click()
    input_valor.clear()
    input_valor.send_keys(vlr)
    driver.find_element(By.ID, 'caption_1_53').click()
    input_digitado = driver.find_element(By.ID, 'control_64')
    input_digitado.click()
    input_digitado.clear()
    input_digitado.send_keys(digit)
    try:
        input_centro_custo = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.ID, 'codigoRed_66')))
        input_centro_custo.click()
        input_centro_custo.clear()
        input_centro_custo.send_keys(centro_c)
    except:
        pass
    try:
        input_conta_corrente = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.ID, 'control_66')))
        input_conta_corrente.click()
        input_conta_corrente.send_keys(Keys.TAB)
    except:
         pass
    botao_gravar = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.ID, 'control_16'))).click()

# Retomar o lançamento
def retoma_lancamento(driver):
    print("retornando ao lançamento.\n")
    driver.find_element(By.ID,'a_28').click() # Botão Retornar

# Função para verificar se o status é "OK"
def verificar_status():
    global status_script, driver, df, error_response
    if status_script != "OK":
        try:
            #regional = obtem_regional(arquivo)
            driver = abre_browser()
            driver = entra_login(driver, url_login)
            aceita_alerta()
            driver = faz_login(driver, user_sicon, pwd_sicon)
            time.sleep(5)
            driver = entra_lac_contab(driver)
            time.sleep(5)
            driver = seleciona_dropdown(driver)
            time.sleep(5)
            driver = seleciona_empresa(driver, opcoes_empresa[var_regional])
            itens = len(df)
            for i in range(itens):
                print(f"Processando linha {i + 1} de {itens}")
                if i > 0:
                    retoma_lancamento(driver) #Gravar
                if i < len(df):
                    cadastra_1a_parte(driver, df, i)
                    abre_lancamento(driver)
                    efetua_lancamento(driver, i) #Gravar
                else:
                    print(f"Índice {i} está fora do intervalo das listas.")
                    continue
            print('Foram atualizados todos os dados da planilha')
            # Gravar Excel ==============================================
            gravar_excel('Ok', posicao_linha_excel)
            mensagem_email = (
                "Olá,\n"
                "Aqui é o Vecna Robo de Contabilização,\n"
                "Foram enviados todos os processos da planilha."
            )
            driver = faz_logoff(driver)
            ("Conclusão da rotina", mensagem_email)
        except Exception as e:
            mensagem = f"Erro na rotina.\n{e}"
            print(mensagem)
            status_script = "Not"
            print("Favor fazer contato com o administrador.")
            send_email("Erro na rotina", mensagem)
            error_response = mensagem
        finally:
            try:
                driver.close()
            except:
                pass
            driver.quit()
    
#envia mensagem de final

# Chama a função para verificar o status
try:
    verificar_status()
    
    
except Exception as e:
    mensagem = f"Erro na rotina.\n{e}"
    print(mensagem)
    status_script = "Not"
    print("Favor fazer contato com o administrador.")
    send_email("Erro na rotina", mensagem)
    error_response = mensagem


end = time.time()
time_spended = round(end - start,2)
print("Finalizado! ->" + str(time_spended) + "s")
gravar_log_database.gravar_log_database(codigo_script_log, time_spended, status_script, sys.argv, error_response)
