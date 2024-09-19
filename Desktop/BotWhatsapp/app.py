#INSTALAR BIBLIOTECA PILLOW
import openpyxl #Biblioteca para gerenciamento de planilha / precisa de instalação
from urllib.parse import quote #Biblioteca para formatação de mensagem
import webbrowser #Biblioteca para abrir o navegador
from time import sleep #Biblioteca para o cronometro
import pyautogui #Biblioteca para movimentar o mouse e clicar na tela / precisa de instalação
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service

profile_path = r"C:\Users\User\AppData\Local\Google\Chrome\User Data\Default"
options = Options()
options.add_argument(f"user-data-dir={profile_path}")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
driver.get('https://web.whatsapp.com/')
sleep(20)

planilha = openpyxl.load_workbook('numeros.xlsx') #Carrega a Planilha de dados
pagina_clientes = planilha['Sheet1'] #Indica a Pagina da Planilha 

cont = 0

for linha in pagina_clientes.iter_rows(min_row=2): #Roda todas as linhas a partir da segunda
    nome = linha[0].value #coluna Nome
    if nome == None:
        continue
    
    telefone = linha[1].value #coluna Telefone
    aniversario = linha[2].value #coluna Data
    cpf = linha[3].value #coluna CPF
    mensagem = f'''Olá {nome}, tudo bem? ☕

Sou correspondente bancário do FGTS. 😁
Estou entrando em contato para realizarmos uma consulta de antecipação no cpf: {str(cpf)} e te ajudar no processo para sacar, caso tenha interesse! 🥳

Responda uma das opções:
1 - Estou interessado(a).
2 - Não estou interessado(a).

> Tenha um ótimo dia!''' #Mensagem que será mandada
    
    try: #Executa uma tentativa de encontrar erros
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}' #link para envio de mensagem
        driver.get(link_mensagem_whatsapp) #executa o navegador direto no contato
        wait = WebDriverWait(driver, 10) #
        botao = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@aria-label="Enviar"]')))
        # Clicar no botão
        botao.click()
        sleep(5)
        cont = cont + 1
        if cont == 6:
            sleep(600)
            cont = 0
    except: #Gera uma mensagem de erro caso não funcione
        print(f'Não foi possível enviar mensagem para {nome}')
        with open('erros.csv','a',newline='',encoding='utf-8') as arquivo: #joga o nome e o numero do cliente que não foi possivel enviar a mensagem para uma outra planilha
            arquivo.write(f'{nome},{telefone}')


