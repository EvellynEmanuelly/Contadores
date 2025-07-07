import time
import os
from selenium.webdriver.common.by import By
from selenium import webdriver
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
# import xlsxwriter
# from retrying import retry
#DB----------------------------------------------------------
#Etapa de criação e conexão com Banco de Dados:
import sqlite3

import xlsxwriter #Importando a biblioteca.

conn = sqlite3.connect('contadores.db')#conectando ao DB local.
cursor = conn.cursor()#Executar comandos

#Comando SQL para criar tabela:
criar_tabela = """
CREATE TABLE IF NOT EXISTS Contadores(
    id INTEGER PRIMARY KEY,
    nSerie TEXT NOT NULL,
    contador INTEGER
);
"""
cursor.execute(criar_tabela)#Executando comando SQL que está dentro da variável.
conn.commit()#garantir que as alterações foram aplicadas.
#------------------------------------------------------------

# Configuração do WebDriver
#servico = Service(GeckoDriverManager().install())
#navegador = webdriver.Firefox(service=servico)
navegador = webdriver.Firefox()


def xpath_to_filename(xpath): 
    return ''.join(c if c.isalnum() else '_' for c in xpath)


#Função para ignorar Alerta de Certificado inválido RICOH.
def verifCerti():
    if "Alerta:" in navegador.page_source:
        print("Ignorando alerta de certificado.")
        navegador.find_element(by=By.ID, value="advancedButton").click()
        time.sleep(1)
        navegador.find_element(by=By.ID, value="exceptionDialogButton").click()
        time.sleep(3)
        

#Função para coletar os dados da impressora Ricoh para inserir no Banco de Dados:
def ricoh():
    print("< Ricoh [Inserindo Infos em DB] >")
    click1 = navegador.find_element(by=By.LINK_TEXT, value="Machine Information")
    click1.click()
    time.sleep(1)
    nSerie = navegador.find_element(by=By.XPATH, value="/html/body/form/div/table/tbody/tr/td[2]/table[2]/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[4]/td[3]").text
    time.sleep(2)
    click2 = navegador.find_element(by=By.LINK_TEXT, value="Counter")
    click2.click()
    numerooo = navegador.find_element(by=By.XPATH, value="/html/body/div/table/tbody/tr/td[2]/table[2]/tbody/tr[2]/td/table/tbody/tr/td[2]/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[2]/td[3]").text
    #Inserindo número de série e contador no banco de dados:
    nContador = (nSerie, numerooo)
    inserir_contador = "INSERT INTO Contadores (nSerie, contador) VALUES (?, ?)"
    cursor.execute(inserir_contador, nContador)
    conn.commit()

#Função para gerar print do navegador:
def gerar_print():
    linksite = navegador.find_element(by=By.XPATH, value="/html/body/p").text
    #link = navegador.get(link)
    print("< Gerando print... >")
    time.sleep(2)
    textoserial = linksite.split('/')[-1]
    screenshot_filename = f'{xpath_to_filename(textoserial)}.png'
    pasta_cont = 'Contadores WEB'
    screen_caminho = os.path.join(pasta_cont, screenshot_filename)
    navegador.save_screenshot(screen_caminho)
    print(f"< Print {textoserial} gerada com sucesso >")
    print("-----------------------------------------------------")

#Função para coletar dados da impressora Lexmark:
def lexmark():
    if navegador.find_element(by=By.TAG_NAME, value="center").text.split()[0] == "Lexmark":
        print("< Lexmark [Inserindo Infos em DB] >")
        nSerie = navegador.find_element(by=By.TAG_NAME, value="table").text.split()[27]
        numerooo = navegador.find_element(by=By.TAG_NAME, value="table").text.split()[6]
        nContador = (nSerie, numerooo)
        inserir_contador = "INSERT INTO Contadores (nSerie, contador) VALUES (?, ?)"
        cursor.execute(inserir_contador, nContador)
        conn.commit()
    else:
        print("< Lexmark: falha ao inserir Infos em DB >")


#Função que irá abrir cada link contido na lista.
def abrir_links_um_por_vez(links):
    for link in links:
        try:
            navegador.get(link) #Abrindo link no navegador.
            print("-----------------------------------------------------")
            print("Abrindo link>>>>>>")
            time.sleep(5)

    
            #Tentando achar o elemento no site para ativar a condição Samsung.
            if "Status trb" in navegador.page_source: 
                print(f"< Samsung detectada: {link} >")
                infoI = navegador.find_element(by=By.ID, value="ext-gen249")
                infoI.click()
                time.sleep(2)
                contadorNumero = navegador.find_element(by=By.LINK_TEXT, value="Contadores de uso")
                contadorNumero.click()
                time.sleep(2)    
                #DB-------------------------------------------------------------------------------------------------------
                #Abaixo uma condição que irá verificar se é uma impressora Mult ou Mono.
                if navegador.find_element(by=By.ID, value="ext-comp-1055").text.split()[4] == "fax": #Mult
                    print("< Samsung Mult [Inserindo Infos em DB] >")
                    #localizando Número de série no site:
                    nSerie = navegador.find_element(by=By.CLASS_NAME, value="x-form-element").text
                    #localizando Contador no site: 
                    numerooo = navegador.find_element(by=By.CLASS_NAME, value=("x-grid3-row.x-grid3-row-last ")).text.split()[7]
                    #Inserindo número de série e contador no banco de dados:
                    nContador = (nSerie, numerooo)
                    inserir_contador = "INSERT INTO Contadores (nSerie, contador) VALUES (?, ?)"
                    cursor.execute(inserir_contador, nContador)
                    conn.commit()
                else: #Mono
                    print("< Samsung Mono [Inserindo Infos em DB] >")
                    #localizando Número de série no site:
                    nSerie = navegador.find_element(by=By.CLASS_NAME, value="x-form-element").text
                    #localizando Contador no site: 
                    numerooo = navegador.find_element(by=By.CLASS_NAME, value=("x-grid3-row.x-grid3-row-last ")).text.split()[5]
                    #Inserindo número de série e contador no banco de dados:
                    nContador = (nSerie, numerooo)
                    inserir_contador = "INSERT INTO Contadores (nSerie, contador) VALUES (?, ?)"
                    cursor.execute(inserir_contador, nContador)
                    conn.commit()
                    time.sleep(2)                    
                #DB-------------------------------------------------------------------------------------------------------
                
                # Gerar print para Samsung
                print("< Gerando print... >")
                textoserial = link.split('/')[-1]
                screenshot_filename = f'{xpath_to_filename(textoserial)}.png'
                pasta_cont = 'Contadores WEB'
                screen_caminho = os.path.join(pasta_cont, screenshot_filename)
                navegador.save_screenshot(screen_caminho)
                print(f"< Print SAMSUNG gerada com sucesso do link: {link} >")
                print("-----------------------------------------------------")

            
            #Tentando achar o elemento no site para ativar a condição Brother.
            elif "Copyright(C) 2000-2015 Brother Industries, Ltd. All Rights Reserved." in navegador.page_source: 
                print(f"< Brother detectada: {link}>")
                time.sleep(1)
                print("< Fazendo Login... >")
                pegarElement = navegador.find_element(by=By.ID, value="LogBox")
                time.sleep(1)
                pegarElement.click()
                pegarElement.send_keys("initpass")
                pegarElement.send_keys(Keys.ENTER)
                time.sleep(2)
                maintenenceText = navegador.find_element(by=By.LINK_TEXT, value="Maintenance Information")
                maintenenceText.click()
                time.sleep(2)

                #DB-------------------------------------------------------------------------------------------------------
                #Coletando as informações para inserir no Banco de Dados:
                print("< Brother [Inserindo Infos em DB] >")
                #localizando Número de série no site:
                nSerie = navegador.find_element(by=By.ID, value="pageContents").text.split()[8]
                #localizando Contador no site: 
                numerooo = navegador.find_element(by=By.ID, value="pageContents").text.split()[24]
                #Inserindo número de série e contador no banco de dados:
                nContador = (nSerie, numerooo)
                inserir_contador = "INSERT INTO Contadores (nSerie, contador) VALUES (?, ?)"
                cursor.execute(inserir_contador, nContador)
                conn.commit()
                #DB-------------------------------------------------------------------------------------------------------

                # Gerar print para Brother.
                print("< Gerando print... >")
                time.sleep(2)
                textoserial = link.split('/')[-1]
                screenshot_filename = f'{xpath_to_filename(textoserial)}.png'
                pasta_cont = 'Contadores WEB'
                screen_caminho = os.path.join(pasta_cont, screenshot_filename)
                navegador.save_screenshot(screen_caminho)
                print(f"< Print BROTHER gerada com sucesso do link: {link} >")
                print("-----------------------------------------------------")

            
            #LEX / RICOH----------------------------------------------------------------------------
            elif link == 'http://192.168.23.132': # Com base no link, direciona para a pagina da Lexmark.
                url = "http://192.168.23.132/cgi-bin/dynamic/printer/config/reports/deviceinfo.html"
                navegador.get(url)
                print(f"< Lermark detectada: {link} >")
                lexmark()#Chamando a função que insere as informações no Banco de Dados.

                #Gerando a print da impressora:
                print("< Gerando print... >")
                time.sleep(2)
                textoserial = link.split('/')[-1]
                screenshot_filename = f'{xpath_to_filename(textoserial)}.png'
                pasta_cont = 'Contadores WEB'
                screen_caminho = os.path.join(pasta_cont, screenshot_filename)
                navegador.save_screenshot(screen_caminho)
                print(f"< Print {textoserial} gerada com sucesso >")
                print("-----------------------------------------------------")

            elif link == 'http://192.168.23.101': #Com base no link é redirecionado para um html com as infos da impressora.
                print(f"< Ricoh detectada: {link} >")
                verifCerti()
                ricoh()
                url = "file:///C:/Users/evellyn.queiroz/Desktop/new/23-101.html/"
                navegador.get(url)
                time.sleep(2)
                gerar_print()

            elif link == 'http://192.168.23.119': #Com base no link é redirecionado para um html com as infos da impressora.
                print(f"< Ricoh detectada: {link} >")
                verifCerti()
                ricoh()
                url = "file:///C:/Users/evellyn.queiroz/Desktop/new/23-119.html/"
                navegador.get(url)
                time.sleep(2)
                gerar_print()

            elif link == 'http://192.168.23.107': #Com base no link é redirecionado para um html com as infos da impressora.
                print(f"< Ricoh detectada: {link} >")
                verifCerti()
                ricoh()
                url = "file:///C:/Users/evellyn.queiroz/Desktop/new/23-107.html/"
                navegador.get(url)
                time.sleep(2)
                gerar_print() 

            elif link == 'http://192.168.23.110': #Com base no link é redirecionado para um html com as infos da impressora.
                print(f"< Ricoh detectada: {link} >")
                verifCerti()
                ricoh()
                url = "file:///C:/Users/evellyn.queiroz/Desktop/new/23-110.html/"
                navegador.get(url)
                time.sleep(2)
                gerar_print()

            elif link == 'http://192.168.23.122': #Com base no link é redirecionado para um html com as infos da impressora.
                print(f"< Ricoh detectada: {link} >")
                verifCerti()
                ricoh()
                url = "file:///C:/Users/evellyn.queiroz/Desktop/new/23-122.html/"
                navegador.get(url)
                time.sleep(2)
                gerar_print()                                  

            elif link == 'http://192.168.23.176': #Com base no link é redirecionado para um html com as infos da impressora.
                print(f"< Ricoh detectada: {link} >")
                verifCerti()
                ricoh()
                url = "file:///C:/Users/evellyn.queiroz/Desktop/new/23-176.html/"
                navegador.get(url)
                time.sleep(2)
                gerar_print()     

            elif link == 'http://192.168.23.177': #Com base no link é redirecionado para um html com as infos da impressora.
                print(f"< Ricoh detectada: {link} >")
                verifCerti()
                ricoh()
                url = "file:///C:/Users/evellyn.queiroz/Desktop/new/23-177.html/"
                navegador.get(url)
                time.sleep(2)
                gerar_print()

            elif link == 'http://192.168.23.151': #Com base no link é redirecionado para um html com as infos da impressora.
                print(f"< Ricoh detectada: {link} >")
                verifCerti()
                ricoh()
                url = "file:///C:/Users/evellyn.queiroz/Desktop/new/23-151.html/"
                navegador.get(url)
                time.sleep(2)
                gerar_print()

            else:
                print(f"< Não foi possível identificar o tipo de impressora para o link: {link} >")
        except Exception as e:

            #Tratativa de erros:
            print(f"Ocorreu um erro: {str(e)}")
            print("-----------------------------------------------------")
            print(f"NÃO FOI POSSÍVEL TIRAR O CONTADOR DO LINK: {link}. Por favor, verificar.")
            print("-----------------------------------------------------")

#Função para tranformar as informações do Bando de Dados em uma planilha.
def planilha():
    #XLSX--------------------------------------------------------
    print("< Gerando planilha! >")
    book = xlsxwriter.Workbook("Contadores.xlsx") #Definindo o nome da planilha.
    sheet = book.add_worksheet("dataimported")

    path = 'contadores.db' #Caminho do Bando de Dados.
    db = sqlite3.connect(path) #Conectando ao banco de dados.
    cursor.execute('''SELECT * FROM Contadores''') #Selecionando os elementos da tabela Contadores.
    all_rows = cursor.fetchall() #Reunindo todos os resultados.

    row = 0
    col = 0

    column_Values = ['ID','Numero de Série','Contador'] #Definindo as colunas que terão na planilha.

    for heading in column_Values:
        sheet.write(row,col,heading)
        col+=1

    for entry in all_rows:
        row += 1
        col = 0
        for data_val in entry:
            sheet.write(row,col,data_val)
            col += 1
    book.close()
    cursor.close()
    db.commit()
    db.close()
    print("< Planilha gerada com sucesso! >")
    #XLSX--------------------------------------------------------

# Nome do arquivo contendo os links
nome_arquivo = 'links.txt'
try:
    with open(nome_arquivo, 'r') as arquivo:
        links = [linha.strip() for linha in arquivo if linha.strip()]

    if links:
        abrir_links_um_por_vez(links) #Executando a função.
        planilha() #Executando a função.

    else:
        print(f'O arquivo {nome_arquivo} não contém links válidos ou está vazio.') 
except FileNotFoundError:
    #Tratativa de erro:
    print(f'O arquivo {nome_arquivo} não foi encontrado.')
except Exception as e:
    #Tratativa de erro:
    print(f'Ocorreu um erro ao tentar ler o arquivo: {str(e)}')

time.sleep(1)
print("-----------------------------------------------------")
print("< Finalizando código... >")
time.sleep(2)
print("< Fechando Navegador. >")
print("-----------------------------------------------------")
navegador.quit()
#Final do código.
