#!/usr/bin/env python
# coding: utf-8

# In[21]:


# Importa as bibliotecas do selenium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

# Importa a biblioteca de data
from datetime import datetime, timedelta

#Importa as bibliotecas do pyautogui
import pyautogui as tempo
import pyautogui as teclas

#By para trabalhar com os computadores mais recentes
from selenium.webdriver.common.by import By

#Importa a biblioteca openpyxl para trabalhar com excel
from openpyxl import load_workbook

# Pegamos o caminho do arquivo na pasta do computador
nome_arquivo_contatos = "C:\\Users\\lemes\\OneDrive\\Área de Trabalho\\Projeto\\Contatos.xlsx"
planilhaContatos = load_workbook(nome_arquivo_contatos)


# Selecionando a planilha "Dados"
sheet_selecionada = planilhaContatos['Dados']

# Emulando o navegador
navegador = webdriver.Chrome()

#Abrindo o navegador com o link do whats
navegador.get("https://web.whatsapp.com/")

while len(navegador.find_elements(By.ID,"side")) < 1:
    tempo.sleep(3)
    
tempo.sleep(20)



for linha in range(2, len(sheet_selecionada["A"])+1): 
    nomeContato = sheet_selecionada["A%s" % linha].value
    sobrenomeContato = sheet_selecionada["B%s" % linha].value
    responsavel = sheet_selecionada["C%s" % linha].value
    numeroContato = sheet_selecionada["D%s" % linha].value
    data_da_consulta = sheet_selecionada["E%s" % linha].value
    if isinstance(data_da_consulta, str):
        data_da_consulta = datetime.strptime(data_da_consulta, '%d/%m/%Y')
    amanha = datetime.now() + timedelta(days=1)
    amanha = datetime(amanha.year, amanha.month, amanha.day)
    horario_consulta = sheet_selecionada["F%s" % linha].value

    if data_da_consulta == amanha:
        if responsavel == nomeContato:
            mensagemContato = f"{nomeContato}, você tem uma consulta *amanhã*, dia *{data_da_consulta.strftime('%d/%m/%Y')}* às *{horario_consulta}* com a Fonoaudióloga Francine. Por favor, caso precise remarcar, entre em contato. Obrigada."
        else:
            mensagemContato = f"{responsavel},  o paciente {nomeContato} possui uma consulta *amanhã*, dia *{data_da_consulta.strftime('%d/%m/%Y')}* às *{horario_consulta}* com a Fonoaudióloga Francine. Por favor, caso precise remarcar, entre em contato. Obrigada."
        
        # Busca pelo elemento XPATH e escreve o nomeContato no campo de pesquisa
        navegador.find_element(By.XPATH, '//*[@id="side"]/div[1]/div/div/div[2]/div/div[2]').send_keys(numeroContato)
        tempo.sleep(3)
        # Pressiona enter
        teclas.press("enter")
        tempo.sleep(3)
        # Envia a mensagem
        caixa_mensagem = navegador.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]')
        caixa_mensagem.send_keys(mensagemContato)
        caixa_mensagem.send_keys(Keys.RETURN)
        tempo.sleep(3)


# In[ ]:





# In[ ]:




