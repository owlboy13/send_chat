import os
import random
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
import pandas as pd
import editacodigo_whats
import time
import unicodedata

nome_planilha = input('Qual a planilha? ')
format = '.xlsx'
planilha_format = nome_planilha + format
planilha = pd.read_excel(planilha_format)

# Carregar navegador
driver = editacodigo_whats.carregar_pagina_whatsapp('zap/sessao', 'https://web.whatsapp.com/')
time.sleep(8)  # Aguardar o carregamento inicial

# Função para gerar intervalos aleatórios entre min_seconds e max_seconds
def gerar_intervalo_aleatorio(min_seconds, max_seconds):
    return random.uniform(min_seconds, max_seconds)

# Contador de mensagens
contador_mensagens = 0

for indice, lista in planilha.iterrows():
    nome = lista['NOME']
    telefones = lista['TELEFONE']
    mensagens = lista['MENSAGEM']
    correct = unicodedata.normalize("NFD", telefones)
    
    # Enviar mensagem
    try:
        caixa_msg = driver.find_element(By.XPATH, '//*[@id="side"]/div[1]/div/div[2]/div[2]/div/div[1]/p')
        caixa_msg.send_keys(correct)
        caixa_msg.send_keys(Keys.ENTER)
        mensagem = driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p')
        mensagem.send_keys(f'Olá {nome}, tudo bem?')
        enviar = driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span')
        enviar.click()
        time.sleep(2)
        mensagem = driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p')
        mensagem.send_keys(mensagens)
        enviar = driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span')
        enviar.click()

        contador_mensagens += 1
        
        # Intervalo aleatório entre 5 e 10 segundos
        time.sleep(gerar_intervalo_aleatorio(5, 10))
        
        # Fazer pausa de 2 a 5 minutos após cada 40 mensagens
        if contador_mensagens % 40 == 0:
            print("Pausa de 2 a 5 minutos após 40 mensagens...")
            time.sleep(gerar_intervalo_aleatorio(120, 300))

    except (NoSuchElementException, StaleElementReferenceException) as e:
        print(f"Erro ao enviar mensagem para {nome}: {e}")
        continue

input()
