import pyautogui 
import time
import pandas 

pyautogui.PAUSE = 1
# Passo 1: Entrar no sistema da empresa (planilha do Excel no Drive)
pyautogui.press('win')
pyautogui.write('chrome')
pyautogui.press('enter')
pyautogui.write("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pyautogui.press('enter')

time.sleep(5)
# Passo 2: Navegar até o Local do relatório (Entrar na pasta Exportar)
pyautogui.click(x=429, y=477, clicks= 2)
# Passo 3: Exportar o relatório (Fazer download da planilha)
pyautogui.click(x=436, y=482)
pyautogui.click(x=1664, y=241)
pyautogui.click(x=1404, y=787, clicks=2)

time.sleep(5)
# Passo 4: Calcular os indicadores (Faturamento e quantidade de produtos)
import pandas as pd

tabela = pandas.read_excel(r'C:\Users\rafae\Downloads/Vendas - Dez.xlsx')

print(tabela)

faturamento = tabela['Valor Final'].sum()
quantidade = tabela['Quantidade'].sum()

print(faturamento)
print(quantidade)
# Passo 5: Enviar um e-mail para a diretoria
# abrir aba e entrar no gmail
pyautogui.hotkey('ctrl', 't')
pyautogui.write("https://mail.google.com/mail/u/0/?tab=rm&ogbl#inbox")
pyautogui.press('enter')

time.sleep(3)
# clicar no botão escrever
pyautogui.click(x=58, y=252, clicks=1)
# preencher as informações do e-email
pyautogui.click(x=1214, y=445)
pyautogui.write('Taavares.diego@gmail.com')
pyautogui.press('tab')

pyautogui.click(x=1236, y=501)
pyautogui.write('Relatorio de Vendas')
pyautogui.press('tab')

time.sleep(2)
texto = f"""Prezados, boa noite! 
O faturamento de ontem foi de: {faturamento}
A quantidade de produtos vendidos foi de: {quantidade}

Abs, Diego Tavares"""
time.sleep(2)

pyautogui.write(texto)
# enviar o e-mail
pyautogui.hotkey('ctrl', 'enter')
