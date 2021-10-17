from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import pyautogui
from time import sleep
from pyperclip import copy
pyautogui.PAUSE = 2
navegador = webdriver.Chrome()
navegador.get('https://ri.magazineluiza.com.br')
navegador.find_element(By.XPATH, '/html/body/form/div[11]/div/div/div[1]/div/div[4]/div/a/img').click()
navegador.find_element(By.ID, 'bWtQ7n6RcQdDDDCgCcH3yg==').click()
sleep(7)
tabela_de_vendas = pd.read_excel(r'C:/Users/User/Downloads/RESULTADO_2T21_POR.xlsx', sheet_name='12. Vendas e Lojas por Canal')
tabela_de_vendas = tabela_de_vendas.dropna(axis=1, how='all')
tabela_de_vendas = tabela_de_vendas.dropna(axis=0)
tabela_de_vendas["soma"] = tabela_de_vendas.sum(axis=1)
soma_vendas_lojas = tabela_de_vendas.loc[7, 'soma']
print(soma_vendas_lojas)
navegador.quit()
pyautogui.click(24,734)
copy('Chrome')
pyautogui.hotkey('ctrl', 'v')
pyautogui.keyDown('enter')
sleep(2)
copy('https://www.google.com/gmail/')
pyautogui.hotkey('ctrl', 'v')
pyautogui.keyDown('enter')
sleep(3)
pyautogui.click(98,171)
copy('universenerd249+diretoria@gmail.com')
pyautogui.hotkey('ctrl', 'v')
pyautogui.keyDown('tab')
# pyautogui.keyDown('tab')
copy('Vendas Totais')
pyautogui.hotkey('ctrl', 'v')
pyautogui.keyDown('tab')
msg = f'''Prezado, Lucas,
        As vendas dos últimos anos resultaram em: {soma_vendas_lojas}
        Qualquer dúvida entrar em contato
Att. Lucas Feitosa'''
copy(msg)
pyautogui.hotkey('ctrl', 'v')
pyautogui.keyDown('tab')
pyautogui.keyDown('enter')
