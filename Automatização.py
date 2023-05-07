import pyautogui
import time
import pyperclip
import pandas
import numpy
import openpyxl

#Baixar o tabela de dados

pyautogui.PAUSE = 0.5

pyautogui.press('win')
pyautogui.write('cr')
pyautogui.press('enter')
pyautogui.write('link para tabela de dados')
pyautogui.press('enter')

time.sleep(3)



pyautogui.press('tab')
pyautogui.write('Nicolas')
pyautogui.press('tab')
pyautogui.write('123')
pyautogui.press('tab')
pyautogui.press('enter')

time.sleep(3)

pyautogui.click(x=396, y=386)
pyautogui.click(x=1713, y=191)
pyautogui.click(x=1465, y=593)

time.sleep(3)

tabela = pandas.read_csv(r"C:\Users\XXXX\Downloads\Compras.csv", sep = ';')
print(tabela)
total_gasto = tabela['ValorFinal'].sum()
quantidade = tabela['Quantidade'].sum()
preco_medio = total_gasto / quantidade
print(total_gasto)
print(quantidade)
print(preco_medio)

#pyautogui.click(x=714, y=1057) 
pyautogui.press('win')
pyautogui.write('out')
pyautogui.press('enter')

time.sleep(15)

#pyautogui.click(x=129, y=107)
pyautogui.hotkey('ctrl', 'n')

time.sleep(3)

pyautogui.write('email-l para envio')
pyautogui.press('tab')
pyautogui.press('tab')
pyperclip.copy('Gestão de pedidos')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab') 

texto = f'''
Prezado chefe, segue abaixo o relatório do dia de hoje:

Total gasto:{total_gasto}
Quantidade:{quantidade}
Preço médio:{preco_medio}

Qualquer dúvida, enviar uma mensagem para esse mesmo e-mail. 
Atenciosamente, Nicolas.
'''
pyperclip.copy(texto)
pyautogui.hotkey('ctrl', 'v')
pyautogui.hotkey('ctrl', 'enter')
