# -*- coding: utf-8 -*-


# Módulos
import pyperclip
import pyautogui
import time
import fun



# mensagem de inicio
pyautogui.alert('Realização do Fluxo Caixa - Celesc G')


# Dados de entrada
# diretorio = 'K:\\DPCL\\_Comum\\Realização do Fluxo de Caixa\\Celesc G\\'
# empresa = 'CE03'
# data_ini, data_fim, mes, ano = fun.obter_data_competencia()
# login, senha = fun.obter_dados_sap()
# nome_pasta = str(ano) + '.' + str(mes)


# Teste
diretorio = 'K:\\DPCL\\_Comum\\Realização do Fluxo de Caixa\\Celesc G\\'
empresa = 'CE03'
login = 'e018266'
senha = 'EaD#5514'
data_ini = '01.12.2023'
data_fim = '31.12.2023'
mes = 12
ano = 2023
nome_pasta = str(ano) + '.' + str(mes)


# Abrir SAP e fazer login
pyautogui.press('win')
time.sleep(1)
pyautogui.write('SAP Logon')
pyautogui.press('enter')
time.sleep(4)
pyautogui.hotkey('win','up')
time.sleep(1)
pyautogui.press('enter')
time.sleep(1)
pyautogui.write(login)
pyautogui.hotkey('down')
time.sleep(1)
pyautogui.write(senha)
pyautogui.press('enter')
time.sleep(3)

# Transacao
pyautogui.write('FBL3N')
pyautogui.press('enter')
time.sleep(1)


### Passo 1) Emitir razao dos bancos ### 
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('enter')
time.sleep(0.5)
pyautogui.write('101*')
pyautogui.press('down')
pyautogui.write('103*')
pyautogui.press('down')
pyautogui.write('111*')
pyautogui.hotkey('f8')
time.sleep(0.5)
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.write(empresa)
pyautogui.press('enter')
time.sleep(0.5)
pyautogui.press('down')
pyautogui.press('down')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('down')
pyautogui.press('tab')
pyautogui.write(data_ini)
pyautogui.press('tab')
pyautogui.write(data_fim)
pyautogui.hotkey('f8')
time.sleep(5)


# Salvar relatorio

pyautogui.hotkey('alt')
pyautogui.write('x')
pyautogui.write('a')
time.sleep(1)
pyautogui.write('1 Bancos')
# pyautogui.hotkey('esc')
time.sleep(1)
for _ in range(6): pyautogui.hotkey('tab')
pyautogui.hotkey('enter')
nova_pasta = fun.criar_diretorio(diretorio, nome_pasta)
pyperclip.copy(nova_pasta)
pyautogui.hotkey("ctrl", "v")
time.sleep(1)
pyautogui.hotkey('enter')
for _ in range(9): pyautogui.hotkey('tab')
pyautogui.hotkey('enter')
time.sleep(3)
for _ in range(4): pyautogui.hotkey('tab')
time.sleep(1)
pyautogui.hotkey('enter')




pyautogui.alert('FIM')


