# -*- coding: utf-8 -*-

# OBS: 
# Verificar acesso ao (K:)
# Fechar todas as janelas do SAP
# Nao mexer no computador enquanto durante o processo
# Certificar-se de que o usuario tem acesso a transacao FBL3N


# Módulos
import pyperclip
import pyautogui
import sys
import datetime
import time
import fun


# Mensagem de início
pyautogui.alert('Realização do Fluxo Caixa - Celesc G')


# Dados de entrada
diretorio = 'K:\\DPCL\\_Comum\\Realização do Fluxo de Caixa\\Celesc G\\'
empresa = "CE03"
data = datetime.datetime.today()
# ano = data.year
# mes = data.month
ano = 2023
mes = 12
# login, senha = fun.obter_dados_sap()
login, senha = 'e018266', 'EaD#5514'


# Definir variaveis
data_inicio, data_fim = fun.calcular_datas(mes, ano)
pasta = str(ano) + '.' + str(mes)


# Acessar o SAP e fazer login
pyautogui.PAUSE = 1
pyautogui.press('win')
pyautogui.write('SAP Logon')
pyautogui.press('enter')
time.sleep(3)
fun.verificar_e_maximizar_janela('SAP Logon')
# pyautogui.hotkey('win','up')
pyautogui.press('enter')
time.sleep(3)
pyautogui.write(login)
pyautogui.hotkey('down')
pyautogui.write(senha)
pyautogui.press('enter')
time.sleep(3)


# Inserir transação
pyautogui.write('FBL3N')
pyautogui.press('enter')


# Chamar variante
pyautogui.PAUSE = 1 
pyautogui.hotkey('shift', 'f5')
for _ in range(2): pyautogui.hotkey('down')
pyautogui.PAUSE = 0
for _ in range(7): pyautogui.hotkey('del')
pyautogui.PAUSE = 0.5
pyautogui.write('e016321')
pyautogui.hotkey('f8')
pyautogui.hotkey('left')
pyautogui.hotkey('f2')


# Passo 1 - Emitir razao bancos
pyautogui.hotkey('down')
pyautogui.PAUSE = 0
for _ in range(4): pyautogui.hotkey('del')
pyautogui.PAUSE = 1
pyautogui.write(empresa)
for _ in range(12): pyautogui.hotkey('tab')
pyautogui.write(data_inicio)
pyautogui.hotkey('tab')
pyautogui.write(data_fim)
pyautogui.press('f8')
time.sleep(5)


# Salvar relatorio
pyautogui.hotkey('alt')
pyautogui.write('x')
pyautogui.write('a')
pyautogui.write('1 Bancos')
for _ in range(6): pyautogui.hotkey('tab')
pyautogui.hotkey('enter')
novo_diretorio = fun.criar_diretorio(diretorio, pasta)
pyperclip.copy(novo_diretorio)
pyautogui.hotkey("ctrl", "v")
pyautogui.hotkey('enter')
for _ in range(9): pyautogui.hotkey('tab')
pyautogui.hotkey('enter')
time.sleep(3)
for _ in range(4): pyautogui.hotkey('tab')
time.sleep(1)
pyautogui.hotkey('enter')


pyautogui.alert('FIM')





