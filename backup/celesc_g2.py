# -*- coding: utf-8 -*-

# INSTRUÇÕES:
# 1 - Verificar acesso ao disco (K:)
# 2 - Abrir o SAP e fazer login
# 3 - Certificar-se de que o usuario tem acesso a transacao FBL3N
# 4 - Nao mexer no computador durante o processo
# 5 - Configurar computador para não entrar em modo de espera




# Módulos
import pyperclip
import pyautogui
import datetime
import time
import fun




# Mensagem de início
pyautogui.alert('Realização do Fluxo Caixa - Celesc G')




# Dados de entrada - Celesc G
diretorio = 'K:\\DPCL\\_Comum\\Realização do Fluxo de Caixa\\Celesc G\\'
empresa = 'CE03'




# Definir variaveis
data = datetime.datetime.today()
# ano = data.year
# mes = data.month
ano = 2023
mes = 12
data_inicio, data_fim = fun.calcular_datas(mes, ano)
pasta = str(ano) + '.' + str(mes)
novo_diretorio = fun.criar_diretorio(diretorio, pasta)




# Maximizar janela SAP 
fun.verificar_e_maximizar_janela('SAP Easy Access')



# Configurar tempo entre ações
pyautogui.PAUSE = 1




# Transação
time.sleep(3)
pyautogui.write('FBL3N')
pyautogui.press('enter')




# Chamar variante
pyautogui.hotkey('shift', 'f5')
pyautogui.write('/FL DIARIO')
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.hotkey('backspace')
pyautogui.write('e016321')
pyautogui.hotkey('f8')




# Passo 1 - Emitir Razao Bancos
pyautogui.hotkey('shift', 'f4')
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.hotkey('enter')
pyautogui.write('101*')
pyautogui.press('down')
pyautogui.write('103*')
pyautogui.press('down')
pyautogui.write('111*')
pyautogui.hotkey('f8')
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.hotkey('backspace')
pyautogui.write(empresa)
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.write(data_inicio)
pyautogui.hotkey('tab')
pyautogui.write(data_fim)
pyautogui.press('f8')
time.sleep(5)




# Maximizar janela Relatório de partidas individuais contas do Razão
fun.verificar_e_maximizar_janela('Relatório de partidas individuais contas do Razão')




# Salvar Relatorio Bancos
pyautogui.hotkey('alt')
pyautogui.write('x')
pyautogui.write('a')
time.sleep(5)
pyautogui.write('1 Bancos')
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.hotkey('enter')
pyperclip.copy(novo_diretorio)
pyautogui.hotkey('ctrl', 'v')
pyautogui.hotkey('enter')
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.hotkey('enter')
time.sleep(3)
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.hotkey('enter')
time.sleep(5)





# Maximizar e fechar janela 1 Bancos - Excel
fun.verificar_e_maximizar_janela('1 Bancos - Excel')
pyautogui.hotkey('alt', 'f4')




# Maximizar janela Relatório de partidas individuais contas do Razão
fun.verificar_e_maximizar_janela('Relatório de partidas individuais contas do Razão')
pyautogui.hotkey('esc')





# Emitir Contrapartida Bancos
pyautogui.press('win')
pyautogui.write('bloco de notas')
pyautogui.press('enter')
pyautogui.hotkey('win','up')
fun.contrapartida_bancos(novo_diretorio)
pyautogui.hotkey('ctrl', 'v')
pyautogui.hotkey('ctrl', 'a')
pyautogui.hotkey('ctrl', 'c')   # copiar para área de transferência
pyautogui.hotkey('alt', 'f4')
pyautogui.hotkey('tab')
pyautogui.hotkey('enter')
fun.verificar_e_maximizar_janela('Relatório de partidas individuais contas do Razão')
pyautogui.hotkey('shift', 'f4')
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.hotkey('enter')
pyautogui.hotkey('shift', 'f4')
pyautogui.hotkey('shift', 'tab')
pyautogui.hotkey('shift', 'tab')
pyautogui.hotkey('shift', 'tab')
pyautogui.hotkey('right')
pyautogui.hotkey('right')
pyautogui.hotkey('enter')
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.hotkey('tab')
pyautogui.write('101*')
pyautogui.press('down')
pyautogui.write('103*')
pyautogui.press('down')
pyautogui.write('111*')
pyautogui.hotkey('f8')
pyautogui.hotkey('shift', 'tab')
pyautogui.hotkey('shift', 'tab')
pyautogui.hotkey('enter')
pyautogui.hotkey('shift', 'f12')
pyautogui.hotkey('f8')
pyautogui.hotkey('f8')






# Salvar Contrapartida Bancos



pyautogui.alert('FIM')












