# -*- coding: utf-8 -*-

# INSTRUÇÕES:
# 1 - Verificar acesso ao disco (K:)
# 2 - Abrir o SAP e fazer login
# 3 - Certificar-se de que o usuario tem acesso a transacao FBL3N
# 4 - Nao mexer no computador durante o processo
# 5 - Configurar computador para não entrar em modo de espera


# Módulos
import pyautogui
import datetime
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

# Configurar tempo entre etapas
pyautogui.PAUSE = 0.5

# Transação
fun.transacao()

# Chamar variante
fun.chamar_variante()

# Passo 1 - Emitir Razao Bancos
fun.emitir_bancos(empresa, data_inicio, data_fim)

# Maximizar janela Relatório de partidas individuais contas do Razão
fun.verificar_e_maximizar_janela('Relatório de partidas individuais contas do Razão')

# Salvar Relatorio Bancos
fun.salvar_arquivo(novo_diretorio, '1 Bancos')

# Maximizar e fechar janela 1 Bancos - Excel
fun.verificar_e_maximizar_janela('1 Bancos - Excel')
pyautogui.hotkey('alt', 'f4')

# Maximizar janela Relatório de partidas individuais contas do Razão
fun.verificar_e_maximizar_janela('Relatório de partidas individuais contas do Razão')
pyautogui.hotkey('esc')

# Voltar para tela principal
pyautogui.hotkey('f3')

# Maximizar janela SAP Easy Access
fun.verificar_e_maximizar_janela('SAP Easy Access')

# Transação
fun.transacao()

# Chamar variante
fun.chamar_variante()

# Emitir Contrapartida Bancos
fun.emitir_contrapartida_bancos(empresa, data_inicio, data_fim, novo_diretorio)

# Salvar Contrapartida Bancos
fun.salvar_arquivo(novo_diretorio, '2 Contrapartida Bancos')

# Maximizar e fechar janela 1 Bancos - Excel
fun.verificar_e_maximizar_janela('2 Contrapartida Bancos - Excel')
pyautogui.hotkey('alt', 'f4')

# Check Contrapartida Bancos




pyautogui.alert('FIM')












