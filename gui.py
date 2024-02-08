






### Criar novo diretório para salvar os arquivos ###
def novo_diretorio(diretorio, nome_pasta):
    
    import os
    
    # Obtém o nome base da pasta sem a extensão (caso seja um diretório)
    nome_base, extensao = os.path.splitext(nome_pasta)
    
    # Inicializa o contador de sequência
    sequencia = 1
    
    # Enquanto existir uma pasta com o mesmo nome, incrementa a sequência
    while os.path.exists(os.path.join(diretorio, f'{nome_base}{extensao} ({sequencia})')):
        sequencia += 1
    
    # Cria a nova pasta com a sequência no nome
    novo_nome_pasta = f'{nome_base}{extensao} ({sequencia})'
    novo_caminho_pasta = os.path.join(diretorio, novo_nome_pasta)
    os.makedirs(novo_caminho_pasta)
    
    return novo_caminho_pasta
            


















### Data de competência ###
def obter_data_competencia():
    
    import pyautogui
    import datetime
    from tkinter import simpledialog
    
    while True:
        
        try:
            
            # Solicitar ao usuário o mês e ano de competência
            mes = simpledialog.askstring('Competência', 'Digite o mês de competência:                      ') 
            ano = simpledialog.askstring('Competência', 'Digite o ano de competência:                      ') 
            mes = int(mes); ano = int(ano)

             # Verificar se o mês está no intervalo válido
            if 1 <= mes <= 12:
                
                # Verificar se o ano está no intervalo válido
                ano_atual = datetime.datetime.now().year
                
                if 2016 <= ano <= ano_atual:
                    
                    # Construir a data de início do mês
                    data_ini = datetime.datetime(ano, mes, 1)
                    
                    # Calcular o último dia do mês
                    ultimo_dia_mes = (datetime.datetime(ano, mes % 12 + 1, 1) - datetime.timedelta(days = 1)).day
                    
                    # Construir a data de fim do mês
                    data_fim = datetime.datetime(ano, mes, ultimo_dia_mes)
                    
                    # Formatar data
                    data_ini = data_ini.strftime('%d.%m.%Y')
                    data_fim = data_fim.strftime('%d.%m.%Y')
                    
                    return (data_ini, data_fim, mes, ano)
                else:
                    pyautogui.alert('Ano inválido. Por favor, insira um ano anterior ou igual ao ano atual.')
            else:
                pyautogui.alert('Mês inválido. Por favor, insira um número de mês entre 1 e 12.')
        
        except ValueError:
            pyautogui.alert('Entrada inválida. Certifique-se de inserir números inteiros para mês e ano.')














### Dados para login SAP ###
def obter_dados_sap():
    
    import pyautogui
    from tkinter import simpledialog
    
    while True:
        
        try:
            
            # Solicitar ao usuário o dados para login do SAP
            login = simpledialog.askstring('Login', 'Digite seu Usuário SAP          ')
            senha = simpledialog.askstring('Login', 'Digite sua Senha SAP            ')

            # Verificar login
            if len(login) < 5 or login == '':
                pyautogui.alert('Login inválido.')
            else:
                return (login, senha)
        
        except ValueError:
            pyautogui.alert('Entrada inválida. Certifique-se de inserir dados corretos.')



















### Salvar Arquivo ###
def salvar_arquivo(novo_diretorio, tempo_inicial, tempo_fim, nome_pasta):
    
    import pyautogui
    import pyperclip
    import time
    import gui
    
    # Maximizar janela Relatório de partidas individuais contas do Razão
    gui.verificar_e_maximizar_janela('Relatório de partidas individuais contas do Razão')
    
    time.sleep(2)
    pyautogui.hotkey('enter')   # ativar tela
    pyautogui.hotkey('alt')
    pyautogui.write('x')
    pyautogui.write('a')
    time.sleep(5)               # tempo para abrir janela de salvar
    pyautogui.write(nome_pasta)
    pyautogui.PAUSE = tempo_fim
    for _ in range(6): pyautogui.hotkey('tab')
    pyautogui.hotkey('enter')
    pyperclip.copy(novo_diretorio)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.hotkey('enter')
    for _ in range(9): pyautogui.hotkey('tab')
    pyautogui.PAUSE = tempo_inicial
    pyautogui.hotkey('enter')
    time.sleep(3)               # tempo para aparecer mensagem de confirmação
    pyautogui.hotkey('shift', 'tab')
    pyautogui.hotkey('enter')
    time.sleep(10)               # tempo para arquivo excel abrir
    
    # Maximizar e fechar janela 1 Bancos - Excel
    gui.verificar_e_maximizar_janela(nome_pasta)
    pyautogui.hotkey('alt', 'f4')
    
    # Voltar para tela principal
    gui.verificar_e_maximizar_janela('Relatório de partidas individuais contas do Razão')
    pyautogui.hotkey('f3')
    gui.verificar_e_maximizar_janela('Relatório de partidas individuais contas do Razão')
    pyautogui.hotkey('f3')
    
    print('Salvar Arquivo OK!')
    














    
    
def copiar_dados(dados):
    
    import pyautogui
    import pyperclip
    import gui

    pyautogui.press('win')
    pyautogui.write('bloco de notas')
    pyautogui.press('enter')
    pyautogui.hotkey('win','up')
    gui.verificar_e_maximizar_janela('Sem título - Bloco de Notas')
    pyperclip.copy(dados)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')
    pyautogui.hotkey('alt', 'f4')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('enter')
    gui.verificar_e_maximizar_janela('Relatório de partidas individuais contas do Razão')

    print('Copia OK!')

    
    















### Calcula data de início e fim do mês atual
def calcular_datas(mes, ano):
    
    """
    Recebe dois numeros referentes a mês e ano e retorna
    as datas de início e fim no formato dd.mm.aaaa
    """
    
    # Determina o número de dias no mês
    if mes == 2:
        if ano % 4 == 0 and (ano % 100 != 0 or ano % 400 == 0):
            dias_no_mes = 29  # Ano bissexto
        else:
            dias_no_mes = 28
    elif mes in [4, 6, 9, 11]:
        dias_no_mes = 30
    else:
        dias_no_mes = 31
    
    # Formatação da data de início
    data_inicio = f"01.{mes:02d}.{ano}"
    
    # Formatação da data de fim
    data_fim = f"{dias_no_mes:02d}.{mes:02d}.{ano}"
    
    return data_inicio, data_fim


# Teste
# mes = 2
# ano = 2023
# data_inicio, data_fim = calcular_datas(mes, ano)















### Localiza e maximiza a janela ###
def verificar_e_maximizar_janela(titulo_janela):
    
    import pyautogui
    import pygetwindow
    import sys
    import time
    
    time.sleep(2)
    janelas = pygetwindow.getWindowsWithTitle(titulo_janela)
    if janelas:
        janela = janelas[0]
        janela.activate()           # Ativa a janela
        janela.maximize()           # Maximiza a janela
        
    else:
        pyautogui.alert(f"ERRO: A janela '{titulo_janela}' está fechada.")
        sys.exit()

# Teste
# titulo_desejado = "Calculadora"
# verificar_e_maximizar_janela(titulo_desejado)
# import pygetwindow 
# pygetwindow.getAllTitles()






















### Transação ###
def transacao():
    
    import pyautogui
    import gui
    
    # Maximizar janela SAP 
    gui.verificar_e_maximizar_janela('SAP Easy Access')
    
    pyautogui.press('enter')
    pyautogui.write('FBL3N')
    pyautogui.press('enter')
    
    print('Transação OK!')





















### Chamar Variante ###
def chamar_variante(tempo_inicial, tempo_fim):
    
    import pyautogui
    import gui
    
    # Maximizar janela SAP 
    gui.verificar_e_maximizar_janela('Relatório de partidas individuais contas do Razão')
    
    pyautogui.hotkey('shift', 'f5')
    pyautogui.PAUSE = tempo_fim
    pyautogui.write('/FL DIARIO')
    for _ in range(4): pyautogui.hotkey('tab')
    pyautogui.hotkey('backspace')
    pyautogui.write('e016321')
    pyautogui.PAUSE = tempo_inicial
    pyautogui.hotkey('f8')
    
    print('Variante OK!')
















### Passo 1 - Emitir Razao Bancos ###
def emitir_bancos(empresa, data_inicio, data_fim, tempo_inicial, tempo_fim, tempo_bancos):

    import pyautogui
    import time
    import gui
    
    # Maximizar janela SAP 
    gui.verificar_e_maximizar_janela('Relatório de partidas individuais contas do Razão')
    
    pyautogui.hotkey('shift', 'f4')
    pyautogui.PAUSE = tempo_fim
    for _ in range(2): pyautogui.hotkey('tab')
    pyautogui.hotkey('enter')
    pyautogui.write('101*')
    pyautogui.press('down')
    pyautogui.write('103*')
    pyautogui.press('down')
    pyautogui.write('111*')
    pyautogui.hotkey('f8')
    for _ in range(2): pyautogui.hotkey('tab')
    pyautogui.hotkey('backspace')
    pyautogui.write(empresa)
    for _ in range(12): pyautogui.hotkey('tab')
    pyautogui.write(data_inicio)
    pyautogui.hotkey('tab')
    pyautogui.write(data_fim)
    pyautogui.PAUSE = tempo_inicial
    pyautogui.press('f8')
    time.sleep(tempo_bancos)   # tempo para relatório abrir no SAP
    
    print('Bancos OK!')
    













### Passo 2 - Emitir Contrapartida Bancos ###
def emitir_contrapartida_bancos(empresa, data_inicio, data_fim, novo_diretorio, tempo_inicial, tempo_fim, tempo_contrap_bancos):

    import pyautogui
    import time
    import gui
    import fun
    
    # Maximizar janela SAP 
    gui.verificar_e_maximizar_janela('Relatório de partidas individuais contas do Razão')
    
    pyautogui.hotkey('shift', 'f4')
    pyautogui.hotkey('shift', 'tab')
    pyautogui.hotkey('enter')
    pyautogui.PAUSE = tempo_fim
    num_documento = fun.dados_contrapartida_bancos(novo_diretorio)
    gui.copiar_dados(num_documento)
    gui.verificar_e_maximizar_janela('Relatório de partidas individuais contas do Razão')
    pyautogui.hotkey('enter')           # ativar tela
    pyautogui.hotkey('shift', 'f12')    # colar valores
    pyautogui.hotkey('f8')
    for _ in range(2): pyautogui.hotkey('tab')
    pyautogui.hotkey('enter')
    pyautogui.hotkey('shift', 'f4')
    for _ in range(3): pyautogui.hotkey('shift', 'tab')
    for _ in range(2): pyautogui.hotkey('right')
    pyautogui.hotkey('enter')
    for _ in range(3): pyautogui.hotkey('tab')
    pyautogui.write('101*')
    pyautogui.press('down')
    pyautogui.write('103*')
    pyautogui.press('down')
    pyautogui.write('111*')
    pyautogui.hotkey('f8')
    for _ in range(2): pyautogui.hotkey('tab')
    pyautogui.write(empresa)
    for _ in range(12): pyautogui.hotkey('tab')
    pyautogui.write(data_inicio)
    pyautogui.hotkey('tab')
    pyautogui.write(data_fim)
    pyautogui.PAUSE = tempo_inicial
    pyautogui.hotkey('f8')
    time.sleep(tempo_contrap_bancos)

    print('Contrapartida Bancos OK!')
    
    
    
    
    
    
    
    
    
    
    
    
    
    
### Passo 3 - Emitir Compensação ###
def emitir_compensacao(empresa, data_inicio, data_fim, novo_diretorio, tempo_inicial, tempo_fim, tempo_compensacao):
        
    import pyautogui
    import time
    import gui
    import fun
    
    # Maximizar janela SAP 
    gui.verificar_e_maximizar_janela('Relatório de partidas individuais contas do Razão')
    
    pyautogui.hotkey('shift', 'f4')
    pyautogui.PAUSE = tempo_fim
    for _ in range(3): pyautogui.hotkey('shift', 'tab')
    pyautogui.hotkey('enter')
    pyautogui.hotkey('shift', 'f4')
    doc_compensacao = fun.dados_compensacao(novo_diretorio)
    gui.copiar_dados(doc_compensacao)
    gui.verificar_e_maximizar_janela('Relatório de partidas individuais contas do Razão')
    pyautogui.hotkey('enter')           # ativar tela
    pyautogui.hotkey('shift', 'f12')    # colar valores
    pyautogui.hotkey('f8')
    for _ in range(5): pyautogui.hotkey('tab')
    pyautogui.write(empresa)
    for _ in range(11): pyautogui.hotkey('tab')
    pyautogui.hotkey('up')
    pyautogui.hotkey('tab')
    pyautogui.write(data_inicio)
    pyautogui.hotkey('tab')
    pyautogui.write(data_fim)
    pyautogui.hotkey('f8')
    time.sleep(tempo_compensacao)
    pyautogui.PAUSE = tempo_inicial
    pyautogui.hotkey('f8')
    time.sleep(tempo_compensacao)

    print('Compensação OK!')