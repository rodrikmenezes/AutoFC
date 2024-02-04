


### Criar novo diretório para salvar os arquivos ###
def criar_diretorio(diretorio, nome_pasta):
    
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
                    pyautogui.alert('Ano inválido. Por favor, insira um ano anterior ao ano atual.')
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
def salvar_arquivo(novo_diretorio, nome_pasta):
    
    import pyautogui
    import pyperclip
    import time
    import fun
    
    # Maximizar janela Relatório de partidas individuais contas do Razão
    fun.verificar_e_maximizar_janela('Relatório de partidas individuais contas do Razão')
    
    pyautogui.hotkey('alt')
    pyautogui.write('x')
    pyautogui.write('a')
    time.sleep(3)
    pyautogui.write(nome_pasta)
    for _ in range(6): pyautogui.hotkey('tab')
    pyautogui.hotkey('enter')
    pyperclip.copy(novo_diretorio)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.hotkey('enter')
    for _ in range(9): pyautogui.hotkey('tab')
    pyautogui.hotkey('enter')
    time.sleep(3)
    for _ in range(4): pyautogui.hotkey('tab')
    pyautogui.hotkey('enter')
    time.sleep(8)   # tempo para arquivo excel abrir
    
    # Maximizar e fechar janela 1 Bancos - Excel
    fun.verificar_e_maximizar_janela(nome_pasta)
    pyautogui.hotkey('alt', 'f4')
    
    print('Salvar Arquivo OK!')
    
















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
        janela.activate()  # Ativa a janela
        janela.maximize()  # Maximiza a janela
    else:
        pyautogui.alert(f"A janela '{titulo_janela}' está fechada.")
        sys.exit()

# Teste
# titulo_desejado = "Calculadora"
# verificar_e_maximizar_janela(titulo_desejado)
# import pygetwindow 
# pygetwindow.getAllTitles()














### Copia coluna Num Documentos do arquivo Bancos ###
def dados_contrapartida_bancos(pasta):
    
    """
    Para funcionar é necessário que exista o arquivo 1 Bancos salvo no diretório
    """
    
    import pandas
    
    # Nome o arquivo a ser importado
    arquivo = pasta + '\\1 Bancos.XLSX'
    
    # Importar planilha
    planilha = pandas.read_excel(arquivo)
    
    # Selecionar coluna Num Documento e eliminar duplicados
    num_documento = planilha['Nº documento'].drop_duplicates()
    
    # Copiar os valores da coluna para a área de transferência
    num_documento = '\n'.join(num_documento.astype(str))
    
    return num_documento


# Teste
# pasta = 'K:\\DPCL\\_Comum\\Realização do Fluxo de Caixa\\Celesc G\\2023.12 (1)'
# teste = dados_contrapartida_bancos(pasta)
# print(teste)
# print(type(teste))




### Copiar dados para serem colados
def copiar_dados(dados):
    
    import pyautogui
    import pyperclip
    
    pyautogui.press('win')
    pyautogui.write('bloco de notas')
    pyautogui.press('enter')
    pyautogui.hotkey('win','up')
    pyperclip.copy(dados)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c') 
    pyautogui.hotkey('alt', 'f4')
    
    print('Copiar dados OK!')


















    
    
    
    
    
    
    
    
    
    
