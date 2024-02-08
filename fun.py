











### Copia coluna Num Documentos do arquivo Bancos ###
def dados_contrapartida_bancos(diretorio):
    
    """
    Para funcionar é necessário que exista o arquivo 1 Bancos salvo no diretório
    """
    
    import pandas
    
    # Nome do arquivo a ser importado
    arquivo = diretorio + '\\1 Bancos.XLSX'
    
    # Importar planilha
    planilha = pandas.read_excel(arquivo)
    
    # Selecionar coluna Num Documento e eliminar duplicados
    dados = planilha['Nº documento'].drop_duplicates()
    
    # Copiar os valores da coluna para a área de transferência
    dados = '\n'.join(dados.astype(str))
    
    return dados

# Teste
# pasta = 'K:\\DPCL\\_Comum\\Realização do Fluxo de Caixa\\Celesc G\\2023.12 (1)'
# teste = dados_contrapartida_bancos(pasta)
# print(teste)
# print(type(teste))









### Copia coluna Num Documentos do arquivo Bancos ###
def dados_compensacao(diretorio):
    
    """
    Para funcionar é necessário que exista o arquivo 2 Contrapartida Bancos salvo no diretório
    """
    
    import pandas
    
    # Nome do arquivo a ser importado
    arquivo = diretorio + '\\2 Contrapartida Bancos.XLSX'
    
    # Importar planilha
    planilha = pandas.read_excel(arquivo)
    
    # Selecionar coluna Num Documento, eliminar duplicados e converter em valor numérico
    dados = planilha.dropna(subset=['Doc.compensação'])['Doc.compensação'].drop_duplicates().astype(int)
    
    # Copiar os valores da coluna para a área de transferência
    dados = '\n'.join(dados.astype(str))
    
    return dados
    

# Teste
# pasta = 'K:\\DPCL\\_Comum\\Realização do Fluxo de Caixa\\Celesc G\\2023.12 (8)'
# teste = dados_contrapartida_bancos(pasta)
# print(teste)
# print(type(teste))
    
    
    
    
    







def check_contrapatida_bancos(diretorio):
    
    import pandas
    import pyautogui
    import sys
    
    # Nome dos arquivos a serem importados
    arquivo_bancos = diretorio + '\\1 Bancos.XLSX'
    arquivo_contrapartida_bancos = diretorio + '\\2 Contrapartida Bancos.XLSX'
    
    # Importar arquivos
    bancos = pandas.read_excel(arquivo_bancos)
    contrapartida_bancos = pandas.read_excel(arquivo_contrapartida_bancos)
    
    # Check
    check = round(sum(bancos['Montante em moeda interna'], 0)) + round(sum(contrapartida_bancos['Montante em moeda interna'], 0))
    
    if check == 0:
        print('Check OK!')
    else:
        pyautogui.alert('ERRO: Montante Nacos diferentes de Contrapartida Bancos ({})'.format(check))
        sys.exit()





