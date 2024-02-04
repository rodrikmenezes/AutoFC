








### Transação ###
def transacao():
    
    import pyautogui
    import fun
    
    # Maximizar janela SAP 
    fun.verificar_e_maximizar_janela('SAP Easy Access')
    
    # time.sleep(2)
    pyautogui.write('FBL3N')
    pyautogui.press('enter')
    
    print('Transação OK!')












### Chamar Variante ###
def chamar_variante():
    
    import pyautogui
    import fun
    
    # Maximizar janela SAP 
    fun.verificar_e_maximizar_janela('Relatório de partidas individuais contas do Razão')
    
    pyautogui.hotkey('shift', 'f5')
    pyautogui.write('/FL DIARIO')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('backspace')
    pyautogui.write('e016321')
    pyautogui.hotkey('f8')
    
    print('Variante OK!')
















### Passo 1 - Emitir Razao Bancos ###
def emitir_bancos(empresa, data_inicio, data_fim):

    import pyautogui
    import time
    import fun
    
    # Maximizar janela SAP 
    fun.verificar_e_maximizar_janela('Relatório de partidas individuais contas do Razão')
    
    pyautogui.hotkey('shift', 'f4')
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
    pyautogui.press('f8')
    time.sleep(5)   # tempo para relatório abrir no SAP
    
    print('Bancos OK!')
    









### Passo 2 - Emitir Contrapartida Bancos ###
def emitir_contrapartida_bancos(empresa, data_inicio, data_fim, novo_diretorio):

    import pyautogui
    import pyperclip
    import time
    import fun
    
    num_documento = fun.dados_contrapartida_bancos(novo_diretorio)
    # pyautogui.press('win')
    # pyautogui.write('bloco de notas')
    # pyautogui.press('enter')
    # pyautogui.hotkey('win','up')
    # pyperclip.copy(num_documento)
    # pyautogui.hotkey('ctrl', 'v')
    # pyautogui.hotkey('ctrl', 'a')
    # pyautogui.hotkey('ctrl', 'c') 
    # pyautogui.hotkey('alt', 'f4')
    pyautogui.hotkey('tab')
    pyautogui.hotkey('enter')
    fun.verificar_e_maximizar_janela('Relatório de partidas individuais contas do Razão')
    pyautogui.hotkey('shift', 'f4')
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
    for _ in range(17): pyautogui.hotkey('shift', 'tab')
    pyautogui.hotkey('enter')
    pyautogui.hotkey('shift', 'f12')
    pyautogui.hotkey('f8')
    pyautogui.hotkey('f8')
    time.sleep(40)

    print('Contrapartida Bancos OK!')
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    