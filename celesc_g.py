#!/usr/bin/env python
# coding: utf-8

# # **Realizar Fluxo de Caixa - Celesc G**

# ## Instruções:
# 
# 1. Verificar acesso ao disco (K:)
# 2. Abrir o SAP e fazer login
# 3. Certificar-se de que o usuario tem acesso a transacao FBL3N
# 4. Nao mexer no computador durante o processo
# 5. Configurar computador para não entrar em modo de espera

# ## Parâmetros de entrada

# Carregar módulos

# In[1]:


import pyautogui
import datetime
import fun
import gui


# Mensagem de início

# In[2]:


pyautogui.alert('Realização do Fluxo Caixa - Celesc G')


# Dados de entrada
# > *caso a pasta do mês atual já exista, uma nova pasta é criada no diretório e os arquivos serão salvos nela*

# In[3]:


diretorio = 'K:\\DPCL\\_Comum\\Realização do Fluxo de Caixa\\Celesc G\\'
empresa = 'CE03'
data = datetime.datetime.today()
# ano = data.year
# mes = data.month
ano = 2023
mes = 12
data_inicio, data_fim = gui.calcular_datas(mes, ano)
pasta = str(ano) + '.' + str(mes)
novo_diretorio = gui.criar_diretorio(diretorio, pasta)
# novo_diretorio = 'K:\\DPCL\\_Comum\Realização do Fluxo de Caixa\\Celesc G\\2023.12 (1)'


# Exibir parâmetros

# In[4]:


# print("Mês: {} \nAno: {} \nIni: {} \nFim: {} \nDir: '{}'".format(mes, ano, data_inicio, data_fim, novo_diretorio))


# Configurar tempo entre etapas

# In[5]:


# Tempo padrão de execução
tempo_inicial = pyautogui.PAUSE = 1

# Tempo de celeridade de processos
tempo_fim = pyautogui.PAUSE = 0.3

# Carrgar tempo de execução padrão
tempo_inicial


# ## Passo 1) Razão Bancos

# Inserir transação SAP

# In[6]:


gui.transacao()


# Chamar variante

# In[7]:


gui.chamar_variante(tempo_inicial, tempo_fim)


# Emitir relatório no SAP

# In[8]:


gui.emitir_bancos(empresa, data_inicio, data_fim, tempo_inicial, tempo_fim)


# Salvar Relatorio Bancos    

# In[ ]:


gui.salvar_arquivo(novo_diretorio, tempo_inicial, tempo_fim, '1 Bancos')


# ## Passo 2) Contrapartida Bancos

# Inserir transação SAP

# In[ ]:


gui.transacao()


# Chamar variante

# In[ ]:


gui.chamar_variante(tempo_inicial, tempo_fim)


# Emitir relatório no SAP

# In[ ]:


gui.emitir_contrapartida_bancos(empresa, data_inicio, data_fim, novo_diretorio, tempo_inicial, tempo_fim)


# Salvar Relatorio Bancos    

# In[ ]:


gui.salvar_arquivo(novo_diretorio, tempo_inicial, tempo_fim, '1 Contrapartida Bancos')

