# %% [markdown]
# ## Importação das Bibliotecas 

# %%
import xlwings as xw
import datetime as dt
import os
import shutil as st
import random as rd
import time as tm
import re
import pandas as pd
import sys

# %% [markdown]
# ## Definicação de variáveis

# %% [markdown]
# Define as variáveis relacionada ao arquivo excel de orçamentação

# %%
arquivo_orc = r'C:\Users\davif\OneDrive\Davi\Faculdade e Cursos\Python\portfolio-1\Orçamentos\Orçamento-1.xlsx'

#Se certifica da quantidade de versões presentes no arquivo original
arquivo_original = xw.App(visible = False).books.open(arquivo_orc)
qtd_versoes_arq_original = len(arquivo_original.sheets)
arquivo_original.close()

cel_cliente = 'B1'
cel_data = 'B2'
cel_preco = 'D5'
cel_impostos = 'D6'
cel_materiais = 'D8'
cel_servicos = 'D9'
cel_custo_fin = 'D11'
cel_comissao = 'D12'

# %% [markdown]
# Verifica qual é o próximo número de orçamento a ser utilizado

# %%
#Cria uma variável contendo o diretório onde os arquivos de orçamentação estão alocados  
pasta_arquivo_orc = os.path.dirname(arquivo_orc)

#Padrão regex de numeração dos arquivos
padrao_num = r'-(\d+)'

num_orcs_gerados = []

for arq in os.listdir(pasta_arquivo_orc):

    if arq.startswith('Orçamento') & arq.endswith('.xlsx'):
        #Extrai o número de orçamento dos arquivos e o joga numa lista  
        num_orcs_gerados.append(int(re.search(padrao_num, arq).group(1)))
        
proximo_num_orc = max(num_orcs_gerados) + 1

# %% [markdown]
# Define as demais variáveis

# %%
qtd_orcamentos = 6000

mult_preco_min = 20000
mult_preco_max = 500000

opcoes_carga_tributaria = list(range(1, 13))

fx_custo_material_min = 30
fx_custo_material_max = 40

fx_custo_servicos_min = 10
fx_custo_servicos_max = 20

opcoes_perc_comissao = list(range(1, 4))

opcoes_perc_custo_fin = list(range(1, 4))

opcoes_clientes = list(range(5, 25))

opcoes_qtd_versoes_exc = list(range(1, 6))

opcoes_status_orc = list(range(1, 5))

data_min = dt.datetime(2024, 1, 2)
data_max = dt.datetime(2024, 4, 15)

nome_arq_data = 'Arquivo Datas.xlsx'
aba_data = 'Datas'
aba_qtd_orcs = 'Qtd Orcs'

# %% [markdown]
# ## Criação do Range de Datas

# %% [markdown]
# Cria as estruturas de dados

# %%
range_datas = []
cont_orc_dia = {}

# %% [markdown]
# Executa a criação do Range

# %%
caminho_arq_datas = pasta_arquivo_orc + '\\' + nome_arq_data

#Verifica se o arquivo de Datas x Qtd Orçs existe...
if os.path.exists(caminho_arq_datas) == True:
    arquivo_datas_orcs = pd.read_excel(caminho_arq_datas, usecols = [aba_data, aba_qtd_orcs])
    cont_orc_dia = arquivo_datas_orcs.to_dict(orient = 'records')
    cont_orc_dia = {item[aba_data]: item[aba_qtd_orcs] for item in cont_orc_dia}

    for data in cont_orc_dia.keys():
        range_datas.append(data)

#Caso contrário, cria um dicionário para tal informação
else:
    data_corrente = data_min

    #Cria um range de datas para escolha, bem como do dicionário "cont_orc_dia"
    #O dicionário "cont_orc_dia" guardará a informação de quantos orçamentos já foram criados em cada dia
    while data_corrente <= data_max:
        data_formatada = data_corrente.strftime('%d-%m-%Y')
        range_datas.append(data_formatada)
        cont_orc_dia[data_formatada] = 0
        data_corrente += dt.timedelta(days = 1)

# %% [markdown]
# ## Criação e configuração dos novos arquivos Excel de orçamentação

# %%
for num_orc in range(proximo_num_orc, qtd_orcamentos + 1):
    
    #Define os valores das variáveis de precificação
    mult_preco = rd.randrange(mult_preco_min, mult_preco_max, 100) 
    cod_carga_tributaria = rd.choice(opcoes_carga_tributaria)
    fx_custo_material = rd.randrange(fx_custo_material_min, fx_custo_material_max) / 100
    fx_custo_servicos = rd.randrange(fx_custo_servicos_min, fx_custo_servicos_max) / 100
    cod_perc_comissao = rd.choice(opcoes_perc_comissao)
    cod_perc_custo_fin = rd.choice(opcoes_perc_custo_fin)
    cod_cliente = rd.choice(opcoes_clientes)

    #Define qual será o status do orçamento (apenas se status_orc for igual a 3 o orçamento em questão será consolidado)
    status_orc = rd.choice(opcoes_status_orc)
    
    #Define a data para o novo orçamento
    #Pela regra de negócios estabelecida, 20 é o limite máximo de orçamentos para um único dia
    qtd_orcs = 99
    #Se a quantidade de orçamentos para o dia escolhido aleatoriamente for superior a 20, escolhe-se uma nova data 
    while qtd_orcs > 20:
        data_orc = rd.choice(range_datas)
        qtd_orcs = cont_orc_dia[data_orc] + 1

    cont_orc_dia[data_orc] += 1

    #Cria o nome do novo arquivo de orçamentação
    if status_orc == 3:
        sufixo = ' (Consolidado).xlsx'
    else:
        sufixo = '.xlsx'

    nome_novo_orc = f'Orçamento-{num_orc}{sufixo}'

    caminho_novo_orc = pasta_arquivo_orc + '\\' + nome_novo_orc

    #Excuta a cópia do arquivo principal para o novo arquivo
    st.copy(arquivo_orc, caminho_novo_orc)

    tm.sleep(2)
    #Cria um dicionário contendo quais são as células das variáveis de precificação (dentro do arquivo excel), bem como por quais valores elas devem ser preenchidas
    variaveis_arquivo = {}

    variaveis_arquivo = {
        cel_cliente: cod_cliente, 
        cel_data: data_orc,
        cel_preco: mult_preco, 
        cel_impostos: cod_carga_tributaria, 
        cel_materiais: fx_custo_material, 
        cel_servicos: fx_custo_servicos, 
        cel_custo_fin: cod_perc_custo_fin, 
        cel_comissao: cod_perc_comissao
    }

    #Acessa o novo arquivo criado
    with xw.App(visible = False).books.open(caminho_novo_orc) as arq:
        abas = arq.sheets
        aba_orc = abas['V1']
        
        try:
            #Altera o valor das variáveis de precificão
            for k, v in variaveis_arquivo.items():
                aba_orc[k].value = v
        
        except Exception as e:
            print(f'Houve erro na configuração de preços (Variável: {k} / Valor: {v}). Erro reportado: "{e}".')

        else:
            tm.sleep(1)
            #Define a quantidade de versões que permanecerá no novo arquivo
            qtd_versoes = len(abas)

            if qtd_versoes != qtd_versoes_arq_original:
                raise KeyError(f'A quantidade de abas no novo arquivo difere do original (O novo arquivo possui {qtd_versoes}).')
                sys.exit()

            qtd_versoes_excluir = rd.choice(opcoes_qtd_versoes_exc)
            versoes_restantes = qtd_versoes - qtd_versoes_excluir
            
            try:
                #Realiza a exclusão de abas
                for i in range(qtd_versoes, versoes_restantes, -1):
                    versao = f'V{i}'
                    aba_deletar = abas[versao]
                    aba_deletar.delete()

            except Exception as e:
                print(f'Houve erro no processo de exclusão da versão {versao} do orçamento. Erro reportado: "{e}".')  

            else:          
                arq.save()

    tm.sleep(1)

    #Por resguardo, força o fechamento das instâncias em aberto do Excel
    xw.apps.active.quit()
    tm.sleep(1)

# %% [markdown]
# ## Salvamento do arquivos de Datas x Orçamentos Criados 

# %%
df_datas_qtd_orcs = pd.DataFrame(list(cont_orc_dia.items()), columns = [aba_data, aba_qtd_orcs])

df_datas_qtd_orcs.to_excel(caminho_arq_datas, sheet_name = 'Dados', index = None)


