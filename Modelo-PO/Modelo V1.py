import pandas as pd
from pulp import *

file_name = 'Dados Louvor Shalom.xlsx'

sheets_dict = pd.read_excel(file_name, sheet_name=['Instrumento', 'Peso dos Instrumentos', 'Peso dos Dias','Disponibilidade', 'Dias por Mês'])

# Acesse cada sheet pelo nome
#Dias por Mês
P = sheets_dict['Instrumento'].set_index('Pessoas')

L = sheets_dict['Peso dos Instrumentos'].set_index('Instrumentos').squeeze()
D = sheets_dict['Peso dos Dias'].set_index('Dias')

Disp = sheets_dict['Disponibilidade'].set_index('Dias')
DPM = sheets_dict['Dias por Mês'].set_index('Pessoas')

# Convertendo os índices para o tipo datetime
D.index = pd.to_datetime(D.index)
Disp.index = pd.to_datetime(Disp.index)

# Formatando os índices para o formato desejado
D.index = D.index.strftime('%d_%m_%Y')
Disp.index = Disp.index.strftime('%d_%m_%Y')

#-----------------------------------------------------------------------------------------------------------------------------------------------
# Definir o problema
prob = LpProblem("Alocacao_de_Musicos", LpMaximize)

# Variáveis de decisão
X = LpVariable.dicts("X", (P.index, P.columns, D.index), cat='Binary')

# Função objetivo
prob += lpSum([X[i][j][k] * Disp[i][k] * P.loc[i, j] * L[j] * D['Peso'][k] for i in P.index for j in P.columns for k in D.index])

# Restrições

# Um instrumento só é tocado por uma pessoa em cada dia K
for j in P.columns: #para todo instrumento
    for k in D.index: #para todo dia
        prob += lpSum([X[i][j][k] for i in P.index]) <= 1 #Cada pessoa 

# Cada pessoa toca somente um instrumento cada dia
for i in P.index: #Cada pessoa 
    for k in D.index: #para todo dia
        prob += lpSum([X[i][j][k] for j in P.columns]) <= 1 #para todo instrumento


# Cada pessoa toca somente o numero de dias que se ofereceu por mês
for i in P.index: 
    for k in range(len(D.index) - 3):
        prob += lpSum([X[i][j][D.index[k]] + X[i][j][D.index[k+1]] + X[i][j][D.index[k+2]] + X[i][j][D.index[k+3]] for j in P.columns]) <= DPM['Quantidade'][i]


# A pessoa não será alocada se não tocar o instrumento específico
for i in P.index: #Pessoas
    for j in P.columns: #Instrumentos
        for k in D.index: #Dias
            if P.loc[i, j] == 0: #Tabela de Capacidades. Pessoa i toca o instrumento j.
                prob += X[i][j][k] == 0

# Resolver o problema
prob.solve()

#-----------------------------------------------------------------------------------------------------------------------------------------------

# Imprimir o status da solução
print("Status:", LpStatus[prob.status])

# Imprimir o valor da função objetivo
print("Valor da função objetivo =", value(prob.objective))

# Dicionário para armazenar os resultados
resultados = {}

# Iterar sobre cada variável de decisão
for v in prob.variables():
    if v.varValue != 0:
        # Extrair o nome da pessoa, o instrumento e o dia
        _, pessoa, instrumento, dia = v.name.split("_", 3)

        # Se o dia ainda não está no dicionário, adicione-o
        if dia not in resultados:
            resultados[dia] = {}

        # Adicione a pessoa e o instrumento ao dia correspondente
        resultados[dia][instrumento] = pessoa

df = pd.DataFrame(columns=P.columns, index=D.index)
escala_instrumental = pd.DataFrame(0, columns=P.index, index=D.index)

# Preencher o DataFrame com os resultados
for dia in resultados:
    for instrumento in resultados[dia]:
        row_index = dia
        column_name = resultados[dia][instrumento]
        escala_instrumental.loc[row_index, column_name] = 1
        df.loc[dia, instrumento] = resultados[dia][instrumento]

# # Substituir NaN por ''
df = df.fillna('')

# Convertendo os índices para o tipo datetime
df.index = pd.to_datetime(df.index, format='%d_%m_%Y')

# Formatando os índices para o formato desejado e criando um novo índice
formatted_index = df.index.strftime('%d/%m/%Y')

# Atribuindo o novo índice formatado ao DataFrame
df.index = formatted_index

df.to_excel('Resultados.xlsx', index=True)