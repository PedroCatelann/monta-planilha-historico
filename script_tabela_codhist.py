import pandas as pd
import os

caminho = input("Digite o caminho no qual o arquivo está localizado: ")

caminho = caminho.replace('\\', '/').replace('"', '')

#"C:\\Users\\pedro\\Downloads\\Levantamento_Historicos_ASA_PlanilhaInicial.xlsx"
print(caminho)
dados = pd.read_excel(caminho)
try:
    colunas_desejadas = ['DESCRIÇÃO', 'CODHIST', 'EVENTO', 'FINALIDADE', 'NATUREZA', 'NATUREZA_DIMP', 'VAI PRO DIMP?']
    dados_selecionados = dados[colunas_desejadas]
except Exception as e:
    print("Erro nos nomes das colunas")
    print("As colunas devem possuir os nomes da seguinte forma: DESCRIÇÃO, CODHIST, EVENTO, FINALIDADE, NATUREZA, NATUREZA_DIMP, VAI PRO DIMP?")
    print('\n')
    print(str(e))

# Filtra os registros apenas de natureza crédito
dados_filtrados = dados.query("NATUREZA == 'C'")

# Troca os tipos de vazio, DOC, TED, PIX para 0, 1, 2, 3 respectivamente 
dados_filtrados['EVENTO'] = dados_filtrados['EVENTO'].map(lambda x: 0 if pd.isna(x) else (1 if x == 'DOC' else (2 if x == 'TED' else (3 if x == 'PIX' else x))))

# Coloca branco na finalidade se for nula
dados_filtrados['FINALIDADE'] = dados_filtrados['FINALIDADE'].map(lambda x: '' if pd.isna(x) or 0 else (x))


# Cria um novo dataframe somente com dados que não possuem código histórico e finalidade repetidos
dados_sem_duplicatas = dados_filtrados.drop_duplicates(subset=['CODHIST', 'FINALIDADE'], keep='first')

# Insere na primeira coluna a informação do código coligada 001 (padrão)
dados_sem_duplicatas.insert(0, 'CODCOLIGADA','001')

# Transforma o dataframe em um excel. Esse excel possuirá somente o tratamento do CODHIST e FINALIDADE, o arquivo terá os indicativos de VAI PRO DIMP? SIM ou NÃO
dados_sem_duplicatas.to_excel('./Levantamento_Historicos_Ajustados.xlsx', index=False)

# Cria um novo dataframe adicionando somente os registros que estão como sim
df_filtrado = dados_sem_duplicatas.loc[(dados_sem_duplicatas['VAI PRO DIMP?'] == 'S') | (dados_sem_duplicatas['VAI PRO DIMP?'] == 'SIM')]

# Cria um novo dataframe deixando somente as colunas que serão úteis na DIMP
dados_com_colunas_do_dimp = df_filtrado.drop(['OBS.:', 'VAI PRO DIMP?', 'NATUREZA'], axis=1)

# Cria um arquivo excel com a planilha pronta para a importação
dados_com_colunas_do_dimp.to_excel('./Levantamento_Historicos_Importacao_DIMP.xlsx', index=False)

# Cria o script para inserção dos registros diretamente no SQL
insert_statements = []
insert_statements.append("USE AB_DIMP")
insert_statements.append("GO")
for index, row in dados_com_colunas_do_dimp.iterrows():
    
    insert_statement = f"INSERT INTO CONFIGURACAO_TRANSACOES_ENVIADAS (CODCOLIGADA, DESCRICAO, CODHISTORICOCC, CODEVENTO, FINALIDADE, NATUREZA, DATAHORAINCLUSAO, USUARIOINCLUSAO) VALUES ('{row['CODCOLIGADA']}', '{row['DESCRIÇÃO']}', '{row['CODHIST']}', {row['EVENTO']}, '{row['FINALIDADE'] if row['FINALIDADE'] == '' else int(float(row['FINALIDADE']))}', {row['NATUREZA_DIMP']}, GETDATE(), 'CARGA');"
    insert_statements.append(insert_statement)

with open('PM_INSERE_HISTORICOS.sql', 'w') as file:
    for statement in insert_statements:
        file.write(statement + '\n')

script_directory = os.path.dirname(os.path.abspath(__file__))


root_directory = os.path.dirname(script_directory)

##print(dados_sem_duplicatas)
##print("Arquivo gerado no diretório: " + root_directory + '\Levantamento_Historicos_Ajustados.xlsx')
print('\n')
print('\n')
print("ARQUIVO SEM DUPLICATAS DE HISTÓRICOS E FINALIDADES, MAS COM COLUNAS SIM E NÃO")
print('./Levantamento_Historicos_Ajustados.xlsx')
print('\n')
print('\n')
print("ARQUIVO SEM DUPLICATAS DE HISTÓRICOS E FINALIDADES, SOMENTE COM AS COLUNAS VAI PRO DIMP SIM E SOMENTE COM AS COLUNAS ACEITÁVEIS PARA ENTRAR NA BASE")
print('./Levantamento_Historicos_Importacao_DIMP.xlsx')
print('\n')
print('\n')

input("Pressione Enter para sair...")

 