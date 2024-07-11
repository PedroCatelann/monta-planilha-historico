import pandas as pd
import os

caminho = input("Digite o caminho no qual o arquivo está localizado: ")

caminho = caminho.replace('\\', '/').replace('"', '')

#"C:\\Users\\pedro\\Downloads\\Levantamento_Historicos_ASA_PlanilhaInicial.xlsx"
print(caminho)
dados = pd.read_excel(caminho)

colunas_desejadas = ['DESCRIÇÃO', 'CODHIST', 'EVENTO', 'FINALIDADE', 'NATUREZA', 'NATUREZA_DIMP', 'VAI PRO DIMP?']
dados_selecionados = dados[colunas_desejadas]


dados_filtrados = dados.query("NATUREZA == 'C'")
dados_filtrados['EVENTO'] = dados_filtrados['EVENTO'].map(lambda x: 0 if pd.isna(x) else (1 if x == 'DOC' else (2 if x == 'TED' else (3 if x == 'PIX' else x))))
dados_filtrados['FINALIDADE'] = dados_filtrados['FINALIDADE'].map(lambda x: '' if pd.isna(x) or 0 else (x))



dados_sem_duplicatas = dados_filtrados.drop_duplicates(subset=['CODHIST', 'FINALIDADE'], keep='first')


dados_sem_duplicatas.insert(0, 'CODCOLIGADA','001')

dados_sem_duplicatas.to_excel('./Levantamento_Historicos_Ajustados.xlsx', index=False)

script_directory = os.path.dirname(os.path.abspath(__file__))

root_directory = os.path.dirname(script_directory)

print(dados_sem_duplicatas)
print("Arquivo gerado no diretório: " + root_directory + '\Levantamento_Historicos_Ajustados.xlsx')

input("Pressione Enter para sair...")

 