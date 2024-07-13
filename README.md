# monta-planilha-historico-dimp


### Descrição
Aplicação criada para auxiliar no tratamento dos códigos históricos no início da implantação do sistema DIMP.

### Uso
Dentro do diretório `dist/` há um arquivo executável que deve ser instalado e executado. Inicialmente a aplicação irá solicitar um caminho para o arquivo da planilha (indicando os históricos que vão ou não para a DIMP) 
que o cliente envia nas RAs.
O caminho deve ser copiado e colado no cmd e em seguida deve-se pressionar a tecla enter. Feito isso, o sistema irá realizar todos os filtros necessários para tratamento dos códigos históricos

### Entrada
O usuário deve-se atentar para que o formato das colunas da planilha esteja da maneira correta, caso contrário será gerado um erro!

Exemplo de formato esperado:

| DESCRIÇÃO | CODHIST | EVENTO | FINALIDADE | NATUREZA | NATUREZA_DIMP | VAI PRO DIMP? |
| --------- | ------- | ------ | ---------- | -------- | ------------- | ------------- |
| Liquidação Cobrança   | 74 | 0 | 10 | C | 4 | S |

### Saída
A aplicação irá gerar 3 arquivos de saída:

`Levantamento_Historicos_Ajustados.xlsx` → Arquivo sem duplicatas de históricos e finalidades, mas com colunas sim e não

`Levantamento_Historicos_Importacao_DIMP.xlsx` → Arquivo sem duplicatas de históricos e finalidades, somente com as colunas vai pro dimp sim e somente com as colunas aceitáveis para entrar na base

`PM_INSERE_HISTORICOS.sql` → Script sql com os históricos filtrados para inserção direta no banco de dados
 
