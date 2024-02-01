# Desafio_Tunts_Rock


## Certifique-se de atender aos seguintes requisitos antes de executar o script:

    Python e Dependências:
        Certifique-se de ter o Python instalado (versão 3.6 ou superior).
        Instale as dependências necessárias usando o seguinte comando:

        bash

        pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client

    Credenciais da API Google Sheets:
        Obtenha um arquivo JSON de credenciais da API do Google Sheets.
        Salve o arquivo de credenciais como Cred.json no mesmo diretório do script.

    Permissões:
        Certifique-se de ter permissões adequadas de leitura e gravação na planilha especificada.

# Configuração

Antes de executar o script, é necessário configurar algumas informações no próprio script:

    ID da Planilha: Substitua o valor de SAMPLE_SPREADSHEET_ID pelo ID da sua planilha do Google Sheets.
    Intervalo da Planilha: Substitua o valor de SAMPLE_RANGE_NAME pelo intervalo específico da sua planilha.

Execução

Execute o script main.py e siga as instruções para autenticar o acesso à planilha usando o navegador. Caso seja necessário, insira o código de autorização fornecido.
Funcionamento

## O script realiza as seguintes operações:

    Obtém ou atualiza as credenciais do usuário.
    Conecta-se à API do Google Sheets.
    Obtém os valores da planilha no intervalo especificado.
    Calcula a situação acadêmica de cada aluno com base em faltas e notas.
    Atualiza a planilha com as novas informações calculadas.

Observações

    Certifique-se de que a planilha contém os dados necessários nas colunas especificadas no intervalo (engenharia_de_software!C4:F27).
    O script considera uma situação de "Exame Final" para médias entre 5 e 7, calculando o NAF correspondente.
