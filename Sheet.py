import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Escopos necessários para acessar a planilha
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# ID da planilha no Google Sheets
SAMPLE_SPREADSHEET_ID = "1R4p914UFIXQuDuCPogAjekNye7oUk3_aDCY-C0KPFo8"

# Intervalo específico da planilha onde os dados são lidos e atualizados
SAMPLE_RANGE_NAME = "engenharia_de_software!C4:F27"


def get_credentials():
    # Verifica se há um arquivo de token existente
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    # Se não houver credenciais ou elas estiverem inválidas, obtém novas credenciais
    if not creds or not creds.valid:
        creds = refresh_or_request_credentials(creds)

    return creds


def refresh_or_request_credentials(creds):
    # Atualiza as credenciais se expiradas, caso contrário, solicita novas
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file("Cred.json", SCOPES)
        creds = flow.run_local_server(port=0)

    # Salva as novas credenciais
    with open("token.json", "wb") as token:
        token.write(creds.to_json())

    return creds


def get_values_from_spreadsheet(service, spreadsheet_id, range_name):
    # Obtém os valores da planilha no intervalo especificado
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=range_name
    ).execute()
    return result.get("values", [])


def calculate_situation(row_values, absence_threshold):
    try:
        # Obtém o número de faltas diretamente da terceira coluna (índice 2)
        total_absences = float(row_values[2]) if row_values[2].replace(".", "", 1).isdigit() else 0

        # Verifica se o total de faltas é maior que o limiar definido
        if total_absences > absence_threshold:
            return "Reprovado por Faltas", 0

        # Converte as notas de string para números
        notas = [float(valor) if valor.replace(".", "", 1).isdigit() else valor for valor in row_values]

        # Calcula a média das três primeiras notas
        media = sum(notas[:3]) / len(notas[:3])

        # Verifica se a média é menor que 5
        if media < 5:
            return "Reprovado por nota", 0
        # Se a média estiver entre 5 e 7, calcula o NAF
        elif 5 <= media < 7:
            naf = max(0, 2 * (5 - media))
            naf_arredondado = round(naf)
            print(f"NAF Calculado: {naf}, NAF Arredondado: {naf_arredondado}")
            return "Exame Final", naf_arredondado
        # Se nenhum dos casos anteriores se aplicar, significa que a média é 7 ou maior
        else:
            return "Aprovado", 0

    except (IndexError, ValueError) as e:
        return f"Erro: {str(e)}", 0


def atualizar_valores_na_planilha(service, spreadsheet_id, range_name, indice_linha, valores):
    # Atualiza as células na planilha com os novos valores
    intervalo_atualizacao = f"{range_name.split('!')[0]}!G{indice_linha + 4}:H{indice_linha + 4}"
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=intervalo_atualizacao,
        body={"values": [valores]},
        valueInputOption="RAW",
    ).execute()


def main():
    # Obtém as credenciais
    creds = get_credentials()

    try:
        # Constrói o serviço Google Sheets
        service = build("sheets", "v4", credentials=creds)

        # Obtém os valores da planilha no intervalo especificado
        valores = get_values_from_spreadsheet(service, SAMPLE_SPREADSHEET_ID, SAMPLE_RANGE_NAME)

        # Define o limiar de faltas como 15
        limiar_faltas = 15

        # Itera sobre cada linha da planilha
        for i, linha in enumerate(valores):
            try:
                # Calcula a situação e o NAF para cada aluno
                situacao, naf = calculate_situation(linha, limiar_faltas)

                # Imprime a situação e o NAF calculados
                print(f"Aluno {i + 4} - Situação: {situacao}, NAF: {naf}")

                # Atualiza a planilha com os novos valores
                if situacao == "Exame Final":
                    atualizar_valores_na_planilha(service, SAMPLE_SPREADSHEET_ID, SAMPLE_RANGE_NAME, i, [situacao, naf])
                else:
                    atualizar_valores_na_planilha(service, SAMPLE_SPREADSHEET_ID, SAMPLE_RANGE_NAME, i, [situacao, 0])

            except (IndexError, ValueError) as e:
                print(f"Erro processando valores para linha {i + 4}: {linha}. Detalhes: {str(e)}")

        print("Valores atualizados com sucesso!")

    except HttpError as err:
        print(f"Erro HTTP: {err.content}")


if __name__ == "__main__":
    main()
