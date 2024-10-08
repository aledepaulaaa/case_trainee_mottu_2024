import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from datetime import datetime

def process_dates(df, date_column):
    """Processa as datas em um DataFrame pandas."""

    try:
        # Remover datas especificadas
        dates_to_remove = ['04/08/24', '11/08/24', '18/08/24', '25/08/24']
        df = df[~df[date_column].isin(dates_to_remove)]

    except KeyError as e:
        print(f"Erro: Coluna '{date_column}' não encontrada. Verifique o nome da coluna.")
        return df
    except Exception as e:
        print(f"Um erro inesperado ocorreu: {e}")
        return df

    return df

# Ler o arquivo CSV. Assumindo que a coluna de data é a primeira coluna e se chama 'Data'
try:
    df = pd.read_csv('tabela_aluguel_vendas.csv', encoding='iso-8859-1')
    if 'Data' in df.columns:
        date_column = 'Data'
    else:
        date_column = df.columns[0]
        print(f"A coluna 'Data' não foi encontrada, usando a primeira coluna '{date_column}' como coluna de datas.")

    df = process_dates(df, date_column)

    # Criar arquivo Excel com formatação
    wb = Workbook()
    ws = wb.active

    # Copiar dados do DataFrame para a planilha.  As linhas removidas serão exibidas como vazias.
    for row in df.values.tolist():
        ws.append(row)

    wb.save('faturamento_atualizado.xlsx')
    print("Arquivo Excel atualizado e exportado com sucesso!")

except FileNotFoundError:
    print("Erro: Arquivo 'tabela_aluguel_vendas.csv' não encontrado.")
except pd.errors.EmptyDataError:
    print("Erro: Arquivo 'tabela_aluguel_vendas.csv' está vazio.")
except pd.errors.ParserError:
    print("Erro: Erro ao analisar o arquivo 'tabela_aluguel_vendas.csv'. Verifique o formato do arquivo.")
except UnicodeDecodeError as e:
    print(f"Erro de decodificação: {e}. Tente especificar uma codificação diferente (ex: encoding='latin-1').")
except Exception as e:
    print(f"Erro desconhecido: {e}")