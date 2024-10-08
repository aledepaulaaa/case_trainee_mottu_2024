import pandas as pd
import locale

def analisa_desempenho(caminho_arquivo):
    """Analisa o desempenho das filiais da Mottu em agosto de 2024."""

    try:
        # Lê o arquivo Excel, pulando a primeira linha (que contém as descrições das colunas)
        df = pd.read_excel(caminho_arquivo, header=None)

        # Define os nomes das colunas
        df.columns = ['Data', 'A/V', 'Estado', 'Valor de Entrada', 'Caução']

        # Converte a coluna 'Data' para o tipo datetime para facilitar a filtragem
        df['Data'] = pd.to_datetime(df['Data'], format='%d/%m/%y', errors='coerce')

        # Filtra as linhas referentes ao mês de agosto de 2024
        df_agosto = df[(df['Data'].dt.month == 8) & (df['Data'].dt.year == 2024)]

        # Calcula a receita total por estado, considerando vendas e alugueis
        df_agosto['Receita'] = df_agosto.apply(lambda row: row['Valor de Entrada'] if pd.notna(row['Valor de Entrada']) else row['Caução'], axis=1)
        receita_por_estado = df_agosto.groupby('Estado')['Receita'].sum()

        # Ordena os estados por receita total, do maior para o menor
        receita_por_estado_ordenada = receita_por_estado.sort_values(ascending=False)

        # Define o estado com melhor desempenho
        melhor_desempenho = receita_por_estado_ordenada.index[0]
        melhor_receita = receita_por_estado_ordenada.iloc[0]

        # Configura a localização para usar o formato de moeda brasileiro
        locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8') #ou 'pt_BR.utf8'

        # Formata os números com R$ e separador de milhar
        melhor_receita_formatada = locale.currency(melhor_receita, grouping=True)


        # Gera o arquivo txt com os resultados, com formatação melhorada
        with open('analise_desempenho.txt', 'w', encoding='utf-8') as f:
            f.write("Análise de Desempenho das Filiais da Mottu - Agosto de 2024\n\n")
            f.write("Metodologia:\n")
            f.write("1. Foram considerados apenas os dados de agosto de 2024.\n")
            f.write("2. A receita total por estado foi calculada somando os valores de 'Valor de Entrada' (para vendas) ou 'Caução' (para alugueis).\n")
            f.write("3. Os estados foram ordenados pela receita total, do maior para o menor.\n\n")

            f.write("Resultados:\n")
            f.write(f"Filial com melhor desempenho: {melhor_desempenho}\n")
            f.write(f"Receita total: {melhor_receita_formatada}\n\n")

            f.write("Desempenho de todas as filiais (ordenado):\n")
            # Formata a coluna Receita com R$ e separador de milhar
            receita_por_estado_ordenada = receita_por_estado_ordenada.apply(lambda x: locale.currency(x, grouping=True))
            f.write(str(receita_por_estado_ordenada))

        print("Arquivo 'analise_desempenho.txt' gerado com sucesso!")

    except FileNotFoundError:
        print(f"Erro: Arquivo '{caminho_arquivo}' não encontrado.")
    except Exception as e:
        print(f"Erro durante a análise: {e}")


if __name__ == "__main__":
    arquivo_excel = 'faturamento_atualizado.xlsx'
    analisa_desempenho(arquivo_excel)