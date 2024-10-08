import pandas as pd
import locale

def analisa_faturamento(caminho_arquivo):
    """Analisa o faturamento das filiais da Mottu, considerando as regras de pagamento."""
    try:
        # Lê o arquivo Excel, pulando a primeira linha (descrições das colunas)
        df = pd.read_excel(caminho_arquivo, header=None)
        df.columns = ['Data', 'A/V', 'Estado', 'Valor de Entrada', 'Caução']
        df['Data'] = pd.to_datetime(df['Data'], format='%d/%m/%y', errors='coerce')

        # Define os valores de faturamento (considerando a relação entre aluguel semanal e venda mensal)
        faturamento_venda_avista = 2000
        faturamento_venda_parcelado = 2500
        faturamento_aluguel_semanal = 125  # 2500 / 4 / 4 (aluguel semanal = venda mensal/4/4)
        faturamento_aluguel_parcelado_semanal = 175  # 2500 / 4 /4

        # Função para calcular o faturamento por linha
        def calcula_faturamento(row):
            if row['A/V'] == 'Venda':
                return faturamento_venda_avista if pd.isna(row['Valor de Entrada']) else faturamento_venda_parcelado
            elif row['A/V'] == 'Aluguel':
                return faturamento_aluguel_semanal * 4 if pd.isna(row['Caução']) else faturamento_aluguel_parcelado_semanal * 4
            else:
                return 0  # Caso haja algum valor diferente de 'Venda' ou 'Aluguel'

        # Aplica a função para calcular o faturamento para cada linha
        df['Faturamento'] = df.apply(calcula_faturamento, axis=1)

        # Agrupa os dados por estado e soma o faturamento
        faturamento_por_estado = df.groupby('Estado')['Faturamento'].sum()

        # Ordena os estados pelo faturamento total
        faturamento_por_estado_ordenado = faturamento_por_estado.sort_values(ascending=False)

        # Define o estado com melhor faturamento
        melhor_faturamento_estado = faturamento_por_estado_ordenado.index[0]
        melhor_faturamento_valor = faturamento_por_estado_ordenado.iloc[0]

        # Configura a localização para usar o formato de moeda brasileiro
        locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

        # Formata os números com R$ e separador de milhar
        melhor_faturamento_formatado = locale.currency(melhor_faturamento_valor, grouping=True)

        # Gera o arquivo txt com os resultados
        with open('analise_faturamento.txt', 'w', encoding='utf-8') as f:
            f.write("Análise de Faturamento das Filiais da Mottu\n\n")
            f.write("Metodologia:\n")
            f.write("1. O faturamento mensal por moto foi considerado igual para aluguel e venda.\n")
            f.write("2. O faturamento semanal do aluguel foi calculado como sendo quatro vezes menor que o faturamento mensal da venda.\n")
            f.write("3. O faturamento total por estado foi calculado somando o faturamento de todas as transações em cada estado.\n")
            f.write("4. Os estados foram ordenados pelo faturamento total, do maior para o menor.\n\n")

            f.write("Resultados:\n")
            f.write(f"Filial com melhor faturamento: {melhor_faturamento_estado}\n")
            f.write(f"Faturamento total: {melhor_faturamento_formatado}\n\n")

            f.write("Faturamento de todas as filiais (ordenado):\n")
            faturamento_por_estado_ordenado = faturamento_por_estado_ordenado.apply(lambda x: locale.currency(x, grouping=True))
            f.write(str(faturamento_por_estado_ordenado))

        print("Arquivo 'analise_faturamento.txt' gerado com sucesso!")

    except FileNotFoundError:
        print(f"Erro: Arquivo '{caminho_arquivo}' não encontrado.")
    except Exception as e:
        print(f"Erro durante a análise: {e}")


if __name__ == "__main__":
    arquivo_excel = 'faturamento_atualizado.xlsx'
    analisa_faturamento(arquivo_excel)