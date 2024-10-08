import pandas as pd

def separa_tabelas(caminho_arquivo_entrada, caminho_arquivo_vendas, caminho_arquivo_aluguel):
    """Separa a tabela de faturamento em duas tabelas: vendas e aluguéis."""
    try:
        df = pd.read_excel(caminho_arquivo_entrada, header=None)
        df.columns = ['Data', 'A/V', 'Estado', 'Valor de Entrada', 'Caução']

        # Cria DataFrames para vendas e aluguéis
        df_vendas = df[df['A/V'] == 'Venda']
        df_aluguel = df[df['A/V'] == 'Aluguel']


        # Salva as novas tabelas como arquivos CSV. Você pode alterar para Excel se preferir.
        df_vendas.to_csv(caminho_arquivo_vendas, index=False)
        df_aluguel.to_csv(caminho_arquivo_aluguel, index=False)

        print(f"Tabelas separadas com sucesso em '{caminho_arquivo_vendas}' e '{caminho_arquivo_aluguel}'")

    except FileNotFoundError:
        print(f"Erro: Arquivo '{caminho_arquivo_entrada}' não encontrado.")
    except Exception as e:
        print(f"Erro durante a separação das tabelas: {e}")


if __name__ == "__main__":
    arquivo_excel = 'faturamento_atualizado.xlsx'
    arquivo_vendas = 'vendas_mottu.csv'
    arquivo_aluguel = 'aluguel_mottu.csv'
    separa_tabelas(arquivo_excel, arquivo_vendas, arquivo_aluguel)