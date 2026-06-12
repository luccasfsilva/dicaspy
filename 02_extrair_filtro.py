# Extrair linhas baseadas em um filtro (ex: vendas > 1000)
import pandas as pd

def filtrar_excel(arquivo_entrada, planilha, coluna, valor_minimo, arquivo_saida):
    """Salva apenas as linhas onde a coluna numérica é maior que valor_minimo."""
    df = pd.read_excel(arquivo_entrada, sheet_name=planilha)
    df_filtrado = df[df[coluna] > valor_minimo]
    df_filtrado.to_excel(arquivo_saida, index=False)
    print(f"Linhas extraídas: {len(df_filtrado)} de {len(df)}")

# Exemplo: manter vendas acima de 1000
filtrar_excel("vendas.xlsx", "Sheet1", "Valor", 1000, "vendas_acima_1000.xlsx")
