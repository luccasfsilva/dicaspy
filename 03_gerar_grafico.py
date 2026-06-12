#  Gerar gráfico automático a partir dos dados
import pandas as pd
import xlsxwriter

def criar_grafico_excel(arquivo_dados, planilha_dados, coluna_categorias, coluna_valores, arquivo_saida):
    """Lê dados de um Excel e gera um novo arquivo com gráfico de barras."""
    df = pd.read_excel(arquivo_dados, sheet_name=planilha_dados)
    
    workbook = xlsxwriter.Workbook(arquivo_saida)
    worksheet = workbook.add_worksheet("Resultado")
    worksheet.add_chart()  # placeholder
    
    # Escreve os dados
    for i, (cat, val) in enumerate(zip(df[coluna_categorias], df[coluna_valores])):
        worksheet.write(i+1, 0, cat)
        worksheet.write(i+1, 1, val)
    worksheet.write(0, 0, coluna_categorias)
    worksheet.write(0, 1, coluna_valores)
    
    # Cria gráfico
    chart = workbook.add_chart({'type': 'column'})
    chart.add_series({
        'categories': f'=Resultado!$A$2:$A${len(df)+1}',
        'values': f'=Resultado!$B$2:$B${len(df)+1}',
        'name': coluna_valores,
    })
    chart.set_title({'name': 'Gráfico automático'})
    worksheet.insert_chart('D2', chart)
    workbook.close()
    print(f"Gráfico salvo em {arquivo_saida}")

criar_grafico_excel("vendas_por_mes.xlsx", "Sheet1", "Mês", "Vendas", "relatorio_com_grafico.xlsx")
