import pandas as pd
from pathlib import Path

def unir_excels(pasta_origem, arquivo_destino):
    """Junta todos os arquivos .xlsx de uma pasta em um único Excel (uma planilha por arquivo)."""
    caminhos = Path(pasta_origem).glob("*.xlsx")
    with pd.ExcelWriter(arquivo_destino, engine='openpyxl') as writer:
        for caminho in caminhos:
            df = pd.read_excel(caminho)
            nome = caminho.stem  # nome do arquivo sem extensão
            df.to_excel(writer, sheet_name=nome[:31], index=False)  # limite de 31 caracteres
    print(f"União concluída: {arquivo_destino}")

if __name__ == "__main__":
    unir_excels("dados/entrada", "dados/saida/unificado.xlsx")
