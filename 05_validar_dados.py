# Validação de dados (CPF, e-mail, números negativos)
import pandas as pd
import re

def validar_cpf(cpf):
    cpf = re.sub(r'\D', '', str(cpf))
    if len(cpf) != 11 or cpf == cpf[0]*11:
        return False
    # cálculo dos dois dígitos verificadores (simplificado)
    for i in range(9, 11):
        soma = sum(int(cpf[num]) * (i+1 - num) for num in range(i))
        digito = (soma * 10) % 11
        if digito == 10: digito = 0
        if digito != int(cpf[i]):
            return False
    return True

def validar_email(email):
    return re.match(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$', str(email)) is not None

def validar_planilha(arquivo, planilha, coluna_cpf, coluna_email, coluna_valor):
    """Verifica CPF, e-mail e valores negativos; gera relatório de erros."""
    df = pd.read_excel(arquivo, sheet_name=planilha)
    erros = []
    
    for idx, row in df.iterrows():
        linha = idx + 2  # porque pandas index começa em 0 mas Excel linha 1 é cabeçalho
        if not validar_cpf(row[coluna_cpf]):
            erros.append(f"Linha {linha}: CPF inválido {row[coluna_cpf]}")
        if not validar_email(row[coluna_email]):
            erros.append(f"Linha {linha}: E-mail inválido {row[coluna_email]}")
        if row[coluna_valor] < 0:
            erros.append(f"Linha {linha}: Valor negativo {row[coluna_valor]}")
    
    if erros:
        pd.DataFrame(erros, columns=["Erro encontrado"]).to_excel("erros_validacao.xlsx", index=False)
        print(f"Foram encontrados {len(erros)} erros. Veja 'erros_validacao.xlsx'")
    else:
        print("Tudo ok! Nenhum erro encontrado.")

validar_planilha("clientes.xlsx", "Plan1", "CPF", "Email", "Valor_Compra")
