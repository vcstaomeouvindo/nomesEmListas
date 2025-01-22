import fitz
import pandas as pd
import unicodedata

# Abas da planilha a serem analisadas (no nosso caso, cada aba possui alunos de uma turma)
abas = ["Turma1", "Turma2", "Turma3", "Turma4"]
planilha_nomes = "planilha.xlsx" # Caminho da planilha com extensão .xlsx

# Função para remover acentos de uma string
def remove_accents(input_str):
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    return ''.join([c for c in nfkd_form if not unicodedata.combining(c)])

# Lista de PDFs a serem analisados
pdf_files = ["endereço1.pdf","endereço2.pdf"]

# Dicionário para armazenar resultados
search_results = []
i = 0

for aba in abas:
    i += 1
    df = pd.read_excel(planilha_nomes, sheet_name=aba)

    # Aplicar a função para remover acentos na coluna "Nome"
    df['Nome'] = df['Nome'].apply(remove_accents)
    # Converter "CPF ajustado" para string
    df["CPF ajustado"] = df["CPF ajustado"].astype(str)

    # Loop por cada PDF
    n = 0
    for pdf_path in pdf_files:
        n += 1
        with fitz.open(pdf_path) as pdf:
            # Iterar por cada página do PDF
            for page_num in range(pdf.page_count):
                page = pdf[page_num]
                blocks = page.get_text("block")  # Extrair texto da página
                
                # Dividir o texto em linhas para análise detalhada
                lines = blocks.split('\n')
                
                # Procurar cada nome e CPF correspondente na mesma linha
                for idx, row in df.iterrows():
                    name = row["Nome"]
                    cpf = row["CPF ajustado"]
                    turma = aba
                    
                    found_name = False
                    found_cpf = False

                    for line in lines:
                        line_lower = line.lower()
                        if name.lower() in line_lower:
                            found_name = True
                        if cpf in line_lower:
                            found_cpf = True

                        # Se ambos são encontrados na mesma linha, salvar o resultado
                        if found_name and found_cpf:
                            search_results.append({
                                "Nome": name,
                                "Arquivo": pdf_path,
                                "Página": page_num + 1,  # Contagem de páginas começa em 1
                                "Turma": turma,
                                "CPF": "OK"
                            })
                            break

                    # Caso apenas o nome seja encontrado, mas o CPF não
                    if found_name and not found_cpf:
                        search_results.append({
                            "Nome": name,
                            "Arquivo": pdf_path,
                            "Página": page_num + 1,
                            "Turma": turma,
                            "CPF": "Não encontrado"
                        })
        # Imprimir para acompanhar o processo de checagem
        print(f"ok! {n}.{i}")

# Converter resultados para um DataFrame para facilitar o manuseio
results_df = pd.DataFrame(search_results)
# Exportar para um arquivo Excel
results_df.to_excel(f'Resultado.xlsx', index=False)

