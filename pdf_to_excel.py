import pdfplumber
import pandas as pd
import sys
from collections import defaultdict


def consolidate_data(all_data):
    """
    Consolida dados onde a primeira coluna está vazia, 
    agrupando as informações com a linha anterior
    """
    consolidated = []
    
    for row in all_data:
        if not row or not row[0] or (isinstance(row[0], str) and not row[0].strip()):
            # Se a primeira coluna está vazia, agrupar com a linha anterior
            if consolidated:
                # Adicionar as informações desta linha à linha anterior
                for i in range(1, len(row)):
                    if row[i] and isinstance(row[i], str) and row[i].strip():
                        # Juntar com quebra de linha se houver conteúdo
                        if consolidated[-1][i]:
                            consolidated[-1][i] += '\n' + row[i]
                        else:
                            consolidated[-1][i] = row[i]
        else:
            # Se a primeira coluna tem valor, é uma nova linha
            consolidated.append(row[:])
    
    return consolidated


def transpose_d_values(all_data):
    """
    Transpõe valores da 3ª coluna que terminam em 'D' para uma 4ª coluna
    """
    transposed = []
    
    for row in all_data:
        if len(row) >= 3:
            col3_value = row[2]
            
            # Verificar se o valor termina em 'D'
            if isinstance(col3_value, str) and col3_value.strip().endswith('D'):
                # Adicionar a 4ª coluna se não existir
                while len(row) < 4:
                    row.append('')
                
                # Mover valor da 3ª para 4ª coluna
                row[3] = col3_value
                row[2] = ''
        
        transposed.append(row)
    
    return transposed


def pdf_to_excel(pdf_file_path, excel_file_path):
    """
    Extrai dados do PDF usando posições para separar em colunas estruturadas
    """
    all_data = []
    
    with pdfplumber.open(pdf_file_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            # Tentar extrair tabelas estruturadas primeiro
            tables = page.extract_tables()
            
            if tables:
                print(f"  Página {page_num}: Usando {len(tables)} tabela(s) estruturada(s)")
                for table in tables:
                    for row in table:
                        all_data.append(row)
            else:
                # Usar posições das palavras para reconstruir colunas
                print(f"  Página {page_num}: Usando análise de posições para extrair colunas")
                
                words = page.extract_words()
                if not words:
                    continue
                
                # Agrupar palavras por linha (similar top position)
                lines = defaultdict(list)
                line_threshold = 3  # pixels de tolerância
                
                for word in words:
                    # Encontrar a linha mais próxima
                    found = False
                    for line_top in list(lines.keys()):
                        if abs(word['top'] - line_top) < line_threshold:
                            lines[line_top].append(word)
                            found = True
                            break
                    
                    if not found:
                        lines[word['top']].append(word)
                
                # Processar cada linha
                for line_top in sorted(lines.keys()):
                    line_words = sorted(lines[line_top], key=lambda w: w['x0'])
                    
                    # Agrupar palavras por coluna (similar x0 position)
                    columns = []
                    col_threshold = 20  # pixels de tolerância
                    
                    for word in line_words:
                        found = False
                        for col in columns:
                            if abs(word['x0'] - col[0]['x0']) < col_threshold:
                                col.append(word)
                                found = True
                                break
                        
                        if not found:
                            columns.append([word])
                    
                    # Juntar as palavras de cada coluna
                    row_data = []
                    for col in columns:
                        col_text = ' '.join(w['text'] for w in col)
                        row_data.append(col_text)
                    
                    # Adicionar apenas linhas com conteúdo significativo
                    if row_data and any(len(cell.strip()) > 0 for cell in row_data):
                        all_data.append(row_data)
    
    # Consolidar dados onde primeira coluna está vazia
    all_data = consolidate_data(all_data)
    
    # Transpor valores da 3ª coluna que terminam em 'D' para 4ª coluna
    all_data = transpose_d_values(all_data)
    
    # Criar DataFrame e salvar em Excel
    if all_data:
        # Encontrar o número máximo de colunas
        max_cols = max(len(row) for row in all_data)
        
        # Preencher linhas com menos colunas
        for row in all_data:
            while len(row) < max_cols:
                row.append('')
        
        # Criar DataFrame com colunas
        columns = [f'Coluna {i+1}' for i in range(max_cols)]
        df = pd.DataFrame(all_data, columns=columns)
        
        # Salvar em Excel
        df.to_excel(excel_file_path, sheet_name='Dados', index=False)
        print(f"\n✓ Arquivo '{excel_file_path}' criado com sucesso!")
        print(f"  Total de linhas: {len(all_data)}")
        print(f"  Total de colunas: {max_cols}\n")
    else:
        print("✗ Nenhum conteúdo encontrado no PDF")


# Permitir usar como script ou importar como módulo
if __name__ == "__main__":
    if len(sys.argv) > 1:
        # Se passado argumento, usar esse arquivo
        pdf_file = sys.argv[1]
        excel_file = sys.argv[2] if len(sys.argv) > 2 else pdf_file.replace('.pdf', '_Convertido.xlsx')
    else:
        # Caso contrário, converter ambos os PDFs
        print("Convertendo PDFs para Excel...\n")
        
        # Converter MercadoPago.pdf
        print("1. Convertendo MercadoPago.pdf...")
        pdf_to_excel('MercadoPago.pdf', 'MercadoPago_Convertido.xlsx')
        
        # Converter sicoob
        import os
        sicoob_file = None
        for f in os.listdir('.'):
            if f.startswith('sicoob') and f.endswith('.pdf'):
                sicoob_file = f
                break
        
        if sicoob_file:
            print(f"2. Convertendo {sicoob_file}...")
            excel_name = sicoob_file.replace('.pdf', '_Convertido.xlsx')
            pdf_to_excel(sicoob_file, excel_name)
        else:
            print("2. Nenhum arquivo sicoob encontrado")
        
        sys.exit(0)
    
    print(f"Convertendo {pdf_file} para {excel_file}...\n")
    pdf_to_excel(pdf_file, excel_file)
