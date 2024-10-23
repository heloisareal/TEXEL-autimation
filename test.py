import os
from openpyxl import load_workbook

# Criar a pasta 'arquivos' se ela não existir
if not os.path.exists('arquivos'):
    os.makedirs('arquivos')

# Abrir o arquivo com os dados da tabela
with open('table.txt', 'r') as file:
    lines = file.readlines()

# Criar listas para armazenar os valores de cada coluna
start_node_values = []
end_node_values = []
span_loss_values = []
eol_att_values = []

# Iterar sobre as linhas do arquivo e extrair os valores das colunas relevantes
for line in lines[1:]:  # Ignorar o cabeçalho
    columns = line.split()  # Separar as colunas por espaços
    if len(columns) == 4:
        start_node_values.append(columns[0])  # Pegar a primeira coluna (Start node)
        end_node_values.append(columns[1])    # Pegar a segunda coluna (End node)
        span_loss_values.append(columns[2])   # Pegar a terceira coluna (Span loss)
        eol_att_values.append(columns[3])     # Pegar a quarta coluna (EOL ATT THR)

# Reordenar os valores de eol_att_values e span_loss_values de acordo com o padrão solicitado
def reorder_values(values):
    reordered_values = []
    for i in range(0, len(values), 4):
        if i < len(values):
            reordered_values.append(values[i])   # Linha 1
        if i + 2 < len(values):
            reordered_values.append(values[i + 2])  # Linha 3
        if i + 1 < len(values):
            reordered_values.append(values[i + 1])  # Linha 2
        if i + 3 < len(values):
            reordered_values.append(values[i + 3])  # Linha 4
    return reordered_values

reordered_eol_values = reorder_values(eol_att_values)
reordered_span_loss_values = reorder_values(span_loss_values)

# Salvar os valores extraídos em arquivos na pasta 'arquivos'
with open('arquivos/start_node.txt', 'w') as start_file:
    for value in start_node_values:
        start_file.write(value + '\n')

with open('arquivos/end_node.txt', 'w') as end_file:
    for value in end_node_values:
        end_file.write(value + '\n')

with open('arquivos/span_loss.txt', 'w') as span_file:
    for value in reordered_span_loss_values:
        span_file.write(value + '\n')

with open('arquivos/eol.txt', 'w') as eol_file:
    for value in reordered_eol_values:
        eol_file.write(value + '\n')

print("Valores extraídos e salvos na pasta 'arquivos' com sucesso, com EOL ATT e Span Loss reordenados.")

# Carregar o arquivo eol.txt para pegar os valores a serem inseridos no Excel
with open('arquivos/eol.txt', 'r') as eol_file:
    eol_values = eol_file.readlines()

# Remover possíveis quebras de linha
eol_values = [value.strip() for value in eol_values]

# Carregar os valores de span_loss.txt
with open('arquivos/span_loss.txt', 'r') as span_file:
    span_loss_values = span_file.readlines()

# Remover possíveis quebras de linha
span_loss_values = [value.strip() for value in span_loss_values]

# Abrir o arquivo controle-potencia-poprecife.xlsx
excel_path = 'controle-potencia-poprecife.xlsx'
workbook = load_workbook(excel_path)

# Selecionar a planilha ativa (ou a planilha específica, se souber o nome)
sheet = workbook.active

# Posições-alvo para inserir os dados
positions = [(2, 'F', 'J'), (6, 'F', 'J'), (10, 'F', 'J'), (14, 'F', 'J'), (18, 'F', 'J')]

# Preencher as células F2, J2, F6, J6, etc. com os valores de eol.txt
for i, (row, col_f, col_j) in enumerate(positions):
    index = i * 2  # Pegamos 2 valores de cada vez
    if index < len(eol_values):  # Certificar que ainda há valores a serem usados
        sheet[f'{col_f}{row}'] = eol_values[index]      # Ex: F2, F6...
    if index + 1 < len(eol_values):  # Certificar que há um segundo valor
        sheet[f'{col_j}{row}'] = eol_values[index + 1]  # Ex: J2, J6...

# Preencher as células F3, J3, F7, J7, etc. com os valores de span_loss.txt
for i, (row, col_f, col_j) in enumerate(positions):
    index = i * 2  # Pegamos 2 valores de cada vez
    if index < len(span_loss_values):  # Certificar que ainda há valores a serem usados
        sheet[f'{col_f}{row + 1}'] = span_loss_values[index]      # Ex: F3, F7...
    if index + 1 < len(span_loss_values):  # Certificar que há um segundo valor
        sheet[f'{col_j}{row + 1}'] = span_loss_values[index + 1]  # Ex: J3, J7...

# Salvar o arquivo Excel sem sobrescrever o conteúdo existente
workbook.save(excel_path)

print(f"Valores de eol.txt e span_loss.txt adicionados ao arquivo {excel_path} com sucesso.")
