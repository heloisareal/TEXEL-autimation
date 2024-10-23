import openpyxl
from openpyxl.styles import PatternFill

# Função para extrair dados do arquivo txt
def extrair_dados(txt_path):
    with open(txt_path, 'r') as file:
        content = file.readlines()

    coluna1 = []  # Para armazenar os números da primeira coluna
    coluna2 = []  # Para armazenar os números da segunda coluna

    for line in content:
        if line.strip() and line[0].isdigit():  # Verifica se a linha começa com um número
            numeros = line.split()
            if len(numeros) >= 2:  # Verifica se há pelo menos 2 números
                coluna1.append(int(numeros[0]))  # Adiciona o primeiro número à coluna 1 como inteiro
                coluna2.append(int(numeros[1]))  # Adiciona o segundo número à coluna 2 como inteiro

    return coluna1, coluna2

# Função para adicionar dados em um Excel existente
def adicionar_dados_no_excel(coluna1, coluna2, excel_path):
    # Carregar o arquivo Excel existente
    workbook = openpyxl.load_workbook(excel_path)
    sheet = workbook.active

    # Inicializa a linha de escrita, considerando a última linha preenchida
    linha_atual = sheet.max_row + 1  # Começa a partir da última linha preenchida

    # Escreve os dados em grupos de três
    for i in range(0, len(coluna1), 3):
        # Escreve até três linhas de dados
        for j in range(3):
            if i + j < len(coluna1):  # Verifica se ainda há dados para escrever
                # Preenche os valores
                sheet[f'H{linha_atual}'] = coluna1[i + j]  # Coluna H
                sheet[f'J{linha_atual}'] = coluna2[i + j]  # Coluna J

                # Aplica a cor baseada no valor
                for col in [f'H{linha_atual}', f'J{linha_atual}']:
                    if sheet[col].value > 10:
                        sheet[col].fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")  # Verde
                    else:
                        sheet[col].fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")  # Vermelho

                linha_atual += 1  # Move para a próxima linha
        # Pula quatro linhas após cada grupo de três
        linha_atual += 4

    # Salvar o arquivo Excel
    workbook.save(excel_path)

# Caminhos dos arquivos
txt_path = 'C:\\Users\\hreal\\OneDrive - Infinera\\Desktop\\aut-rotas\\teste.txt'
excel_path = 'C:\\Users\\hreal\\OneDrive - Infinera\\Desktop\\aut-rotas\\done.xlsx'

# Extração dos dados e adição no Excel existente
coluna1, coluna2 = extrair_dados(txt_path)
adicionar_dados_no_excel(coluna1, coluna2, excel_path)

# Mensagem de confirmação
print("Os dados foram extraídos e adicionados ao Excel com sucesso! Programa finalizado!")
