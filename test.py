import tkinter as tk
from tkinter import filedialog, messagebox
import os
import openpyxl

# Função para carregar o arquivo .txt
def load_txt_file():
    file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
    if file_path:
        process_txt_file(file_path)

# Função para processar o arquivo .txt e gerar a tabela Excel
def process_txt_file(input_file):
    # Caminho para o arquivo Excel modelo
    excel_modelo = 'modelo-alterarnome.xlsx'
    
    # Criando o diretório de saída 'downloads' (será a pasta de Downloads do usuário)
    output_dir = os.path.expanduser("~/Downloads")
    
    # Caminho do novo arquivo Excel a ser salvo
    new_excel_file = os.path.join(output_dir, f"modelo_preenchido_{len(os.listdir(output_dir)) + 1}.xlsx")

    # Carregando os dados do arquivo .txt
    data = []
    try:
        with open(input_file, 'r') as file:
            lines = file.readlines()

        for line in lines:
            parts = line.split()
            if len(parts) < 4:
                continue
            try:
                order = int(parts[0])
                end_node = parts[1]
                span_loss = float(parts[2])
                eol_att = float(parts[3])
                data.append((order, end_node, span_loss, eol_att))
            except ValueError:
                continue

        # Ordenando os dados pela coluna 'Ordem'
        data_sorted = sorted(data, key=lambda x: x[0])

        # Extraindo os valores de End node, EOL ATT THR e Span loss
        endnode_values = [item[1] for item in data_sorted]
        eolatt_values = [str(item[3]) for item in data_sorted]
        spanloss_values = [str(item[2]) for item in data_sorted]

        # Parte para inserir os dados no arquivo Excel
        workbook = openpyxl.load_workbook(excel_modelo)
        sheet = workbook.active

        # Inserindo os valores de EOL ATT THR
        row = 2
        for i in range(0, len(eolatt_values), 2):
            sheet[f'F{row}'] = eolatt_values[i]
            if i + 1 < len(eolatt_values):
                sheet[f'J{row}'] = eolatt_values[i + 1]
            row += 4

        # Inserindo os valores de Span Loss
        row = 3
        for i in range(0, len(spanloss_values), 2):
            sheet[f'F{row}'] = spanloss_values[i]
            if i + 1 < len(spanloss_values):
                sheet[f'J{row}'] = spanloss_values[i + 1]
            row += 4

        # Inserindo os valores de End node
        row = 2
        for i in range(0, len(endnode_values), 2):
            sheet[f'D{row}'] = endnode_values[i]
            if i + 1 < len(endnode_values):
                sheet[f'L{row}'] = endnode_values[i + 1]
            row += 4

        # Salvando o arquivo Excel preenchido
        workbook.save(new_excel_file)
        
        # Mostrando a mensagem de sucesso
        messagebox.showinfo("Sucesso", f"Arquivo Excel gerado com sucesso!\nSalvo em: {new_excel_file}")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao processar o arquivo: {e}")

# Função para exibir instruções
def show_instructions():
    instructions = """
    1. Clique em 'Carregar Tabela' para selecionar o arquivo table.txt com os dados.
    2. O programa irá processar os dados e gerar um arquivo Excel preenchido.
    3. O arquivo gerado será salvo na sua pasta de Downloads com um nome único.
    4. O arquivo Excel modelo (modelo-alterarnome.xlsx) permanecerá inalterado para futuras edições.
    """
    messagebox.showinfo("Instruções", instructions)

# Criando a janela principal da interface
root = tk.Tk()
root.title("Processador de Dados para Excel")

# Definindo o tamanho da janela
root.geometry("400x300")

# Adicionando um botão para carregar a tabela
load_button = tk.Button(root, text="Carregar Tabela", command=load_txt_file, width=30)
load_button.pack(pady=20)

# Adicionando um botão para mostrar as instruções
instructions_button = tk.Button(root, text="Instruções", command=show_instructions, width=30)
instructions_button.pack(pady=10)

# Rodando a interface
root.mainloop()
