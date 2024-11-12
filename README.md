# TEXCEL - TNMS to EXCEL Automation

**Descrição**  
Este projeto automatiza o processo de extração de dados de atenuação de rotas do TNMS e a inserção desses dados em uma planilha Excel. O objetivo é otimizar o tempo gasto com tarefas manuais, permitindo que o usuário faça o upload de um arquivo .txt extraído do TNMS, e o sistema processa automaticamente as informações, gerando uma planilha Excel pronta para uso.

**Tecnologias Usadas**  
- Python 3.11 ou superior  
- Bibliotecas: `pandas`, `openpyxl`, `tkinter`, `PyInstaller`  
- VS Code ou outro editor Python

## Funcionalidades
- Upload automático de arquivos .txt com dados do TNMS
- Processamento e extração dos valores de atenuação
- Inserção dos dados extraídos em uma planilha Excel
- Geração de relatórios automatizados
- Criação de novos arquivos Excel para cada execução, sem sobrescrever os anteriores
- Interface gráfica simples com Tkinter
- Redução do tempo de execução de 15 minutos para 1 minuto por rota

## Como Usar no VS Code

### Pré-requisitos
- Instale o [Python 3.11](https://www.python.org/downloads/) ou superior
- Instale as bibliotecas necessárias com o seguinte comando no terminal do VS Code:
  ```bash
  pip install pandas openpyxl tkinter pyinstaller

### Rodar o código e obter resposta no terminal

## Licença
Este projeto está licenciado sob a Licença MIT - consulte o arquivo [LICENSE](LICENSE) para mais detalhes.

