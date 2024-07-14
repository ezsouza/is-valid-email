import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import pandas as pd
import re
import requests
from dotenv import load_dotenv
import os

# Carregar a chave API do Hunter.io do arquivo .env
load_dotenv()
HUNTER_API_KEY = os.getenv('HUNTER_API_KEY')

# Função para validar a estrutura do email
def is_valid_email(email):
    regex = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(regex, email)

# Função para verificar emails usando Hunter.io
def verify_email_with_hunter(api_key, email):
    url = f"https://api.hunter.io/v2/email-verifier?email={email}&api_key={api_key}"
    response = requests.get(url)
    if response.status_code == 200:
        return response.json()
    return None

# Função para processar o arquivo Excel
def process_excel(file_path):
    df = pd.read_excel(file_path)

    # Verificando duplicados e estrutura de email
    df['reports'] = 'ok'
    df['reports'] = df['reports'].mask(df.duplicated(subset=['email']), 'duplicado')
    df['reports'] = df['reports'].mask(df['email'].apply(lambda x: not is_valid_email(x)), 'estrutura incorreta')

    # Lista de colunas que serão adicionadas ao DataFrame
    additional_columns = ['status', 'result', 'webmail', 'regexp', 'disposable', 'mx_records', 'smtp_server', 'smtp_check', 'accept_all', 'block', 'sources']
    for column in additional_columns:
        df[column] = None

    total_rows = len(df)
    progress_bar['maximum'] = total_rows

    # Verificando emails válidos e não duplicados com Hunter.io
    for index, row in df.iterrows():
        if df.at[index, 'reports'] == 'ok':
            email = row['email']
            result = verify_email_with_hunter(HUNTER_API_KEY, email)
            if result:
                hunter_data = result.get('data', {})
                for key, value in hunter_data.items():
                    if key in df.columns:
                        df.at[index, key] = value
                df.at[index, 'status'] = hunter_data.get('status', 'desconhecido')

        # Atualizando o progresso
        progress_bar['value'] = index + 1
        progress_label.config(text=f"Processando linha {index + 1} de {total_rows} ({(index + 1) / total_rows * 100:.2f}%)")
        root.update_idletasks()
        print(f"Processando linha {index + 1} de {total_rows}")

    # Salvando o arquivo atualizado
    output_file = file_path.replace('.xlsx', '_verified.xlsx')
    df.to_excel(output_file, index=False)
    print(f'Arquivo processado e salvo em: {output_file}')

    # Filtrando e salvando os emails válidos
    valid_emails_df = df[df['status'] == 'valid'].drop(columns=['reports'])
    valid_output_file = file_path.replace('.xlsx', '_valid_emails.xlsx')
    valid_emails_df.to_excel(valid_output_file, index=False)
    print(f'Emails válidos salvos em: {valid_output_file}')

# Função para abrir o dialog de seleção de arquivo e iniciar o processamento
def select_file():
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if filepath:
        file_path.set(filepath)
        file_label.config(text=f"Arquivo selecionado: {os.path.basename(filepath)}")
        process_excel(filepath)

# Configuração da GUI
root = tk.Tk()
root.title("Verificador de Emails")
root.resizable(False, False)

file_path = tk.StringVar()

tk.Button(root, text="Selecionar Arquivo Excel", command=select_file).grid(row=0, column=0, columnspan=2, padx=10, pady=10)

file_label = tk.Label(root, text="Nenhum arquivo selecionado")
file_label.grid(row=1, column=0, columnspan=2, padx=10, pady=10)

progress_label = tk.Label(root, text="")
progress_label.grid(row=2, column=0, columnspan=2, padx=10, pady=10)

progress_bar = ttk.Progressbar(root, length=300)
progress_bar.grid(row=3, column=0, columnspan=2, padx=10, pady=10)

root.mainloop()