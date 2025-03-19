import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime
import os

class DataMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Merger")

        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()

        tk.Label(root, text="Arquivo RFID (CSV):").pack()
        tk.Entry(root, textvariable=self.file1_path, width=50).pack()
        tk.Button(root, text="Selecionar", command=self.load_file1).pack()

        tk.Label(root, text="Arquivo de Dados (XLSX):").pack()
        tk.Entry(root, textvariable=self.file2_path, width=50).pack()
        tk.Button(root, text="Selecionar", command=self.load_file2).pack()

        tk.Button(root, text="Processar", command=self.process_data).pack(pady=10)

    def load_file1(self):
        path = filedialog.askopenfilename(filetypes=[["CSV Files", "*.csv"]])
        if path:
            self.file1_path.set(path)

    def load_file2(self):
        path = filedialog.askopenfilename(filetypes=[["Excel Files", "*.xlsx"]])
        if path:
            self.file2_path.set(path)

    def process_data(self):
        try:
            # Importando planilha RFID e convertendo CSV em colunas
            with open(self.file1_path.get(), 'r') as file:
                content = file.readlines()

            # Detectando o separador (vírgula ou ponto e vírgula)
            delimiter = ',' if ',' in content[0] else ';'

            df1 = pd.read_csv(self.file1_path.get(), delimiter=delimiter, dtype=str)
            if not {'Employee ID', 'Badge ID'}.issubset(df1.columns):
                messagebox.showerror("Erro", "Planilha1 deve conter as colunas 'Employee ID' e 'Badge ID'")
                return

            # Importando planilha de dados
            df2 = pd.read_excel(self.file2_path.get(), dtype=str)
            if not {'Nome', 'Employee ID'}.issubset(df2.columns):
                messagebox.showerror("Erro", "Planilha2 deve conter as colunas 'Nome' e 'Employee ID'")
                return

            # Removendo valores nulos antes de dividir os nomes
            df2 = df2.dropna(subset=['Nome'])

            # Separando nomes (tratando nomes sem sobrenome)
            df2[['First Name', 'Last Name']] = df2['Nome'].str.split(' ', n=1, expand=True)
            df2['Last Name'].fillna('', inplace=True)

            # Procv para buscar Badge ID
            merged = pd.merge(df2, df1[['Employee ID', 'Badge ID']], on='Employee ID', how='left')

            # Selecionando colunas finais
            result = merged[['First Name', 'Last Name', 'Employee ID', 'Badge ID']]

            # **Definindo o caminho fixo para salvar na área de trabalho**
            desktop_path = r"C:\Users\vitor\Desktop"
            filename = os.path.join(desktop_path, f"merged_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")

            # Salvando arquivo
            result.to_excel(filename, index=False)

            messagebox.showinfo("Sucesso", f"Arquivo salvo em: {filename}")

        except Exception as e:
            messagebox.showerror("Erro", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = DataMergerApp(root)
    root.mainloop()
