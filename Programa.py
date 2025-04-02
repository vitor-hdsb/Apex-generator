import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime
import os

class AplicativoMesclagemDados:
    def __init__(self, raiz):
        self.raiz = raiz
        self.raiz.title("Mesclagem de Dados")

        self.caminho_arquivo1 = tk.StringVar()
        self.caminho_arquivo2 = tk.StringVar()

        # Interface simples e direta
        tk.Label(raiz, text="Arquivo RFID (CSV):").pack()
        tk.Entry(raiz, textvariable=self.caminho_arquivo1, width=50).pack()
        tk.Button(raiz, text="Selecionar", command=self.carregar_arquivo1).pack()

        tk.Label(raiz, text="Arquivo de Dados (XLSX):").pack()
        tk.Entry(raiz, textvariable=self.caminho_arquivo2, width=50).pack()
        tk.Button(raiz, text="Selecionar", command=self.carregar_arquivo2).pack()

        tk.Button(raiz, text="Processar", command=self.processar_dados).pack(pady=10)

    def carregar_arquivo1(self):
        caminho = filedialog.askopenfilename(filetypes=[("Arquivos CSV", "*.csv")])
        if caminho:
            self.caminho_arquivo1.set(caminho)

    def carregar_arquivo2(self):
        caminho = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
        if caminho:
            self.caminho_arquivo2.set(caminho)

    def processar_dados(self):
        try:
            # Lendo o CSV e verificando o separador correto
            with open(self.caminho_arquivo1.get(), 'r') as arquivo:
                conteudo = arquivo.readlines()

            separador = ',' if ',' in conteudo[0] else ';'

            df1 = pd.read_csv(self.caminho_arquivo1.get(), delimiter=separador, dtype=str)
            if not {'Employee ID', 'Badge ID'}.issubset(df1.columns):
                messagebox.showerror("Erro", "Planilha1 precisa ter 'Employee ID' e 'Badge ID'")
                return

            # Agora carregando o XLSX
            df2 = pd.read_excel(self.caminho_arquivo2.get(), dtype=str)
            if not {'Nome', 'Employee ID'}.issubset(df2.columns):
                messagebox.showerror("Erro", "Planilha2 precisa ter 'Nome' e 'Employee ID'")
                return

            # Limpando dados para evitar erros
            df2 = df2.dropna(subset=['Nome'])

            # Separando nomes (caso não tenha sobrenome, deixa vazio)
            df2[['Primeiro Nome', 'Sobrenome']] = df2['Nome'].str.split(' ', n=1, expand=True)
            df2['Sobrenome'].fillna('', inplace=True)

            # Mesclando dados para buscar o Badge ID
            mesclado = pd.merge(df2, df1[['Employee ID', 'Badge ID']], on='Employee ID', how='left')

            # Selecionando apenas as colunas importantes
            resultado = mesclado[['Primeiro Nome', 'Sobrenome', 'Employee ID', 'Badge ID']]

            # Salvando na área de trabalho
            caminho_area_trabalho = r"C:\Users\vitor\Desktop"
            nome_arquivo = os.path.join(caminho_area_trabalho, f"dados_mesclados_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
            resultado.to_excel(nome_arquivo, index=False)

            messagebox.showinfo("Sucesso", f"Arquivo salvo em: {nome_arquivo}")

        except Exception as e:
            messagebox.showerror("Erro", str(e))

if __name__ == "__main__":
    raiz = tk.Tk()
    aplicativo = AplicativoMesclagemDados(raiz)
    raiz.mainloop()
