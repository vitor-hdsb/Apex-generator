import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
import pandas as pd
from fpdf import FPDF
from datetime import datetime
import os

class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 14)
        self.cell(0, 10, 'TERMO DE CONSENTIMENTO', ln=True, align='C')
        self.ln(10)

    def footer(self):
        pass

def gerar_pdf_termos(resultados, df1, caminho_saida_pdf):
    pdf = PDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    for _, row in resultados.iterrows():
        nome = f"{row['First Name']} {row['Last Name']}"
        employee_id = row['Employee ID']
        data_atual = datetime.now().strftime('%d/%m/%Y')

        login_info = df1.loc[df1['Employee ID'] == employee_id, 'Login']
        if not login_info.empty:
            login = login_info.values[0]
        else:
            login = "NÃO ENCONTRADO"

        texto_termo = (
            f"TERMO DE CONSENTIMENTO PARA UTILIZAÇÃO DE SISTEMAS INTERNOS\n\n"
            f"Eu, {nome}, portador do login {login} e Employee ID {employee_id}, "
            f"declaro que recebi treinamento sobre a utilização dos sistemas internos da empresa, "
            f"estando ciente das políticas de segurança da informação, responsabilidade sobre o acesso, "
            f"sigilo de dados e boas práticas de utilização.\n\n"
            f"Comprometo-me a utilizar os sistemas de forma ética, responsável e de acordo com as normas "
            f"estabelecidas pela companhia, ciente de que qualquer desvio poderá acarretar medidas disciplinares "
            f"e legais cabíveis.\n\n"
            f"Por ser verdade, firmo o presente termo para que produza seus efeitos legais.\n\n"
            f"Data: {data_atual}\n\n"
            f"Assinatura: ________________________________________________"
        )

        pdf.add_page()
        pdf.set_font('Arial', '', 12)
        pdf.multi_cell(0, 10, texto_termo)

    pdf.output(caminho_saida_pdf)

class DataMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Amazon WHS - Apex Importer")
        self.root.geometry("550x650")

        # Criar notebook para abas
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # Aba 1: Geração Original
        self.tab1 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab1, text="Apex Generator")

        # Aba 2: Atualização da Base
        self.tab2 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab2, text="Atualização da Base")

        # Variáveis para Aba 1 (Original)
        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()

        # Variáveis para Aba 2 (Atualização)
        self.csv_current_path = tk.StringVar()
        self.xlsx_apex_base_path = tk.StringVar()

        self.setup_tab1()
        self.setup_tab2()
        self.load_logo()

    def load_logo(self):
        try:
            logo_path = r"C:\Users\vitor\Desktop\Apex-generator\image.png"

            if not os.path.exists(logo_path):
                raise FileNotFoundError(f"Logo não encontrada em: {logo_path}")

            logo_image = Image.open(logo_path)

            if hasattr(Image, 'Resampling'):
                logo_image = logo_image.resize((200, 80), Image.Resampling.LANCZOS)
            else:
                logo_image = logo_image.resize((200, 80), Image.ANTIALIAS)
  
            self.logo = ImageTk.PhotoImage(logo_image)
            
            # Adicionar logo nas duas abas
            tk.Label(self.tab1, image=self.logo).pack(pady=10)
            tk.Label(self.tab2, image=self.logo).pack(pady=10)

        except Exception as e:
            print("Erro ao carregar logo:", str(e))
            tk.Label(self.tab1, text="[Logo não carregada]").pack()
            tk.Label(self.tab2, text="[Logo não carregada]").pack()

    def setup_tab1(self):
        """Configurar a aba original (Apex Generator)"""
        # Título
        tk.Label(self.tab1, text="Geração de Arquivos Apex", font=("Arial", 14, "bold")).pack(pady=10)

        # Arquivo RFID
        tk.Label(self.tab1, text="Arquivo RFID (CSV):").pack()
        tk.Entry(self.tab1, textvariable=self.file1_path, width=60).pack()
        tk.Button(self.tab1, text="Selecionar arquivo", command=self.load_file1).pack(pady=5)

        # Planilha Associados
        tk.Label(self.tab1, text="Planilha Associados (XLSX):").pack()
        tk.Entry(self.tab1, textvariable=self.file2_path, width=60).pack()
        tk.Button(self.tab1, text="Selecionar arquivo", command=self.load_file2).pack(pady=5)

        # Botão processar
        tk.Button(self.tab1, text="Mesclar dados", command=self.process_data, 
                 bg="#0077cc", fg="white", font=("Arial", 10, "bold")).pack(pady=20)

    def setup_tab2(self):
        """Configurar a nova aba de atualização"""
        # Título
        tk.Label(self.tab2, text="Atualização da Base Apex", font=("Arial", 14, "bold")).pack(pady=10)

        # Instruções
        instruction_text = (
            "Esta função atualiza a base do Apex comparando:\n"
            "• Arquivo CSV atual com funcionários ativos\n"
            "• Base de dados atual do Apex\n\n"
            "Resultado: Nova planilha com status atualizados"
        )
        tk.Label(self.tab2, text=instruction_text, justify="left", 
                font=("Arial", 9), wraplength=500).pack(pady=10)

        # Arquivo CSV Atual
        tk.Label(self.tab2, text="Arquivo CSV Atual (Funcionários Ativos):").pack()
        tk.Entry(self.tab2, textvariable=self.csv_current_path, width=60).pack()
        tk.Button(self.tab2, text="Selecionar CSV", command=self.load_csv_current).pack(pady=5)

        # Base Apex
        tk.Label(self.tab2, text="Base de Dados Apex (XLSX):").pack()
        tk.Entry(self.tab2, textvariable=self.xlsx_apex_base_path, width=60).pack()
        tk.Button(self.tab2, text="Selecionar Base Apex", command=self.load_xlsx_apex_base).pack(pady=5)

        # Botão processar
        tk.Button(self.tab2, text="Atualizar Base", command=self.update_apex_base, 
                 bg="#00AA44", fg="white", font=("Arial", 10, "bold")).pack(pady=20)

    # Métodos da Aba 1 (Original)
    def load_file1(self):
        path = filedialog.askopenfilename(filetypes=[["CSV Files", "*.csv"]])
        if path:
            self.file1_path.set(path)

    def load_file2(self):
        path = filedialog.askopenfilename(filetypes=[["Excel Files", "*.xlsx"]])
        if path:
            self.file2_path.set(path)

    # Métodos da Aba 2 (Atualização)
    def load_csv_current(self):
        path = filedialog.askopenfilename(filetypes=[["CSV Files", "*.csv"]])
        if path:
            self.csv_current_path.set(path)

    def load_xlsx_apex_base(self):
        path = filedialog.askopenfilename(filetypes=[["Excel Files", "*.xlsx"]])
        if path:
            self.xlsx_apex_base_path.set(path)

    def process_data(self):
        """Processo original da Aba 1"""
        try:
            with open(self.file1_path.get(), 'r') as file:
                content = file.readlines()
                delimiter = ',' if ',' in content[0] else ';'

            df1 = pd.read_csv(self.file1_path.get(), delimiter=delimiter, dtype=str)
            if not {'Employee ID', 'Badge ID', 'Login'}.issubset(df1.columns):
                messagebox.showerror("Erro", "Arquivo CSV deve conter 'Employee ID', 'Badge ID' e 'Login'")
                return

            df2 = pd.read_excel(self.file2_path.get(), skiprows=1, dtype=str)
            df2 = df2.drop(columns=[col for col in df2.columns if col.strip().lower() == 'employee id'], errors='ignore')
            if not {'Nome', 'Login'}.issubset(df2.columns):
                messagebox.showerror("Erro", "Arquivo Excel deve conter 'Nome' e 'Login'")
                return

            df2 = df2.dropna(subset=['Nome'])

            df2[['First Name', 'Last Name']] = df2['Nome'].str.split(' ', n=1, expand=True)
            df2['Last Name'].fillna('', inplace=True)

            merged = pd.merge(df2, df1[['Login', 'Employee ID', 'Badge ID']], on='Login', how='left')
            result = merged[['First Name', 'Last Name', 'Employee ID', 'Badge ID']]

            result = result.rename(columns={'Badge ID': 'Badge #'})

            result['Language'] = 'English'
            result['Active'] = 'Y'
            result['Reporting Group'] = 'Associados'
            result['User group 2'] = ''
            result['User group 3'] = ''
            result['User group 4'] = ''
            result['User group 5'] = ''
            result['Expiration Date'] = ''

            result = result[
                ['Employee ID', 'First Name', 'Last Name', 'Badge #', 
                 'Language', 'Active', 'Reporting Group', 
                 'User group 2', 'User group 3', 'User group 4', 'User group 5', 'Expiration Date']
            ]

            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], initialfile=f"planilha_completa_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
            if not save_path:
                messagebox.showwarning("Aviso", "Arquivo não foi salvo.")
                return

            result.to_excel(save_path, index=False)

            erros = merged[merged['Badge ID'].isna()]
            if not erros.empty:
                error_path = os.path.splitext(save_path)[0] + "_erros.xlsx"
                erros.to_excel(error_path, index=False)

            pdf_path = os.path.splitext(save_path)[0] + "_termos.pdf"
            gerar_pdf_termos(result, df1, pdf_path)

            instance_df = pd.DataFrame({
                'Instance Name': result['First Name'] + ' ' + result['Last Name'],
                'Instance Num': result['Employee ID'],
                'Desc': ''
            })
            instance_path = os.path.splitext(save_path)[0] + "_Instances.xlsx"
            instance_df.to_excel(instance_path, index=False)

            messagebox.showinfo(
                "Concluído",
                f"Arquivo Excel salvo em:\n{save_path}\n\n"
                f"Arquivo Instances salvo em:\n{instance_path}\n\n"
                f"Termo de consentimento salvo em:\n{pdf_path}"
            )

        except Exception as e:
            messagebox.showerror("Erro", str(e))

    def update_apex_base(self):
        """Novo processo da Aba 2 - Atualização da Base"""
        try:
            # Validar se os arquivos foram selecionados
            if not self.csv_current_path.get() or not self.xlsx_apex_base_path.get():
                messagebox.showerror("Erro", "Selecione ambos os arquivos!")
                return

            # Carregar CSV atual (funcionários ativos)
            with open(self.csv_current_path.get(), 'r') as file:
                content = file.readlines()
                delimiter = ',' if ',' in content[0] else ';'

            df_csv_current = pd.read_csv(self.csv_current_path.get(), delimiter=delimiter, dtype=str)
            
            # Verificar se tem as colunas necessárias no CSV
            required_csv_cols = {'Employee ID', 'First Name', 'Last Name', 'Badge ID'}
            if not required_csv_cols.issubset(df_csv_current.columns):
                # Se não tem First Name e Last Name, verificar se tem Nome para separar
                if 'Nome' in df_csv_current.columns:
                    df_csv_current[['First Name', 'Last Name']] = df_csv_current['Nome'].str.split(' ', n=1, expand=True)
                    df_csv_current['Last Name'].fillna('', inplace=True)
                else:
                    messagebox.showerror("Erro", "CSV deve conter 'Employee ID', 'First Name', 'Last Name', 'Badge ID' ou 'Nome'")
                    return

            # Carregar base do Apex (XLSX)
            df_apex_base = pd.read_excel(self.xlsx_apex_base_path.get(), dtype=str)
            
            # Verificar colunas da base Apex - CORRIGIDO com asteriscos
            apex_required_cols = {'Employee ID*', 'First Name*', 'Last Name*', 'Badge #', 'Active*'}
            if not apex_required_cols.issubset(df_apex_base.columns):
                messagebox.showerror("Erro", "Base Apex deve conter 'Employee ID*', 'First Name*', 'Last Name*', 'Badge #', 'Active*'")
                return

            # Criar DataFrame resultado baseado na estrutura da base Apex
            result_df = df_apex_base.copy()

            # 1. Registros no CSV mas não na base Apex (NOVOS USUÁRIOS)
            csv_employee_ids = set(df_csv_current['Employee ID'].fillna('').astype(str))
            apex_employee_ids = set(df_apex_base['Employee ID*'].fillna('').astype(str))

            novos_usuarios = csv_employee_ids - apex_employee_ids
            
            if novos_usuarios:
                # Filtrar novos usuários do CSV
                df_novos = df_csv_current[df_csv_current['Employee ID'].isin(novos_usuarios)].copy()
                
                # Renomear colunas para corresponder ao formato Apex com asteriscos
                df_novos = df_novos.rename(columns={
                    'Employee ID': 'Employee ID*',
                    'First Name': 'First Name*',
                    'Last Name': 'Last Name*',
                    'Badge ID': 'Badge #'
                })
                
                # Adicionar colunas padrão para novos usuários
                df_novos['Language'] = 'English'
                df_novos['Active*'] = 'Y'
                df_novos['Reporting Group'] = 'Associados'
                df_novos['User Group 2'] = ''
                df_novos['User Group 3'] = ''
                df_novos['User Group 4'] = ''
                df_novos['User Group 5'] = ''
                df_novos['Expiration Date'] = ''

                # Garantir que as colunas estejam na ordem correta
                colunas_ordenadas = [
                    'Employee ID*', 'First Name*', 'Last Name*', 'Badge #', 
                    'Language', 'Active*', 'Reporting Group', 
                    'User Group 2', 'User Group 3', 'User Group 4', 'User Group 5', 'Expiration Date'
                ]
                
                # Adicionar colunas que possam estar faltando
                for col in colunas_ordenadas:
                    if col not in df_novos.columns:
                        df_novos[col] = ''

                df_novos = df_novos[colunas_ordenadas]
                
                # Adicionar novos usuários ao resultado
                result_df = pd.concat([result_df, df_novos], ignore_index=True)

            # 2. Registros não no CSV mas na base Apex (DESATIVAR)
            usuarios_desativar = apex_employee_ids - csv_employee_ids
            result_df.loc[result_df['Employee ID*'].isin(usuarios_desativar), 'Active*'] = 'N'

            # 3. Registros no CSV e na base Apex mas inativos (REATIVAR)
            usuarios_comuns = csv_employee_ids & apex_employee_ids
            mask_reativar = (result_df['Employee ID*'].isin(usuarios_comuns)) & (result_df['Active*'] == 'N')
            result_df.loc[mask_reativar, 'Active*'] = 'Y'

            # Remover duplicatas baseadas em Employee ID*
            result_df = result_df.drop_duplicates(subset=['Employee ID*'], keep='last')

            # Salvar resultado
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx", 
                filetypes=[("Excel files", "*.xlsx")], 
                initialfile=f"apex_base_atualizada_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            )
            
            if not save_path:
                messagebox.showwarning("Aviso", "Arquivo não foi salvo.")
                return

            result_df.to_excel(save_path, index=False)

            # Gerar relatório de mudanças
            total_registros = len(result_df)
            novos_count = len(novos_usuarios) if novos_usuarios else 0
            desativados_count = len(usuarios_desativar) if usuarios_desativar else 0
            reativados_count = len(result_df[mask_reativar]) if 'mask_reativar' in locals() else 0

            messagebox.showinfo(
                "Atualização Concluída",
                f"Base do Apex atualizada com sucesso!\n\n"
                f"Arquivo salvo em:\n{save_path}\n\n"
                f"Resumo das alterações:\n"
                f"• Total de registros: {total_registros}\n"
                f"• Novos usuários: {novos_count}\n"
                f"• Usuários desativados: {desativados_count}\n"
                f"• Usuários reativados: {reativados_count}"
            )

        except Exception as e:
            messagebox.showerror("Erro", f"Erro no processamento: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()

    messagebox.showinfo(
        "Instruções - Amazon WHS Apex Importer",
        "🔹 ABA APEX GENERATOR:\n"
        "1. Selecione o arquivo CSV com 'Employee ID', 'Badge ID' e 'Login'\n"
        "2. Selecione a planilha Excel com 'Nome' e 'Login'\n"
        "3. Clique em 'Mesclar dados' para gerar os arquivos\n\n"
        "🔹 ABA ATUALIZAÇÃO DA BASE:\n"
        "1. Selecione o CSV atual com funcionários ativos\n"
        "2. Selecione a base de dados atual do Apex (XLSX)\n"
        "3. Clique em 'Atualizar Base' para sincronizar\n\n"
        "OBS: A planilha Excel deve ter cabeçalho na linha 2."
    )
    
    root.deiconify()
    app = DataMergerApp(root)
    root.mainloop()