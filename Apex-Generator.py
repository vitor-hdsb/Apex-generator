import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import pandas as pd
from fpdf import FPDF
from datetime import datetime
import os

class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 14)
        self.cell(0, 10, 'TERMO DE CONSENTIMENTO PARA USO DE SISTEMAS INTERNOS', ln=True, align='C')
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
            f"Eu, {nome}, portador do login {login} e Employee ID {employee_id}, declaro estar ciente que: \n"
            f"1. A empresa utiliza um sistema eletrônico de distribuição e registro de entrega de EPI's (Equipamentos individuais de proteção), através de máquinas específicas para este fim.\n"
            f"2. O sistema grava num banco de dados as seguintes informações, sempre que for processada a retirada de um EPI: Nome, matrícula, data, quantidade, descrição e CA do EPI. \n"
            f"3. A retirada de EPI's e, consequentemente o seu registro, é realizada com a utilização do crachá de identificação e de uma senha pessoal e instransferível, cadastrada por mim. \n"
            f"4. É proibido o uso, bem como o empréstimo do crachá de identificação e da senha pessoal para a retirada de EPI's para outra pessoa. \n"
            f"5. Assumo total responsabilidade pelo uso da senha, não devendo em hipótese alguma fornecer a mesma para outra pessoa, sob pena de aplicação de punições disciplinares, inclusive demissão por justa causa. \n"
            f"6. Declaro ter sido treinado e orientado sobre o processo de retirada e registro de entrega dos EPIs através do processo utilizado na empresa.  legais cabíveis.\n"
            f"7. Declaro ter recebido os EPIs descritos em relatório anexo, em perfeitas condições de uso, devidamente aprovados pelo órgão competente, com a devida orientação e treinamento quanto ao uso correto, guarda, conservação, substituição e limitações de proteção dos mesmos.\n"
            f"8. Estou ciente da obrigatoriedade de utilização sempre que exercer atividades que demandem tais equipamentos, conforme orientações da empresa e legislação vigente (NR 6 da portaria 3.241/78 do MTE) "
            f"Data de ciência deste termo e cadastro da senha: {data_atual}\n\n"
            f"Assinatura: ________________________________________________"
        )

        pdf.add_page()
        pdf.set_font('Arial', '', 8)
        pdf.multi_cell(0, 10, texto_termo)

    pdf.output(caminho_saida_pdf)

class DataMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Amazon WHS - Apex Importer")
        self.root.geometry("500x600")

        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()

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
            tk.Label(root, image=self.logo).pack(pady=10)

        except Exception as e:
            print("Erro ao carregar logo:", str(e))
            tk.Label(root, text="[Logo não carregada]").pack()

        tk.Label(root, text="Arquivo RFID (CSV):").pack()
        tk.Entry(root, textvariable=self.file1_path, width=60).pack()
        tk.Button(root, text="Selecionar arquivo", command=self.load_file1).pack(pady=5)

        tk.Label(root, text="Planilha Associados (XLSX):").pack()
        tk.Entry(root, textvariable=self.file2_path, width=60).pack()
        tk.Button(root, text="Selecionar arquivo", command=self.load_file2).pack(pady=5)

        tk.Button(root, text="Mesclar dados", command=self.process_data, bg="#0077cc", fg="white").pack(pady=20)

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
            with open(self.file1_path.get(), 'r') as file:
                content = file.readlines()
                delimiter = ',' if ',' in content[0] else ';'

            df1 = pd.read_csv(self.file1_path.get(), delimiter=delimiter, dtype=str)
            if not {'Employee ID', 'Badge ID', 'Login'}.issubset(df1.columns):
                messagebox.showerror("Erro", "Arquivo CSV deve conter 'Employee ID', 'Badge ID' e 'Login'")
                return

            df2 = pd.read_excel(self.file2_path.get(), skiprows=1, dtype=str)
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

            # Gerar o arquivo Instances
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

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()

    messagebox.showinfo(
        "Instruções",
        "1. Selecione o arquivo CSV com 'Employee ID', 'Badge ID' e 'Login'.\n"
        "2. Selecione a planilha Excel com 'Nome' e 'Login'.\n"
        "3. Clique em 'Mesclar dados' para gerar os arquivos.\n\n"
        "OBS: A planilha Excel deve ter cabeçalho na linha 2."
    )
    root.deiconify()
    app = DataMergerApp(root)
    root.mainloop()
