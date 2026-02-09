import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

# Adiciona o diretório raiz do projeto ao sys.path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
#import utils

from processor import extrair_dados_word, preencher_planilha_excel

class DocxToExcelAutomator:
    def __init__(self, root):
        self.root = root
        self.root.title("Aut Lista de Material - DOCX to Excel Automator")
        self.root.geometry("600x400")

        # File selection frame
        file_frame = tk.Frame(root, padx=10, pady=10)
        file_frame.pack(fill=tk.X)

        # DOCX file selection
        self.docx_path = tk.StringVar()
        docx_label = tk.Label(file_frame, text="Lista de Material (.docx):")
        docx_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 5))

        self.docx_entry = tk.Entry(file_frame, textvariable=self.docx_path, width=70)
        self.docx_entry.grid(row=1, column=0, sticky=tk.EW)

        self.browse_docx_button = tk.Button(file_frame, text="Procurar...", command=self.browse_docx_file)
        self.browse_docx_button.grid(row=1, column=1, padx=(5, 0))

        # Excel file selection
        self.excel_path = tk.StringVar()
        excel_label = tk.Label(file_frame, text="Planilha Excel (.xlsx):")
        excel_label.grid(row=2, column=0, sticky=tk.W, pady=(10, 5))

        self.excel_entry = tk.Entry(file_frame, textvariable=self.excel_path, width=70)
        self.excel_entry.grid(row=3, column=0, sticky=tk.EW)

        self.browse_excel_button = tk.Button(file_frame, text="Procurar...", command=self.browse_excel_file)
        self.browse_excel_button.grid(row=3, column=1, padx=(5, 0))

        file_frame.grid_columnconfigure(0, weight=1)

        # Start button
        self.start_button = tk.Button(root, text="Iniciar Script", command=self.start_automation, font=("Helvetica", 12, "bold"))
        self.start_button.pack(pady=10, padx=10, fill=tk.X, ipady=5)

        # Log frame
        log_frame = tk.Frame(root, padx=10, pady=10)
        log_frame.pack(fill=tk.BOTH, expand=True)

        log_label = tk.Label(log_frame, text="Log de Eventos:")
        log_label.pack(anchor=tk.W)

        self.log_text = ScrolledText(log_frame, state='disabled', height=10, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        self.add_placeholder(self.docx_entry, "Clique em 'Procurar...' para selecionar a lista de material")
        self.add_placeholder(self.excel_entry, "Clique em 'Procurar...' para selecionar a planilha de aço")

    def add_placeholder(self, widget, placeholder):
        widget.insert(0, placeholder)
        widget.config(fg='grey')
        widget.bind('<FocusIn>', lambda event: self.on_focus_in(event, placeholder))
        widget.bind('<FocusOut>', lambda event: self.on_focus_out(event, placeholder))

    def on_focus_in(self, event, placeholder):
        widget = event.widget
        if widget.get() == placeholder:
            widget.delete(0, tk.END)
            widget.config(fg='black')

    def on_focus_out(self, event, placeholder):
        widget = event.widget
        if not widget.get():
            widget.insert(0, placeholder)
            widget.config(fg='grey')

    def log(self, message):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.config(state='disabled')
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def browse_docx_file(self):
        filepath = filedialog.askopenfilename(
            title="Selecione um arquivo DOCX",
            filetypes=(("Word Documents", "*.docx"), ("All files", "*.*"))
        )
        if filepath:
            self.docx_path.set(filepath)
            self.log(f"Arquivo DOCX selecionado: {os.path.basename(filepath)}")
            # Clear focus to trigger placeholder logic if entry is cleared and refocused
            self.root.focus()


    def browse_excel_file(self):
        filepath = filedialog.askopenfilename(
            title="Selecione um arquivo Excel",
            filetypes=(("Excel Spreadsheets", "*.xlsx"), ("All files", "*.*"))
        )
        if filepath:
            self.excel_path.set(filepath)
            self.log(f"Planilha Excel selecionada: {os.path.basename(filepath)}")
            # Clear focus to trigger placeholder logic if entry is cleared and refocused
            self.root.focus()

    def start_automation(self):
        arquivo_word = self.docx_path.get()
        planilha_excel = self.excel_path.get()

        # Validate that the fields are not empty or with placeholder text
        if not arquivo_word or "selecionar a lista de material" in arquivo_word:
            messagebox.showerror("Erro", "Por favor, selecione um arquivo .docx primeiro.")
            return

        if not planilha_excel or "selecionar a planilha de aço" in planilha_excel:
            messagebox.showerror("Erro", "Por favor, selecione uma planilha Excel.")
            return
            
        if not os.path.exists(planilha_excel):
            messagebox.showerror("Erro", f"Planilha Excel não encontrada:\n{planilha_excel}")
            return

        self.log("Iniciando o processo...")
        
        try:
            dados_extraidos = extrair_dados_word(arquivo_word)
            if dados_extraidos:
                self.log(f"Dados extraídos de {os.path.basename(arquivo_word)} com sucesso.")
                self.log("Preenchendo a planilha Excel...")
                preencher_planilha_excel(planilha_excel, dados_extraidos)
                messagebox.showinfo("Sucesso", "A planilha Excel foi atualizada com sucesso.")
            else:
                self.log("Nenhum dado extraído do arquivo Word. Verifique o arquivo.")
                messagebox.showwarning("Aviso", "Nenhum dado foi extraído do arquivo Word. Por favor, verifique o arquivo.")
        except Exception as e:
            self.log(f"Erro: {e}")
            messagebox.showerror("Erro na Automação", f"Ocorreu um erro: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = DocxToExcelAutomator(root)
    root.mainloop()
