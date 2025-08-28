import os
import customtkinter as ctk
from tkinter import filedialog, messagebox
import win32com.client as win32

# Função de conversão
def converter_arquivos(pasta, log_widget):
    try:
        excel = win32.Dispatch("Excel.Application")
        excel.DisplayAlerts = False  # evita mensagens do Excel

        arquivos_convertidos = 0

        for arquivo in os.listdir(pasta):
            if arquivo.lower().endswith(".xls") and not arquivo.lower().endswith(".xlsx"):
                caminho_arquivo = os.path.join(pasta, arquivo)
                workbook = excel.Workbooks.Open(caminho_arquivo)

                # Novo nome .xlsx
                novo_nome = arquivo.rsplit(".", 1)[0] + ".xlsx"
                caminho_novo = os.path.join(pasta, novo_nome)

                workbook.SaveAs(caminho_novo, FileFormat=51)  # 51 = xlOpenXMLWorkbook (.xlsx)
                workbook.Close()

                log_widget.insert("end", f"Convertido: {arquivo} → {novo_nome}\n")
                log_widget.see("end")
                arquivos_convertidos += 1

        excel.Quit()

        if arquivos_convertidos > 0:
            messagebox.showinfo("Sucesso", f"Conversão concluída!\n{arquivos_convertidos} arquivos convertidos.")
        else:
            messagebox.showinfo("Aviso", "Nenhum arquivo .xls encontrado na pasta.")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

# Função para escolher diretório
def escolher_diretorio():
    pasta = filedialog.askdirectory(title="Selecione a pasta com arquivos .xls")
    if pasta:
        pasta_entry.delete(0, "end")
        pasta_entry.insert(0, pasta)

# Função chamada ao clicar no botão converter
def iniciar_conversao():
    pasta = pasta_entry.get().strip()
    if not pasta or not os.path.isdir(pasta):
        messagebox.showwarning("Atenção", "Selecione uma pasta válida.")
        return
    log_textbox.delete("1.0", "end")  # limpa log
    converter_arquivos(pasta, log_textbox)

# ---------------- GUI ----------------
ctk.set_appearance_mode("System")  # "Dark" ou "Light"
ctk.set_default_color_theme("blue")

app = ctk.CTk()
app.title("Conversor XLS → XLSX")
app.geometry("600x400")

# Frame principal
frame = ctk.CTkFrame(app, corner_radius=15)
frame.pack(padx=20, pady=20, fill="both", expand=True)

# Campo de seleção de pasta
pasta_label = ctk.CTkLabel(frame, text="Pasta dos arquivos:")
pasta_label.pack(anchor="w", padx=10, pady=(10, 0))

pasta_frame = ctk.CTkFrame(frame, fg_color="transparent")
pasta_frame.pack(fill="x", padx=10, pady=5)

pasta_entry = ctk.CTkEntry(pasta_frame, width=400)
pasta_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))

browse_button = ctk.CTkButton(pasta_frame, text="Procurar", command=escolher_diretorio)
browse_button.pack(side="right")

# Botão de conversão
convert_button = ctk.CTkButton(frame, text="Converter Arquivos", command=iniciar_conversao)
convert_button.pack(pady=10)

# Caixa de log
log_textbox = ctk.CTkTextbox(frame, height=150)
log_textbox.pack(fill="both", expand=True, padx=10, pady=10)

app.mainloop()
