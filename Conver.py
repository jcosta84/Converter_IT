import os
import customtkinter as ctk
from tkinter import filedialog, messagebox
import win32com.client as win32

# Função de conversão
def converter_arquivos(pasta, pasta_saida, log_widget, progress_bar):
    try:
        excel = win32.Dispatch("Excel.Application")
        excel.DisplayAlerts = False

        arquivos = [f for f in os.listdir(pasta) if f.lower().endswith(".xls") and not f.lower().endswith(".xlsx")]
        total_arquivos = len(arquivos)
        arquivos_convertidos = 0

        if total_arquivos == 0:
            messagebox.showinfo("Aviso", "Nenhum arquivo .xls encontrado para converter.")
            return

        # Criar pasta de saída caso não exista
        if not os.path.exists(pasta_saida):
            os.makedirs(pasta_saida)

        for idx, arquivo in enumerate(arquivos, start=1):
            caminho_arquivo = os.path.abspath(os.path.join(pasta, arquivo))

            if not os.path.exists(caminho_arquivo):
                log_widget.insert("end", f"Arquivo não encontrado: {caminho_arquivo}\n")
                log_widget.see("end")
                continue

            try:
                wb = excel.Workbooks.Open(caminho_arquivo)

                novo_nome = arquivo.rsplit(".", 1)[0] + ".xlsx"
                caminho_novo = os.path.abspath(os.path.join(pasta_saida, novo_nome))

                wb.SaveAs(caminho_novo, FileFormat=51)  # 51 = .xlsx
                wb.Close(SaveChanges=False)

                log_widget.insert("end", f"Convertido: {arquivo} → {caminho_novo}\n")
                log_widget.see("end")
                arquivos_convertidos += 1

            except Exception as e:
                log_widget.insert("end", f"Erro no arquivo {arquivo}: {e}\n")
                log_widget.see("end")

            # Atualizar barra de progresso
            progress = idx / total_arquivos
            progress_bar.set(progress)
            app.update_idletasks()

        excel.Quit()

        messagebox.showinfo("Sucesso", f"Conversão concluída!\n{arquivos_convertidos}/{total_arquivos} arquivos convertidos.")

    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

# Função para escolher diretório de entrada
def escolher_diretorio_entrada():
    pasta = filedialog.askdirectory(title="Selecione a pasta com arquivos .xls")
    if pasta:
        pasta_entry.delete(0, "end")
        pasta_entry.insert(0, pasta)

# Função para escolher diretório de saída
def escolher_diretorio_saida():
    pasta = filedialog.askdirectory(title="Selecione a pasta de saída")
    if pasta:
        pasta_saida_entry.delete(0, "end")
        pasta_saida_entry.insert(0, pasta)

# Função chamada ao clicar no botão converter
def iniciar_conversao():
    pasta = pasta_entry.get().strip()
    pasta_saida = pasta_saida_entry.get().strip()

    if not pasta or not os.path.isdir(pasta):
        messagebox.showwarning("Atenção", "Selecione uma pasta de entrada válida.")
        return
    if not pasta_saida:
        messagebox.showwarning("Atenção", "Selecione uma pasta de saída válida.")
        return

    log_textbox.delete("1.0", "end")  # limpa log
    progress_bar.set(0)  # resetar barra
    converter_arquivos(pasta, pasta_saida, log_textbox, progress_bar)

# ---------------- GUI ----------------
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

app = ctk.CTk()
app.title("Conversor XLS → XLSX")
app.geometry("650x500")

# Frame principal
frame = ctk.CTkFrame(app, corner_radius=15)
frame.pack(padx=20, pady=20, fill="both", expand=True)

# Campo de seleção de pasta de entrada
pasta_label = ctk.CTkLabel(frame, text="Pasta dos arquivos (.xls):")
pasta_label.pack(anchor="w", padx=10, pady=(10, 0))

pasta_frame = ctk.CTkFrame(frame, fg_color="transparent")
pasta_frame.pack(fill="x", padx=10, pady=5)

pasta_entry = ctk.CTkEntry(pasta_frame, width=400)
pasta_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))

browse_button = ctk.CTkButton(pasta_frame, text="Procurar", command=escolher_diretorio_entrada)
browse_button.pack(side="right")

# Campo de seleção de pasta de saída
pasta_saida_label = ctk.CTkLabel(frame, text="Pasta de saída (.xlsx):")
pasta_saida_label.pack(anchor="w", padx=10, pady=(10, 0))

pasta_saida_frame = ctk.CTkFrame(frame, fg_color="transparent")
pasta_saida_frame.pack(fill="x", padx=10, pady=5)

pasta_saida_entry = ctk.CTkEntry(pasta_saida_frame, width=400)
pasta_saida_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))

browse_saida_button = ctk.CTkButton(pasta_saida_frame, text="Procurar", command=escolher_diretorio_saida)
browse_saida_button.pack(side="right")

# Botão de conversão
convert_button = ctk.CTkButton(frame, text="Converter Arquivos", command=iniciar_conversao)
convert_button.pack(pady=10)

# Barra de progresso
progress_bar = ctk.CTkProgressBar(frame)
progress_bar.pack(fill="x", padx=10, pady=5)
progress_bar.set(0)

# Caixa de log
log_textbox = ctk.CTkTextbox(frame, height=200)
log_textbox.pack(fill="both", expand=True, padx=10, pady=10)

app.mainloop()
