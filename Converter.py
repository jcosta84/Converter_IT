import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

def converter_csv_para_kml():
    file_path = filedialog.askopenfilename(
        title="Selecione o ficheiro CSV exportado do Excel",
        filetypes=[("CSV files", "*.csv")]
    )
    if not file_path:
        return

    try:
        # Ler CSV com pandas (tolerante)
        df = pd.read_csv(
            file_path,
            sep=";",
            engine="python",
            on_bad_lines="skip"
        )

        # Mostrar colunas originais
        print("Colunas originais:", list(df.columns))
        messagebox.showinfo("Colunas encontradas", f"{list(df.columns)}")

        # 🔹 Normalizar colunas (remove espaços, deixa minúsculo)
        colunas_normalizadas = {c.strip().lower(): c for c in df.columns}

        # 🔹 Padronizar colunas importantes
        mapeamento = {
            "roteiro": "Roteiro",
            "itinerario": "Itinerário",
            "itinerário": "Itinerário",
            "zona": "Zona",
            "rua": "Rua",
            "cliente": "Cliente",
            "ponto de medida": "Ponto de Medida",
            "cil": "CIL",
            "número": "Número",
            "numero": "Número",
            "latitude": "Latitude",
            "longitude": "Longitude"
        }

        # Renomear colunas do dataframe conforme mapeamento
        df.rename(columns=lambda x: mapeamento.get(x.strip().lower(), x.strip()), inplace=True)

        # Verificar colunas obrigatórias
        if not {"Latitude", "Longitude", "CIL"}.issubset(df.columns):
            messagebox.showerror(
                "Erro",
                f"O CSV precisa ter colunas 'Latitude', 'Longitude' e 'CIL'.\nColunas encontradas: {list(df.columns)}"
            )
            return

        # Criar conteúdo KML
        kml_content = """<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
<Document>
"""

        for i, row in df.iterrows():
            cil = row.get("CIL", f"{i+1}")
            lat = row["Latitude"]
            lon = row["Longitude"]

            # Montar descrição com os campos desejados
            detalhes = ""
            for campo in ["Roteiro", "Itinerário", "Zona", "Rua", "Cliente", "Ponto de Medida", "CIL", "Número"]:
                if campo in df.columns:
                    valor = str(row.get(campo, "")).strip()
                    if valor and valor != "nan":
                        detalhes += f"{campo}: {valor}\n"

            kml_content += f"""
<Placemark>
    <name>{cil}</name>
    <description><![CDATA[{detalhes.strip()}]]></description>
    <Point>
        <coordinates>{lon},{lat},0</coordinates>
    </Point>
</Placemark>
"""

        kml_content += """
</Document>
</kml>
"""

        # Guardar com mesmo nome do CSV, mas extensão .kml
        base_name = os.path.splitext(file_path)[0]
        save_path = f"{base_name}.kml"

        with open(save_path, "w", encoding="utf-8") as f:
            f.write(kml_content)

        messagebox.showinfo("Sucesso", f"Ficheiro guardado em:\n{save_path}")

    except Exception as e:
        messagebox.showerror("Erro", str(e))


# Interface Tkinter
root = tk.Tk()
root.title("Conversor CSV → KML")

btn = tk.Button(root, text="Converter CSV para KML", command=converter_csv_para_kml, width=40, height=3)
btn.pack(pady=20)

root.mainloop()
