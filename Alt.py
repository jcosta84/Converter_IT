import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

def converter_csv_para_kml():
    file_path = filedialog.askopenfilename(
        title="Selecione o ficheiro CSV exportado do Excel",
        filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx;*.xls")]
    )
    if not file_path:
        return

    try:
        # Ler CSV ou Excel
        if file_path.endswith(".csv"):
            df = pd.read_csv(file_path, sep=";", engine="python", on_bad_lines="skip")
        else:
            df = pd.read_excel(file_path)

        # Mostrar colunas originais
        print("Colunas originais:", list(df.columns))
        messagebox.showinfo("Colunas encontradas", f"{list(df.columns)}")

        # Padronizar colunas importantes
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

        df.rename(columns=lambda x: mapeamento.get(x.strip().lower(), x.strip()), inplace=True)

        # Verificar colunas obrigatórias
        if not {"Latitude", "Longitude", "CIL", "Roteiro", "Itinerário"}.issubset(df.columns):
            messagebox.showerror(
                "Erro",
                f"O ficheiro precisa ter colunas 'Latitude', 'Longitude', 'CIL', 'Roteiro' e 'Itinerário'.\nColunas encontradas: {list(df.columns)}"
            )
            return

        # Criar pasta base de saída
        base_dir = os.path.dirname(file_path)
        pasta_saida = os.path.join(base_dir, "KML_Saida")
        os.makedirs(pasta_saida, exist_ok=True)

        # Separar por Roteiro + Itinerário
        for (roteiro, itinerario), grupo in df.groupby(["Roteiro", "Itinerário"]):
            kml_content = """<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
<Document>
"""

            for i, row in grupo.iterrows():
                cil = row.get("CIL", f"{i+1}")
                lat = row["Latitude"]
                lon = row["Longitude"]

                # Montar descrição
                detalhes = ""
                for campo in ["Roteiro", "Itinerário", "Zona", "Rua", "Cliente", "Ponto de Medida", "CIL", "Número"]:
                    if campo in grupo.columns:
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

            # Criar subpasta com nome do roteiro
            safe_roteiro = str(roteiro).replace("/", "-").replace("\\", "-").strip()
            pasta_roteiro = os.path.join(pasta_saida, safe_roteiro)
            os.makedirs(pasta_roteiro, exist_ok=True)

            # Nome do ficheiro → número do Itinerário
            safe_itinerario = str(itinerario).replace("/", "-").replace("\\", "-").strip()
            save_path = os.path.join(pasta_roteiro, f"{safe_itinerario}.kml")

            with open(save_path, "w", encoding="utf-8") as f:
                f.write(kml_content)

        messagebox.showinfo("Sucesso", f"KMLs guardados em:\n{pasta_saida}")

    except Exception as e:
        messagebox.showerror("Erro", str(e))


# Interface Tkinter
root = tk.Tk()
root.title("Conversor CSV/Excel → KML")

btn = tk.Button(root, text="Converter CSV/Excel para KML", command=converter_csv_para_kml, width=40, height=3)
btn.pack(pady=20)

root.mainloop()
