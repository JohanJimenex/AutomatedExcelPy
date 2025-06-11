# limpieza_txt.py

import tkinter as tk
from tkinter import filedialog
import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

def procesar_archivos():
    rutas_archivos = filedialog.askopenfilenames(
        title="Selecciona uno o varios archivos .txt",
        filetypes=[("Archivos de texto", "*.txt")]
    )

    if not rutas_archivos:
        return

    for ruta_archivo in rutas_archivos:
        with open(ruta_archivo, 'r') as f:
            lineas = f.readlines()

        datos = []
        
        for linea in lineas:
            entidad = linea[0:4]
            centralta = linea[4:8]
            cuenta = linea[8:20]
            producto = linea[20:22]
            subproducto = linea[22:26]
            pan = linea[26:45]
            fecha_alta = linea[45:59].strip()
            bloqueo = linea[59:61]
            fecha_baja = linea[61:71]
            hora = linea[71:].strip()
            datos.append([
                entidad, centralta, cuenta, producto, subproducto, pan,
                fecha_alta, bloqueo, fecha_baja, hora
            ])

        columnas = [
            "ENTIDAD", "CENTALTA", "CUENTA", "PRODUCTO", "SUBPRODUCTO", "PAN",
            "FECHA ALTA", "BLOQUEO", "FECHA BAJA", "HORA"
        ]
        df = pd.DataFrame(datos, columns=columnas)
        
        # Convertir columnas de fecha a datetime para que Excel las trate como fechas reales
        df["FECHA ALTA"] = pd.to_datetime(df["FECHA ALTA"].str.strip(), format="%d-%m-%Y", errors="coerce")
        # df["FECHA BAJA"] = pd.to_datetime(df["FECHA BAJA"], format="%d-%m-%Y", errors="coerce")
        df["FECHA BAJA"] = pd.to_datetime(df["FECHA BAJA"].str.strip(), format="%d-%m-%Y", errors="coerce").dt.date

        # df["HORA"] = pd.to_datetime(df["HORA"], format="%H:%M:%S", errors="coerce").dt.time

        nombre_archivo = os.path.splitext(os.path.basename(ruta_archivo))[0]
        ruta_salida = os.path.join(os.path.dirname(ruta_archivo), f"{nombre_archivo}_procesado.xlsx")

        with pd.ExcelWriter(ruta_salida, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name="Original", index=False)

        wb = load_workbook(ruta_salida)
        ws = wb["Original"]
        
        from openpyxl.styles import numbers

        # Formatear columnas de FECHA ALTA (columna G) y FECHA BAJA (columna I)
        for fila in range(2, len(df) + 2):  # Desde fila 2 (sin encabezado)
            ws[f"G{fila}"].number_format = "DD/MM/YYYY"
            ws[f"I{fila}"].number_format = "DD/MM/YYYY"


        rango = f"A1:J{len(df) + 1}"  # Ahora son 10 columnas (A hasta J)
        tabla = Table(displayName="TablaDatos", ref=rango)
        estilo = TableStyleInfo(
            name="TableStyleMedium9", showFirstColumn=False,
            showLastColumn=False, showRowStripes=True, showColumnStripes=False
        )
        tabla.tableStyleInfo = estilo
        ws.add_table(tabla)

        wb.save(ruta_salida)
        print(f"Archivo generado con filtros y estilo: {ruta_salida}")

# ===============================
# Interfaz gráfica principal
# ===============================
ventana = tk.Tk()
ventana.title("Procesar Archivos de Tarjetas")
ventana.geometry("300x100")

boton = tk.Button(ventana, text="Seleccionar archivos .txt", command=procesar_archivos)
boton.pack(pady=30)

ventana.mainloop()
