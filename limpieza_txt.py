# limpieza_txt.py

import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import sys
import os
import threading
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import numbers

cancelado = False  # Variable global de control

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def cancelar_proceso():
    global cancelado
    cancelado = True
    estado_label.config(text="Cancelado.")
    barra.stop()
    barra.pack_forget()
    boton_cancelar.pack_forget()
    boton_seleccionar.pack(pady=20)

def procesar_archivos_async():
    threading.Thread(target=procesar_archivos).start()

def procesar_archivos():
    global cancelado
    cancelado = False
    rutas_archivos = filedialog.askopenfilenames(
        title="Selecciona uno o varios archivos .txt",
        filetypes=[("Archivos de texto", "*.txt")]
    )
    if not rutas_archivos:
        return

    boton_seleccionar.pack_forget()
    boton_cancelar.pack(pady=20)
    barra.pack(pady=10)
    barra.start()
    estado_label.config(text="Procesando...")

    for ruta_archivo in rutas_archivos:
        if cancelado:
            break
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
        df["FECHA ALTA"] = pd.to_datetime(df["FECHA ALTA"].str.strip(), format="%d-%m-%Y", errors="coerce")
        df["FECHA BAJA"] = pd.to_datetime(df["FECHA BAJA"].str.strip(), format="%d-%m-%Y", errors="coerce").dt.date

        nombre_archivo = os.path.splitext(os.path.basename(ruta_archivo))[0]
        ruta_salida = os.path.join(os.path.dirname(ruta_archivo), f"{nombre_archivo}_procesado.xlsx")

        with pd.ExcelWriter(ruta_salida, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name="Original", index=False)

        wb = load_workbook(ruta_salida)
        ws = wb["Original"]

        for fila in range(2, len(df) + 2):
            ws[f"G{fila}"].number_format = "DD/MM/YYYY"
            ws[f"I{fila}"].number_format = "DD/MM/YYYY"

        rango = f"A1:J{len(df) + 1}"
        tabla = Table(displayName="TablaDatos", ref=rango)
        estilo = TableStyleInfo(
            name="TableStyleMedium9", showFirstColumn=False,
            showLastColumn=False, showRowStripes=True, showColumnStripes=False
        )
        tabla.tableStyleInfo = estilo
        ws.add_table(tabla)

        wb.save(ruta_salida)

    barra.stop()
    barra.pack_forget()
    boton_cancelar.pack_forget()
    boton_seleccionar.pack(pady=20)

    if cancelado:
        estado_label.config(text="Cancelado por el usuario.")
    else:
        estado_label.config(text="¡Completado!")
        messagebox.showinfo("Proceso finalizado", "✅ Archivos procesados correctamente.")

# ===============================
# Interfaz gráfica principal
# ===============================
ventana = tk.Tk()
ventana.title("Procesar Archivos")
ventana.geometry("300x200")
ventana.resizable(False, False)
ventana.iconbitmap(resource_path("assets/icono.ico"))

# Centrar ventana
ventana.update_idletasks()
x = (ventana.winfo_screenwidth() // 2) - (300 // 2)
y = (ventana.winfo_screenheight() // 2) - (200 // 2)
ventana.geometry(f"300x200+{x}+{y}")

descripcion = tk.Label(
    ventana,
    text="Convierte archivos .txt en hojas de Excel con formato automáticamente.",
    wraplength=280,
    justify="center",
    fg="#888888"
)
descripcion.pack(pady=(10, 5))

boton_seleccionar = ttk.Button(ventana, text="Seleccionar archivos", command=procesar_archivos_async)
boton_seleccionar.pack(pady=20)

barra = ttk.Progressbar(ventana, mode="indeterminate")

estado_label = tk.Label(ventana, text="", fg="#555555")
estado_label.pack()

boton_cancelar = ttk.Button(ventana, text="Cancelar", command=cancelar_proceso)

ventana.mainloop()
