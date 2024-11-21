import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
import time

# Función para aplicar la conversión de precios
def convertir_precio(precio, escala_actual, escala_deseada):
    factor_conversion = {
        (1, 1): 1, (1, 100): 0.01, (1, 1000): 0.001,
        (100, 1): 100, (100, 100): 1, (100, 1000): 0.1,
        (1000, 1): 1000, (1000, 100): 10, (1000, 1000): 1
    }
    return precio / factor_conversion[(escala_actual, escala_deseada)]

# Función para seleccionar archivos
def seleccionar_archivo_compra():
    archivo = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if archivo:
        entry_compra.delete(0, tk.END)
        entry_compra.insert(0, archivo)

def seleccionar_archivo_productos():
    archivo = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if archivo:
        entry_productos.delete(0, tk.END)
        entry_productos.insert(0, archivo)

# Función para mostrar mensajes en el Text Widget
def mostrar_mensaje(mensaje):
    text_area.insert(tk.END, mensaje + "\n")
    text_area.see(tk.END)  # Auto-scroll al final

# Función para iniciar el procesamiento completo
def iniciar_procesamiento():
    try:
        start_time = time.time()

        # Obtener rutas de archivos
        archivo_compra = entry_compra.get()
        archivo_productos = entry_productos.get()

        # Obtener escalas
        escala_actual = int(var_escala_actual.get())
        escala_deseada = int(var_escala_deseada.get())

        # Obtener rango de filas
        inicio_fila = int(entry_inicio.get()) - 2
        fin_fila = int(entry_fin.get()) - 1

        # Cargar los archivos de Excel
        compra_df = pd.read_excel(archivo_compra)
        productos_df = pd.read_excel(archivo_productos)

        # Filtrar el DataFrame de productos
        productos_df = productos_df.iloc[inicio_fila:fin_fila]

        # Convertir los precios
        productos_df['Precio'] = productos_df['Precio'].apply(lambda x: convertir_precio(x, escala_actual, escala_deseada))

        # Definir el tamaño del bloque y calcular el número total de archivos
        tamaño_bloque = 10
        num_archivos = (len(productos_df) + tamaño_bloque - 1) // tamaño_bloque
        progress_bar["maximum"] = num_archivos

        # Variable para acumular el total de precios copiados
        total_precios_copiados = 0

        # Loop para crear cada archivo de resultado en bloques de 10 filas
        for i in range(num_archivos):
            inicio_fila_bloque = i * tamaño_bloque
            fin_fila_bloque = inicio_fila_bloque + tamaño_bloque
            productos_filtrados = productos_df.iloc[inicio_fila_bloque:fin_fila_bloque]

            # Realizar el merge para actualizar el "Precio Compra"
            compra_actualizada = compra_df.merge(
                productos_filtrados[['Código', 'Precio']],
                how='left', on='Código'
            )

            # Reemplazar los valores en "Precio Compra"
            compra_actualizada['Precio Compra'] = compra_actualizada['Precio'].combine_first(compra_actualizada['Precio Compra'])

            # Contar cuántos precios fueron copiados
            precios_copiados = compra_actualizada['Precio Compra'].notna().sum() - compra_df['Precio Compra'].notna().sum()
            total_precios_copiados += precios_copiados

            # Asignar "0" a "Cantidad" y "Descuento (%)"
            compra_actualizada.loc[compra_actualizada['Precio'].notna(), ['Cantidad', 'Descuento (%)']] = 0

            # Eliminar la columna extra 'Precio'
            compra_actualizada = compra_actualizada.drop(columns=['Precio'])

            # Guardar el archivo de Excel resultante
            nombre_archivo = f"excelCompras/Compra{i + 1}.xlsx"
            compra_actualizada.to_excel(nombre_archivo, index=False)

            # Obtener el primer y último código copiado
            codigos_copiados = productos_filtrados['Código'].tolist()
            primer_codigo = codigos_copiados[0] if codigos_copiados else "N/A"
            ultimo_codigo = codigos_copiados[-1] if codigos_copiados else "N/A"

            # Mostrar mensaje en el Text Widget
            mostrar_mensaje(f"Archivo '{nombre_archivo}' creado. Primer código: {primer_codigo}, Último código: {ultimo_codigo}.")

            # Actualizar la barra de progreso
            progress_bar["value"] = i + 1
            root.update_idletasks()

        # Mostrar mensaje final
        end_time = time.time()
        elapsed_time = end_time - start_time
        mostrar_mensaje(f"Procesamiento completo. Total de precios copiados: {total_precios_copiados}.")
        mostrar_mensaje(f"Tiempo total de procesamiento: {elapsed_time:.2f} segundos.")

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error: {str(e)}")

# Crear la ventana principal
root = tk.Tk()
root.title("Procesador de Excel con Tkinter")

# Valores por defecto
ruta_compra_default = "excelIngreso/compra.xlsx"
ruta_productos_default = "C:/Users/User/OneDrive/Documentos/ProductosParaIngresar.xlsx"

# Campos para seleccionar archivos
tk.Label(root, text="Archivo de Compra:").grid(row=0, column=0, padx=10, pady=5)
entry_compra = tk.Entry(root, width=50)
entry_compra.insert(0, ruta_compra_default)
entry_compra.grid(row=0, column=1)
tk.Button(root, text="Seleccionar", command=seleccionar_archivo_compra).grid(row=0, column=2)

tk.Label(root, text="Archivo de Productos:").grid(row=1, column=0, padx=10, pady=5)
entry_productos = tk.Entry(root, width=50)
entry_productos.insert(0, ruta_productos_default)
entry_productos.grid(row=1, column=1)
tk.Button(root, text="Seleccionar", command=seleccionar_archivo_productos).grid(row=1, column=2)

# Menús desplegables para escalas
tk.Label(root, text="Escala Actual:").grid(row=2, column=0, padx=10, pady=5)
var_escala_actual = tk.StringVar(value="1")
tk.OptionMenu(root, var_escala_actual, "1", "100", "1000").grid(row=2, column=1)

tk.Label(root, text="Escala Deseada:").grid(row=3, column=0, padx=10, pady=5)
var_escala_deseada = tk.StringVar(value="1")
tk.OptionMenu(root, var_escala_deseada, "1", "100", "1000").grid(row=3, column=1)

# Entradas para rango de filas
tk.Label(root, text="Inicio de filas:").grid(row=4, column=0)
entry_inicio = tk.Entry(root, width=10)
entry_inicio.insert(0, "1")
entry_inicio.grid(row=4, column=1)

tk.Label(root, text="Fin de filas:").grid(row=5, column=0)
entry_fin = tk.Entry(root, width=10)
entry_fin.insert(0, "10")
entry_fin.grid(row=5, column=1)

# Barra de progreso y área de texto para mensajes
progress_bar = ttk.Progressbar(root, length=400)
progress_bar.grid(row=6, column=0, columnspan=3, pady=10)

text_area = tk.Text(root, height=10, width=60)
text_area.grid(row=7, column=0, columnspan=3)

# Botón para iniciar el procesamiento
tk.Button(root, text="Iniciar Procesamiento", command=iniciar_procesamiento).grid(row=8, column=0, columnspan=3, pady=20)

# Ejecutar la aplicación
root.mainloop()
