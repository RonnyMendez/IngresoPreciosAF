import tkinter as tk
from tkinter import filedialog, Toplevel, Checkbutton, IntVar, messagebox
from tkinter import ttk
from processor import iniciar_procesamiento

# Variable global para almacenar los archivos generados
archivos_generados = []

# Función para seleccionar archivos
def seleccionar_archivo(entry):
    archivo = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if archivo:
        entry.delete(0, tk.END)
        entry.insert(0, archivo)

# Función para mostrar el modal con checkboxes
def mostrar_modal_archivos():
    if not archivos_generados:
        messagebox.showinfo("Revisión", "No hay archivos para revisar.")
        return

    modal = Toplevel(root)
    modal.title("Archivos Generados")
    modal.geometry("400x500")

    checkbox_vars = []
    for archivo in archivos_generados:
        var = IntVar()
        checkbox = Checkbutton(modal, text=archivo, variable=var)
        checkbox.pack(anchor='w', padx=10, pady=5)
        checkbox_vars.append(var)

    tk.Button(modal, text="Cerrar", command=modal.destroy).pack(pady=20)

# Función para iniciar procesamiento
def iniciar_y_guardar_archivos():
    global archivos_generados
    tamaño_bloque = int(entry_tamaño_bloque.get())
    archivos_generados = iniciar_procesamiento(
        entry_compra.get(),
        entry_productos.get(),
        int(var_escala_actual.get()),
        int(var_escala_deseada.get()),
        int(entry_inicio.get()) - 2,
        int(entry_fin.get()) - 1,
        tamaño_bloque,
        text_area,
        progress_bar
    )
    if archivos_generados:
        messagebox.showinfo("Éxito", "El procesamiento ha finalizado. Puedes revisar los archivos creados.")
    else:
        messagebox.showerror("Error", "No se generaron archivos.")

# Crear la ventana principal
root = tk.Tk()
root.title("Procesador de Excel con Tkinter")
root.geometry("800x600")

entry_compra = tk.Entry(root, width=70)
entry_compra.insert(0, "excelIngreso/compra.xlsx")
entry_compra.grid(row=0, column=1, padx=10, pady=5)
tk.Button(root, text="Seleccionar Compra", command=lambda: seleccionar_archivo(entry_compra)).grid(row=0, column=2)

entry_productos = tk.Entry(root, width=70)
entry_productos.insert(0, "C:/Users/User/OneDrive/Documentos/ProductosParaIngresar.xlsx")
entry_productos.grid(row=1, column=1, padx=10, pady=5)
tk.Button(root, text="Seleccionar Productos", command=lambda: seleccionar_archivo(entry_productos)).grid(row=1, column=2)

var_escala_actual = tk.StringVar(value="1")
tk.OptionMenu(root, var_escala_actual, "1", "100", "1000").grid(row=2, column=1)
var_escala_deseada = tk.StringVar(value="1")
tk.OptionMenu(root, var_escala_deseada, "1", "100", "1000").grid(row=3, column=1)

entry_inicio = tk.Entry(root, width=10)
entry_inicio.insert(0, "1")
entry_inicio.grid(row=4, column=1)
entry_fin = tk.Entry(root, width=10)
entry_fin.insert(0, "10")
entry_fin.grid(row=5, column=1)

# Entrada para el tamaño del bloque
tk.Label(root, text="Tamaño del bloque:").grid(row=6, column=0, padx=10, pady=5)
entry_tamaño_bloque = tk.Entry(root, width=10)
entry_tamaño_bloque.insert(0, "10")
entry_tamaño_bloque.grid(row=6, column=1)

progress_bar = ttk.Progressbar(root, length=600)
progress_bar.grid(row=7, column=0, columnspan=3, pady=10)
text_area = tk.Text(root, height=15, width=100)
text_area.grid(row=8, column=0, columnspan=3)

tk.Button(root, text="Iniciar Procesamiento", command=iniciar_y_guardar_archivos).grid(row=9, column=0, columnspan=3, pady=10)
tk.Button(root, text="Revisión", command=mostrar_modal_archivos).grid(row=10, column=0, columnspan=3, pady=10)

root.mainloop()
