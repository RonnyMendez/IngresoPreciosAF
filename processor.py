import pandas as pd
import time
from utils import mostrar_mensaje

def iniciar_procesamiento(archivo_compra, archivo_productos, escala_actual, escala_deseada, inicio_fila, fin_fila, tamaño_bloque, text_area, progress_bar):
    try:
        start_time = time.time()
        compra_df = pd.read_excel(archivo_compra)
        productos_df = pd.read_excel(archivo_productos)

        productos_df = productos_df.iloc[inicio_fila:fin_fila]
        productos_df['Precio'] = productos_df['Precio'].apply(lambda x: convertir_precio(x, escala_actual, escala_deseada))

        num_archivos = (len(productos_df) + tamaño_bloque - 1) // tamaño_bloque
        progress_bar["maximum"] = num_archivos

        archivos_generados = []
        total_precios_copiados = 0

        for i in range(num_archivos):
            inicio_fila_bloque = i * tamaño_bloque
            fin_fila_bloque = inicio_fila_bloque + tamaño_bloque
            productos_filtrados = productos_df.iloc[inicio_fila_bloque:fin_fila_bloque]

            compra_actualizada = compra_df.merge(
                productos_filtrados[['Código', 'Precio']],
                how='left', on='Código'
            )

            compra_actualizada['Precio Compra'] = compra_actualizada['Precio'].combine_first(compra_actualizada['Precio Compra'])
            nombre_archivo = f"excelCompras/Compra{i + 1}.xlsx"
            compra_actualizada.to_excel(nombre_archivo, index=False)

            codigos_copiados = productos_filtrados['Código'].tolist()
            primer_codigo = codigos_copiados[0] if codigos_copiados else "N/A"
            ultimo_codigo = codigos_copiados[-1] if codigos_copiados else "N/A"
            precios_copiados = compra_actualizada['Precio'].notna().sum()
            total_precios_copiados += precios_copiados

            mostrar_mensaje(text_area, f"Archivo '{nombre_archivo}' creado. Primer código: {primer_codigo}, Último código: {ultimo_codigo}, Códigos copiados: {precios_copiados}.", "green")

            archivos_generados.append(nombre_archivo)
            progress_bar["value"] = i + 1
            text_area.update_idletasks()

        elapsed_time = time.time() - start_time
        mostrar_mensaje(text_area, f"Procesamiento completo. Total de códigos copiados: {total_precios_copiados}.", "blue")
        mostrar_mensaje(text_area, f"Tiempo total de procesamiento: {elapsed_time:.2f} segundos.", "blue")
        return archivos_generados

    except Exception as e:
        mostrar_mensaje(text_area, f"Error: {str(e)}", "red")
        return []

def convertir_precio(precio, escala_actual, escala_deseada):
    factor_conversion = {
        (1, 1): 1, (1, 100): 0.01, (1, 1000): 0.001,
        (100, 1): 100, (100, 100): 1, (100, 1000): 0.1,
        (1000, 1): 1000, (1000, 100): 10, (1000, 1000): 1
    }
    return precio / factor_conversion.get((escala_actual, escala_deseada), 1)