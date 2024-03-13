import pandas as pd
from barcode import EAN13
from barcode.writer import ImageWriter
from PIL import Image
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as xlImage

__author__      = "Andrés L. Ospina"
__copyright__   = "Copyright 2024"
__license__     = "GPL"
__version__     = "1.0.1"
__maintainer__  = "SyGmA"
__email__       = "116124867+SyGmAV3@users.noreply.github.com"
__status__      = "Test"

def generar_codigo_barras(ean, numero_fila):
    print(f"Generando código de barras para EAN: {ean}, Fila: {numero_fila}")

    # Agregar ceros a la izquierda si el código EAN tiene menos de 13 dígitos
    ean = str(ean).zfill(13)

    # Generar el código de barras con el EAN
    codigo = EAN13(ean, writer=ImageWriter())
    codigo_barra_path = f'c:/BARCODES/codigos_de_barras/{numero_fila}'  # No incluir la extensión en la ruta
    codigo.save(codigo_barra_path, options={'dpi': 300, 'module_width': 0.2})  # Ajustar las opciones según sea necesario

    print(f"Código de barras generado: {codigo_barra_path}")

    return codigo_barra_path

def main():
    # Leer el archivo Excel original
    df = pd.read_excel('c:/BARCODES/articulos.xlsx', header=None, dtype=str)

    # Crear una nueva hoja de cálculo en el archivo de Excel
    wb = load_workbook('c:/BARCODES/articulos.xlsx')
    ws = wb.active

    # Iterar sobre las filas del DataFrame
    for index, row in df.iterrows():
        if index < 1:
            continue  # Saltar las primeras filas (encabezados y fila 1)

        ean = row[1]
        numero_fila = index + 1  # Sumar 1 para empezar desde la fila 2

        # Generar el código de barras y obtener la ruta de la imagen generada
        codigo_barra_path = generar_codigo_barras(ean, numero_fila)

        # Insertar la imagen del código de barras en la hoja de cálculo
        img = Image.open(codigo_barra_path + ".png")  # Agregar la extensión aquí
    #   img = img.resize((413, 236))  # Redimensionar la imagen
        ws.add_image(xlImage(codigo_barra_path + ".png"), f'D{index+1}')

    # Guardar el archivo de Excel
    wb.save('c:/BARCODES/archivo_con_codigos.xlsx')

if __name__ == "__main__":
    main()   