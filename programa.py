import pandas as pd
from fpdf import FPDF, XPos, YPos
import os

# Ruta al archivo Excel
archivo_excel = "Datos futplay.xlsx"

# Leer el archivo Excel
df = pd.read_excel(archivo_excel, header=0)

# Ruta de la fuente DejaVuSans.ttf (asegúrate de que exista)
ruta_fuente = "C:/Users/joaqu/Ficha Medica/DejaVuSans.ttf/ttf/DejaVuSans.ttf"

# Ruta al escritorio y carpeta de certificados
escritorio = os.path.join(os.path.expanduser("~"), "Desktop")
carpeta_salida = os.path.join(escritorio, "certificados")
os.makedirs(carpeta_salida, exist_ok=True)

# Crear una instancia base del PDF para registrar la fuente UNA VEZ
pdf_base = FPDF()
pdf_base.add_font("DejaVu", "", ruta_fuente)

# Generar un PDF por persona
for index, fila in df.iterrows():
    nombre = str(fila.get("Nombre", f"persona_{index}")).replace(" ", "_")
    pdf = FPDF()
    pdf.add_page()
    pdf.add_font("DejaVu", "", ruta_fuente,uni=True) 
    pdf.set_font("DejaVu", "", 12)

      # Marco
    pdf.set_draw_color(0, 102, 204)
    pdf.set_line_width(1.5)
    pdf.rect(10, 10, 190, 277)

    # Encabezado
    pdf.set_fill_color(220, 230, 241)  # Azul claro
    pdf.set_text_color(0, 51, 102)
    pdf.set_font("DejaVu", "", 18)
    pdf.cell(0, 15, f"Informe Personal de Salud", 0, 1, 'C', fill=True)
    pdf.ln(10)

    # Nombre grande centrado
    pdf.set_font("DejaVu", "", 20)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(0, 10, "FutPlay", 0, 1, 'C')
    pdf.ln(10)

    # Línea decorativa
    pdf.set_draw_color(180, 180, 180)
    pdf.set_line_width(0.5)
    pdf.line(20, pdf.get_y(), 190, pdf.get_y())
    pdf.ln(5)

    # Tabla de datos
    pdf.set_font("DejaVu", "", 12)
    pdf.set_fill_color(245, 245, 245)
    pdf.set_draw_color(200, 200, 200)

    alternar = False
    for columna, valor in fila.items():
        texto_col = str(columna)
        texto_val = str(valor if pd.notna(valor) else "-")
        pdf.set_fill_color(245, 245, 245) if alternar else pdf.set_fill_color(255, 255, 255)
        pdf.cell(120, 8, texto_col, border=1, fill=True)
        pdf.cell(70, 8, texto_val, border=1, fill=True)
        pdf.ln()
        alternar = not alternar

    # Guardar el PDF
    nombre_archivo = os.path.join(carpeta_salida, f"{nombre}.pdf")
    pdf.output(nombre_archivo)

print("✅ ¡PDFs generados correctamente en la carpeta 'certificados' del escritorio!")
