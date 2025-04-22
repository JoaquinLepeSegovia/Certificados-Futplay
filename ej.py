import pandas as pd
from fpdf import FPDF
import os
import unicodedata

# Ruta al archivo Excel
archivo_excel = "Datos futplay.xlsx"

# Leer el archivo Excel
df = pd.read_excel(archivo_excel, header=0)


# Función para limpiar nombres de columnas
def limpiar_columna(col):
    # Eliminar espacios al principio y al final
    col = col.strip()
    # Normalizar unicode (quita acentos, etc.)
    col = unicodedata.normalize("NFKC", col)
    return col

# Aplicar la limpieza a todas las columnas
df.columns = [limpiar_columna(c) for c in df.columns]

# Ruta de la fuente DejaVuSans.ttf (asegúrate de que exista)
ruta_fuente = "C:/Users/joaqu/Ficha Medica/DejaVuSans.ttf/ttf/DejaVuSans.ttf"

# Ruta al escritorio y carpeta de certificados
escritorio = os.path.join(os.path.expanduser("~"), "Ficha Medica")
carpeta_salida = os.path.join(escritorio, "certificados")
os.makedirs(carpeta_salida, exist_ok=True)

# Crear una instancia base del PDF para registrar la fuente UNA VEZ
pdf_base = FPDF()
pdf_base.add_font("DejaVu", "", ruta_fuente, True)

# Definir manualmente las columnas por sección según el nuevo orden
columnas_secciones = {
    "INFORMACIÓN BÁSICA": [
        "Nombre", "Edad", "Tel. Emergencia", "Sede", "Estado civil",
        "Nivel Educacional", "Ocupacion", "¿Enfermedades?",
        "Enfermedades Familiares", "Farmacos", "Alergias",
        "Consumo de sustancias", "x̄  actividad fisica semanal",
        "cirugias u hospitalizaciones", "¿Alguna Lesion?"
    ],
    "ANTROPOMETRÍA": [
        "Peso (kg)", "Talla (cm)", "FC (bpm)", "Sat. (%)", "Pa (mmHg)",
        "IMC", "Estado segun IMC", "Circunferencia cintura",
        "Circunferencia brazo", "Lateralidad"
    ],
    "TESTS FÍSICOS": [
        "Test de Lunge", "Test de Lunge 2", "Single Leg Squad 1",
        "SingLe Leg Squad2", "Salto1", "Salto2", "Illinois (segundos)"
    ]
}

# Generar un PDF por persona
for index, fila in df.iterrows():
    nombre = str(fila.get("Nombre", f"persona_{index}")).replace(" ", "_")
    pdf = FPDF()
    pdf.add_page()
    pdf.add_font("DejaVu", "", ruta_fuente, uni=True)
    pdf.add_font("DejaVu", "B", "DejaVu_Sans/DejaVuSans-bold.ttf", uni=True)
    pdf.set_font("DejaVu", "", 12)

    # Encabezado
    pdf.set_fill_color(49, 124, 209)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("DejaVu", "", 18)
    pdf.cell(0, 15, f"Informe Personal de Salud", 0, 1, 'C', fill=True)
    pdf.ln(10)

    # Nombre grande centrado
    pdf.set_font("DejaVu", "B", 40)
    pdf.set_text_color(206, 131, 46)
    pdf.cell(0, 10, "FutPlay", 0, 1, 'C')
    pdf.ln(10)

    # Línea decorativa
    pdf.set_draw_color(180, 180, 180)
    pdf.line(20, pdf.get_y(), 190, pdf.get_y())
    pdf.ln(5)

    # Procesar cada sección
    for seccion, columnas in columnas_secciones.items():
        pdf.set_font("DejaVu", "B", 14)
        pdf.set_text_color(49, 124, 209)
        pdf.cell(0, 10, seccion, 0, 1, 'L')
        pdf.ln(5)

        pdf.set_font("DejaVu", "", 12)
        pdf.set_fill_color(245, 245, 245)
        pdf.set_draw_color(180, 180, 180)

        alternar = False
        for columna in columnas:
            if columna in df.columns:
                if pdf.get_y() > 260:  # Si se acerca al final de la página, agrega nueva
                    pdf.add_page()
                texto_col = str(columna)
                texto_val = str(fila[columna]) if pd.notna(fila[columna]) else "-"
                pdf.set_fill_color(245, 245, 245) if alternar else pdf.set_fill_color(255, 255, 255)
                pdf.cell(100, 6, texto_col, border=1, fill=True)
                pdf.cell(90, 6, texto_val, border=1, fill=True)
                pdf.ln()
                alternar = not alternar
            else:
                print(f"Advertencia: La columna '{columna}' no se encuentra en los datos.")
                pdf.ln(5)

        pdf.ln(5) 

        # Agregar título "Observaciones"
        pdf.set_font("DejaVu", "B", 12)
        pdf.set_text_color(206, 131, 46)
        pdf.cell(0, 8, "Observaciones", ln=True)

        # Recuadro para observaciones
        x_obs = 20
        y_obs = pdf.get_y()
        ancho_obs = 170
        alto_obs = 8 * 6  # Aproximadamente 15 líneas
        pdf.set_draw_color(0, 0, 0)
        pdf.set_line_width(0.3)
        pdf.rect(x_obs, y_obs, ancho_obs, alto_obs)

        pdf.set_xy(x_obs + 2, y_obs + 2)
        pdf.set_font("DejaVu", "", 12)
        pdf.multi_cell(ancho_obs - 4, 8, "")

        pdf.set_y(y_obs + alto_obs + 4)
        pdf.ln(5)

    # Guardar PDF
    nombre_archivo = os.path.join(carpeta_salida, f"{nombre}.pdf")
    pdf.output(nombre_archivo)

print("\u2705 \u00a1PDFs generados correctamente con subsecciones ajustadas!")