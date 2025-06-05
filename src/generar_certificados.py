from docx2pdf import convert
import pandas as pd
import os
import time
import win32com.client
import re

from certificados_utils import (
    obtener_curso_completo,
    fpuntoscedula,
    formatear_fecha,
    obtener_parametro
)

# Rutas relativas
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_PATH = os.path.join(BASE_DIR, "templates", "plantillasitocar.docx")
EXCEL_PATH = os.path.join(BASE_DIR, "data", "datos_certificados.xlsx")
OUTPUT_DIR = os.path.join(BASE_DIR, "data", "certificados")
DOCX_DIR = os.path.join(OUTPUT_DIR, "docx")
PDF_DIR = os.path.join(OUTPUT_DIR, "pdf")
os.makedirs(DOCX_DIR, exist_ok=True)
os.makedirs(PDF_DIR, exist_ok=True)

def editardatos(nombre, numerocedula, curse, date, cargo, ri, output_file):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(TEMPLATE_PATH)

    # Leer valores actuales de la plantilla
    lname = ldate = lcurse = Lri = Lcargo = parameters = None
    for shape in doc.Shapes:
        try:
            if shape.TextFrame.HasText:
                text = shape.TextFrame.TextRange.Text
                lines = text.replace('\r', '\n').strip().split('\n')
                if shape.Name == "nombre y cc":
                    lname = lines[0].strip()
                elif shape.Name == "fecha":
                    ldate = lines[0].strip()
                elif shape.Name == "nombre curso":
                    lcurse = lines[0].strip()
                elif shape.Name == "RI":
                    Lri = text
                elif shape.Name == "cargo":
                    Lcargo = text
                elif shape.Name == "parametros":
                    parameters = text
        except AttributeError:
            continue

    # Preparar datos
    cedula2 = fpuntoscedula(str(numerocedula))
    curse_completo = obtener_curso_completo(curse)
    fecha_formateada = formatear_fecha(date)
    parametro = obtener_parametro(curse_completo)

    # Expresiones regulares
    pcedula = re.compile(r"C\.C\.\s[\d\.]+")
    pdate = re.compile(r"EXPEDIDO EN BOGOTÁ EL DÍA \d{2} DE [A-Z]+ DE \d{4}")
    pcurso = re.compile(r"ASISTIÓ Y APROBÓ EL CURSO DE (.+)")

    # Editar los cuadros de texto
    for shape in doc.Shapes:
        if shape.TextFrame.HasText:
            text = shape.TextFrame.TextRange.Text
            nc = shape.Name
            if nc == "nombre y cc":
                text = text.replace(lname, nombre)
                if pcedula.search(text):
                    text = pcedula.sub(f"C.C. {cedula2}", text)
            elif nc == "fecha":
                if pdate.search(text):
                    text = pdate.sub(f"EXPEDIDO EN BOGOTÁ EL DÍA {fecha_formateada}", text)
            elif nc == "nombre curso":
                if pcurso.search(text):
                    text = pcurso.sub(f"ASISTIÓ Y APROBÓ EL CURSO DE {curse_completo} \nCON UNA INTENSIDAD DE CUARENTA (40) HORAS", text)
            elif nc == "parametros":
                text = parametro
            elif nc == "cargo":
                text = cargo
            elif nc == "registro":
                text = "R.I. " + str(ri)
            shape.TextFrame.TextRange.Text = text.strip()

    # Guardar y cerrar
    ruta_docx = os.path.join(DOCX_DIR, output_file)
    doc.SaveAs(ruta_docx)
    doc.Close()
    word.Quit()
    time.sleep(1)

    # Convertir a PDF
    if os.path.exists(ruta_docx):
        try:
            archivo_pdf = os.path.join(PDF_DIR, output_file.replace(".docx", ".pdf"))
            convert(ruta_docx, archivo_pdf)
            print(f"✅ Certificado generado: {output_file} y PDF.")
        except Exception as e:
            print(f"❌ Error al convertir a PDF: {e}")

def main():
    try:
        df = pd.read_excel(EXCEL_PATH)
    except Exception as e:
        print(f"❌ Error al leer el archivo Excel: {e}")
        exit()
        return

    for index, row in df.iterrows():
        try:
            nombre = str(row["APELLIDOS Y NOMBRES COMPLETO"]) if pd.notna(row["APELLIDOS Y NOMBRES COMPLETO"]) else "NOMBRE"
            numerocedula = str(row["# DOCUMENTO"]) if pd.notna(row["# DOCUMENTO"]) else "1234567890"
            curse = str(row["CURSO DICTADO"]) if pd.notna(row["CURSO DICTADO"]) else "XXXXXX"
            date = str(row["FECHA EXPEDICIÓN"]) if pd.notna(row["FECHA EXPEDICIÓN"]) else "XXXXXXX"
            cargo = str(row["PROFESIÓN"]) if pd.notna(row["PROFESIÓN"]) else "XXXXXX"
            ri = str(row["REGISTRO INTERNO (RI)"]) if pd.notna(row["REGISTRO INTERNO (RI)"]) else "0000"
            output_file = f"certificado_{ri}.docx"
            editardatos(nombre, numerocedula, curse, date, cargo, ri, output_file)
        except Exception as e:
            print(f"❌ Error al procesar la fila {index + 1}: {e}")
            exit()

if __name__ == "__main__":
    main()