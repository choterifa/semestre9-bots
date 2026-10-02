#!/usr/bin/env python3
"""
Módulo automatizado para generar documentos Word (.docx) institucionales
conservando portada, membrete oficial (header2/footer2) y estilos nativos del ITSP.
"""
import json
import os
import sys
import zipfile
import re
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")

def cargar_config():
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {
        "estudiante": {
            "nombre_completo": "ESTUDIANTE ITSP",
            "matricula": "04220000",
            "grado_grupo": "9° - Grupo 1",
            "carrera": "INGENIERIA EN SISTEMAS COMPUTACIONALES"
        }
    }

def obtener_plantilla(materia="ia"):
    if materia.lower() in ["ia", "inteligencia artificial"]:
        return os.path.join(TEMPLATES_DIR, "Portada_Inteligencia_Artificial_Base.docx")
    elif materia.lower() in ["forense", "informatica forense", "informática forense"]:
        return os.path.join(TEMPLATES_DIR, "Portada_Informatica_Forense_Base.docx")
    elif materia.lower() in ["pentesting", "holzen", "olsen"]:
        return os.path.join(TEMPLATES_DIR, "Portada_Pentesting_Base.docx")
    else:
        # Por defecto IA
        return os.path.join(TEMPLATES_DIR, "Portada_Inteligencia_Artificial_Base.docx")

def preparar_documento_base(materia, titulo_tarea, ruta_salida, fecha=None):
    """
    Toma la plantilla institucional y personaliza los campos XML de la portada
    con los datos del estudiante del config.json, título y fecha.
    Conserva al 100% logotipos, membrete (header2/footer2) y saltos de página.
    """
    config = cargar_config()
    estudiante = config["estudiante"]["nombre_completo"]
    matricula = config["estudiante"].get("matricula", "04220000")
    
    if not fecha:
        meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", 
                 "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
        now = datetime.now()
        fecha = f"{now.day} de {meses[now.month - 1]} de {now.year}"

    plantilla_path = obtener_plantilla(materia)
    if not os.path.exists(plantilla_path):
        raise FileNotFoundError(f"No se encontró la plantilla en {plantilla_path}")

    # Reemplazo limpio a nivel XML para no alterar gráficos ni relaciones
    os.makedirs(os.path.dirname(os.path.abspath(ruta_salida)), exist_ok=True)
    
    with zipfile.ZipFile(plantilla_path, 'r') as zin, zipfile.ZipFile(ruta_salida, 'w') as zout:
        for item in zin.infolist():
            content = zin.read(item.filename)
            if item.filename == "word/document.xml":
                xml_text = content.decode("utf-8")
                
                # Sustituir título de la tarea
                xml_text = re.sub(r'\[Nombre de la Tarea / Actividad\]', titulo_tarea, xml_text)
                
                # Sustituir estudiante si estaba Carlos o genérico
                xml_text = re.sub(r'Carlos Eduardo Valencia Hernández', estudiante, xml_text)
                xml_text = re.sub(r'ESTUDIANTE ITSP', estudiante, xml_text)
                
                # Sustituir fecha si aplica
                xml_text = re.sub(r'Septiembre de 2026', fecha, xml_text)
                # Reemplazo robusto para fechas divididas en múltiples nodos <w:t> en la portada
                xml_text = re.sub(r'<w:t>Septiembre</w:t></w:r><w:r[^>]*>(?:<w:rPr>.*?</w:rPr>)?<w:t[^>]*>\s*de 2026</w:t>', f'<w:t>{fecha}</w:t>', xml_text)
                
                content = xml_text.encode("utf-8")
            zout.writestr(item, content)

    print(f"✅ Documento base institucional generado en: {ruta_salida}")
    print(f"   👤 Estudiante: {estudiante} ({matricula})")
    print(f"   📋 Tarea: {titulo_tarea}")
    print(f"   📅 Fecha: {fecha}")
    return ruta_salida

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Uso: python3 generar_word.py <materia: ia|forense> <titulo_tarea> [ruta_salida]")
        sys.exit(1)
        
    materia = sys.argv[1]
    titulo = sys.argv[2]
    salida = sys.argv[3] if len(sys.argv) > 3 else os.path.expanduser(f"~/Downloads/{titulo.replace(' ', '_')}.docx")
    preparar_documento_base(materia, titulo, salida)
