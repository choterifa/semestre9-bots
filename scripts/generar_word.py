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
            elif item.filename == "word/settings.xml":
                xml_settings = content.decode("utf-8")
                if "updateFields" not in xml_settings:
                    xml_settings = xml_settings.replace("</w:settings>", '<w:updateFields xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="true"/></w:settings>')
                content = xml_settings.encode("utf-8")
            zout.writestr(item, content)

    print(f"✅ Documento base institucional generado en: {ruta_salida}")
    print(f"   👤 Estudiante: {estudiante} ({matricula})")
    print(f"   📋 Tarea: {titulo_tarea}")
    print(f"   📅 Fecha: {fecha}")
    return ruta_salida

def insertar_tabla_contenido_nativa(doc):
    """
    Inserta la Tabla de Contenidos nativa de Microsoft Word (Referencias -> Tabla de contenido)
    en la Hoja 2 del documento.
    Elimina cualquier banner repetido o texto previo de la portada en Hoja 2,
    dejando únicamente el encabezado 'Contenido' y el bloque SDT nativo de Word
    enlazado a los estilos Heading 1, Heading 2 y Heading 3.
    """
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls

    sdt_xml = (
        f'<w:sdt {nsdecls("w")}>\n'
        f'  <w:sdtPr>\n'
        f'    <w:docPartObj>\n'
        f'      <w:docPartGallery w:val="Table of Contents"/>\n'
        f'      <w:docPartUnique/>\n'
        f'    </w:docPartObj>\n'
        f'  </w:sdtPr>\n'
        f'  <w:sdtContent>\n'
        f'    <w:p>\n'
        f'      <w:pPr>\n'
        f'        <w:pStyle w:val="TtuloTDC"/>\n'
        f'        <w:spacing w:before="120" w:after="240"/>\n'
        f'      </w:pPr>\n'
        f'      <w:r>\n'
        f'        <w:rPr>\n'
        f'          <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>\n'
        f'          <w:b/>\n'
        f'          <w:color w:val="0F4761"/>\n'
        f'          <w:sz w:val="32"/>\n'
        f'        </w:rPr>\n'
        f'        <w:t>Contenido</w:t>\n'
        f'      </w:r>\n'
        f'    </w:p>\n'
        f'    <w:p>\n'
        f'      <w:pPr>\n'
        f'        <w:pStyle w:val="TDC1"/>\n'
        f'        <w:tabs>\n'
        f'          <w:tab w:val="right" w:leader="dot" w:pos="9350"/>\n'
        f'        </w:tabs>\n'
        f'      </w:pPr>\n'
        f'      <w:r>\n'
        f'        <w:fldChar w:fldCharType="begin"/>\n'
        f'      </w:r>\n'
        f'      <w:r>\n'
        f'        <w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText>\n'
        f'      </w:r>\n'
        f'      <w:r>\n'
        f'        <w:fldChar w:fldCharType="separate"/>\n'
        f'      </w:r>\n'
        f'      <w:r>\n'
        f'        <w:fldChar w:fldCharType="end"/>\n'
        f'      </w:r>\n'
        f'    </w:p>\n'
        f'  </w:sdtContent>\n'
        f'</w:sdt>'
    )
    sdt = parse_xml(sdt_xml)
    if len(doc.paragraphs) > 2:
        doc.paragraphs[2]._p.addprevious(sdt)
        for p in list(doc.paragraphs[2:5]):
            if p._p.getparent() is not None:
                p._p.getparent().remove(p._p)
        for p in list(doc.paragraphs[2:]):
            if p.text == '' and len(p._p.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}br')) == 0:
                if p._p.getparent() is not None:
                    p._p.getparent().remove(p._p)
    else:
        doc._body._body.append(sdt)

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Uso: python3 generar_word.py <materia: ia|forense> <titulo_tarea> [ruta_salida]")
        sys.exit(1)
        
    materia = sys.argv[1]
    titulo = sys.argv[2]
    salida = sys.argv[3] if len(sys.argv) > 3 else os.path.expanduser(f"~/Downloads/{titulo.replace(' ', '_')}.docx")
    preparar_documento_base(materia, titulo, salida)
