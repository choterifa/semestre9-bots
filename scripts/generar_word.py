#!/usr/bin/env python3
"""
Módulo automatizado para generar documentos Word (.docx) institucionales
conservando portada, membrete oficial (header2/footer2) y estilos nativos del ITSP.
Regla obligatoria: Elimina la segunda hoja duplicada del template e integra
el índice nativo de Word enlazado con Heading 1 (Ttulo1) y Heading 2 (Ttulo2).
"""
import json
import os
import sys
import zipfile
import re
from datetime import datetime
import docx
from docx.shared import Pt, RGBColor, Inches
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")

AZUL_INSTITUCIONAL = RGBColor(15, 71, 97) # #0F4761
GRIS_TEXTO = RGBColor(34, 34, 34)

def cargar_config():
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "r", encoding="utf-8-sig") as f:
            return json.load(f)
    return {
        "estudiante": {
            "nombre_completo": "Carlos Eduardo Valencia Hernández",
            "matricula": "04220057",
            "grado_grupo": "9° - Grupo 1",
            "carrera": "INGENIERIA EN SISTEMAS COMPUTACIONALES"
        }
    }

def obtener_plantilla(materia="ia"):
    m = materia.lower()
    if m in ["ia", "inteligencia artificial"]:
        return os.path.join(TEMPLATES_DIR, "Portada_Inteligencia_Artificial_Base.docx")
    elif m in ["forense", "informatica forense", "informática forense"]:
        return os.path.join(TEMPLATES_DIR, "Portada_Informatica_Forense_Base.docx")
    elif m in ["pentesting", "holzen", "olsen"]:
        return os.path.join(TEMPLATES_DIR, "Portada_Pentesting_Base.docx")
    else:
        return os.path.join(TEMPLATES_DIR, "Portada_Pentesting_Base.docx")

def preparar_documento_base(materia, titulo_tarea, ruta_salida, fecha=None):
    """
    Toma la plantilla institucional y personaliza los campos XML de la portada
    con los datos del estudiante del config.json, título y fecha.
    Conserva al 100% logotipos, membrete (header2/footer2) y saltos de página.
    """
    config = cargar_config()
    estudiante = config["estudiante"]["nombre_completo"]
    matricula = config["estudiante"].get("matricula", "04220057")
    
    if not fecha:
        meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", 
                 "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
        now = datetime.now()
        fecha = f"{now.day} de {meses[now.month - 1]} de {now.year}"

    plantilla_path = obtener_plantilla(materia)
    if not os.path.exists(plantilla_path):
        raise FileNotFoundError(f"No se encontró la plantilla en {plantilla_path}")

    os.makedirs(os.path.dirname(os.path.abspath(ruta_salida)), exist_ok=True)
    
    with zipfile.ZipFile(plantilla_path, 'r') as zin, zipfile.ZipFile(ruta_salida, 'w') as zout:
        for item in zin.infolist():
            content = zin.read(item.filename)
            if item.filename == "word/document.xml":
                xml_text = content.decode("utf-8")
                
                # Sustituir título de la tarea
                xml_text = re.sub(r'\[Nombre de la Tarea / Actividad\]', titulo_tarea, xml_text)
                
                # Sustituir estudiante
                xml_text = re.sub(r'Carlos Eduardo Valencia Hernández', estudiante, xml_text)
                xml_text = re.sub(r'ESTUDIANTE ITSP', estudiante, xml_text)
                
                # Sustituir fecha
                xml_text = re.sub(r'Septiembre de 2026', fecha, xml_text)
                xml_text = re.sub(r'<w:t>Septiembre</w:t></w:r><w:r[^>]*>(?:<w:rPr>.*?</w:rPr>)?<w:t[^>]*>\s*de 2026</w:t>', f'<w:t>{fecha}</w:t>', xml_text)
                
                content = xml_text.encode("utf-8")
            elif item.filename == "word/settings.xml":
                xml_settings = content.decode("utf-8")
                if "updateFields" not in xml_settings:
                    xml_settings = xml_settings.replace("</w:settings>", '<w:updateFields xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="true"/></w:settings>')
                content = xml_settings.encode("utf-8")
            zout.writestr(item, content)

    print(f"[OK] Documento base institucional generado en: {ruta_salida}")
    print(f"   [INFO] Estudiante: {estudiante} ({matricula})")
    print(f"   [INFO] Tarea: {titulo_tarea}")
    print(f"   [INFO] Fecha: {fecha}")
    return ruta_salida

def purgar_hoja_duplicada_template(doc):
    """
    REGLA ESTRICTA UNIVERSAL DE SEMESTRE 9:
    Elimina permanentemente los párrafos sobrantes de la plantilla base (P[1:])
    que contienen el encabezado repetido:
    'Reporte de práctica / [Materia] / Ingeniería en Sistemas Computacionales...'
    Garantiza que NUNCA aparezca esa segunda hoja vacía o repetida.
    """
    parrafos_sobrantes = [p for p in doc.paragraphs[1:]]
    for p in parrafos_sobrantes:
        p._p.getparent().remove(p._p)

def insertar_tabla_contenido_nativa(doc):
    """
    Inserta la Tabla de Contenidos nativa de Microsoft Word (Referencias -> Tabla de contenido)
    en la Hoja 2 del documento, vinculada formalmente a Heading 1 (Ttulo1) y Heading 2 (Ttulo2).
    """
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
        f'        <w:instrText xml:space="preserve"> TOC \\o "1-2" \\h \\z \\u </w:instrText>\n'
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
    doc._body._body.append(sdt)

def iniciar_documento_institucional(materia, titulo_tarea, ruta_salida, fecha=None):
    """
    Función maestra: Inicializa el documento, purga la hoja duplicada del template,
    crea el salto a la Hoja 2, inyecta la Tabla de Contenidos nativa y añade
    el salto de página a la Hoja 3 para iniciar el desarrollo del trabajo.
    Retorna el objeto Document de python-docx listo para redactar.
    """
    preparar_documento_base(materia, titulo_tarea, ruta_salida, fecha)
    doc = docx.Document(ruta_salida)
    
    # 1. Purgar párrafos no deseados de la segunda hoja del template
    purgar_hoja_duplicada_template(doc)
    
    # 2. Salto de página tras la portada hacia la Hoja 2
    p_br1 = doc.add_paragraph()
    p_br1.paragraph_format.space_before = Pt(0)
    p_br1.paragraph_format.space_after = Pt(0)
    p_br1.add_run().add_break(docx.enum.text.WD_BREAK.PAGE)
    
    # 3. Insertar Índice Nativo de Word en la Hoja 2
    insertar_tabla_contenido_nativa(doc)
    
    # 4. Salto de página tras el índice hacia la Hoja 3 (Desarrollo)
    p_br2 = doc.add_paragraph()
    p_br2.paragraph_format.space_before = Pt(0)
    p_br2.paragraph_format.space_after = Pt(0)
    p_br2.add_run().add_break(docx.enum.text.WD_BREAK.PAGE)
    
    return doc

def agregar_encabezado_1(doc, texto):
    """Agrega un Título de Nivel 1 con el estilo nativo Heading 1 (Ttulo1) enlazado al índice."""
    p = doc.add_paragraph(style='Heading 1')
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(texto)
    r.font.name = "Calibri"
    r.font.size = Pt(17)
    r.font.bold = True
    r.font.color.rgb = AZUL_INSTITUCIONAL
    return p

def agregar_encabezado_2(doc, texto):
    """Agrega un Subtítulo de Nivel 2 con el estilo nativo Heading 2 (Ttulo2) enlazado al índice."""
    p = doc.add_paragraph(style='Heading 2')
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(texto)
    r.font.name = "Calibri"
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = AZUL_INSTITUCIONAL
    return p

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Uso: python3 generar_word.py <materia: ia|forense|pentesting> <titulo_tarea> [ruta_salida]")
        sys.exit(1)
        
    materia = sys.argv[1]
    titulo = sys.argv[2]
    salida = sys.argv[3] if len(sys.argv) > 3 else os.path.expanduser(f"~/Downloads/{titulo.replace(' ', '_')}.docx")
    iniciar_documento_institucional(materia, titulo, salida)
