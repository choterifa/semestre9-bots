#!/usr/bin/env python3
"""
Módulo automatizado para generar documentos Word (.docx) institucionales
conservando portada, membrete oficial (header2/footer2) y estilos nativos del ITSP.

REGLAS UNIVERSALES ESTRICTAS DE SEMESTRE 9:
1. Portada Oficial Completa en Hoja 1: Conserva P[0] y P[1] con logotipos, institución,
   estudiante, docente, carrera, grado y grupo, tarea y fecha. Purgado absoluto de P[2:].
2. Hoja 2 Exclusiva para la Tabla de Contenidos: Bloque SDT nativo oficial de Word
   (Referencias -> Tabla de contenido) con campo TOC, título 'Contenido' en #0F4761,
   marcadores nativos (_Toc...), tabuladores con puntos de relleno y números visibles.
3. Salto de Página antes de cada Título de Nivel 1: Toda sección principal (Heading 1)
   a partir de la Sección 2 en adelante inicia obligatoriamente al principio de una nueva hoja.
4. Vinculación formal de estilos: Heading 1 (Ttulo1) y Heading 2 (Ttulo2).
"""
import json
import os
import sys
import zipfile
import re
import html
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
    Conserva al 100% logotipos, membrete (header2/footer2), formas y salto a la Hoja 2.
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
    Conserva P[0] y P[1] que componen la portada oficial completa institucional.
    Elimina cualquier párrafo residual a partir de P[2:] que pudiera provenir
    de la plantilla base ('Reporte de práctica...').
    """
    if len(doc.paragraphs) > 2:
        for p in list(doc.paragraphs[2:]):
            p._p.getparent().remove(p._p)

def finalizar_indice_nativo(doc, pagina_map=None):
    """
    Construye y reemplaza el bloque SDT oficial nativo de Microsoft Word
    (Referencias -> Tabla de contenido) en la Hoja 2 del documento.
    Precarga todas las entradas registradas con estilos TDC1 / TDC2,
    tabulaciones con puntos de relleno (leader='dot'), hipervínculos a los marcadores
    _Toc y números de página calculados, manteniendo el campo TOC activo.
    """
    if not hasattr(doc, '_p_ancla_toc') or doc._p_ancla_toc is None:
        return
    
    secciones = getattr(doc, '_secciones_toc', [])
    if not secciones:
        return
    
    # Estimación de páginas
    # La Sección 1 inicia en la Página 3 tras la Portada (Pág 1) y el Índice (Pág 2).
    # Cada H1 posterior a la Sección 1 agrega un salto de página garantizado.
    pagina_actual = 3
    sdt_entries = []
    
    for sec in secciones:
        bm = sec['bm']
        txt = html.escape(sec['texto'])
        
        if sec['nivel'] == 1:
            if sec.get('salto_aplicado', False):
                pagina_actual += 1
            if pagina_map and sec['texto'] in pagina_map:
                pag_mostrar = pagina_map[sec['texto']]
            else:
                pag_mostrar = pagina_actual
                
            sdt_entries.append(f'''
    <w:p>
      <w:pPr>
        <w:pStyle w:val="TDC1"/>
        <w:tabs>
          <w:tab w:val="right" w:leader="dot" w:pos="9350"/>
        </w:tabs>
      </w:pPr>
      <w:hyperlink w:anchor="{bm}" w:history="1">
        <w:r>
          <w:rPr>
            <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
            <w:b/>
            <w:color w:val="0F4761"/>
            <w:sz w:val="22"/>
          </w:rPr>
          <w:t>{txt}</w:t>
        </w:r>
        <w:r><w:tab/></w:r>
        <w:r>
          <w:fldChar w:fldCharType="begin"/>
        </w:r>
        <w:r>
          <w:instrText xml:space="preserve"> PAGEREF {bm} \\h </w:instrText>
        </w:r>
        <w:r>
          <w:fldChar w:fldCharType="separate"/>
        </w:r>
        <w:r>
          <w:rPr>
            <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
            <w:b/>
            <w:color w:val="0F4761"/>
            <w:sz w:val="22"/>
          </w:rPr>
          <w:t>{pag_mostrar}</w:t>
        </w:r>
        <w:r>
          <w:fldChar w:fldCharType="end"/>
        </w:r>
      </w:hyperlink>
    </w:p>''')
        else:
            if pagina_map and sec['texto'] in pagina_map:
                pag_mostrar = pagina_map[sec['texto']]
            else:
                pag_mostrar = pagina_actual
                
            sdt_entries.append(f'''
    <w:p>
      <w:pPr>
        <w:pStyle w:val="TDC2"/>
        <w:ind w:left="360"/>
        <w:tabs>
          <w:tab w:val="right" w:leader="dot" w:pos="9350"/>
        </w:tabs>
      </w:pPr>
      <w:hyperlink w:anchor="{bm}" w:history="1">
        <w:r>
          <w:rPr>
            <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
            <w:color w:val="2A6F97"/>
            <w:sz w:val="20"/>
          </w:rPr>
          <w:t>{txt}</w:t>
        </w:r>
        <w:r><w:tab/></w:r>
        <w:r>
          <w:fldChar w:fldCharType="begin"/>
        </w:r>
        <w:r>
          <w:instrText xml:space="preserve"> PAGEREF {bm} \\h </w:instrText>
        </w:r>
        <w:r>
          <w:fldChar w:fldCharType="separate"/>
        </w:r>
        <w:r>
          <w:rPr>
            <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
            <w:color w:val="2A6F97"/>
            <w:sz w:val="20"/>
          </w:rPr>
          <w:t>{pag_mostrar}</w:t>
        </w:r>
        <w:r>
          <w:fldChar w:fldCharType="end"/>
        </w:r>
      </w:hyperlink>
    </w:p>''')

    entries_xml = "\n".join(sdt_entries)

    sdt_full_xml = f'''<w:sdt {nsdecls("w")}>
  <w:sdtPr>
    <w:docPartObj>
      <w:docPartGallery w:val="Table of Contents"/>
      <w:docPartUnique/>
    </w:docPartObj>
  </w:sdtPr>
  <w:sdtContent>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="TtuloTDC"/>
        <w:spacing w:before="120" w:after="240"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
          <w:b/>
          <w:color w:val="0F4761"/>
          <w:sz w:val="32"/>
        </w:rPr>
        <w:t>Contenido</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="TDC1"/>
        <w:tabs>
          <w:tab w:val="right" w:leader="dot" w:pos="9350"/>
        </w:tabs>
      </w:pPr>
      <w:r>
        <w:fldChar w:fldCharType="begin"/>
      </w:r>
      <w:r>
        <w:instrText xml:space="preserve"> TOC \\o "1-2" \\h \\z \\u </w:instrText>
      </w:r>
      <w:r>
        <w:fldChar w:fldCharType="separate"/>
      </w:r>
    </w:p>
{entries_xml}
    <w:p>
      <w:r>
        <w:fldChar w:fldCharType="end"/>
      </w:r>
    </w:p>
  </w:sdtContent>
</w:sdt>'''

    sdt_element = parse_xml(sdt_full_xml)
    doc._p_ancla_toc._p.addprevious(sdt_element)
    doc._p_ancla_toc._p.getparent().remove(doc._p_ancla_toc._p)
    doc._p_ancla_toc = None

def iniciar_documento_institucional(materia, titulo_tarea, ruta_salida, fecha=None):
    """
    Función maestra: Inicializa el documento, purga la hoja duplicada del template,
    ubica la Hoja 2 para la Tabla de Contenidos nativa y añade el salto de página
    hacia la Hoja 3 para iniciar el desarrollo del trabajo.
    Retorna el objeto Document de python-docx configurado con autoguardado de índice.
    """
    preparar_documento_base(materia, titulo_tarea, ruta_salida, fecha)
    doc = docx.Document(ruta_salida)
    
    # 1. Purgar párrafos residuales de la plantilla base (conservando P[0] y P[1])
    purgar_hoja_duplicada_template(doc)
    
    # 2. Inicializar estructuras de control para el índice
    doc._ruta_salida = ruta_salida
    doc._secciones_toc = []
    doc._h1_count = 0
    doc._toc_id_counter = 1000
    
    # 3. Párrafo ancla en la Hoja 2 para la Tabla de Contenidos
    # Recordar: P[1] ya finaliza con un salto de página, por lo que estamos en la Hoja 2.
    doc._p_ancla_toc = doc.add_paragraph()
    doc._p_ancla_toc.paragraph_format.space_before = Pt(0)
    doc._p_ancla_toc.paragraph_format.space_after = Pt(0)
    
    # 4. Salto de página tras el índice hacia la Hoja 3 (Desarrollo)
    p_salto_desarrollo = doc.add_paragraph()
    p_salto_desarrollo.paragraph_format.space_before = Pt(0)
    p_salto_desarrollo.paragraph_format.space_after = Pt(0)
    p_salto_desarrollo.add_run().add_break(docx.enum.text.WD_BREAK.PAGE)
    
    # Interceptar save() para garantizar que el índice nativo se finalice siempre
    _original_save = doc.save
    def safe_save(path_or_stream=None):
        target = path_or_stream or getattr(doc, '_ruta_salida', ruta_salida)
        finalizar_indice_nativo(doc)
        _original_save(target)
    doc.save = safe_save
    
    return doc

def agregar_encabezado_1(doc, texto, salto_pagina=True):
    """
    Agrega un Título de Nivel 1 con el estilo nativo Heading 1 (Ttulo1) enlazado al índice.
    REGLA UNIVERSAL: Si salto_pagina=True y ya existe una sección previa en el documento,
    inserta un salto de página previo para iniciar al principio de una nueva hoja.
    """
    if not hasattr(doc, '_h1_count'):
        doc._h1_count = 0
    if not hasattr(doc, '_toc_id_counter'):
        doc._toc_id_counter = 1000
    if not hasattr(doc, '_secciones_toc'):
        doc._secciones_toc = []
        
    salto_aplicado = False
    if salto_pagina and doc._h1_count > 0:
        doc.add_page_break()
        salto_aplicado = True
        
    doc._h1_count += 1
    doc._toc_id_counter += 1
    bm_name = f"_Toc{doc._toc_id_counter}"
    
    p = doc.add_paragraph(style='Heading 1')
    p.paragraph_format.space_before = Pt(16)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.keep_with_next = True
    
    # Bookmark start
    bm_start = parse_xml(f'<w:bookmarkStart {nsdecls("w")} w:id="{doc._toc_id_counter}" w:name="{bm_name}"/>')
    bm_end = parse_xml(f'<w:bookmarkEnd {nsdecls("w")} w:id="{doc._toc_id_counter}"/>')
    p._p.append(bm_start)
    
    r = p.add_run(texto)
    r.font.name = "Calibri"
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = AZUL_INSTITUCIONAL
    
    p._p.append(bm_end)
    doc._secciones_toc.append({
        'nivel': 1,
        'texto': texto,
        'bm': bm_name,
        'p': p,
        'salto_aplicado': salto_aplicado
    })
    return p

def agregar_encabezado_2(doc, texto):
    """
    Agrega un Subtítulo de Nivel 2 con el estilo nativo Heading 2 (Ttulo2) enlazado al índice.
    """
    if not hasattr(doc, '_toc_id_counter'):
        doc._toc_id_counter = 1000
    if not hasattr(doc, '_secciones_toc'):
        doc._secciones_toc = []
        
    doc._toc_id_counter += 1
    bm_name = f"_Toc{doc._toc_id_counter}"
    
    p = doc.add_paragraph(style='Heading 2')
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    
    bm_start = parse_xml(f'<w:bookmarkStart {nsdecls("w")} w:id="{doc._toc_id_counter}" w:name="{bm_name}"/>')
    bm_end = parse_xml(f'<w:bookmarkEnd {nsdecls("w")} w:id="{doc._toc_id_counter}"/>')
    p._p.append(bm_start)
    
    r = p.add_run(texto)
    r.font.name = "Calibri"
    r.font.size = Pt(12.5)
    r.font.bold = True
    r.font.color.rgb = AZUL_INSTITUCIONAL
    
    p._p.append(bm_end)
    doc._secciones_toc.append({
        'nivel': 2,
        'texto': texto,
        'bm': bm_name,
        'p': p,
        'salto_aplicado': False
    })
    return p

def guardar_documento(doc, ruta_salida=None, pagina_map=None):
    """
    Finaliza formalmente la Tabla de Contenidos nativa y guarda el archivo Word.
    """
    finalizar_indice_nativo(doc, pagina_map)
    target = ruta_salida or getattr(doc, '_ruta_salida', None)
    if not target:
        raise ValueError("No se especificó la ruta de salida para guardar el documento.")
    doc.save(target)
    print(f"✅ Documento Word guardado exitosamente en: {target}")
    return target

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Uso: python3 generar_word.py <materia: ia|forense|pentesting> <titulo_tarea> [ruta_salida]")
        sys.exit(1)
        
    materia = sys.argv[1]
    titulo = sys.argv[2]
    salida = sys.argv[3] if len(sys.argv) > 3 else os.path.expanduser(f"~/Downloads/{titulo.replace(' ', '_')}.docx")
    doc = iniciar_documento_institucional(materia, titulo, salida)
    doc.save()
