import os
import sys
import shutil
import subprocess
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docxcompose.composer import Composer
import fitz

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'))

def set_cell_margins(cell, top=140, bottom=140, left=180, right=180):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'''
        <w:tcMar {nsdecls("w")}>
            <w:top w:w="{top}" w:type="dxa"/>
            <w:bottom w:w="{bottom}" w:type="dxa"/>
            <w:left w:w="{left}" w:type="dxa"/>
            <w:right w:w="{right}" w:type="dxa"/>
        </w:tcMar>
    ''')
    tcPr.append(tcMar)

def create_task_full_subdoc(src_path, task_num, task_title, instructions_text, dates_info, out_path):
    """
    Creates a complete subdoc for a task:
    1. Forced page break so separator is ALWAYS on its own fresh page.
    2. Task Separator Page (Title, Indicaciones box, screenshot placeholder).
    3. Page break.
    4. Full task content with:
       - local w:sdt (mini-TOC) completely removed.
       - initial 3 cover lines removed.
       - internal headings demoted to Normal (bold) so TOC ignores them.
    """
    doc_sep = docx.Document()
    
    # Task Title (Heading 2) with page_break_before = True
    p_title = doc_sep.add_paragraph(style='Heading 2')
    p_title.paragraph_format.page_break_before = True
    r_title = p_title.add_run(f"Tarea {task_num}: {task_title}")
    r_title.font.name = "Arial"
    r_title.font.size = Pt(16)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(0x8B, 0x1E, 0x3F)
    p_title.paragraph_format.space_before = Pt(20)
    p_title.paragraph_format.space_after = Pt(12)
    
    # Indicaciones Label
    p_ind = doc_sep.add_paragraph()
    r_ind = p_ind.add_run("Indicaciones de la actividad / Rúbrica institucional Moodle:")
    r_ind.font.name = "Arial"
    r_ind.font.size = Pt(12)
    r_ind.font.bold = True
    r_ind.font.color.rgb = RGBColor(16, 44, 87)
    p_ind.paragraph_format.space_after = Pt(6)
    
    # Box table (1x1)
    tbl = doc_sep.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    set_cell_background(cell, "F8F9FA")
    set_cell_margins(cell, top=160, bottom=160, left=200, right=200)
    
    # Instructions text
    p_inst = cell.paragraphs[0]
    r_inst = p_inst.add_run(instructions_text)
    r_inst.font.name = "Arial"
    r_inst.font.size = Pt(10)
    p_inst.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_inst.paragraph_format.line_spacing = 1.15
    p_inst.paragraph_format.space_after = Pt(10)
    
    # Metadata dates
    p_meta = cell.add_paragraph()
    r_meta = p_meta.add_run(dates_info)
    r_meta.font.name = "Arial"
    r_meta.font.size = Pt(9.5)
    r_meta.font.italic = True
    r_meta.font.color.rgb = RGBColor(70, 70, 70)
    p_meta.paragraph_format.space_after = Pt(16)
    
    # Screenshot placeholder box
    p_ph = cell.add_paragraph()
    p_ph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_ph = p_ph.add_run(
        "------------------------------------------------------------------------------------------------------------------------\n"
        "[ Espacio reservado para captura de pantalla de indicaciones / rúbrica de Moodle ]\n"
        "\n\n\n\n"
        "------------------------------------------------------------------------------------------------------------------------"
    )
    r_ph.font.name = "Arial"
    r_ph.font.size = Pt(9)
    r_ph.font.bold = True
    r_ph.font.color.rgb = RGBColor(120, 120, 120)
    
    # Page Break after separator page
    doc_sep.add_page_break()
    
    # Load and clean task body
    d_src = docx.Document(src_path)
    
    # Remove all w:sdt (mini local Table of Contents)
    for sdt in d_src._body._element.xpath('.//w:sdt'):
        sdt.getparent().remove(sdt)
        
    # Remove first 3 header paragraphs
    for _ in range(3):
        if len(d_src.paragraphs) > 0:
            p = d_src.paragraphs[0]
            p._element.getparent().remove(p._element)
            
    # Demote all internal headings in task body to Normal (bold) so TOC ignores them
    for p in d_src.paragraphs:
        txt = p.text.strip()
        if p.style.name.startswith('Heading 1'):
            p.style = d_src.styles['Normal']
            for r in p.runs:
                r.bold = True
                r.font.name = "Arial"
                r.font.size = Pt(13)
                r.font.color.rgb = RGBColor(16, 44, 87)
            p.paragraph_format.space_before = Pt(14)
            p.paragraph_format.space_after = Pt(6)
        elif p.style.name.startswith('Heading 2'):
            p.style = d_src.styles['Normal']
            for r in p.runs:
                r.bold = True
                r.font.name = "Arial"
                r.font.size = Pt(11.5)
                r.font.color.rgb = RGBColor(16, 44, 87)
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(4)
        elif p.style.name.startswith('Heading'):
            p.style = d_src.styles['Normal']
            for r in p.runs:
                r.bold = True
                r.font.name = "Arial"
                r.font.size = Pt(11)
                r.font.color.rgb = RGBColor(16, 44, 87)
                
    # Append task body to separator doc
    composer = Composer(doc_sep)
    composer.append(d_src)
    composer.save(out_path)
    return out_path

def create_unit2_divider_doc(out_path):
    """Creates the dedicated Unit Divider Page matching Image 2, centered vertically and horizontally."""
    doc = docx.Document()
    
    # Enable vertical centering on this section
    sectPr = doc.sections[0]._sectPr
    sectPr.append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>'))
    
    p = doc.add_paragraph(style='Heading 1')
    p.paragraph_format.page_break_before = True
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(18)
    
    r1 = p.add_run("Parcial 2\n")
    r1.font.name = "Arial"
    r1.font.size = Pt(26)
    r1.font.bold = True
    r1.font.color.rgb = RGBColor(0x8B, 0x1E, 0x3F)
    
    r2 = p.add_run("Unidad 2: Fase de Reconocimiento y Escaneo de Vulnerabilidades")
    r2.font.name = "Arial"
    r2.font.size = Pt(18)
    r2.font.bold = True
    r2.font.color.rgb = RGBColor(0x8B, 0x1E, 0x3F)
    
    doc.save(out_path)
    return out_path

def create_u2_evaluations_and_references_doc(out_path):
    """Creates the 3 evaluation elements of Unidad 2 and consolidated references."""
    doc = docx.Document()
    
    # 1. Coevaluación final (parcial 2)
    p_coev = doc.add_paragraph(style='Heading 2')
    p_coev.paragraph_format.page_break_before = True
    r_coev = p_coev.add_run("Coevaluación final (parcial 2)")
    r_coev.font.name = "Arial"
    r_coev.font.size = Pt(16)
    r_coev.font.bold = True
    r_coev.font.color.rgb = RGBColor(0x8B, 0x1E, 0x3F)
    p_coev.paragraph_format.space_before = Pt(16)
    p_coev.paragraph_format.space_after = Pt(6)
    
    p_form = doc.add_paragraph()
    r_form = p_form.add_run("Formato para Co-evaluación")
    r_form.font.name = "Arial"
    r_form.font.size = Pt(12)
    r_form.font.bold = True
    p_form.paragraph_format.space_after = Pt(4)
    
    p_desc = doc.add_paragraph()
    r_desc = p_desc.add_run("Tomando como referencia el desempeño de tus compañeros de grupo, marca con una 'X' o asigna la puntuación correspondiente a cada una de las competencias indicadas en la tabla.")
    r_desc.font.name = "Arial"
    r_desc.font.size = Pt(9)
    r_desc.font.bold = True
    p_desc.paragraph_format.space_after = Pt(8)
    
    # Scale Table
    scales = ["Nunca 0", "Ocasionalmente 2", "Frecuentemente 4", "Siempre 6"]
    t_scale = doc.add_table(rows=1, cols=4)
    t_scale.alignment = WD_TABLE_ALIGNMENT.CENTER
    for c_idx, s_text in enumerate(scales):
        cell = t_scale.cell(0, c_idx)
        cell.text = s_text
        set_cell_background(cell, "F2F4F7")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].font.name = "Arial"
        p.runs[0].font.size = Pt(9)
        p.runs[0].font.bold = True
        
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    
    coev_rows = [
        ["COMPETENCIA", "Puntuación por Integrante", "Puntuación por Integrante", "Puntuación por Integrante", "Puntuación por Integrante", "Puntuación por Integrante", "Puntuación por Integrante"],
        ["COMPETENCIA", "Erick Espinosa", "Osvaldo Canul", "Lopez Acosta", "Coba Cuevas", "", "Promedio"],
        ["Respeta las ideas de los otros miembros del grupo.", "6", "6", "6", "6", "", "6"],
        ["Desempeña un papel activo en la búsqueda de información relevante y la comparte con el grupo.", "6", "6", "6", "6", "", "6"],
        ["Comparte la información que encuentra con los otros miembros del grupo.", "6", "6", "6", "6", "", "6"],
        ["Presenta sus ideas de una manera coherente.", "6", "6", "6", "6", "", "6"],
        ["Su desempeño en el rol asignado ha contribuido a cumplir con las tareas.", "6", "6", "6", "6", "", "6"],
        ["Su participación permite el desarrollo de óptimas relaciones interpersonales.", "6", "6", "6", "6", "", "6"],
        ["Es puntual a las reuniones para desarrollar las actividades realizadas.", "6", "6", "6", "6", "", "6"],
        ["Es capaz de reconocer y enmendar sus errores.", "6", "6", "6", "6", "", "6"],
        ["Puntuación Total", "Puntuación Total", "Puntuación Total", "Puntuación Total", "Puntuación Total", "Puntuación Total", "48"]
    ]
    
    t_coev = doc.add_table(rows=len(coev_rows), cols=7)
    t_coev.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r_idx, row_data in enumerate(coev_rows):
        for c_idx, val in enumerate(row_data):
            cell = t_coev.cell(r_idx, c_idx)
            cell.text = val
            p = cell.paragraphs[0]
            if len(p.runs) > 0:
                p.runs[0].font.name = "Arial"
                p.runs[0].font.size = Pt(8.5)
                if r_idx < 2 or r_idx == 10:
                    p.runs[0].font.bold = True
            if r_idx < 2:
                set_cell_background(cell, "EAECEF")
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif r_idx == 10:
                set_cell_background(cell, "F2F4F7")
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                if c_idx > 0:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
    # 2. Autoevaluación final (parcial 2)
    doc.add_page_break()
    p_auto = doc.add_paragraph(style='Heading 2')
    p_auto.paragraph_format.page_break_before = True
    r_auto = p_auto.add_run("Autoevaluación final (parcial 2)")
    r_auto.font.name = "Arial"
    r_auto.font.size = Pt(16)
    r_auto.font.bold = True
    r_auto.font.color.rgb = RGBColor(0x8B, 0x1E, 0x3F)
    p_auto.paragraph_format.space_before = Pt(16)
    p_auto.paragraph_format.space_after = Pt(6)
    
    p_form2 = doc.add_paragraph()
    r_form2 = p_form2.add_run("Formato para autoevaluación")
    r_form2.font.name = "Arial"
    r_form2.font.size = Pt(12)
    r_form2.font.bold = True
    p_form2.paragraph_format.space_after = Pt(4)
    
    p_desc2 = doc.add_paragraph()
    r_desc2 = p_desc2.add_run("Tomando como referencia tu desempeño en el salón de clases, marca con una 'X' o asigna la puntuación correspondiente en la columna que mejor describa tu actuar académico durante el segundo parcial.")
    r_desc2.font.name = "Arial"
    r_desc2.font.size = Pt(9)
    r_desc2.font.bold = True
    p_desc2.paragraph_format.space_after = Pt(8)
    
    t_scale2 = doc.add_table(rows=1, cols=4)
    t_scale2.alignment = WD_TABLE_ALIGNMENT.CENTER
    for c_idx, s_text in enumerate(scales):
        cell = t_scale2.cell(0, c_idx)
        cell.text = s_text
        set_cell_background(cell, "F2F4F7")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].font.name = "Arial"
        p.runs[0].font.size = Pt(9)
        p.runs[0].font.bold = True
        
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    
    auto_rows = [
        ["COMPETENCIA", "Autoevaluación", "Autoevaluación", "Autoevaluación", "Autoevaluación"],
        ["COMPETENCIA", "Nunca", "Ocasionalmente", "Frecuentemente", "Siempre"],
        ["Respeto las ideas de los otros integrantes del salón de clases.", "", "", "", "6"],
        ["Desempeño un papel activo en la búsqueda de información relevante y la comparto con el grupo.", "", "", "", "6"],
        ["Soy puntual a las reuniones para desarrollar las actividades encomendadas.", "", "", "", "6"],
        ["Soy capaz de reconocer y enmendar mis errores.", "", "", "", "6"],
        ["Llego con puntualidad a las clases.", "", "", "", "6"],
        ["Participo de manera activa con preguntas o comentarios en la clase.", "", "", "", "6"],
        ["Puntuación Total", "Puntuación Total", "Puntuación Total", "", "36"]
    ]
    
    t_auto = doc.add_table(rows=len(auto_rows), cols=5)
    t_auto.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r_idx, row_data in enumerate(auto_rows):
        for c_idx, val in enumerate(row_data):
            cell = t_auto.cell(r_idx, c_idx)
            cell.text = val
            p = cell.paragraphs[0]
            if len(p.runs) > 0:
                p.runs[0].font.name = "Arial"
                p.runs[0].font.size = Pt(8.5)
                if r_idx < 2 or r_idx == 8:
                    p.runs[0].font.bold = True
            if r_idx < 2:
                set_cell_background(cell, "EAECEF")
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif r_idx == 8:
                set_cell_background(cell, "F2F4F7")
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                if c_idx > 0:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
    # 3. Experiencias de aprendizaje y comentarios finales (parcial 2)
    doc.add_page_break()
    p_exp_hdr = doc.add_paragraph(style='Heading 2')
    p_exp_hdr.paragraph_format.page_break_before = True
    r_exp_hdr = p_exp_hdr.add_run("Experiencias de aprendizaje y comentarios finales (parcial 2)")
    r_exp_hdr.font.name = "Arial"
    r_exp_hdr.font.size = Pt(16)
    r_exp_hdr.font.bold = True
    r_exp_hdr.font.color.rgb = RGBColor(0x8B, 0x1E, 0x3F)
    p_exp_hdr.paragraph_format.space_before = Pt(16)
    p_exp_hdr.paragraph_format.space_after = Pt(6)
    
    p_sub1 = doc.add_paragraph()
    r_sub1 = p_sub1.add_run("Experiencia de aprendizaje")
    r_sub1.font.name = "Arial"
    r_sub1.font.size = Pt(12)
    r_sub1.font.bold = True
    p_sub1.paragraph_format.space_after = Pt(4)
    
    p_txt1 = doc.add_paragraph()
    r_txt1 = p_txt1.add_run(
        "A lo largo de este segundo parcial en la materia de Pentesting, centrado en la Fase de Reconocimiento y Escaneo de Vulnerabilidades, "
        "consolidé habilidades técnicas indispensables para mi formación como futuro ingeniero en sistemas computacionales. Aprendí a diferenciar "
        "con claridad y rigor metodológico entre la recolección pasiva de inteligencia de fuentes abiertas (OSINT) y el escaneo activo de puertos "
        "y servicios. Mediante el uso de técnicas de Google Dorking avanzado, la consulta de respaldos históricos en la Wayback Machine y la correlación "
        "automatizada de entidades en SpiderFoot, pude identificar superficies de ataque expuestas sin enviar un solo paquete directo al objetivo, "
        "comprendiendo cómo los atacantes recopilan credenciales y activos desprotegidos. En la parte activa, el despliegue del laboratorio virtual con "
        "OpenVAS (Greenbone Vulnerability Management) sobre Metasploitable 2 me permitió analizar en profundidad más de 112 NVTs y vulnerabilidades críticas "
        "como VSFTPD v2.3.4 (backdoor Smiley), Samba 'username map script' y puertos RMI desprotegidos, contrastando la severidad CVSS v3.1 con los impactos "
        "operativos reales. Paralelamente, el estudio del marco jurídico mexicano (Artículos 211 bis 1 al 211 bis 7 del Código Penal Federal), el Convenio de "
        "Budapest y el análisis de derecho comparado con el caso Scott Moulton me enseñaron la trascendencia de contar siempre con instrumentos de blindaje "
        "legal (Rules of Engagement, Acuerdos de Confidencialidad y Cartas de Autorización) antes de ejecutar cualquier escaneo."
    )
    r_txt1.font.name = "Arial"
    r_txt1.font.size = Pt(11)
    p_txt1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_txt1.paragraph_format.line_spacing = 1.15
    p_txt1.paragraph_format.space_after = Pt(12)
    
    p_sub2 = doc.add_paragraph()
    r_sub2 = p_sub2.add_run("Comentarios finales")
    r_sub2.font.name = "Arial"
    r_sub2.font.size = Pt(12)
    r_sub2.font.bold = True
    p_sub2.paragraph_format.space_after = Pt(4)
    
    p_txt2 = doc.add_paragraph()
    r_txt2 = p_txt2.add_run(
        "Esta unidad representó un salto de madurez significativo en mi perspectiva sobre la ciberseguridad. Comprobé que un escaneo de vulnerabilidades "
        "no es una prueba de penetración completa, sino una herramienta de diagnóstico que requiere criterio humano, análisis forense y contextualización "
        "para no generar falsos positivos ni causar caídas de servicio por denegación involuntaria (DDoS). La redacción del informe final para la alta "
        "dirección me demostró que el verdadero valor de un auditor de seguridad radica en su capacidad de comunicar riesgos técnicos complejos en términos "
        "de continuidad de negocio y pérdidas financieras para los directivos, proporcionando recomendaciones de hardening estructuradas en fases a corto, "
        "mediano y largo plazo. Agradezco la dinámica práctica y analítica impartida por el docente, ya que estas competencias son las que realmente diferencian "
        "a un profesional ético en el mercado laboral actual."
    )
    r_txt2.font.name = "Arial"
    r_txt2.font.size = Pt(11)
    p_txt2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_txt2.paragraph_format.line_spacing = 1.15
    p_txt2.paragraph_format.space_after = Pt(16)
    
    # 4. Referencias Bibliográficas Consolidadas (APA 7ª Edición)
    doc.add_page_break()
    p_ref_hdr = doc.add_paragraph(style='Heading 2')
    p_ref_hdr.paragraph_format.page_break_before = True
    r_ref_hdr = p_ref_hdr.add_run("Referencias Bibliográficas Consolidadas (APA 7ª Edición)")
    r_ref_hdr.font.name = "Arial"
    r_ref_hdr.font.size = Pt(16)
    r_ref_hdr.font.bold = True
    r_ref_hdr.font.color.rgb = RGBColor(0x8B, 0x1E, 0x3F)
    p_ref_hdr.paragraph_format.space_before = Pt(16)
    p_ref_hdr.paragraph_format.space_after = Pt(12)
    
    references = [
        "Association for Computing Machinery. (2018). ACM Code of Ethics and Professional Conduct. Association for Computing Machinery. https://www.acm.org/code-of-ethics",
        "Cámara de Diputados del H. Congreso de la Unión. (1931/2024). Código Penal Federal: Título Noveno - Revelación de secretos y acceso ilícito a sistemas y equipos de informática (Artículos 211 bis 1 al 211 bis 7). Diario Oficial de la Federación.",
        "Consejo de Europa. (2001). Convenio sobre la Ciberdelincuencia (Convenio de Budapest), Serie de Tratados Europeos No. 185. Budapest.",
        "First.org. (2019). Common Vulnerability Scoring System v3.1: Specification Document. Forum of Incident Response and Security Teams. https://www.first.org/cvss/v3.1/specification-document",
        "Greenbone Networks GmbH. (2024). Greenbone Enterprise Appliance & OpenVAS Architecture Documentation (Version 22.4). Greenbone AG. https://docs.greenbone.net/",
        "MITRE Corporation. (2024). ATT&CK Framework: Reconnaissance & Discovery Tactics (TA0043, TA0007). MITRE Enterprise ATT&CK. https://attack.mitre.org/",
        "National Institute of Standards and Technology. (2008). Technical Guide to Information Security Testing and Assessment (NIST Special Publication 800-115). U.S. Department of Commerce. https://doi.org/10.6028/NIST.SP.800-115",
        "National Institute of Standards and Technology. (2018). Framework for Improving Critical Infrastructure Cybersecurity (Version 1.1). U.S. Department of Commerce.",
        "OWASP Foundation. (2023). Web Security Testing Guide (WSTG v4.2). Open Web Application Security Project. https://owasp.org/www-project-web-security-testing-guide/",
        "Penetration Testing Execution Standard. (2017). PTES Technical Guidelines: Intelligence Gathering & Vulnerability Analysis. http://www.pentest-standard.org/",
        "Rapid7. (2024). Metasploitable 2 Exploitability Guide & Vulnerability Verification. Rapid7 Cybersecurity Resources.",
        "Scarfone, K., Souppaya, M., Cody, A., & Orebaugh, A. (2018). Guide to Storage Encryption Technologies for End User Devices (NIST SP 800-111). NIST."
    ]
    
    for ref in references:
        p_ref = doc.add_paragraph(style='List Paragraph')
        r_ref = p_ref.add_run(ref)
        r_ref.font.name = "Arial"
        r_ref.font.size = Pt(10)
        p_ref.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_ref.paragraph_format.line_spacing = 1.15
        p_ref.paragraph_format.space_after = Pt(6)
        
    doc.save(out_path)
    return out_path

def main():
    print("=== INICIANDO COMPILACION MASTER FINAL SIN RESIDUOS ===")
    
    scratch_dir = r'C:\Users\carlos\.gemini\antigravity\scratch\portafolio_v2'
    os.makedirs(scratch_dir, exist_ok=True)
    
    # 1. Base Document (Unidad 1)
    base_src = r'C:\Users\carlos\OneDrive - Instituto Tecnológico Superior Progreso\Todo\6. Escuela\Semestres\Semestre 9\Pentesting\Unidad 1\Portafolio de evidencias Unidad 1_Carlos Eduardo Valencia.docx'
    clean_base = os.path.join(scratch_dir, '01_clean_u1_base.docx')
    
    doc = docx.Document(base_src)
    
    # Remove any existing w:sdt in base doc if any
    for sdt in doc._body._element.xpath('.//w:sdt'):
        sdt.getparent().remove(sdt)
        
    # Remove paragraph 6 if empty to prevent extra blank page between Cover and TOC
    if len(doc.paragraphs) > 6 and doc.paragraphs[6].text.strip() == '':
        p6 = doc.paragraphs[6]
        p6._element.getparent().remove(p6._element)
        
    # Insert Table of Contents right before Reglas de comportamiento
    for i, p in enumerate(doc.paragraphs):
        if 'Reglas de comportamiento' in p.text:
            p_tit = p.insert_paragraph_before(style='Heading 1')
            r_tit = p_tit.add_run("Índice de Contenido")
            r_tit.font.name = "Arial"
            r_tit.font.size = Pt(20)
            r_tit.font.bold = True
            r_tit.font.color.rgb = RGBColor(0x8B, 0x1E, 0x3F)
            p_tit.paragraph_format.space_before = Pt(12)
            p_tit.paragraph_format.space_after = Pt(14)
            
            p_fld = p.insert_paragraph_before()
            fld_xml = f'''
            <w:p {nsdecls("w")}>
              <w:pPr>
                <w:spacing w:after="240"/>
              </w:pPr>
              <w:r>
                <w:fldChar w:fldCharType="begin"/>
              </w:r>
              <w:r>
                <w:instrText xml:space="preserve"> TOC \\o "1-2" \\h \\z </w:instrText>
              </w:r>
              <w:r>
                <w:fldChar w:fldCharType="separate"/>
              </w:r>
              <w:r>
                <w:fldChar w:fldCharType="end"/>
              </w:r>
            </w:p>
            '''
            p_fld._p.addnext(parse_xml(fld_xml))
            
            p_br2 = p.insert_paragraph_before()
            p_br2.add_run().add_break(docx.enum.text.WD_BREAK.PAGE)
            break
            
    # Clean up Unidad 1 headings:
    for p in doc.paragraphs:
        txt = p.text.strip()
        if '1- Informe teórico-histórico-reflexivo' in txt:
            p.text = 'Tarea 1: Informe teórico-histórico-reflexivo acerca del pentesting'
            p.style = doc.styles['Heading 2']
            p.runs[0].font.name = 'Arial'
            p.runs[0].font.size = Pt(15)
            p.runs[0].font.bold = True
            p.runs[0].font.color.rgb = RGBColor(0x8B, 0x1E, 0x3F)
        elif '2- Presentación individual' in txt:
            p.text = 'Tarea 2: Presentación individual acerca de los fundamentos del pentesting'
            p.style = doc.styles['Heading 2']
            p.runs[0].font.name = 'Arial'
            p.runs[0].font.size = Pt(15)
            p.runs[0].font.bold = True
            p.runs[0].font.color.rgb = RGBColor(0x8B, 0x1E, 0x3F)
        elif '3- Proyecto colaborativo: Plan básico' in txt or '3- Proyecto colaborativo' in txt:
            p.text = 'Tarea 4: Proyecto colaborativo Plan básico de pentesting'
            p.style = doc.styles['Heading 2']
            p.runs[0].font.name = 'Arial'
            p.runs[0].font.size = Pt(15)
            p.runs[0].font.bold = True
            p.runs[0].font.color.rgb = RGBColor(0x8B, 0x1E, 0x3F)
        elif p.style.name.startswith('Heading 1'):
            if any(k in txt for k in ['Reglas de comportamiento', 'Parcial 1', 'Índice']):
                continue
            else:
                p.style = doc.styles['Normal']
                for r in p.runs:
                    r.bold = True
                    r.font.name = "Arial"
                    r.font.size = Pt(13)
                    r.font.color.rgb = RGBColor(16, 44, 87)
                p.paragraph_format.space_before = Pt(14)
                p.paragraph_format.space_after = Pt(6)
        elif p.style.name.startswith('Heading 2'):
            if any(k in txt for k in ['Coevaluación final', 'Autoevaluación final', 'Experiencias de aprendizaje']):
                continue
            elif 'Tarea ' in txt:
                continue
            else:
                p.style = doc.styles['Normal']
                for r in p.runs:
                    r.bold = True
                    r.font.name = "Arial"
                    r.font.size = Pt(11.5)
                    r.font.color.rgb = RGBColor(16, 44, 87)
                p.paragraph_format.space_before = Pt(10)
                p.paragraph_format.space_after = Pt(4)
                
    # Add page break at the very end of Unidad 1
    doc.add_page_break()
    doc.save(clean_base)
    print("Base de Unidad 1 preparada y guardada.")
    
    # 2. Unit Divider Document for Unidad 2
    u2_divider_path = os.path.join(scratch_dir, '02_u2_divider.docx')
    create_unit2_divider_doc(u2_divider_path)
    print("Hoja de presentación de Unidad 2 creada.")
    
    # 3. Prepare the 4 tasks for Unidad 2 with their Separator Pages and Indicaciones
    tasks_config = [
        (
            1,
            "Laboratorio de Inteligencia de Fuentes Abiertas (OSINT) y Reconocimiento Perimetral",
            "Realizar una auditoría práctica de reconocimiento pasivo y recolección de inteligencia de fuentes abiertas (OSINT) sobre la infraestructura y el dominio institucional progreso.tecnm.mx. Emplear técnicas avanzadas de Google Dorking para detección de fuga documental sensible (archivos PDF, configuraciones, bases de datos), análisis forense de snapshots históricos en Wayback Machine, y recolección automatizada y correlación masiva de activos externos mediante SpiderFoot 4.0.0. Diseñar matriz pericial de riesgos y plan de hardening defensivo.",
            "Fecha de entrega: 22 de octubre de 2026 | Modalidad: Individual | Asignatura: Pentesting (9° Semestre)",
            r'C:\Users\carlos\Downloads\Pentesting_Practica_OSINT\Reporte_Laboratorio_OSINT_Pentesting.docx',
            os.path.join(scratch_dir, 'task1_osint_subdoc.docx')
        ),
        (
            2,
            "Escaneo de vulnerabilidades con OpenVAS sobre Metasploitable 2",
            "Utilizando la máquina virtual propuesta por el profesor (Metasploitable 2 en red Host-Only), realiza un escaneo exhaustivo con OpenVAS (Greenbone Vulnerability Management) sobre ella. Elabora el reporte técnico de la auditoría aprovechando al máximo las gráficas y estadísticas de severidad proporcionadas por el escáner (Greenbone Security Assistant). Identificar y analizar pericialmente las vulnerabilidades críticas y altas detectadas (CVEs, CVSS v3.1) y proponer medidas de remediación técnica y hardening defensivo.",
            "Fecha de entrega: 29 de octubre de 2026 | Modalidad: Individual | Asignatura: Pentesting (9° Semestre)",
            r'C:\Users\carlos\Downloads\Practica_OpenVAS_Metasploitable2\Reporte_Practica_OpenVAS_Metasploitable2_Carlos_Valencia.docx',
            os.path.join(scratch_dir, 'task2_openvas_subdoc.docx')
        ),
        (
            3,
            "Ensayo sobre implicaciones éticas y legales del escaneo",
            "Elaborar un ensayo analítico riguroso sobre las implicaciones éticas y legales asociadas a la ejecución de escaneos de vulnerabilidades y puertos en redes informáticas. Analizar el marco deontológico del auditor de ciberseguridad, los riesgos operativos colaterales (denegación de servicio, inestabilidad de equipos), el marco jurídico mexicano bajo los Artículos 211 bis 1 al 211 bis 7 del Código Penal Federal, legislación internacional comparada (Convenio de Budapest, CFAA, caso Scott Moulton) y los instrumentos de blindaje legal necesarios (Rules of Engagement, NDA, LoA).",
            "Fecha de entrega: 30 de octubre de 2026 | Modalidad: Individual | Asignatura: Pentesting (9° Semestre)",
            r'C:\Users\carlos\Downloads\Ensayo_Etico_Legal_Escaneo_Carlos_Valencia\Ensayo_Implicaciones_Eticas_Legales_Escaneo_Carlos_Valencia.docx',
            os.path.join(scratch_dir, 'task3_ensayo_subdoc.docx')
        ),
        (
            4,
            "Informe final y presentación profesional Fase de Reconocimiento",
            "Consolidar los resultados periciales y estratégicos obtenidos durante la Fase de Reconocimiento e Inteligencia de Fuentes Abiertas en un informe ejecutivo formal dirigido a la alta dirección corporativa (C-Level). Presentar un Executive Summary gerencial, cartografía de superficie de ataque externa (EASM), inventario pericial de subdominios, auditoría criptográfica SSL/TLS (SANs), identificación de vulnerabilidad crítica de spoofing en correo institucional por ausencia de directiva DMARC (CVSS 7.5), mapeo metodológico al framework MITRE ATT&CK y un plan integral de hardening en tres fases.",
            "Fecha de entrega: 30 de octubre de 2026 | Modalidad: Individual | Asignatura: Pentesting (9° Semestre)",
            r'C:\Users\carlos\Downloads\Informe_Final_Reconocimiento_Pentesting_Carlos_Valencia\Informe_Final_Fase_Reconocimiento_Carlos_Valencia.docx',
            os.path.join(scratch_dir, 'task4_informe_subdoc.docx')
        )
    ]
    
    for t_num, t_title, t_inst, t_dates, t_src, t_out in tasks_config:
        print(f"Generando subdoc limpio con separador para Tarea {t_num}...")
        create_task_full_subdoc(t_src, t_num, t_title, t_inst, t_dates, t_out)
        
    # 4. Prepare Unidad 2 Evaluations and References
    u2_eval_path = os.path.join(scratch_dir, '03_u2_evaluations_and_refs.docx')
    create_u2_evaluations_and_references_doc(u2_eval_path)
    print("Evaluaciones de Parcial 2 y referencias consolidadas creadas.")
    
    # 5. Compose Master Document
    print("Componiendo documento maestro completo...")
    doc_master = docx.Document(clean_base)
    composer = Composer(doc_master)
    
    # Append Unit 2 divider page
    doc_u2_div = docx.Document(u2_divider_path)
    composer.append(doc_u2_div)
    
    # Append the 4 tasks
    for t_num, t_title, t_inst, t_dates, t_src, t_out in tasks_config:
        print(f"Añadiendo Tarea {t_num}...")
        sub_d = docx.Document(t_out)
        composer.append(sub_d)
        
    # Append evaluations & references
    print("Añadiendo evaluaciones de Parcial 2 y referencias...")
    eval_d = docx.Document(u2_eval_path)
    composer.append(eval_d)
    
    final_docx = os.path.join(scratch_dir, 'Portafolio_de_evidencias_Carlos_Eduardo_Valencia.docx')
    final_pdf = os.path.join(scratch_dir, 'Portafolio_de_evidencias_Carlos_Eduardo_Valencia.pdf')
    composer.save(final_docx)
    print(f"Documento DOCX maestro guardado en: {final_docx}")
    
    # 6. Convert to PDF using WordCraft CLI
    print("Convirtiendo a PDF y actualizando campos de Índice (TOC)...")
    res = subprocess.run([r'C:\Program Files\WordCraft\wordcraft-cli.exe', 'convert', final_docx, final_pdf, '--update-fields'], capture_output=True, text=True)
    print("WordCraft stderr:", res.stderr)
    
    # 7. Distribute and Mirror
    # OneDrive Pentesting Root
    one_root = r'C:\Users\carlos\OneDrive - Instituto Tecnológico Superior Progreso\Todo\6. Escuela\Semestres\Semestre 9\Pentesting'
    shutil.copy2(final_docx, os.path.join(one_root, 'Portafolio de evidencias_Carlos Eduardo Valencia.docx'))
    shutil.copy2(final_pdf, os.path.join(one_root, 'Portafolio de evidencias_Carlos Eduardo Valencia.pdf'))
    print("Guardado en OneDrive Pentesting Raíz.")
    
    # OneDrive Unidad 2
    one_u2 = os.path.join(one_root, 'Unidad 2')
    shutil.copy2(final_docx, os.path.join(one_u2, 'Portafolio de evidencias Unidad 2_Carlos Eduardo Valencia.docx'))
    shutil.copy2(final_pdf, os.path.join(one_u2, 'Portafolio de evidencias Unidad 2_Carlos Eduardo Valencia.pdf'))
    print("Guardado en OneDrive Unidad 2.")
    
    # Local Downloads
    down_dir = r'C:\Users\carlos\Downloads\Portafolio_Evidencias_Pentesting_Carlos_Valencia'
    os.makedirs(down_dir, exist_ok=True)
    shutil.copy2(final_docx, os.path.join(down_dir, 'Portafolio_Evidencias_Pentesting_Carlos_Valencia.docx'))
    shutil.copy2(final_pdf, os.path.join(down_dir, 'Portafolio_Evidencias_Pentesting_Carlos_Valencia.pdf'))
    print("Copiado a Downloads.")
    
    # Mac Tailscale (Z:\)
    z_port = r'Z:\Portafolio_Evidencias_Pentesting_Carlos_Valencia'
    z_u2 = r'Z:\Pentesting\Unidad 2'
    if os.path.exists('Z:\\'):
        os.makedirs(z_port, exist_ok=True)
        os.makedirs(z_u2, exist_ok=True)
        shutil.copy2(final_docx, os.path.join(z_port, 'Portafolio_Evidencias_Pentesting_Carlos_Valencia.docx'))
        shutil.copy2(final_pdf, os.path.join(z_port, 'Portafolio_Evidencias_Pentesting_Carlos_Valencia.pdf'))
        shutil.copy2(final_docx, os.path.join(z_u2, 'Portafolio de evidencias Unidad 2_Carlos Eduardo Valencia.docx'))
        shutil.copy2(final_pdf, os.path.join(z_u2, 'Portafolio de evidencias Unidad 2_Carlos Eduardo Valencia.pdf'))
        print("Copiado a Mac (Z:\\).")
        
    print("=== VERIFICANDO RESULTADO EN PDF ===")
    pdf_check = fitz.open(final_pdf)
    print(f"Total de páginas: {len(pdf_check)}")
    print("\n--- TOC en Página 2 ---")
    print(pdf_check[1].get_text()[:600])

if __name__ == '__main__':
    main()
