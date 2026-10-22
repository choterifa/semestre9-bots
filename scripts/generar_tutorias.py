#!/usr/bin/env python3
"""Combina la portada institucional de Tutorías con un plan ya rellenado en PDF."""
import argparse
from pathlib import Path
import shutil
import subprocess
import tempfile
from generar_word import preparar_documento_base, cargar_config


def nombre_archivo(titulo, nombre):
    for valor in (titulo, nombre):
        if not valor.strip() or any(c in valor for c in '<>:"/\\|?*\x00\r\n'):
            raise ValueError('Título y nombre deben ser nombres de archivo válidos, sin separadores.')
    return f'{titulo.strip()}_{nombre.strip()}.pdf'


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('--plan-pdf', required=True, type=Path, help='Plan rellenado y evidencias en un solo PDF')
    p.add_argument('--titulo', required=True, help='Título exacto de la actividad en Moodle')
    p.add_argument('--fecha', required=True, help='Fecha que debe aparecer en la portada')
    p.add_argument('--nombre-entrega', help='Nombre del archivo; no cambia el nombre completo de portada')
    p.add_argument('--salida', type=Path, help='Carpeta de salida')
    a = p.parse_args()
    config = cargar_config()
    alumno = config['estudiante']
    nombre = a.nombre_entrega or alumno.get('nombre_entrega') or alumno['nombre_completo']
    try:
        archivo = nombre_archivo(a.titulo, nombre)
    except ValueError as e:
        p.error(str(e))
    if not a.plan_pdf.is_file() or a.plan_pdf.suffix.lower() != '.pdf':
        p.error('--plan-pdf debe ser un archivo PDF existente.')
    soffice = shutil.which('soffice') or shutil.which('libreoffice')
    if not soffice:
        p.error('Instala LibreOffice y agrega soffice al PATH.')
    from pypdf import PdfReader, PdfWriter
    plan = PdfReader(str(a.plan_pdf))
    if plan.is_encrypted or not plan.pages:
        p.error('El plan debe ser un PDF legible, no cifrado y con páginas.')
    salida = a.salida or Path(config.get('preferencias', {}).get('directorio_salida', '~/Downloads')).expanduser() / a.titulo
    salida.mkdir(parents=True, exist_ok=True)
    destino = salida / archivo
    if destino.resolve() == a.plan_pdf.resolve():
        p.error('La salida no puede reemplazar el PDF de entrada.')
    if destino.exists():
        p.error(f'La salida ya existe: {destino}. Elige otra carpeta para no sobrescribirla.')
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        portada = tmp / 'portada.docx'
        preparar_documento_base('tutorias', a.titulo, str(portada), a.fecha)
        subprocess.run([soffice, f'-env:UserInstallation={tmp.as_uri()}/perfil', '--headless', '--convert-to', 'pdf', '--outdir', str(tmp), str(portada)], check=True, capture_output=True)
        pdf = tmp / 'portada.pdf'
        if not pdf.exists():
            raise RuntimeError('LibreOffice no generó la portada PDF.')
        cover = PdfReader(str(pdf))
        if len(cover.pages) != 1:
            raise RuntimeError('La portada debe ocupar exactamente una página; revisa título y datos.')
        writer = PdfWriter()
        writer.add_page(cover.pages[0])
        for pagina in plan.pages:
            writer.add_page(pagina)
        with destino.open('wb') as f:
            writer.write(f)
    print(f'PDF listo para revisión visual: {destino}')
    return destino


if __name__ == '__main__':
    main()
