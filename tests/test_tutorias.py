import json
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch
from zipfile import ZipFile
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'scripts'))
from generar_word import obtener_plantilla, preparar_documento_base
from generar_tutorias import nombre_archivo
import configurar_perfil


class TutoriasTest(unittest.TestCase):
    def test_aliases(self):
        for alias in ['tutorias', 'Tutorías', 'tutoría']:
            self.assertTrue(obtener_plantilla(alias).endswith('Portada_Tutorias_Base.docx'))
        self.assertTrue(obtener_plantilla('pentesting').endswith('Portada_Pentesting_Base.docx'))

    def test_cover_placeholders_and_xml(self):
        with tempfile.TemporaryDirectory() as tmp:
            output = str(Path(tmp) / 'portada.docx')
            preparar_documento_base('tutorias', 'Elección & empleo', output, '21 de octubre de 2026')
            with ZipFile(output) as z:
                root = ET.fromstring(z.read('word/document.xml'))
                text = ''.join(root.itertext())
                self.assertIn('TUTORÍAS', text)
                self.assertIn('Elección & empleo', text)
                self.assertNotIn('PENTESTING', text)
                self.assertNotIn('[Fecha de entrega]', text)
                self.assertNotIn('Reporte de práctica', text)

    def test_filename(self):
        self.assertEqual(nombre_archivo('Actividad 1', 'Carlos Eduardo Valencia'), 'Actividad 1_Carlos Eduardo Valencia.pdf')
        with self.assertRaises(ValueError):
            nombre_archivo('../actividad', 'Alumno')

    def test_migration_preserves_profile(self):
        with tempfile.TemporaryDirectory() as tmp:
            a, b = Path(tmp) / 'local.json', Path(tmp) / 'global.json'
            original = {'estudiante': {'nombre_completo': 'Alumno de prueba'}, 'materias': {'ia': {'docente': 'Personalizado'}}}
            for ruta in (a, b): ruta.write_text(json.dumps(original))
            with patch.object(configurar_perfil, 'CONFIG_PATH', str(a)), patch.object(configurar_perfil, 'GLOBAL_CONFIG_PATH', str(b)):
                configurar_perfil.actualizar_materias_existentes()
                configurar_perfil.actualizar_materias_existentes()
            result = json.loads(a.read_text())
            self.assertIn('tutorias', result['materias'])
            self.assertEqual(result['materias']['ia']['docente'], 'Personalizado')
            self.assertEqual(result['estudiante']['nombre_entrega'], 'Alumno de prueba')


if __name__ == '__main__':
    unittest.main()
