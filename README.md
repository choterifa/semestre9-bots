# 🤖 Semestre9-Bots

> **Suite Académica Automatizada y Skill de IA para 9° Semestre de ISC en el Instituto Tecnológico Superior Progreso (ITSP)**  
> *Compatible con Gemini / Antigravity, Claude Code, Cursor y Terminal.*

`semestre9-bots` es un asistente inteligente diseñado para que los estudiantes de 9° semestre de Ingeniería en Sistemas Computacionales resuelvan, programen y documenten sus prácticas de laboratorio con los más altos estándares metodológicos y de formato exigidos por cada docente.

---

## 🌟 Materias Soportadas

| Materia | Docente | Entregables Principales | Estándar de Formato |
| :--- | :--- | :--- | :--- |
| **Inteligencia Artificial** | Mtro. Ulises Morales Ramírez | Carpeta con `.ipynb` (Google Colab) + `.py` (script autónomo) | *Clean Code* en español (comentarios en el *por qué*), fórmulas KaTeX y salidas precalculadas. |
| **Informática Forense** | Mtro. Edgar Alejandro Sagundo Duarte | Reporte Técnico Oficial en Word (`.docx`) | Estilos nativos (`Heading 1` / `Heading 2` en Aptos Display 20pt/16pt `#0F4761`), tablas sin color, capturas reales de terminal y conclusión en 1ª persona. |
| **Pentesting** | Docente: Martínez García Holzen | Investigaciones Académicas (APA 7ª) o Reportes de Labs | Búsquedas de papers/estándares (OWASP/MITRE/NIST), citas en texto, 3-5 referencias en español, tablas de remediación y Word oficial. |

| **Tutorías** | Portada basada en Pentesting; confirmar docente del grupo | PDF único: portada + plan rellenado + resultados | `<título Moodle>_Carlos Eduardo Valencia.pdf` para Carlos; sin notas internas ni evidencias inventadas. |

---

## 🚀 Instalación Rápida (1 solo paso)

Clona este repositorio e inicia el instalador interactivo:

```bash
git clone https://github.com/choterifa/semestre9-bots.git
cd semestre9-bots
python3 install.py
```

El instalador:
1. Te preguntará tu **Nombre completo** y **Matrícula** (se guardan de forma local en `config.json`).
2. Te permitirá elegir dónde quieres que se guarden tus tareas (por defecto en tu carpeta de `Descargas`).
3. Registrará automáticamente la **Skill** en **Gemini / Antigravity**, **Claude Code** y **Cursor**.

---

## 💡 ¿Cómo se usa?

Una vez instalado, abre tu asistente de IA habitual (Gemini, Claude Code o Cursor) y dile lo que necesitas:

### Ejemplo 1 (Para Inteligencia Artificial):
```text
Tengo esta tarea para Inteligencia Artificial:
"Construir un recomendador Item-to-Item con el Índice de Jaccard a partir de este dataset..."
```
👉 **Resultado automático:** El bot creará una subcarpeta dedicada en tus Descargas con:
- `practica_jaccard.ipynb` (listo para abrir en Google Colab con portada, explicaciones y resultados).
- `practica_jaccard.py` (script autocontenido listo para correr en terminal).

### Ejemplo 2 (Para Informática Forense):
```text
Resuelve la práctica 4 de Informática Forense con los siguientes requisitos:
[Pegar la rúbrica del Mtro. Sagundo]
```
👉 **Resultado automático:** El bot generará el documento Word `.docx` con la portada institucional oficial, tu nombre, membrete del ITSP, tablas limpias, formato de cadena de custodia y evidencias de terminal.

---

## 📁 Estructura del Repositorio

```text
semestre9-bots/
├── SKILL.md                               # Definición de la Skill unificada (Gemini, Claude, Cursor)
├── config.example.json                    # Plantilla de configuración del alumno
├── install.py                             # Instalador interactivo multiplataforma
├── README.md                              # Guía de uso completa
├── templates/
│   ├── Portada_Inteligencia_Artificial_Base.docx  # Plantilla oficial ITSP (Materia IA)
│   ├── Portada_Informatica_Forense_Base.docx      # Plantilla oficial ITSP (Materia Forense)
│   ├── Portada_Pentesting_Base.docx              # Plantilla oficial ITSP (Pentesting)
│   └── Portada_Tutorias_Base.docx                # Portada única para Tutorías
├── scripts/
│   └── configurar_perfil.py               # Asistente de configuración de perfil
└── examples/
    └── ia_recomendador_jaccard/           # Práctica de referencia completa
        ├── dataset.txt
        ├── recomendador_musical.ipynb
        └── recomendador_musical.py
```

---

## ⚙️ Personalización Manual

Si necesitas cambiar tu nombre o la ruta donde se guardan tus tareas, simplemente ejecuta:

```bash
python3 scripts/configurar_perfil.py
```

O edita directamente el archivo `config.json`.

---

## 🔒 Privacidad y Autoría

Este repositorio es público; no subas datos personales ni credenciales. Todos los scripts y archivos generados permanecen de forma **100% local** en tu computadora. Nada se sincroniza ni sube a servicios externos sin tu confirmación explícita.

## Tutorías: configuración y entrega

Tutorías aparece como cuarta opción en la Skill y se registra en los perfiles nuevos.
Al volver a ejecutar `python3 install.py`, también se agrega a perfiles existentes sin sobrescribir los datos guardados.

1. Descarga el plan de sesión de Moodle y rellena los apartados pedidos, manteniendo las instrucciones del docente.
2. Añade solo evidencias reales de resultados: tablas, respuestas, comparaciones o reflexión. No inventes asistencia, fotos, videos ni enlaces.
3. Exporta el plan rellenado a PDF.
4. Genera la entrega con la portada base al principio:

```bash
python3 -m pip install -r requirements-tutorias.txt
# También se requiere LibreOffice (comando soffice).
python3 scripts/generar_tutorias.py --plan-pdf plan_rellenado.pdf --titulo 'Título exacto Moodle' --fecha '21 de octubre de 2026' --nombre-entrega 'Carlos Eduardo Valencia'
```

El comando no rellena el plan automáticamente ni sube nada a Moodle. Une la portada y el PDF ya rellenado en un único archivo. Usa los datos de `config.json`, que permanece local e ignorado por Git.
Para el nombre abreviado habitual, guarda `estudiante.nombre_entrega` en tu configuración local; no altera el nombre completo de la portada.
No entrega DOCX ni añade notas sobre el asistente o la preparación del documento. El PDF final debe revisarse visualmente antes de subirlo con autorización.
Si el plan pide video y la instrucción de clase parece distinta, confirma el requisito; no borres la consigna ni reemplaces el video con contenido ajeno.

Archivos nuevos: `templates/Portada_Tutorias_Base.docx`, `scripts/generar_tutorias.py`, `requirements-tutorias.txt` y `tests/test_tutorias.py`.
