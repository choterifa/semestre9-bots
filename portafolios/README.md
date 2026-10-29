# Repositorio Institucional de Portafolios de Evidencias - 9° Semestre ISC (ITSP)

Este directorio almacena los **Portafolios de Evidencias** institucionales compilados por unidad y el acumulado master para las materias del 9° Semestre de Ingeniería en Sistemas Computacionales en el Instituto Tecnológico Superior Progreso (ITSP).

---

## Estructura de Directorios

```text
portafolios/
├── README.md
└── pentesting/
    ├── master/
    │   ├── Portafolio_de_evidencias_Carlos_Eduardo_Valencia.docx
    │   └── Portafolio_de_evidencias_Carlos_Eduardo_Valencia.pdf
    ├── unidad_1/
    │   ├── Portafolio de evidencias Unidad 1_Carlos Eduardo Valencia.docx
    │   ├── Portafolio de evidencias Unidad 1_Carlos Eduardo Valencia.pdf
    │   ├── Tarea 1 - Informe teórico-histórico-reflexivo acerca del pentesting.docx (.pdf)
    │   ├── Tarea 2 - Presentación individual acerca de los fundamentos del pentesting.docx (.pdf)
    │   └── Tarea 4 - Proyecto colaborativo Plan básico de pentesting.docx (.pdf)
    └── unidad_2/
        ├── Portafolio de evidencias Unidad 2_Carlos Eduardo Valencia.docx
        ├── Portafolio de evidencias Unidad 2_Carlos Eduardo Valencia.pdf
        ├── Tarea 1 - Reporte Laboratorio OSINT Pentesting.docx (.pdf)
        ├── Tarea 2 - Escaneo de vulnerabilidades con OpenVAS.docx (.pdf)
        ├── Tarea 3 - Ensayo sobre implicaciones eticas y legales del escaneo.docx (.pdf)
        └── Tarea 4 - Informe final y presentacion profesional Fase de Reconocimiento.docx (.pdf)
```

---

## Reglas Estrictas Oficiales de Formato para Portafolios

1. **Portada Oficial y Metadatos Institucionales:**
   - Nombre del estudiante: `Valencia Hernández Carlos Eduardo` (Matrícula: `04220057`).
   - Carrera: `Ingeniería en Sistemas Computacionales` | Semestre: `9° - Grupo 1`.
   - Docente: `Dr. Holzen A. Martínez García` (Pentesting).

2. **Índice de Contenido (Página 2):**
   - Campo nativo de Word: `TOC \o "1-2" \h \z`.
   - **Regla Estricta:** El índice **NO** debe contener subíndices de secciones internas (1.1, 1.2, 2.1, etc.).
   - Solo se listan Unidades (`Heading 1`), Tareas (`Heading 2`) y los 3 elementos de evaluación final (`Heading 2`).
   - Los subtítulos internos de las tareas se configuran en estilo `Normal` (negrita y color `#102C57`) para evitar que se filtren al índice.

3. **Hoja de Presentación / Separador de Unidad:**
   - **Centrado vertical:** Configurado mediante sección con alineación vertical centrada (`<w:vAlign w:val="center"/>`) o espaciado de sección equivalente.
   - **Centrado horizontal:** Texto centrado en color guinda institucional (`#8B1E3F`):
     - `Parcial X` (26 pt, negrita).
     - `Unidad X: [Nombre de la Unidad]` (18 pt, negrita).
   - En página propia y exclusiva con saltos de página antes y después.

4. **Separador de Cada Tarea (`Salto de Hoja` Obligatorio):**
   - Inicia en página nueva (`page_break_before = True`).
   - Título: `Tarea X: [Nombre de la Tarea]` (`Heading 2`, Arial 16 pt, negrita, `#8B1E3F`).
   - Cuadro estilizado con `Indicaciones de la actividad / Rúbrica institucional Moodle`.
   - Cuadro punteado: `[ Espacio reservado para captura de pantalla de indicaciones / rúbrica de Moodle ]`.
   - **Salto de página obligatorio** (`PageBreak`) justo después del separador, comenzando el contenido de la tarea en la página siguiente.
   - **Eliminación obligatoria de mini-índices locales (`w:sdt`)** de cada tarea individual antes de ensamblarlas en el portafolio.

5. **Los 3 Elementos de Evaluación y Conclusiones al Final de CADA Unidad:**
   - `Coevaluación final (parcial X)`: Matriz con los integrantes del equipo (Erick Espinosa, Osvaldo Canul, Lopez Acosta, Coba Cuevas, Promedio: 6.0).
   - `Autoevaluación final (parcial X)`: Matriz personal evaluada en el nivel máximo (*Siempre: 6*).
   - `Experiencias de aprendizaje y comentarios finales (parcial X)`: Redacción reflexiva y auténtica en primera persona singular (*"yo"*).

6. **Referencias Bibliográficas Consolidadas:**
   - Formato APA 7ª edición unificando las fuentes de todas las tareas.

7. **Doble Entregable:**
   - Versión editable `.docx` y versión final `.pdf` compilada con `wordcraft-cli.exe convert <docx> <pdf> --update-fields`.
