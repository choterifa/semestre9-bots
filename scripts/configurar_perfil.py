#!/usr/bin/env python3
"""
Asistente interactivo de configuración de perfil para el estudiante.
Configura nombre, matrícula, directorio de tareas y preferencias de las 3 materias.
"""
import json
import os
import sys

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")
GLOBAL_CONFIG_PATH = os.path.expanduser("~/.semestre9_bots_config.json")

def solicitar_datos():
    print("=" * 70)
    print("🤖 CONFIGURACIÓN INICIAL - SEMESTRE 9 BOTS (ITSP)")
    print("=" * 70)
    print("Ingresa tus datos una sola vez para personalizar automáticamente todas")
    print("tus tareas, cuadernos de Colab, scripts y reportes en Word.\n")

    nombre = input("👤 Tu nombre completo: ").strip()
    while not nombre:
        print("El nombre no puede estar vacío.")
        nombre = input("👤 Tu nombre completo: ").strip()

    # La matrícula tiene base institucional fija 042200
    entrada_mat = input("🆔 Matrícula (completa o últimos dígitos tras 042200): ").strip()
    if not entrada_mat:
        matricula = "04220000"
    elif entrada_mat.startswith("042200"):
        matricula = entrada_mat
    else:
        matricula = f"042200{entrada_mat}"

    # Grado y grupo fijos para toda la generación
    grupo = "9° - Grupo 1"
    carrera = "INGENIERIA EN SISTEMAS COMPUTACIONALES"

    # Detección automática del Sistema Operativo y carpeta de Descargas
    sistema_actual = "Windows" if sys.platform.startswith("win") else ("macOS" if sys.platform == "darwin" else "Linux")
    
    if sistema_actual == "Windows":
        perfil_usuario = os.environ.get("USERPROFILE", "C:\\Users\\Default")
        ruta_salida = os.path.join(perfil_usuario, "Downloads")
    else:
        ruta_salida = os.path.expanduser("~/Downloads")

    os.makedirs(ruta_salida, exist_ok=True)

    config_data = {
        "estudiante": {
            "nombre_completo": nombre,
            "matricula": matricula,
            "grado_grupo": grupo,
            "carrera": carrera
        },
        "materias": {
            "ia": {
                "nombre": "INTELIGENCIA ARTIFICIAL",
                "docente": "MORALES RAMÍREZ ULISES",
                "plantilla": os.path.join(BASE_DIR, "templates", "Portada_Inteligencia_Artificial_Base.docx")
            },
            "forense": {
                "nombre": "INFORMÁTICA FORENSE",
                "docente": "EDGAR ALEJANDRO SAGUNDO DUARTE",
                "plantilla": os.path.join(BASE_DIR, "templates", "Portada_Informatica_Forense_Base.docx")
            },
            "materia3": {
                "nombre": "MATERIA 3",
                "docente": "DOCENTE ASIGNADO",
                "plantilla": ""
            }
        },
        "preferencias": {
            "directorio_salida": ruta_salida,
            "generar_doble_entregable": True,
            "sistema_operativo": sys.platform
        }
    }

    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config_data, f, ensure_ascii=False, indent=2)

    with open(GLOBAL_CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config_data, f, ensure_ascii=False, indent=2)

    print("\n" + "=" * 70)
    print("✅ ¡Perfil configurado exitosamente!")
    print(f"📄 Archivo local: {CONFIG_PATH}")
    print(f"🌍 Archivo global: {GLOBAL_CONFIG_PATH}")
    print(f"📁 Directorio de tareas: {ruta_salida}")
    print("=" * 70 + "\n")
    return config_data

if __name__ == "__main__":
    solicitar_datos()
