#!/usr/bin/env python3
"""
Instalador automático multiplataforma para Semestre9-Bots.
Configura el perfil del estudiante e instala la Skill en Gemini, Claude Code y Cursor.
"""
import os
import shutil
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SKILL_SRC = os.path.join(BASE_DIR, "SKILL.md")
CONFIG_PATH = os.path.join(BASE_DIR, "config.json")

def registrar_skill(destino_dir, nombre_asistente):
    try:
        os.makedirs(destino_dir, exist_ok=True)
        shutil.copy2(SKILL_SRC, os.path.join(destino_dir, "SKILL.md"))
        print(f"   [✓] Instalado en {nombre_asistente}: {destino_dir}")
        return True
    except Exception as e:
        print(f"   [!] No se pudo registrar en {nombre_asistente}: {e}")
        return False

def main():
    print("=" * 70)
    print("🚀 INSTALADOR AUTOMÁTICO DE SEMESTRE9-BOTS (ITSP)")
    print("=" * 70)

    # 1. Configuración de perfil
    if not os.path.exists(CONFIG_PATH):
        from scripts.configurar_perfil import solicitar_datos
        solicitar_datos()
    else:
        print("ℹ️  Perfil de estudiante encontrado en config.json.")

    print("\n📦 Registrando Skill en tus asistentes de Inteligencia Artificial...")

    home = os.path.expanduser("~")
    destinos = [
        (os.path.join(home, ".gemini", "config", "skills", "semestre9-bots"), "Gemini / Antigravity"),
        (os.path.join(home, ".claude", "skills", "semestre9-bots"), "Claude Code"),
        (os.path.join(home, ".cursor", "skills", "semestre9-bots"), "Cursor IDE"),
        (os.path.join(home, ".agents", "skills", "semestre9-bots"), "Agentes Generales (.agents)")
    ]

    instalados = 0
    for ruta, nombre in destinos:
        if registrar_skill(ruta, nombre):
            instalados += 1

    print("\n" + "=" * 70)
    print("🎉 ¡INSTALACIÓN COMPLETADA CON ÉXITO!")
    print("=" * 70)
    print("¿Cómo usar Semestre9-Bots con tus tareas?")
    print("1. Abre tu asistente favorito (Gemini, Claude Code o Cursor).")
    print("2. Escribe tu petición pegando las instrucciones de la tarea:")
    print("   Ejemplo: 'Resuelve esta tarea para Inteligencia Artificial: [instrucciones]'")
    print("   El bot detectará la materia, aplicará el formato del docente y creará")
    print("   automáticamente tu carpeta con el cuaderno de Colab (.ipynb) y script (.py).\n")

if __name__ == "__main__":
    main()
