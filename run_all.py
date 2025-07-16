import sys
import subprocess
import os

def main():
    # Carpeta donde están los scripts
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # Lista de scripts a ejecutar en orden
    scripts = [
        "Puntuaciones teoricas.py",
        "Puntuaciones ELO.py"
    ]

    for fname in scripts:
        path = os.path.join(script_dir, fname)
        print(f"\n▶ Ejecutando {fname} …")
        subprocess.run([sys.executable, path], check=True)

    print("\n✅ ¡Todos los procesos han terminado correctamente!")

if __name__ == "__main__":
    main()
