import os
import sys
from pathlib import Path

# Agregar el directorio src al path para imports
sys.path.append(str(Path(__file__).parent))

from certificados_utils import CertificadoGenerator

def main():
    """Función principal para generar certificados"""
    
    # Rutas relativas desde la raíz del proyecto
    project_root = Path(__file__).parent.parent
    
    template_path = project_root / "data" / "templates" / "plantilla_certificado.docx"
    input_data_path = project_root / "data" / "input" / "datos_certificados.xlsx"
    output_dir = project_root / "data" / "output" / "docx"
    
    # Verificar que existen los archivos necesarios
    if not template_path.exists():
        print(f"Error: No se encuentra la plantilla en {template_path}")
        return
    
    if not input_data_path.exists():
        print(f"Error: No se encuentra el archivo de datos en {input_data_path}")
        print("Crea un archivo Excel con las columnas necesarias en data/input/")
        return
    
    # Crear directorio de salida si no existe
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Generar certificados
    generator = CertificadoGenerator(
        template_path=str(template_path),
        output_dir=str(output_dir)
    )
    
    try:
        generator.generate_from_excel(str(input_data_path))
        print("✅ Certificados generados exitosamente")
    except Exception as e:
        print(f"❌ Error al generar certificados: {e}")

if __name__ == "__main__":
    main()