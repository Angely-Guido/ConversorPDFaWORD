import os
import shutil
import logging
import fitz  # PyMuPDF
from pdf2docx import Converter

# --- CONFIGURACIÓN ---
INPUT_FOLDER = "C:/Docs/Entrada_PDFs"
OUTPUT_FOLDER = "C:/Docs/Salida_Words"
ERROR_FOLDER = os.path.join(OUTPUT_FOLDER, "_ERRORES")

# Asegurar directorios
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(ERROR_FOLDER, exist_ok=True)

# --- CONFIGURACIÓN DE LOGS (Nivel Forense) ---
# Esto guardará fecha, hora, nivel de severidad y el mensaje exacto
logging.basicConfig(
    filename=os.path.join(ERROR_FOLDER, 'log_detallado.log'),
    level=logging.ERROR,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)

def es_pdf_nativo(ruta_pdf):
    """
    Analiza si el PDF tiene texto seleccionable real.
    Retorna True si es texto (nativo), False si es imagen (escaneado).
    """
    try:
        doc = fitz.open(ruta_pdf)
        texto_total = ""
        # Leemos solo las primeras 3 páginas para no perder tiempo
        for i in range(min(3, len(doc))):
            texto_total += doc[i].get_text()
        
        doc.close()
        
        # Si tiene menos de 50 caracteres en 3 páginas, asumimos que es imagen
        if len(texto_total.strip()) < 50:
            return False 
        return True
    except Exception as e:
        # Si ni siquiera puede abrirlo para leer, está corrupto
        raise Exception(f"Error analizando estructura del PDF: {str(e)}")

def procesar_archivo(archivo):
    ruta_pdf = os.path.join(INPUT_FOLDER, archivo)
    nombre_base = os.path.splitext(archivo)[0]
    ruta_docx = os.path.join(OUTPUT_FOLDER, f"{nombre_base}.docx")
    
    print(f"[*] Analizando: {archivo}...")
    
    try:
        # 1. TRIAGE: ¿Es imagen o texto?
        if es_pdf_nativo(ruta_pdf):
            # ESTRATEGIA A: Conversión Directa (Rápida y Perfecta)
            cv = Converter(ruta_pdf)
            cv.convert(ruta_docx, start=0, end=None)
            cv.close()
            print(f"   -> Éxito (Modo Nativo)")
            
        else:
            # ESTRATEGIA B: OCR (Lenta pero necesaria)
            # Aquí iría tu lógica de Tesseract/OCR
            # Por ahora lanzamos una alerta simulada para que veas el manejo de errores
            # si no tienes el OCR montado aún.
            raise Exception("El PDF es una imagen escaneada y el módulo OCR no está activo todavía.")

    except Exception as e:
        print(f"   [!] ERROR: {archivo} -> Movido a cuarentena.")
        
        # 1. Registrar en el Log con detalle técnico
        logging.error(f"Archivo: {archivo} | Error: {str(e)}", exc_info=True)
        
        # 2. Mover el PDF original a la carpeta de errores (Cuarentena)
        ruta_destino_error = os.path.join(ERROR_FOLDER, archivo)
        
        # Si ya existe en errores, lo reemplazamos o renombramos para no chocar
        if os.path.exists(ruta_destino_error):
            os.remove(ruta_destino_error)
        shutil.move(ruta_pdf, ruta_destino_error)

# --- BUCLE PRINCIPAL ---
def main():
    archivos = [f for f in os.listdir(INPUT_FOLDER) if f.lower().endswith('.pdf')]
    print(f"Procesando {len(archivos)} archivos...\n")
    
    for archivo in archivos:
        procesar_archivo(archivo)

if __name__ == "__main__":
    main()