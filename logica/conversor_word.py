import os
import logging
import pythoncom  # Necesario para que funcione en hilos (UI)
from docx2pdf import convert

def convertir_word_a_pdf(ruta_docx, ruta_pdf_salida):
    """
    Convierte DOCX a PDF usando el motor de Microsoft Word.
    NOTA PARA EL FRONTEND: Esta función es bloqueante y lenta (abre Word).
    ¡EJECUTAR EN UN HILO APARTE!
    """
    try:
        # 1. Validaciones básicas
        if not os.path.exists(ruta_docx):
            return {'status': 'error', 'mensaje': 'El archivo DOCX no existe'}

        # 2. Inicializar COM (OBLIGATORIO si esto corre en un hilo secundario)
        # Si no pones esto, la aplicación crasheará al intentar abrir Word desde la UI.
        pythoncom.CoInitialize()

        # 3. Conversión
        # Usamos abspath porque Word a veces se confunde con rutas relativas
        convert(os.path.abspath(ruta_docx), os.path.abspath(ruta_pdf_salida))

        return {'status': 'ok', 'mensaje': 'Conversión exitosa'}

    except Exception as e:
        logging.error(f"Error DOCX->PDF: {e}", exc_info=True)
        return {
            'status': 'error', 
            'mensaje': f"Fallo al invocar Word. ¿Está instalado?: {str(e)}"
        }
        
    finally:
        # Limpieza de recursos COM para no dejar basura en memoria
        try:
            pythoncom.CoUninitialize()
        except:
            pass