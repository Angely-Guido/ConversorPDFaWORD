import os
import shutil
import logging
import fitz  # PyMuPDF
import pytesseract
from pdf2docx import Converter
from pdf2image import convert_from_path
from PIL import Image, ImageOps

# --- CONFIGURACIÓN DE DEPENDENCIAS ---

# 1. Tesseract: Intentamos detectarlo automáticamente
POSIBLES_RUTAS = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    shutil.which("tesseract")
]
# Buscamos la primera ruta válida
TESSERACT_CMD = next((path for path in POSIBLES_RUTAS if path and os.path.exists(path)), None)

if TESSERACT_CMD:
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD
else:
    # No lanzamos error aquí para permitir que funcione el modo nativo, 
    # pero el OCR fallará si se llama.
    logging.warning("ADVERTENCIA: Tesseract no encontrado. El modo OCR no funcionará.")

# 2. Poppler: Se asume que está en el PATH del sistema (gracias a setup_dependencies.ps1)
# Si tu script de instalación falló, aquí podrías forzar la ruta, pero mejor confiar en el PATH.

# --- FUNCIONES INTERNAS (PRIVADAS) ---

def _es_pdf_nativo(ruta_pdf):
    """
    Decide la estrategia de conversión.
    Retorna True si es DIGITAL (Texto/Vectores).
    Retorna False si es ESCANEADO (Imagen plana).
    """
    try:
        doc = fitz.open(ruta_pdf)
        texto_acumulado = 0
        vectores_acumulados = 0
        
        # Muestreo: Analizamos hasta 3 páginas para ser rápidos
        paginas_a_analizar = min(3, len(doc))
        
        for i in range(paginas_a_analizar):
            page = doc[i]
            # Contar caracteres de texto real
            texto_acumulado += len(page.get_text().strip())
            # Contar dibujos vectoriales (líneas de tablas, planos, gráficos)
            vectores_acumulados += len(page.get_drawings())
        
        doc.close()

        # CRITERIOS DE DECISIÓN:
        # 1. Si tiene texto seleccionable decente (>50 caracteres), es nativo.
        if texto_acumulado > 50:
            return True
            
        # 2. Si tiene pocos caracteres pero muchos vectores (>5 trazos), 
        # es probablemente un plano o gráfico digital (Excel exportado). No usar OCR.
        if vectores_acumulados > 5:
            return True
            
        # Si no cumple nada, es una foto pegada en un PDF.
        return False
        
    except Exception as e:
        logging.error(f"Error en triage del PDF {ruta_pdf}: {e}")
        # Ante la duda, asumimos que es imagen para intentar rescatar algo con OCR
        return False

def _corregir_orientacion(imagen):
    """Detecta si la imagen está chueca y la endereza."""
    try:
        # OSD = Orientation and Script Detection
        osd = pytesseract.image_to_osd(imagen)
        rotate_angle = int(osd.split("\nRotate: ")[1].split("\n")[0])
        
        if rotate_angle != 0:
            # Nota: PIL rota Anti-Horario, Tesseract reporta Horario. Ajustamos.
            if rotate_angle == 90:
                return imagen.transpose(Image.ROTATE_270)
            elif rotate_angle == 180:
                return imagen.transpose(Image.ROTATE_180)
            elif rotate_angle == 270:
                return imagen.transpose(Image.ROTATE_90)
    except:
        # Si la imagen es muy ruidosa o blanca, OSD falla. La dejamos igual.
        pass
    return imagen

def _ocr_avanzado_sandwich(pdf_path, docx_path):
    """
    Técnica SANDWICH: 
    Crea un PDF intermedio con capa de texto invisible sobre la imagen original.
    Esto preserva TABLAS, FORMATO y DISEÑO visual.
    """
    if not TESSERACT_CMD:
        raise EnvironmentError("Se requiere Tesseract-OCR para procesar este documento escaneado.")

    # 1. Convertir PDF a imágenes (Requiere Poppler en PATH)
    try:
        images = convert_from_path(pdf_path, dpi=300)
    except Exception as e:
        raise EnvironmentError(f"Error de Poppler (¿Ejecutaste el script setup?): {e}")

    # Archivo temporal intermedio
    pdf_temp_path = pdf_path.replace(".pdf", "_temp_ocr.pdf")
    
    # Creamos un nuevo PDF vacío
    pdf_writer = fitz.open()

    for img in images:
        # A. Enderezar si está rotada
        img = _corregir_orientacion(img)
        
        # B. Tesseract genera una página PDF de una sola capa (Imagen + Texto invisible)
        # lang='spa+eng' permite detectar español e inglés simultáneamente
        pdf_bytes = pytesseract.image_to_pdf_or_hocr(img, extension='pdf', lang='spa+eng')
        
        # C. Insertar esa página en nuestro PDF temporal
        img_doc = fitz.open("pdf", pdf_bytes)
        pdf_writer.insert_pdf(img_doc)
        img_doc.close()

    # Guardamos el PDF "híbrido" temporal
    pdf_writer.save(pdf_temp_path)
    pdf_writer.close()

    # 2. Convertir ese PDF híbrido a Word usando pdf2docx
    # (Ahora pdf2docx verá texto "seleccionable" y podrá armar el Word manteniendo el fondo)
    try:
        cv = Converter(pdf_temp_path)
        cv.convert(docx_path, start=0, end=None, verbose=False)
        cv.close()
    finally:
        # Limpieza: Borrar el archivo temporal pase lo que pase
        if os.path.exists(pdf_temp_path):
            os.remove(pdf_temp_path)

# --- FUNCIÓN PÚBLICA (API) ---

def procesar_pdf(ruta_entrada, ruta_salida):
    """
    Función principal llamada desde el Main o la Interfaz Gráfica.
    """
    try:
        if not os.path.exists(ruta_entrada):
            return {'status': 'error', 'mensaje': 'Archivo no encontrado'}

        # 1. TRIAGE: Decidir estrategia
        es_nativo = _es_pdf_nativo(ruta_entrada)
        metodo_usado = "nativo" if es_nativo else "ocr_sandwich"

        # 2. EJECUCIÓN
        if es_nativo:
            # Estrategia A: Rápida y perfecta para digitales
            cv = Converter(ruta_entrada)
            cv.convert(ruta_salida, start=0, end=None, verbose=False)
            cv.close()
        else:
            # Estrategia B: Lenta pero potente para escaneos
            _ocr_avanzado_sandwich(ruta_entrada, ruta_salida)

        return {'status': 'ok', 'metodo': metodo_usado, 'mensaje': 'Conversión exitosa'}

    except Exception as e:
        return {'status': 'error', 'metodo': 'fallido', 'mensaje': str(e)}