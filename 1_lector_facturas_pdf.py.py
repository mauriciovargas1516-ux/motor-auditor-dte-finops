import os
import re
import logging
import pandas as pd
import pdfplumber
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

# ========================================================
# 0. CONFIGURACIÓN DEL SISTEMA Y MENÚ INTERACTIVO
# ========================================================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s | %(levelname)s | %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    handlers=[
        logging.FileHandler("auditoria_extraccion_pdf.log", encoding='utf-8'),
        logging.StreamHandler()
    ]
)

# Iniciamos el motor gráfico oculto de Windows
root = tk.Tk()
root.withdraw() 

# Lanzamos el menú inteligente antes de hacer cualquier cosa
MODO_OPERACION = simpledialog.askinteger(
    "Motor FinOps - Selección de Módulo",
    "¿Qué flujo de caja deseas auditar hoy?\n\n"
    "[ 0 ] = Cuentas por Pagar (Proveedores)\n"
    "[ 1 ] = Cuentas por Cobrar (Clientes)\n\n"
    "Ingresa 0 o 1:",
    minvalue=0, maxvalue=1
)

# Si el usuario cierra la ventana o presiona cancelar
if MODO_OPERACION is None:
    logging.error("❌ Operación cancelada por el usuario.")
    exit()

nombre_reporte = 'Cuentas_Por_Pagar_Consolidado.xlsx' if MODO_OPERACION == 0 else 'Cuentas_Por_Cobrar_Consolidado.xlsx'
tipo_operacion = "CUENTAS POR PAGAR" if MODO_OPERACION == 0 else "CUENTAS POR COBRAR"

logging.info("="*75)
logging.info(f"INICIANDO MOTOR FINOPS V6.0 - MODO: {tipo_operacion}")
logging.info("="*75)

# ========================================================
# 1. OPTIMIZACIÓN EXTREMA: REGEX PRE-COMPILADAS (CHILE DTE)
# ========================================================
P_RUT = re.compile(r'(\d{1,2}[.\s]*\d{3}[.\s]*\d{3}\s*[\-–—−_]\s*[0-9Kk])', re.IGNORECASE)
P_FOLIO = re.compile(r'(?:FOLIO|N°|NÚMERO|FACTURA N°|Nº)\s*[:#]?\s*(\d{2,10})', re.IGNORECASE)
P_TOTAL = re.compile(r'(?:TOTAL|TOTAL A PAGAR|MONTO TOTAL)\s*[:\$]?\s*([\d\.]+)', re.IGNORECASE)
P_FECHA = re.compile(r'Fecha Emision:\s*(\d{1,2}\s+de\s+[A-Za-z]+\s+del\s+\d{4})', re.IGNORECASE)

def limpiar_monto(monto_str):
    try:
        return int(str(monto_str).replace('.', '').strip())
    except:
        return 0

# ========================================================
# 2. SELECCIÓN DE CARPETA
# ========================================================
logging.info("Esperando que el usuario seleccione la carpeta con los PDFs...")
carpeta_pdfs = filedialog.askdirectory(title=f"Selecciona la carpeta de PDFs ({tipo_operacion})")

if not carpeta_pdfs:
    logging.error("Operación cancelada. No se seleccionó ninguna carpeta.")
    exit()

# ========================================================
# 3. MOTOR DE LECTURA Y LIMPIEZA MILITAR
# ========================================================
datos_facturas = []
archivos_pdf = [f for f in os.listdir(carpeta_pdfs) if f.lower().endswith('.pdf')]
errores = 0

logging.info(f"Se detectaron {len(archivos_pdf)} archivos PDF. Iniciando lectura...")

for archivo in archivos_pdf:
    ruta_pdf = os.path.join(carpeta_pdfs, archivo)
    
    try:
        with pdfplumber.open(ruta_pdf) as pdf:
            texto_factura = pdf.pages[0].extract_text()
            
            if not texto_factura:
                raise ValueError("PDF sin texto vectorial seleccionable.")

            todos_los_ruts = P_RUT.findall(texto_factura)
            match_folio = P_FOLIO.search(texto_factura)
            match_total = P_TOTAL.search(texto_factura)
            match_fecha = P_FECHA.search(texto_factura)

            if todos_los_ruts:
                try:
                    rut_sucio = todos_los_ruts[MODO_OPERACION].upper()
                except IndexError:
                    rut_sucio = todos_los_ruts[0].upper()
                
                rut_limpio = re.sub(r'\s+', '', rut_sucio)
                rut = re.sub(r'[\-–—−_]', '-', rut_limpio)
            else:
                rut = "SIN RUT"

            folio = match_folio.group(1) if match_folio else "SIN FOLIO"
            fecha = match_fecha.group(1) if match_fecha else "SIN FECHA"
            monto = limpiar_monto(match_total.group(1)) if match_total else 0

            datos_facturas.append({
                'ARCHIVO_ORIGEN': archivo,
                'RUT_OBJETIVO': rut,
                'N_FACTURA': folio,
                'FECHA_EMISION': fecha,
                'MONTO_FACTURA': monto
            })
            
            logging.info(f"✅ [OK] {archivo} -> RUT: {rut} | Folio: {folio} | Monto: ${monto}")

    except Exception as e:
        errores += 1
        logging.error(f"❌ [ERROR] No se pudo leer {archivo}: {e}")

# ========================================================
# 4. AUDITORÍA FORENSE Y DETECCIÓN DE DUPLICADOS
# ========================================================
if not datos_facturas:
    logging.warning("No se extrajo información válida de los PDFs.")
    exit()

df_facturas = pd.DataFrame(datos_facturas)

df_facturas['ES_DUPLICADO'] = df_facturas.duplicated(subset=['RUT_OBJETIVO', 'N_FACTURA'], keep=False)
duplicados = df_facturas[df_facturas['ES_DUPLICADO'] == True]
cantidad_duplicados = len(duplicados)

if cantidad_duplicados > 0:
    logging.warning(f"🚨 ¡ALERTA ROJA! Se detectaron {cantidad_duplicados} anomalías (duplicados).")
else:
    logging.info("✅ Auditoría superada: Base de datos limpia.")

# ========================================================
# 5. EXPORTACIÓN DEL REPORTE GERENCIAL
# ========================================================
ruta_escritorio = os.path.join(os.path.expanduser('~'), 'OneDrive', 'Escritorio', nombre_reporte)

with pd.ExcelWriter(ruta_escritorio, engine='openpyxl') as writer:
    df_listas = df_facturas[df_facturas['ES_DUPLICADO'] == False].drop(columns=['ES_DUPLICADO'])
    df_listas.to_excel(writer, sheet_name='Data_Validada', index=False)
    
    if cantidad_duplicados > 0:
        duplicados.to_excel(writer, sheet_name='🚨 CUARENTENA_DUPLICADOS', index=False)

logging.info("="*75)
logging.info(f"PROCESO TERMINADO. Se procesaron {len(archivos_pdf) - errores} facturas con éxito.")

messagebox.showinfo(
    "Motor FinOps Terminado", 
    f"Operación: {tipo_operacion}\nFacturas procesadas: {len(archivos_pdf)}\nDuplicados bloqueados: {cantidad_duplicados}\n\nRevisa el reporte en tu Escritorio."
)