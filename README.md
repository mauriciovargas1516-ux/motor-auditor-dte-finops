# Motor Automatizado de Auditoría de DTEs (FinOps Pipeline) 

Este proyecto es un motor de automatización financiera desarrollado en Python, diseñado para optimizar el ciclo de **Cuentas por Pagar (AP)** y **Cuentas por Cobrar (AR)**. El sistema extrae, limpia y audita masivamente Documentos Tributarios Electrónicos (DTE) bajo el estándar del Servicio de Impuestos Internos (SII) de Chile, mitigando el riesgo de pagos duplicados y eliminando el ingreso manual de datos.

## Impacto de Negocio (ROI)
- **Mitigación de Riesgo:** Previene activamente la fuga de liquidez bloqueando facturas duplicadas antes de que ingresen al ERP.
- **Eficiencia Operativa:** Reduce el tiempo de procesamiento (Data Entry) de horas a fracciones de segundo por lote.
- **Escalabilidad:** Capacidad para procesar volúmenes ilimitados de facturas sin fatiga humana ni pérdida de precisión.

##  Características Técnicas Principales
1. **Modo Operativo Híbrido (Switch UI):** Interfaz gráfica inicial (`tkinter`) que permite al usuario seleccionar dinámicamente si auditará la cartera de Proveedores (AP) o de Clientes (AR).
2. **Extracción Vectorial y Regex de Alta Precisión:** Utiliza `pdfplumber` y Expresiones Regulares avanzadas para capturar RUTs, Folios, Fechas y Montos, ignorando ruido visual, espacios invisibles y caracteres anómalos (ej. guiones largos).
3. **Auditoría Forense en Tiempo Real:** Emplea `pandas` para buscar colisiones en los DataFrames (RUT + Folio). Las anomalías se segregan automáticamente a una lista de cuarentena.
4. **Exportación Estructurada:** Genera un reporte consolidado en `.xlsx` con datos limpios y casteados (montos enteros), listo para la importación directa a sistemas contables o cruce con cartolas bancarias.

## Tecnologías Utilizadas
- **Lenguaje:** Python 3.x
- **Librerías Core:** `pandas` (Análisis), `pdfplumber` (Extracción PDF), `re` (Pattern Matching), `openpyxl` (Exportación Excel).
- **GUI:** `tkinter` (Interacción y despliegue de alertas nativas).

## Cómo usar este motor
1. Clona el repositorio e instala las dependencias:
   ```bash
   pip install pandas pdfplumber openpyxl
