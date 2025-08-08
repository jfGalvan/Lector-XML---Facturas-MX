"""Se realizan las siguientes importaciones necesarias"""
import os
import xml.etree.ElementTree as ET
import shutil
import pandas as pd

def procesar_facturas_xml(directorio_xml, archivo_excel_salida):
    datos_facturas = []
    xmlaprocesar = len(os.listdir(directorio_xml))
    
    # Directorio para los archivos procesados
    directorio_procesados = os.path.join(directorio_xml, 'procesados_xml')
    if not os.path.exists(directorio_procesados):
        os.makedirs(directorio_procesados)

    # Definir namespaces para CFDI 4 (cambiar a cfd/3 para CFDI 3.3)
    namespaces = {
        'cfdi': 'http://www.sat.gob.mx/cfd/4',
        'tfd': 'http://www.sat.gob.mx/TimbreFiscalDigital'
    }
    
    # 1. Recorrer y Leer Archivos XML
    for nombre_archivo in os.listdir(directorio_xml):
        if nombre_archivo.endswith('.xml'):
            ruta_completa_archivo = os.path.join(directorio_xml, nombre_archivo)
            try:
                # Parsear el archivo XML
                tree = ET.parse(ruta_completa_archivo)
                root = tree.getroot()
                
                # 2. Extraer Información Relevante
                try:
                    # Datos del comprobante (elemento raíz)
                    numero_factura = root.get('Serie', '') + root.get('Folio', '')
                    fecha = root.get('Fecha')
                    total = float(root.get('Total', 0))
                    subtotal = float(root.get('SubTotal', 0))
                    moneda = root.get('Moneda', 'MXN')
                    tipo_comprobante = root.get('TipoDeComprobante')

                    # Datos de lo comprado
                    descripcion = root.find('cfdi:Conceptos/cfdi:Concepto', namespaces)
                    descripcion_concepto = descripcion.get('Descripcion') if descripcion is not None else ''

                    # Datos del emisor
                    emisor = root.find('cfdi:Emisor', namespaces)
                    emisor_rfc = emisor.get('Rfc') if emisor is not None else ''
                    emisor_nombre = emisor.get('Nombre') if emisor is not None else ''
                    
                    # Datos del receptor
                    receptor = root.find('cfdi:Receptor', namespaces)
                    receptor_rfc = receptor.get('Rfc') if receptor is not None else ''
                    receptor_nombre = receptor.get('Nombre') if receptor is not None else ''
                    receptor_regfiscal = receptor.get('RegimenFiscalReceptor') if receptor is not None else ''  
                    
                    # UUID del timbre fiscal
                    timbre = root.find('.//tfd:TimbreFiscalDigital', namespaces)
                    uuid = timbre.get('UUID') if timbre is not None else ''
                    
                    # Impuestos
                    impuestos = root.find('cfdi:Impuestos', namespaces)
                    total_impuestos = float(impuestos.get('TotalImpuestosTrasladados', 0)) if impuestos is not None else 0

                    #DatosFactura ///////
                    uso_cfdi = receptor.get('UsoCFDI') if receptor is not None else ''
                    metodo_pago = root.get('MetodoPago')
                    forma_pago = root.get('FormaPago')

                    factura_data = {
                        "ArchivoXML": nombre_archivo,
                        "NumeroFactura": numero_factura,
                        "Fecha": fecha,
                        "Total": total,
                        "Subtotal": subtotal,
                        "Moneda": moneda,
                        "TipoComprobante": tipo_comprobante,
                        "EmisorRFC": emisor_rfc,
                        "EmisorNombre": emisor_nombre,
                        "ReceptorRFC": receptor_rfc,
                        "ReceptorNombre": receptor_nombre,
                        "ReceptorRegimenFiscal": receptor_regfiscal,
                        "UUID": uuid,
                        "TotalImpuestos": total_impuestos,
                        "UsoCFDI": uso_cfdi,
                        "MetodoPago": metodo_pago,
                        "FormaPago": forma_pago,
                        "DescripcionConcepto": descripcion_concepto
                    }
                    
                    datos_facturas.append(factura_data)
                    print(f"Procesado: {nombre_archivo}")
                    
                    # --- Mover el archivo a la carpeta de procesados ---
                    shutil.move(ruta_completa_archivo, os.path.join(directorio_procesados, nombre_archivo))
                    # ---------------------------------------------------
                    
                except (AttributeError, ValueError, TypeError) as e:
                    print(f"Error extrayendo datos de {nombre_archivo}: {e}")
                    continue
                    
            except ET.ParseError as e:
                print(f"Error al parsear el archivo {nombre_archivo}: {e}")
            except Exception as e:
                print(f"Ocurrió un error inesperado con {nombre_archivo}: {e}")
    
    # 3. Estructurar la Información en un DataFrame
    df_facturas = pd.DataFrame(datos_facturas)
    
    # 4. Filtrar y Procesar la Información
    if not df_facturas.empty:
        # Convertir fecha a datetime
        df_facturas['Fecha'] = pd.to_datetime(df_facturas['Fecha'])
        
        # Filtrar facturas con un total mayor a $1
        df_filtrado = df_facturas[df_facturas['Total'] > 1]
        
        # Seleccionar columnas para el reporte final
        df_final = df_filtrado[[ 'Fecha', 'EmisorNombre', 'EmisorRFC', 'Total', 'Moneda', 'UUID','NumeroFactura', 'DescripcionConcepto', 'UsoCFDI', 'MetodoPago', 'FormaPago','ReceptorNombre', 'ReceptorRFC', 'ReceptorRegimenFiscal']]
        
        # Ordenar por fecha
        df_final = df_final.sort_values('Fecha')
        
        print(f"Se procesaron {len(df_facturas)} facturas de un total de {xmlaprocesar} archivos XML")
        print(f"Se filtraron {len(df_final)} facturas con total > 1")
        
    else:
        print("No se encontraron facturas válidas para procesar")
        df_final = pd.DataFrame()
    
    # 5. Exportar a Excel
    try:
        if not df_final.empty:
            df_final.to_excel(archivo_excel_salida, index=False)
            print(f"Datos exportados exitosamente a {archivo_excel_salida}")
        else:
            print("No hay datos para exportar")
    except Exception as e:
        print(f"Error al exportar a Excel: {e}")

# --- Uso de la función ---
# Directorio donde se encuentran tus archivos XML de facturas
DIR_FACTURAS = 'facturas_xml'

# Nombre del archivo Excel de salida
EXCEL_SALIDA = 'reporte_facturas.xlsx'

# Asegúrate de que el directorio exista
if not os.path.exists(DIR_FACTURAS):
    os.makedirs(DIR_FACTURAS)
    print(f"Directorio '{DIR_FACTURAS}' creado. Por favor, coloca tus archivos XML aquí.")
else:
    procesar_facturas_xml(DIR_FACTURAS, EXCEL_SALIDA)

# --- Fin del uso de la función --- 