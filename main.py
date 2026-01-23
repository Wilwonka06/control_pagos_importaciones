"""
PASO 1: Copiar archivo Control Pagos de COMERCIO con nombre por fecha en carpeta destino
Automatización de Pagos de Importaciones
"""

"""
1. Paso listo 
"""

import shutil # para copiar archivo
from pathlib import Path # para manejar rutas de archivos y carpetas
from datetime import datetime
import locale # para configurar el idioma de la fecha

#configuración de los meses al español

try:
    locale.setlocale(locale.LC_TIME, 'Spanish_Spain.1252') #Para windows
except locale.Error:
    try:
        locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8') #Para linux
    except locale.Error:
        print("No se pudo establecer la configuración regional a español.")

# clase para copia de archivo (Control Pagos.xlsm)
class CopiarArchivo:
    def __init__(self):
        self.ruta_origen = Path(r"C:\Users\auxtesoreria2\OneDrive - GCO\Escritorio\CONTROL PAGOS.xlsx")
        self.ruta_destino = Path(r"C:\Users\auxtesoreria2\OneDrive - GCO\Escritorio\finanzas\info bancos\proyección semana")

    def obtener_fecha_actual(self):
        fecha = datetime.now()
        print("Fecha actual: ", fecha)
        return fecha
    
    def Crear_nombre_archivo(self, fecha):
        dia = fecha.strftime('%d')
        mes = fecha.strftime('%B').capitalize() # Capitalizar el mes en español
        año = fecha.strftime('%Y')

        nombre_archivo = f"Control Pagos {fecha.strftime('%d %B %Y')}.xlsx"
        print("Nombre del archivo: ", nombre_archivo)
        return nombre_archivo
    
    def estructura_carpetas(self, fecha):
        año_carpeta = f"Año {fecha.strftime('%Y')}"
        mes_carpeta = fecha.strftime('%B').upper() # Mes en mayúsculas

        carpeta_destino = self.ruta_destino / año_carpeta / mes_carpeta
        carpeta_destino.mkdir(parents=True, exist_ok=True)

        print("Carpeta destino creada: ", carpeta_destino)
        return carpeta_destino

    def verficar_existencia_original(self):
        if not self.ruta_origen.exists():
            raise FileNotFoundError(f"El archivo original no existe en la ruta: {self.ruta_origen}")
        print("El archivo original existe.")
        return True
    
    def copia_archivo(self):
        print("Iniciando proceso de copia de archivo...")

        #1. verificar existencia del archivo original
        print("Verificando existencia del archivo original...")
        if not self.verficar_existencia_original():
            return None
        print()

        #2. obtener fecha actual
        print("Obteniendo fecha actual...")
        fecha_actual = self.obtener_fecha_actual()
        print()

        #3. crear nombre del archivo
        print("Creando nombre del archivo...")
        nombre_archivo = self.Crear_nombre_archivo(fecha_actual)
        print()

        #4. crear estructura de carpetas
        print("Creando estructura de carpetas...")
        carpetas_destino = self.estructura_carpetas(fecha_actual)

        #5. crear ruta completa del archivo destino
        print("Creando ruta completa del archivo destino...")
        ruta_archivo_destino = carpetas_destino / nombre_archivo
        print("Ruta del archivo destino: ", ruta_archivo_destino)
        print()

        #6  verificar si el archivo ya existe en la ruta de destino
        print("Verificando si el archivo ya existe en la ruta de destino...")

        if ruta_archivo_destino.exists():
           respuesta = input(f"El archivo ya existe en la ruta de destino: {ruta_archivo_destino}. ¿Desea reemplazarlo? (S/N)")
           if respuesta.lower() != 's':
                print("Operación cancelada. El archivo no se ha reemplazado.")
                return None
        else:
            print("El archivo no existe en la ruta de destino. Se procederá a copiar.")
        print()

        #7. copiar archivo
        print("PASO 7: Copiando archivo...")
        try:
            shutil.copy2(self.ruta_origen, ruta_archivo_destino)
            print(f"Archivo copiado exitosamente")
        except Exception as e:
            print(f"Error al copiar archivo: {str(e)}")
            return None

        print("PROCESO COMPLETADO EXITOSAMENTE")
        print(f"\n📄 Archivo creado: {nombre_archivo}")
        print(f"📁 Ubicación: {carpetas_destino}")
        print(f"📊 Ruta completa: {ruta_archivo_destino}")
        print()
        
        return str(ruta_archivo_destino)

def ejecucion_copiador():
    copiador = CopiarArchivo()
    archivo_creado = copiador.copia_archivo()

    if archivo_creado:
        print(f"Archivo creado exitosamente: {archivo_creado}")
    else:
        print("Error al crear el archivo.")


if __name__ == "__main__":
    ejecucion_copiador()