# Automatización de Control de Pagos - Importaciones

Este proyecto automatiza el flujo de trabajo para la gestión y proyección de pagos de importaciones. Permite generar archivos de proyección semanal y actualizar automáticamente el archivo maestro de control de pagos, asegurando la integridad de los datos y el formato.

## 🚀 Funcionalidades Principales

1.  **Interfaz Gráfica Intuitiva**:
    *   Selección de fecha mediante calendario interactivo.
    *   Cálculo automático del próximo miércoles (día habitual de proyección).

2.  **Generación de Proyección Semanal**:
    *   Copia el archivo base `CONTROL PAGOS.xlsx` (origen).
    *   Filtra los registros cuya `FECHA DE VENCIMIENTO` (o Pago) coincida con la fecha seleccionada y tengan estado 'PAGAR'.
    *   Genera un nuevo archivo Excel con el nombre de la fecha (ej. `04 FEBRERO 2026.xlsx`) en la carpeta correspondiente al año y mes.
    *   Crea una segunda hoja con los datos agrupados por Importador y Proveedor, calculando totales.
    *   **Preservación de Formato**: Utiliza automatización nativa de Excel (COM) para mantener imágenes, estilos y macros del archivo original.

3.  **Actualización del Archivo Maestro**:
    *   Anexa los registros detallados al archivo final `CONTROL PAGOS.xlsx` (destino).
    *   **Expansión Automática de Tabla**: Detecta la tabla de Excel existente y redimensiona el rango automáticamente para incluir los nuevos registros, manteniendo fórmulas y formatos condicionales.

4.  **Validaciones y Seguridad**:
    *   Detección de archivos bloqueados/abiertos con sistema de reintento y alertas al usuario.
    *   Validación de columnas requeridas y limpieza de nombres.

## 📋 Requisitos del Sistema

*   **Sistema Operativo**: Windows (Requerido para la automatización COM de Excel).
*   **Software**: Microsoft Excel instalado.
*   **Python**: 3.8 o superior.

## 🛠️ Instalación y Configuración

1.  **Clonar o descargar el repositorio**.

2.  **Crear un entorno virtual** (recomendado):
    ```bash
    python -m venv venv
    .\venv\Scripts\activate
    ```

3.  **Instalar dependencias**:
    ```bash
    pip install -r requirements.txt
    ```
    *Nota: `pywin32` es crucial para la interacción con Excel.*

## ▶️ Uso

1.  Asegúrese de que el archivo origen `CONTROL PAGOS.xlsx` esté actualizado y cerrado (o guardado).
2.  Ejecute el script principal:
    ```bash
    python control_pagos_1_1.py
    ```
3.  En la ventana emergente, seleccione la fecha para la proyección (por defecto sugiere el próximo miércoles).
4.  Haga clic en **"EJECUTAR PROCESO"**.
5.  El sistema:
    *   Creará la carpeta del mes si no existe.
    *   Generará el archivo de proyección.
    *   Actualizará el archivo maestro.
    *   Mostrará mensajes de confirmación o alerta en caso de errores (ej. archivo abierto).

## 📂 Estructura del Proyecto

*   `control_pagos_1_1.py`: Script principal con toda la lógica de negocio e interfaz gráfica.
*   `requirements.txt`: Lista de librerías Python necesarias.
*   `README.md`: Documentación del proyecto.

## ⚠️ Notas Importantes

*   **Rutas de Archivos**: Las rutas a los archivos de origen y destino están configuradas en el código (`control_pagos_1_1.py`). Asegúrese de que correspondan a su estructura de carpetas local o OneDrive.
*   **Excel Interactivo**: El script abre instancias de Excel en segundo plano. Evite interactuar con otras ventanas de Excel mientras el proceso se ejecuta para prevenir conflictos.
