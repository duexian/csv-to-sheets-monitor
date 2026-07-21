# **🔄 CSV to Sheets Monitor**

**CSV to Sheets Monitor** es una herramienta en Python de segundo plano que monitorea en tiempo real la carpeta de **Descargas** (Downloads) del sistema. Cuando detecta que un archivo .csv específico ha sido descargado, procesa sus datos y los sincroniza automáticamente en un libro de **Google Sheets** preconfigurado, enviando una notificación de escritorio al finalizar.

## **✨ Características Principales**

> * **Monitoreo en tiempo real:** Detecta de forma instantánea nuevas descargas usando watchdog.  
> * **Configuración desacoplada (rules.yaml):** Agrega o modifica reglas para tus archivos sin tocar una sola línea de código.  
> * **Protección de fórmulas:** En lugar de eliminar la pestaña completa, limpia exactamente las columnas utilizadas por el CSV (A..última\_columna), preservando las columnas a la derecha que contengan fórmulas o cálculos.  
> * **Formato seguro de texto:** Agrega un apóstrofe inicial (') a los datos para evitar que Google Sheets altere números largos, IDs, ceros a la izquierda o formatos de fecha.  
> * **Registro de auditoría (Notas en A1):** Escribe una nota en la celda A1 indicando la fecha, hora exacta y nombre del archivo utilizado en la última actualización.  
> * **Notificaciones de escritorio:** Avisos nativos mediante plyer confirmando el éxito de la carga o notificando errores.  
> * **Multi-usuario y portable:** Detecta dinámicamente la ruta Path.home() del sistema operativo actual.

## **📂 Estructura del Repositorio**

`csv-to-sheets-monitor/`  
`├── .gitignore               # Archivo de protección para no subir tokens/credenciales`  
`├── README.md                # Documentación del proyecto`  
`├── requirements.txt         # Librerías necesarias de Python`  
`├── rules.yaml               # Reglas de mapeo (CSV -> Google Sheets)`  
`└── csv_to_sheets_monitor.py # Script principal de ejecución`

## **🛠️ Requisitos e Instalación**

### **1\. Prerrequisitos**

> * **Python 3.8+** instalado.  
> * Una cuenta de Google con acceso a **Google Cloud Console**.

### **2\. Clonar e Instalar Dependencias**

`# 1. Clona el repositorio`  
`git clone https://github.com/duexian/csv-to-sheets-monitor.git`  
`cd csv-to-sheets-monitor`

`# 2. (Opcional) Crea y activa un entorno virtual`  
`python -m venv venv`  
`# En Windows:`  
`venv\Scripts\activate`  
`# En Mac/Linux:`  
`source venv/bin/activate`

`# 3. Instala los paquetes requeridos`  
`pip install -r requirements.txt`

## **⚙️ Configuración**

### **1\. Credenciales de Google Cloud API**

> 1. Ve a [Google Cloud Console](https://console.cloud.google.com/).  
> 2. Crea un proyecto y habilita la API de **Google Sheets** y **Google Drive**.  
> 3. Ve a **Pantalla de consentimiento de OAuth** y configúrala como aplicación externa/interna.  
> 4. En **Credenciales**, crea un **ID de cliente OAuth 2.0** de tipo *Aplicación de escritorio*.  
> 5. Descarga el archivo JSON generado, renómbralo exactamente a credentials.json y colócalo en la raíz de este proyecto.

**🚨 Seguridad:** El archivo .gitignore incluido evita por defecto que credentials.json y token.json se suban a GitHub. **Nunca compartas tus credenciales.**

### **2\. Archivo de Reglas (rules.yaml)**

Define las reglas según el prefijo del nombre de los archivos CSV que descargas. Edita el archivo rules.yaml:  
`# Si el archivo empieza con "Enrolled" (ej: Enrolled_2026.csv)`  
`Enrolled:`  
  `spreadsheet_id: "TU_GOOGLE_SHEET_ID_AQUÍ"`  
  `nombre_hoja: "Enrolled"`  
  `poner_nota_a1: true`

`# Si el archivo empieza con "Participantes" (ej: Participantes_Reporte.csv)`  
`Participantes:`  
  `spreadsheet_id: "TU_OTRO_GOOGLE_SHEET_ID_AQUÍ"`  
  `nombre_hoja: "Participantes"`  
  `poner_nota_a1: true`

## **🚀 Uso**

Abre tu terminal en la carpeta del proyecto y ejecuta:  
`python csv_to_sheets_monitor.py`

### **¿Cómo probarlo?**

> 1. Deja corriendo la consola (verás el mensaje 🚀 Servicio de monitoreo activado con éxito).  
> 2. Descarga un archivo CSV en tu carpeta de **Descargas** de Windows/Mac cuya palabra inicial coincida con las claves configuradas en rules.yaml (ej. Enrolled\_Lunes.csv).  
> 3. Verás la actividad en la terminal y recibirás una notificación de escritorio confirmando la actualización.

## **🔄 Ejecución Automática al Iniciar Windows (Opcional)**

Si deseas que el monitor se ejecute automáticamente sin abrir una ventana de consola cada vez que enciendes la PC:

> 1. Presiona Win \+ R, escribe shell:startup y presiona **Enter**.  
> 2. Haz un clic derecho en la carpeta que se abre \> **Nuevo** \> **Acceso directo**.  
> 3. Para ejecutarlo en segundo plano sin ventana de consola, usa pythonw.exe:  
>    `C:\Ruta\A\Tu\pythonw.exe "C:\Ruta\Al\Proyecto\csv_to_sheets_monitor.py"`  
> 4. Guarda el acceso directo. ¡Listo\!

## **📄 Licencia**

Este proyecto está bajo la Licencia **MIT**. Consulta el archivo LICENSE para más detalles.