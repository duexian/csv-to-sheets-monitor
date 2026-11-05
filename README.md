# 🔄 CSV to Sheets Monitor

**CSV to Sheets Monitor** es una herramienta en Python que corre en segundo plano y vigila tu carpeta de **Descargas**. Cuando descargas un `.csv` cuyo nombre empieza por un prefijo configurado (por ejemplo `Enrolled…` o `Participantes…`), sube su contenido a la pestaña correspondiente de **Google Sheets** y te avisa con una notificación de Windows.

## ✨ Características

* **Monitoreo en tiempo real** con `watchdog`. Detecta descargas de Chrome, Edge y Firefox, y espera a que el archivo termine de escribirse antes de leerlo.
* **Reglas en `rules.yaml`**: agregas o cambias reglas sin tocar el código. Se recargan solas al guardar el archivo; si tienen un error, se siguen usando las anteriores y recibes un aviso.
* **Protección de fórmulas**: solo se tocan las columnas que trae el CSV (A..última). Las columnas a la derecha, con tus fórmulas, no se tocan.
* **La hoja nunca queda vacía**: los datos nuevos se escriben encima de los anteriores y después se limpian las filas que sobran. Si algo falla a mitad de camino, no pierdes la información.
* **Todo como texto**: se usa `RAW`, así que los IDs largos, los ceros a la izquierda y las fechas no se alteran.
* **Nota de auditoría en A1** con fecha, archivo y tamaño de la última carga.
* **Notificaciones nativas de Windows 10/11** con botones **Abrir hoja** y **Ver registro**, mensajes de error claros ("Sin permiso sobre la hoja", "Pestaña no encontrada", "Sin conexión"…) y sin avisos repetidos.
* **Icono en la bandeja del sistema** (opcional) que cambia de color según el estado. Desde su menú puedes abrir la última hoja, reintentar, sincronizar un archivo a mano, pausar, editar las reglas, ver el registro o salir.
* **Robusto**:
  * Reintenta automáticamente ante cortes de internet, errores 5xx o 429 de Google y archivos bloqueados.
  * Si la sesión de Google caduca, vuelve a abrir el inicio de sesión.
  * No procesa dos veces el mismo archivo.
  * No deja ejecutar dos copias a la vez.
* **Registro persistente** en `logs/monitor.log`, también al ejecutarlo con `pythonw`.

## 📂 Estructura

```
enrolled/
├── csv-to-sheets-monitor.py   # Script principal
├── rules.yaml                 # Reglas prefijo → libro/pestaña
├── requirements.txt
├── assets/icon.png            # Icono de notificaciones y bandeja
├── tests/test_monitor.py      # Pruebas automatizadas (pytest)
├── logs/                      # Registro (se crea solo, ignorado por git)
├── credentials.json           # (tuyo, ignorado por git)
└── token.json                 # (se genera solo, ignorado por git)
```

## 🛠️ Instalación

```bash
pip install -r requirements.txt
```

Instálalo con el **mismo Python** que use el acceso directo de inicio (por ejemplo, el `pythonw.exe` global). Sin `pystray`, el monitor funciona igual, solo que sin icono en la bandeja.

### Credenciales de Google

1. En [Google Cloud Console](https://console.cloud.google.com/) habilita **Google Sheets API**.
2. Crea un **ID de cliente OAuth** de tipo *Aplicación de escritorio* y guárdalo como `credentials.json` junto al script.
3. La primera vez se abrirá el navegador para iniciar sesión, y se crea `token.json`.

> 💡 Si tu pantalla de consentimiento OAuth está en modo **"Prueba"**, Google hace caducar la sesión cada 7 días. El monitor te pedirá iniciar sesión de nuevo de forma automática. Para evitarlo, publica la app ("En producción").

## ⚙️ rules.yaml

```yaml
Enrolled:                       # archivos que empiezan por "Enrolled" (sin distinguir mayúsculas)
  spreadsheet_id: "ID_O_URL_DEL_LIBRO"   # se puede pegar la URL completa
  nombre_hoja: "Enrolled"
  poner_nota_a1: true           # opcional (true por defecto)
  incluir_encabezados: false    # opcional (true por defecto), ver abajo

Participantes:
  spreadsheet_id: "OTRO_ID"
  nombre_hoja: "Participantes2026"
```

| Campo | Obligatorio | Qué hace |
|---|---|---|
| `spreadsheet_id` | Sí | ID del libro (o su URL completa) |
| `nombre_hoja` | Sí | Pestaña destino (no distingue mayúsculas) |
| `poner_nota_a1` | No (`true`) | Deja en A1 una nota con fecha, archivo y filas cargadas |
| `incluir_encabezados` | No (`true`) | `true`: sube el CSV completo desde la fila 1. `false`: omite la primera fila del CSV y escribe los datos **desde la fila 2**, sin tocar el encabezado que ya tiene la hoja (útil si siempre es el mismo o le diste formato) |

Si dos prefijos coinciden, gana el más largo (el más específico).

## 🚀 Uso

| Comando | Qué hace |
|---|---|
| `python csv-to-sheets-monitor.py` | Inicia el monitor (con icono en la bandeja) |
| `python csv-to-sheets-monitor.py --verificar` | Revisa las reglas, la sesión de Google y que existan las pestañas |
| `python csv-to-sheets-monitor.py --archivo "ruta.csv"` | Sube un archivo concreto y termina |
| `python csv-to-sheets-monitor.py --archivo "ruta.csv" --simulacion` | Muestra lo que haría, **sin escribir** |
| `--carpeta RUTA` · `--sin-bandeja` · `--debug` | Otra carpeta, sin icono, registro detallado |

### Inicio automático con Windows

Crea un `.vbs` en `shell:startup`:

```vb
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "pythonw.exe ""C:\ruta\enrolled\csv-to-sheets-monitor.py""", 0, False
```

Al arrancar aparecerá la notificación **"Auto-Sync activo"**.

## 🧪 Pruebas

```bash
python -m pytest -q
```

Las pruebas no usan internet ni modifican tus hojas: la API de Google se simula. Cubren reglas, lectura de CSV (BOM, emojis, cp1252, `;`, celdas con saltos de línea), subida por lotes, ampliación de la cuadrícula, reintentos, sesión caducada, notificaciones, instancia única y una descarga simulada de Chrome de principio a fin.

## 🛡️ Qué archivos se suben (y cuáles nunca)

* Solo se suben **descargas nuevas**: archivos creados o modificados después de arrancar el monitor. Los CSV que ya estaban en Descargas no se suben solos, aunque Windows (antivirus, indexador, vista previa) los lea.
* Un mismo contenido **nunca se sube dos veces**: se compara una huella del archivo, no la fecha.
* **Nunca se sube un CSV más antiguo** que otro de la misma regla presente en la carpeta. Como cada CSV reemplaza la hoja entera, uno viejo no puede pisar datos nuevos.
* Para subir a propósito un archivo antiguo: usa `--archivo RUTA` o, en la bandeja, **Sincronizar un archivo…**.

## 📦 Mover a "procesados"

Con `mover_a_procesados: true` (por defecto en cada regla de `rules.yaml`), el CSV se mueve a `Descargas\procesados\Sheets` después de **subirlo con éxito**. Así Descargas queda limpia.

* Si la subida falla, queda pendiente, es una simulación o el CSV no tiene datos, el archivo se queda en Descargas.
* Si ya existe uno con el mismo nombre en `procesados`, se guarda como `Nombre (2).csv`. Nunca se pisa.
* Para desactivarlo en una regla: `mover_a_procesados: false`.
* `logs/ultimas_subidas.json` guarda la fecha del último CSV subido por regla. Gracias a eso, nunca se sube uno más viejo, aunque el más nuevo ya no esté en Descargas.

## 🌐 Red bloqueada (Fortinet) y archivos pendientes

El firewall de la red (**Fortinet / FortiGuard Secure DNS**) bloquea a ratos `oauth2.googleapis.com`. Es el servicio que renueva la sesión de Google cada hora. El error que aparece es `CERTIFICATE_VERIFY_FAILED`. El monitor:

* **renueva la sesión por adelantado**, en cuanto la red lo permite, para no depender de ese servicio justo al llegar un CSV;
* si aun así falla, reintenta varias veces;
* después deja el archivo en `logs/pendientes.json` (uno por regla, el más reciente) y lo reintenta solo cada 3 minutos, también tras reiniciar;
* avisa una sola vez, y otra cuando por fin se sube;
* nunca desactiva la verificación SSL.

La solución definitiva es pedir a TI que agregue `oauth2.googleapis.com` a la lista de permitidos. `--verificar` muestra si está bloqueado en ese momento.

## 🩺 Solución de problemas

1. Abre `logs/monitor.log` (o usa **Ver registro** en la notificación o en la bandeja).
2. Ejecuta `python csv-to-sheets-monitor.py --verificar`.
3. Si la sesión quedó mal, borra `token.json` y vuelve a iniciar el monitor.

## 📄 Licencia

MIT. Consulta el archivo `LICENSE`.
