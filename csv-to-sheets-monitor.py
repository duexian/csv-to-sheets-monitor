from dataclasses import dataclass
from datetime import datetime
import logging
from pathlib import Path
import csv
import time
from typing import Dict, Optional, Protocol
import yaml

# --- Librerías de Google API ---
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import Resource, build

# --- Librerías de Monitoreo y Notificaciones ---
from plyer import notification
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer

# Configuración del Logger de Producción
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler()],
)
logger = logging.getLogger("SheetsAutoSync")


# ==============================================================================
# --- ESTRUCTURAS DE CONFIGURACIÓN ---
# ==============================================================================
@dataclass(frozen=True)
class FileRule:
    """Regla de procesamiento para un tipo específico de archivo CSV."""
    spreadsheet_id: str
    nombre_hoja: str
    poner_nota_a1: bool = True


@dataclass(frozen=True)
class AppConfig:
    """Configuración inmutable de la aplicación.
    
    Diseñada para ser portable entre diferentes usuarios y computadoras.
    """

    script_dir: Path = Path(__file__).resolve().parent
    
    # RUTA DINÁMICA: Detecta la carpeta Downloads del usuario que ejecute el script
    carpeta_origen: Path = Path.home() / "Downloads"
    extension_requerida: str = ".csv"

    # Permisos combinados (Google Sheets + Google Drive)
    scopes: tuple = (
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.file",
    )

    @property
    def ruta_rules(self) -> Path:
        return self.script_dir / "rules.yaml"

    @property
    def ruta_credentials(self) -> Path:
        return self.script_dir / "credentials.json"

    @property
    def ruta_token(self) -> Path:
        return self.script_dir / "token.json"

    @property
    def reglas_archivos(self) -> Dict[str, FileRule]:
        """Carga dinámicamente las reglas desde el archivo rules.yaml."""
        if not self.ruta_rules.exists():
            raise FileNotFoundError(
                f"❌ No se encontró el archivo de reglas: '{self.ruta_rules.name}' "
                f"en {self.script_dir}"
            )

        with open(self.ruta_rules, "r", encoding="utf-8") as file:
            data = yaml.safe_load(file)

        if not data or not isinstance(data, dict):
            raise ValueError(f"El archivo '{self.ruta_rules.name}' está vacío o tiene un formato inválido.")

        return {
            prefijo: FileRule(**config_dict)
            for prefijo, config_dict in data.items()
        }


# ==============================================================================
# --- SERVICIO DE NOTIFICACIONES ---
# ==============================================================================
class NotifierProtocol(Protocol):
    def notify(self, title: str, message: str, timeout: int = 5) -> None: ...


class DesktopNotifier:
    """Implementación de notificaciones de escritorio seguras."""

    def __init__(self, app_name: str = "Sheets Auto-Sync"):
        self.app_name = app_name

    def notify(self, title: str, message: str, timeout: int = 5) -> None:
        safe_message = message[:200] if len(message) > 200 else message
        safe_title = title[:60] if len(title) > 60 else title

        try:
            notification.notify(
                title=safe_title,
                message=safe_message,
                app_name=self.app_name,
                timeout=timeout,
            )
        except Exception as err:
            logger.warning(f"No se pudo desplegar la notificación: {err}")


# ==============================================================================
# --- SERVICIO GOOGLE SHEETS ---
# ==============================================================================
class GoogleSheetsService:
    """Clase encargada de la autenticación y actualización en Google Sheets."""

    def __init__(self, config: AppConfig):
        self.config = config
        self._service: Optional[Resource] = None

    def get_service(self) -> Resource:
        """Obtiene o inicializa el cliente del servicio de Google Sheets."""
        if self._service:
            return self._service

        creds: Optional[Credentials] = None

        if self.config.ruta_token.exists():
            creds = Credentials.from_authorized_user_file(
                str(self.config.ruta_token), list(self.config.scopes)
            )

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                if not self.config.ruta_credentials.exists():
                    raise FileNotFoundError(
                        f"No se encontró el archivo '{self.config.ruta_credentials.name}' "
                        f"en {self.config.script_dir}"
                    )

                flow = InstalledAppFlow.from_client_secrets_file(
                    str(self.config.ruta_credentials), list(self.config.scopes)
                )
                creds = flow.run_local_server(port=0)

            with open(self.config.ruta_token, "w") as token_file:
                token_file.write(creds.to_json())

        self._service = build("sheets", "v4", credentials=creds)
        return self._service

    def _get_sheet_id_by_name(self, service: Resource, spreadsheet_id: str, sheet_name: str) -> int:
        """Obtiene el ID numérico (sheetId) de una pestaña específica."""
        spreadsheet = (
            service.spreadsheets()
            .get(spreadsheetId=spreadsheet_id)
            .execute()
        )
        for sheet in spreadsheet.get("sheets", []):
            if sheet["properties"]["title"].strip().lower() == sheet_name.strip().lower():
                return sheet["properties"]["sheetId"]
        raise ValueError(f"No se encontró la pestaña '{sheet_name}' en el Spreadsheet {spreadsheet_id}.")

    def _add_update_note_to_a1(self, service: Resource, spreadsheet_id: str, sheet_name: str, note_text: str) -> None:
        """Añade/actualiza una nota en la celda A1 de la pestaña indicada."""
        sheet_id = self._get_sheet_id_by_name(service, spreadsheet_id, sheet_name)

        requests = [
            {
                "updateCells": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 0,
                        "endRowIndex": 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": 1,
                    },
                    "rows": [{"values": [{"note": note_text}]}],
                    "fields": "note",
                }
            }
        ]

        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id, body={"requests": requests}
        ).execute()

    def update_sheet_with_csv(self, csv_path: Path, rule: FileRule) -> str:
        """Lee el CSV e inserta los datos en el Spreadsheet y pestaña indicados por la regla."""
        service = self.get_service()

        # 1. Obtener pestañas reales para validar coincidencia exacta
        spreadsheet = service.spreadsheets().get(spreadsheetId=rule.spreadsheet_id).execute()
        pestanas_existentes = [s["properties"]["title"] for s in spreadsheet.get("sheets", [])]

        nombre_hoja_real = None
        for p in pestanas_existentes:
            if p.strip().lower() == rule.nombre_hoja.strip().lower():
                nombre_hoja_real = p
                break

        if not nombre_hoja_real:
            mensaje_error = (
                f"❌ La pestaña '{rule.nombre_hoja}' no existe en el libro {rule.spreadsheet_id}. "
                f"Pestañas disponibles: {pestanas_existentes}"
            )
            logger.error(mensaje_error)
            raise ValueError(mensaje_error)

        # 2. Leer los datos del archivo CSV local
        with open(csv_path, mode="r", encoding="utf-8-sig") as file:
            reader = csv.reader(file)
            raw_values = list(reader)

        if not raw_values:
            logger.warning(f"El archivo {csv_path.name} está vacío.")
            return "El CSV está vacío, no se realizaron cambios."

        # 3. Formatear datos: apóstrofe al inicio para forzar texto
        formatted_values = [
            [f"'{cell}" if cell != "" else "" for cell in row]
            for row in raw_values
        ]

        num_columnas = max(len(row) for row in formatted_values)

        def col_to_letter(col_idx: int) -> str:
            result = ""
            while col_idx > 0:
                col_idx, remainder = divmod(col_idx - 1, 26)
                result = chr(65 + remainder) + result
            return result

        ultima_columna = col_to_letter(num_columnas)

        # 4. Limpiar únicamente las columnas usadas por el CSV (respetando fórmulas vecinas)
        rango_csv = f"'{nombre_hoja_real}'!A1:{ultima_columna}"
        service.spreadsheets().values().clear(
            spreadsheetId=rule.spreadsheet_id,
            range=rango_csv
        ).execute()

        # 5. Escribir los datos
        range_to_update = f"'{nombre_hoja_real}'!A1"
        body = {"values": formatted_values}

        service.spreadsheets().values().update(
            spreadsheetId=rule.spreadsheet_id,
            range=range_to_update,
            valueInputOption="USER_ENTERED",
            body=body
        ).execute()

        # 6. Colocar la nota en A1 si la regla lo requiere
        if rule.poner_nota_a1:
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            note_content = f"Última actualización: {now_str}\nArchivo: {csv_path.name}"
            self._add_update_note_to_a1(service, rule.spreadsheet_id, nombre_hoja_real, note_content)

        logger.info(
            f"✅ Se actualizaron las columnas A..{ultima_columna} en '{nombre_hoja_real}' (ID: {rule.spreadsheet_id})"
        )
        return f"Datos cargados en '{nombre_hoja_real}' (cols A..{ultima_columna})"


# ==============================================================================
# --- MANEJADOR DE EVENTOS DE ARCHIVO ---
# ==============================================================================
class CSVDownloadHandler(FileSystemEventHandler):
    """Manejador dinámico de eventos para el sistema de archivos."""

    def __init__(
        self,
        config: AppConfig,
        sheets_service: GoogleSheetsService,
        notifier: NotifierProtocol,
    ):
        self.config = config
        self.sheets_service = sheets_service
        self.notifier = notifier

    def _get_matching_rule(self, file_path: Path) -> Optional[FileRule]:
        """Devuelve la regla si el archivo cumple las condiciones."""
        if file_path.is_dir() or not file_path.exists():
            return None

        # Ignorar archivos temporales o descargas incompletas del navegador
        if file_path.suffix.lower() in [".crdownload", ".part", ".tmp"]:
            return None

        if file_path.suffix.lower() != self.config.extension_requerida:
            return None

        reglas = self.config.reglas_archivos
        for prefijo, rule in reglas.items():
            if file_path.name.startswith(prefijo):
                return rule

        return None

    def process_file(self, raw_path: str) -> None:
        file_path = Path(raw_path)
        rule = self._get_matching_rule(file_path)

        if not rule:
            return

        try:
            # Esperar 2.5s para asegurar que el navegador soltó el archivo completamente
            time.sleep(2.5)
            if not file_path.exists():
                return

            logger.info(f"📄 Archivo detectado en Descargas: {file_path.name}")

            # Procesar directamente el archivo desde la carpeta de descargas
            result_msg = self.sheets_service.update_sheet_with_csv(file_path, rule)

            self.notifier.notify(
                title="Google Sheets Auto-Sync", message=result_msg
            )

        except PermissionError:
            time.sleep(1.5)
            if file_path.exists():
                result_msg = self.sheets_service.update_sheet_with_csv(file_path, rule)
                self.notifier.notify(
                    title="Google Sheets Auto-Sync", message=result_msg
                )
        except Exception as err:
            logger.error(
                f"❌ Fallo al procesar el archivo {file_path.name}: {err}",
                exc_info=True,
            )
            self.notifier.notify(
                title="❌ Error en Auto-Sync",
                message=f"No se pudo procesar '{file_path.name}'.",
            )

    def on_created(self, event) -> None:
        self.process_file(event.src_path)

    def on_moved(self, event) -> None:
        self.process_file(event.dest_path)


# ==============================================================================
# --- PUNTO DE ENTRADA ---
# ==============================================================================
def main() -> None:
    config = AppConfig()
    sheets_service = GoogleSheetsService(config)
    notifier = DesktopNotifier()

    event_handler = CSVDownloadHandler(
        config=config, sheets_service=sheets_service, notifier=notifier
    )

    observer = Observer()
    observer.schedule(
        event_handler, str(config.carpeta_origen), recursive=False
    )

    logger.info("🚀 Servicio de monitoreo activado con éxito.")
    logger.info(f"📁 Monitoreando origen: {config.carpeta_origen}")
    logger.info("📊 Reglas cargadas desde YAML:")
    for prefijo, regla in config.reglas_archivos.items():
        logger.info(
            f"   • Prefijo: '{prefijo}' ➡️ ID: {regla.spreadsheet_id} | Hoja: '{regla.nombre_hoja}'"
        )

    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        logger.info("🛑 Deteniendo el servicio de monitoreo...")
        observer.stop()

    observer.join()


if __name__ == "__main__":
    main()