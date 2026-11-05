"""
CSV to Sheets Monitor
=====================

Vigila la carpeta de Descargas y, cuando aparece un CSV cuyo nombre empieza por
alguno de los prefijos definidos en ``rules.yaml``, sube su contenido a la
pestaña de Google Sheets correspondiente y muestra una notificación de Windows.

Uso rápido::

    python csv-to-sheets-monitor.py                 # monitor (con icono en la bandeja)
    python csv-to-sheets-monitor.py --verificar     # valida reglas, sesión y pestañas
    python csv-to-sheets-monitor.py --archivo X.csv # sincroniza un archivo y termina
    python csv-to-sheets-monitor.py --archivo X.csv --simulacion   # sin escribir nada

Registro de actividad: ``logs/monitor.log`` (también cuando se ejecuta con pythonw).
"""
from __future__ import annotations

import argparse
import base64
import csv
import dataclasses
import hashlib
import io
import json
import logging
import os
import queue
import re
import shutil
import socket
import ssl
import subprocess
import sys
import threading
import time
import traceback
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Protocol, Sequence, Tuple
from xml.sax.saxutils import escape as xml_escape

import yaml

# --- Librerías de Google API ---
from google.auth.exceptions import RefreshError, TransportError
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# --- Monitoreo de archivos ---
from watchdog.events import FileSystemEvent, FileSystemEventHandler
from watchdog.observers import Observer

APP_NAME = "Sheets Auto-Sync"
SCRIPT_DIR = Path(__file__).resolve().parent
LOG_DIR = SCRIPT_DIR / "logs"
LOG_FILE = LOG_DIR / "monitor.log"
ICON_FILE = SCRIPT_DIR / "assets" / "icon.png"

TEMP_SUFFIXES = {".crdownload", ".part", ".tmp", ".download", ".partial"}

logger = logging.getLogger("SheetsAutoSync")


# ==============================================================================
# --- LOGGING ---
# ==============================================================================
def setup_logging(log_file: Optional[Path] = LOG_FILE, level: int = logging.INFO) -> None:
    """Configura el registro en archivo rotativo (UTF-8) y en consola si existe.

    - Con ``pythonw.exe`` no hay consola (``sys.stderr`` es ``None``): antes los
      errores se perdían. Ahora siempre quedan en ``logs/monitor.log``.
    - En consolas Windows con cp1252 los emojis provocaban ``UnicodeEncodeError``
      ("--- Logging error ---"). Se reconfigura la salida para no fallar nunca.
    """
    formatter = logging.Formatter(
        "%(asctime)s [%(levelname)s] %(message)s", "%Y-%m-%d %H:%M:%S"
    )
    logger.setLevel(level)
    logger.propagate = False
    for handler in list(logger.handlers):
        logger.removeHandler(handler)
        handler.close()

    if log_file is not None:
        try:
            log_file.parent.mkdir(parents=True, exist_ok=True)
            file_handler = RotatingFileHandler(
                log_file, maxBytes=1_000_000, backupCount=5, encoding="utf-8"
            )
            file_handler.setFormatter(formatter)
            logger.addHandler(file_handler)
        except OSError as err:  # pragma: no cover - disco lleno / sin permisos
            print(f"No se pudo crear el archivo de log: {err}", file=sys.stderr)

    for out in (sys.stdout, sys.stderr):
        if out is not None:
            try:
                out.reconfigure(errors="backslashreplace")  # type: ignore[attr-defined]
            except Exception:
                pass

    stream = sys.stderr
    if stream is not None:
        console = logging.StreamHandler(stream)
        console.setFormatter(formatter)
        console.addFilter(lambda record: not getattr(record, "solo_archivo", False))
        logger.addHandler(console)


def log_details(err: BaseException, level: int = logging.INFO) -> None:
    """Traceback completo solo en el archivo de registro (la consola queda limpia)."""
    logger.log(level, "Detalle técnico:\n%s", format_error_details(err), extra={"solo_archivo": True})


# ==============================================================================
# --- ERRORES Y MENSAJES AMIGABLES ---
# ==============================================================================
class ConfigError(Exception):
    """Error en rules.yaml, credentials.json u otra configuración local."""


class AuthError(Exception):
    """La sesión de Google expiró o fue revocada y hace falta volver a iniciarla."""


class SheetNotFoundError(Exception):
    """La pestaña indicada en la regla no existe en el libro."""


@dataclass(frozen=True)
class ErrorInfo:
    titulo: str
    mensaje: str
    transitorio: bool = False


def _error_chain(err: BaseException):
    seen = set()
    current: Optional[BaseException] = err
    while current is not None and id(current) not in seen:
        seen.add(id(current))
        yield current
        current = current.__cause__ or current.__context__


def is_network_block(err: BaseException) -> bool:
    """True si la red interceptó la conexión a Google.

    El filtro del firewall (Fortinet / FortiGuard Secure DNS) responde a ratos con
    su "página de bloqueo" para oauth2.googleapis.com (el servicio que renueva la
    sesión cada hora), con un certificado autofirmado → CERTIFICATE_VERIFY_FAILED.
    Nunca se desactiva la verificación SSL: se reintenta hasta que la red deja pasar.
    """
    for e in _error_chain(err):
        if isinstance(e, ssl.SSLCertVerificationError):
            return True
        text = str(e)
        if "CERTIFICATE_VERIFY_FAILED" in text or "Fortiguard" in text or "Hostname mismatch" in text:
            return True
    return False


def describe_error(err: BaseException) -> ErrorInfo:
    """Traduce una excepción a un mensaje breve y accionable para el usuario."""
    detalle = str(err).strip() or type(err).__name__

    if is_network_block(err):
        return ErrorInfo(
            "Red bloqueó la conexión a Google",
            "El firewall de la red (Fortinet) interceptó la conexión. Se reintentará automáticamente.",
            True,
        )
    if isinstance(err, HttpError):
        status = getattr(err.resp, "status", None)
        try:
            status = int(status) if status is not None else None
        except (TypeError, ValueError):
            status = None
        reason = err._get_reason() if hasattr(err, "_get_reason") else detalle
        if status in (401,):
            return ErrorInfo("Sesión de Google inválida", "Vuelve a iniciar sesión (ejecuta con --verificar).")
        if status == 403:
            return ErrorInfo(
                "Sin permiso sobre la hoja",
                "Tu cuenta no puede editar ese libro. Revisa que esté compartido contigo como editor.",
            )
        if status == 404:
            return ErrorInfo(
                "Libro no encontrado",
                "El spreadsheet_id de rules.yaml no existe o no tienes acceso.",
            )
        if status == 429:
            return ErrorInfo("Límite de Google alcanzado", "Demasiadas peticiones. Se reintentará en unos segundos.", True)
        if status is not None and status >= 500:
            return ErrorInfo("Google Sheets no responde", f"Error temporal del servidor ({status}).", True)
        return ErrorInfo(f"Error de Google Sheets ({status})", str(reason)[:180])

    if isinstance(err, AuthError):
        return ErrorInfo("Inicia sesión de nuevo", detalle)
    if isinstance(err, SheetNotFoundError):
        return ErrorInfo("Pestaña no encontrada", detalle)
    if isinstance(err, ConfigError):
        return ErrorInfo("Revisa la configuración", detalle)
    if isinstance(err, PermissionError):
        return ErrorInfo("Archivo en uso", "Otro programa tiene el CSV abierto. Se reintentará.", True)
    if isinstance(err, (TransportError, ConnectionError, TimeoutError, socket.timeout, socket.gaierror, OSError)) and not isinstance(
        err, (FileNotFoundError, IsADirectoryError)
    ):
        return ErrorInfo("Sin conexión", "No se pudo contactar a Google. Revisa tu internet; se reintentará.", True)
    if isinstance(err, (UnicodeDecodeError, csv.Error)):
        return ErrorInfo("CSV ilegible", f"No se pudo leer el archivo: {detalle[:150]}")
    if isinstance(err, FileNotFoundError):
        return ErrorInfo("Archivo no encontrado", "El CSV se movió o eliminó antes de procesarlo.")
    return ErrorInfo(f"Error inesperado ({type(err).__name__})", detalle[:180])


def format_error_details(err: BaseException) -> str:
    """Tipo, detalle y traceback de la excepción para el log."""
    tb = "".join(traceback.format_exception(type(err), err, err.__traceback__)).strip()
    return f"Tipo: {type(err).__name__}\nDetalles: {str(err).strip() or 'Sin detalles.'}\n{tb}"


# ==============================================================================
# --- CONFIGURACIÓN ---
# ==============================================================================
_SPREADSHEET_URL_RE = re.compile(r"/spreadsheets/d/([a-zA-Z0-9_-]+)")
_RULE_REQUIRED = {"spreadsheet_id", "nombre_hoja"}
_RULE_OPTIONAL = {"poner_nota_a1", "incluir_encabezados", "mover_a_procesados"}


@dataclass(frozen=True)
class FileRule:
    """Regla de procesamiento para un tipo específico de archivo CSV."""

    prefijo: str
    spreadsheet_id: str
    nombre_hoja: str
    poner_nota_a1: bool = True
    # False: se omite la primera fila del CSV y los datos se escriben desde la
    # fila 2, dejando intacto el encabezado que ya tiene la hoja.
    incluir_encabezados: bool = True
    # True: tras subirlo con éxito, el CSV se mueve de Descargas a procesados/Sheets.
    mover_a_procesados: bool = True

    @property
    def url(self) -> str:
        return f"https://docs.google.com/spreadsheets/d/{self.spreadsheet_id}/edit"


def parse_rules(data: Any, origen: str = "rules.yaml") -> List[FileRule]:
    """Valida el contenido de rules.yaml y devuelve las reglas.

    Las reglas se ordenan por prefijo más largo primero, para que el prefijo más
    específico gane si dos coinciden.
    """
    if not data or not isinstance(data, dict):
        raise ConfigError(f"'{origen}' está vacío o no tiene el formato 'Prefijo: {{...}}'.")

    reglas: List[FileRule] = []
    for prefijo, conf in data.items():
        prefijo = str(prefijo)
        if not prefijo.strip():
            raise ConfigError(f"'{origen}': hay una regla con prefijo vacío.")
        if not isinstance(conf, dict):
            raise ConfigError(f"'{origen}': la regla '{prefijo}' debe tener campos (spreadsheet_id, nombre_hoja...).")

        faltan = _RULE_REQUIRED - conf.keys()
        if faltan:
            raise ConfigError(f"'{origen}': a la regla '{prefijo}' le falta: {', '.join(sorted(faltan))}.")
        desconocidos = conf.keys() - _RULE_REQUIRED - _RULE_OPTIONAL
        if desconocidos:
            raise ConfigError(
                f"'{origen}': la regla '{prefijo}' tiene campos desconocidos: {', '.join(sorted(desconocidos))}. "
                f"Válidos: {', '.join(sorted(_RULE_REQUIRED | _RULE_OPTIONAL))}."
            )

        spreadsheet_id = str(conf["spreadsheet_id"] or "").strip()
        match = _SPREADSHEET_URL_RE.search(spreadsheet_id)
        if match:  # Permite pegar la URL completa del libro
            spreadsheet_id = match.group(1)
        if not spreadsheet_id:
            raise ConfigError(f"'{origen}': la regla '{prefijo}' tiene spreadsheet_id vacío.")

        nombre_hoja = str(conf["nombre_hoja"] or "").strip()
        if not nombre_hoja:
            raise ConfigError(f"'{origen}': la regla '{prefijo}' tiene nombre_hoja vacío.")

        booleanos = {}
        for campo in ("poner_nota_a1", "incluir_encabezados", "mover_a_procesados"):
            valor = conf.get(campo, True)
            if not isinstance(valor, bool):
                raise ConfigError(f"'{origen}': en la regla '{prefijo}', {campo} debe ser true o false.")
            booleanos[campo] = valor

        reglas.append(FileRule(prefijo, spreadsheet_id, nombre_hoja, **booleanos))

    reglas.sort(key=lambda r: len(r.prefijo), reverse=True)
    return reglas


class RulesStore:
    """Carga rules.yaml y lo recarga automáticamente cuando cambia.

    Si el archivo editado tiene un error, se conservan las últimas reglas válidas
    y se avisa una sola vez (en vez de fallar con cada descarga).
    """

    def __init__(self, path: Path, on_error: Optional[Callable[[Exception], None]] = None):
        self.path = path
        self.on_error = on_error
        self._rules: Optional[List[FileRule]] = None
        self._mtime: Optional[int] = None
        self._lock = threading.Lock()

    def get(self) -> List[FileRule]:
        with self._lock:
            try:
                mtime = self.path.stat().st_mtime_ns
            except FileNotFoundError:
                if self._rules is not None:
                    return self._rules
                raise ConfigError(f"No se encontró el archivo de reglas: {self.path}")

            if self._rules is not None and mtime == self._mtime:
                return self._rules

            try:
                with open(self.path, "r", encoding="utf-8") as fh:
                    data = yaml.safe_load(fh)
                rules = parse_rules(data, self.path.name)
            except (yaml.YAMLError, ConfigError, OSError, UnicodeDecodeError) as err:
                self._mtime = mtime  # no reintentar hasta que se vuelva a editar
                if not isinstance(err, ConfigError):
                    err = ConfigError(f"No se pudo leer '{self.path.name}': {err}")
                if self._rules is None:
                    raise err
                logger.error("⚠️ %s. Se mantienen las reglas anteriores.", err)
                if self.on_error:
                    self.on_error(err)
                return self._rules

            if self._rules is not None:
                logger.info("🔄 Reglas recargadas (%d).", len(rules))
            self._rules, self._mtime = rules, mtime
            return rules

    def match(self, filename: str) -> Optional[FileRule]:
        """Regla cuyo prefijo coincide con el nombre (sin distinguir mayúsculas)."""
        name = filename.lower()
        for rule in self.get():
            if name.startswith(rule.prefijo.lower()):
                return rule
        return None


def get_downloads_folder() -> Path:
    """Carpeta real de Descargas (respeta redirecciones de OneDrive/Políticas)."""
    if sys.platform == "win32":
        try:
            import ctypes
            from ctypes import wintypes
            import uuid

            folder_id = uuid.UUID("{374DE290-123F-4565-9164-39C4925E467B}")  # FOLDERID_Downloads
            guid = (ctypes.c_byte * 16).from_buffer_copy(folder_id.bytes_le)
            path_ptr = ctypes.c_wchar_p()
            shell32 = ctypes.windll.shell32  # type: ignore[attr-defined]
            if shell32.SHGetKnownFolderPath(ctypes.byref(guid), 0, None, ctypes.byref(path_ptr)) == 0:
                result = Path(path_ptr.value)
                ctypes.windll.ole32.CoTaskMemFree(path_ptr)  # type: ignore[attr-defined]
                if result.is_dir():
                    return result
        except Exception:
            pass
    return Path.home() / "Downloads"


@dataclass(frozen=True)
class AppConfig:
    """Configuración de la aplicación (portable entre usuarios y equipos)."""

    script_dir: Path = SCRIPT_DIR
    carpeta_origen: Path = field(default_factory=get_downloads_folder)
    extension_requerida: str = ".csv"
    scopes: Tuple[str, ...] = (
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.file",
    )
    # Tiempo que el tamaño del archivo debe permanecer estable para darlo por terminado.
    segundos_estabilidad: float = 1.5
    # Tiempo máximo esperando a que el navegador termine de escribir el archivo.
    segundos_espera_maxima: float = 120.0
    # Ignora eventos repetidos del mismo archivo (mismo tamaño y fecha) en esta ventana.
    segundos_ventana_duplicados: float = 60.0
    # Reintentos de cada petición HTTP ante 5xx/429 (los gestiona la librería de Google).
    reintentos: int = 3
    # Esperas entre reintentos inmediatos ante errores temporales (red bloqueada,
    # sin internet, archivo en uso). Si se agotan, el archivo pasa a "pendientes".
    esperas_reintento: Tuple[float, ...] = (10.0, 30.0, 60.0)
    minutos_reintento_pendientes: float = 3.0
    horas_max_pendiente: float = 24.0
    # Renovar la sesión cuando falten menos de estos minutos para que caduque.
    minutos_margen_token: float = 15.0
    filas_por_lote: int = 5000
    # Destino de los CSV ya subidos (relativo a la carpeta vigilada o ruta absoluta).
    # Se activa/desactiva por regla con 'mover_a_procesados' en rules.yaml.
    carpeta_procesados: str = "procesados/Sheets"

    @property
    def ruta_procesados(self) -> Path:
        p = Path(os.path.expandvars(self.carpeta_procesados)).expanduser()
        return p if p.is_absolute() else self.carpeta_origen / p

    @property
    def ruta_ultimas_subidas(self) -> Path:
        return self.script_dir / "logs" / "ultimas_subidas.json"

    @property
    def ruta_pendientes(self) -> Path:
        return self.script_dir / "logs" / "pendientes.json"

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
    def ruta_lock(self) -> Path:
        return self.script_dir / "logs" / "monitor.lock"


# ==============================================================================
# --- NOTIFICACIONES ---
# ==============================================================================
@dataclass(frozen=True)
class NotificationButton:
    etiqueta: str
    destino: str  # URL (https://...) o ruta local


class NotifierProtocol(Protocol):
    def notify(
        self,
        title: str,
        message: str,
        *,
        level: str = "info",
        buttons: Sequence[NotificationButton] = (),
        tag: Optional[str] = None,
    ) -> None: ...


def _to_uri(destino: str) -> str:
    if re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*:(//|/)", destino) and not re.match(r"^[a-zA-Z]:[\\/]", destino):
        return destino
    return Path(destino).resolve().as_uri()


class WindowsToastNotifier:
    """Notificaciones nativas de Windows 10/11 (centro de actividades).

    - Botones de acción ("Abrir hoja", "Ver registro").
    - Icono propio y sonido distinto para errores.
    - El XML viaja en base64 y el script con ``-EncodedCommand``: ningún carácter
      del mensaje (&, <, $, comillas, emojis) puede romper la notificación.
    - No abre ventanas de consola aunque se ejecute con pythonw.
    """

    MAX_TITLE = 80
    MAX_MESSAGE = 300

    def __init__(
        self,
        app_id: str = APP_NAME,
        icon: Optional[Path] = ICON_FILE,
        runner: Callable[..., Any] = subprocess.run,
    ):
        self.app_id = app_id
        self.icon = icon if icon and icon.exists() else None
        self.runner = runner

    @staticmethod
    def is_supported() -> bool:
        return sys.platform == "win32"

    def build_xml(
        self,
        title: str,
        message: str,
        level: str = "info",
        buttons: Sequence[NotificationButton] = (),
    ) -> str:
        title = _truncate(title, self.MAX_TITLE)
        message = _truncate(message, self.MAX_MESSAGE)
        attr = lambda s: xml_escape(s, {'"': "&quot;"})  # noqa: E731

        launch = ""
        if buttons:
            launch = f' activationType="protocol" launch="{attr(_to_uri(buttons[0].destino))}"'
        duration = ' duration="long"' if level == "error" else ""

        lines = message.splitlines() or [""]
        body = lines[0]
        extra = " ".join(l.strip() for l in lines[1:] if l.strip())
        texts = [f"<text hint-maxLines=\"1\">{xml_escape(title)}</text>", f"<text>{xml_escape(body)}</text>"]
        if extra:
            texts.append(f"<text>{xml_escape(extra)}</text>")
        hora = datetime.now().strftime("%H:%M")
        texts.append(f'<text placement="attribution">{xml_escape(APP_NAME)} · {hora}</text>')

        image = ""
        if self.icon:
            image = f'<image placement="appLogoOverride" src="{attr(self.icon.resolve().as_uri())}"/>'

        actions = "".join(
            f'<action content="{attr(b.etiqueta)}" activationType="protocol" arguments="{attr(_to_uri(b.destino))}"/>'
            for b in buttons[:5]
        )
        actions_xml = f"<actions>{actions}</actions>" if actions else ""
        sound = {
            "error": "ms-winsoundevent:Notification.Reminder",
            "warning": "ms-winsoundevent:Notification.Default",
        }.get(level, "ms-winsoundevent:Notification.Default")
        audio = '<audio silent="true"/>' if level == "silent" else f'<audio src="{sound}"/>'

        return (
            f"<toast{launch}{duration}><visual><binding template=\"ToastGeneric\">"
            f"{''.join(texts)}{image}</binding></visual>{actions_xml}{audio}</toast>"
        )

    def build_script(self, xml: str, tag: Optional[str] = None) -> str:
        xml_b64 = base64.b64encode(xml.encode("utf-8")).decode("ascii")
        app_id = self.app_id.replace("'", "''")
        tag_line = ""
        if tag:
            safe_tag = re.sub(r"[^A-Za-z0-9_.-]", "_", tag)[:60]
            tag_line = f"$t.Tag = '{safe_tag}'; $t.Group = 'SheetsAutoSync'\n"
        return (
            "$ErrorActionPreference = 'Stop'\n"
            "[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null\n"
            "[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] | Out-Null\n"
            "$x = New-Object Windows.Data.Xml.Dom.XmlDocument\n"
            f"$x.LoadXml([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{xml_b64}')))\n"
            "$t = [Windows.UI.Notifications.ToastNotification]::new($x)\n"
            f"{tag_line}"
            f"[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('{app_id}').Show($t)\n"
        )

    def notify(
        self,
        title: str,
        message: str,
        *,
        level: str = "info",
        buttons: Sequence[NotificationButton] = (),
        tag: Optional[str] = None,
    ) -> None:
        script = self.build_script(self.build_xml(title, message, level, buttons), tag)
        encoded = base64.b64encode(script.encode("utf-16-le")).decode("ascii")
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        result = self.runner(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
             "-EncodedCommand", encoded],
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.PIPE,
            timeout=30,
            creationflags=creationflags,
        )
        if getattr(result, "returncode", 0) != 0:
            err = (getattr(result, "stderr", b"") or b"").decode("utf-8", "replace").strip()
            raise RuntimeError(f"PowerShell devolvió {result.returncode}: {err[:300]}")


class PlyerNotifier:
    """Respaldo multiplataforma (globo clásico). Sin botones."""

    def notify(self, title: str, message: str, *, level: str = "info",
               buttons: Sequence[NotificationButton] = (), tag: Optional[str] = None) -> None:
        from plyer import notification  # import diferido: solo si se usa

        notification.notify(
            title=_truncate(title, 63),
            message=_truncate(message, 255),
            app_name=APP_NAME,
            app_icon=str(ICON_FILE.with_suffix(".ico")) if ICON_FILE.with_suffix(".ico").exists() else "",
            timeout=10 if level == "error" else 5,
        )


class SafeNotifier:
    """Envía la notificación en segundo plano, con respaldo y sin lanzar nunca.

    - Si la notificación nativa falla, usa la siguiente disponible.
    - Evita repetir el mismo error una y otra vez (``throttle_seconds``).
    - Nunca bloquea la sincronización (se envía en un hilo aparte).
    """

    def __init__(self, backends: Sequence[NotifierProtocol], throttle_seconds: float = 120.0,
                 asynchronous: bool = True, clock: Callable[[], float] = time.monotonic):
        self.backends = list(backends)
        self.throttle_seconds = throttle_seconds
        self.asynchronous = asynchronous
        self.clock = clock
        self._recent: Dict[str, float] = {}
        self._lock = threading.Lock()

    def _is_throttled(self, key: str) -> bool:
        now = self.clock()
        with self._lock:
            last = self._recent.get(key)
            if last is not None and now - last < self.throttle_seconds:
                return True
            self._recent[key] = now
            return False

    def notify(self, title: str, message: str, *, level: str = "info",
               buttons: Sequence[NotificationButton] = (), tag: Optional[str] = None) -> None:
        if level in ("error", "warning") and self._is_throttled(f"{title}|{message}"):
            logger.debug("Notificación repetida omitida: %s", title)
            return
        if self.asynchronous:
            threading.Thread(
                target=self._send, args=(title, message, level, tuple(buttons), tag),
                name="notifier", daemon=True,
            ).start()
        else:
            self._send(title, message, level, tuple(buttons), tag)

    def _send(self, title, message, level, buttons, tag) -> None:
        for backend in self.backends:
            try:
                backend.notify(title, message, level=level, buttons=buttons, tag=tag)
                return
            except Exception as err:
                logger.warning("Notificación con %s falló (%s: %s).", type(backend).__name__, type(err).__name__, err)
        logger.info("[Notificación no mostrada] %s — %s", title, message)


def build_default_notifier() -> SafeNotifier:
    backends: List[NotifierProtocol] = []
    if WindowsToastNotifier.is_supported():
        backends.append(WindowsToastNotifier())
    backends.append(PlyerNotifier())
    return SafeNotifier(backends)


def _truncate(text: str, limit: int) -> str:
    text = (text or "").strip()
    return text if len(text) <= limit else text[: limit - 1].rstrip() + "…"


# ==============================================================================
# --- LECTURA DEL CSV ---
# ==============================================================================
def read_csv_values(path: Path, encodings: Sequence[str] = ("utf-8-sig", "cp1252")) -> List[List[str]]:
    """Lee el CSV probando varias codificaciones y detectando el separador.

    Elimina filas vacías al final. Devuelve una lista de filas (listas de str).
    """
    raw = path.read_bytes()
    text: Optional[str] = None
    last_error: Optional[Exception] = None
    for enc in encodings:
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError as err:
            last_error = err
    if text is None:
        assert last_error is not None
        raise last_error
    text = text.lstrip("﻿")
    if not text.strip():
        return []

    delimiter = _detect_delimiter(text)
    # StringIO(newline="") respeta saltos de línea dentro de celdas entre comillas.
    rows = list(csv.reader(io.StringIO(text, newline=""), delimiter=delimiter))
    while rows and not any(cell.strip() for cell in rows[-1]):
        rows.pop()
    return rows


def _detect_delimiter(text: str) -> str:
    """',' por defecto; ';' o tabulador si la cabecera claramente los usa."""
    header = next(csv.reader([text.splitlines()[0]]), [])
    if len(header) > 1:
        return ","
    first = text.splitlines()[0]
    for candidate in (";", "\t", "|"):
        if candidate in first:
            return candidate
    return ","


def col_to_letter(col_idx: int) -> str:
    """1 -> A, 26 -> Z, 27 -> AA."""
    if col_idx < 1:
        raise ValueError("El índice de columna debe ser >= 1")
    result = ""
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(65 + remainder) + result
    return result


def quote_sheet_name(name: str) -> str:
    """Nombre de pestaña entre comillas para notación A1 (escapa apóstrofes)."""
    return "'" + name.replace("'", "''") + "'"


# ==============================================================================
# --- SERVICIO GOOGLE SHEETS ---
# ==============================================================================
@dataclass(frozen=True)
class SheetInfo:
    title: str
    sheet_id: int
    row_count: int
    column_count: int


@dataclass(frozen=True)
class SyncResult:
    archivo: str
    hoja: str
    filas: int
    columnas: int
    ultima_columna: str
    url: str
    segundos: float
    vacio: bool = False
    simulacion: bool = False
    con_encabezado: bool = True  # False: el encabezado de la hoja (fila 1) no se tocó

    @property
    def resumen(self) -> str:
        if self.vacio:
            return f"'{self.archivo}' no tiene datos; no se hicieron cambios."
        prefijo = "[Simulación] Se cargarían" if self.simulacion else "Se cargaron"
        datos = self.filas - 1 if self.con_encabezado else self.filas
        filas_datos = f"{max(datos, 0):,}".replace(",", ".")  # 7.129
        extra = "(+ encabezado)" if self.con_encabezado else "desde la fila 2 (encabezado de la hoja intacto)"
        return f"{prefijo} {filas_datos} filas {extra} en '{self.hoja}', columnas A–{self.ultima_columna}."


class GoogleSheetsService:
    """Autenticación y escritura en Google Sheets (seguro entre hilos)."""

    def __init__(
        self,
        config: AppConfig,
        interactive: bool = True,
        on_login_required: Optional[Callable[[], None]] = None,
        service_builder: Callable[..., Any] = build,
    ):
        self.config = config
        self.interactive = interactive
        self.on_login_required = on_login_required
        self.service_builder = service_builder
        self._service: Any = None
        self._creds: Optional[Credentials] = None
        self._lock = threading.RLock()

    # ---------------------------- autenticación -----------------------------
    def _save_token(self, creds: Credentials) -> None:
        tmp = self.config.ruta_token.with_suffix(".json.tmp")
        tmp.write_text(creds.to_json(), encoding="utf-8")
        os.replace(tmp, self.config.ruta_token)  # escritura atómica

    def _load_credentials(self) -> Credentials:
        scopes = list(self.config.scopes)
        creds: Optional[Credentials] = None
        token = self.config.ruta_token

        if token.exists():
            try:
                creds = Credentials.from_authorized_user_file(str(token), scopes)
            except (ValueError, json.JSONDecodeError, KeyError) as err:
                logger.warning("token.json inválido (%s). Se pedirá iniciar sesión.", err)
                creds = None

        if creds and creds.valid:
            return creds

        if creds and creds.refresh_token:
            try:
                creds.refresh(Request())
                self._save_token(creds)
                return creds
            except RefreshError as err:
                # invalid_grant: el token caducó o se revocó (en apps de Google Cloud en
                # modo "Prueba" los tokens caducan cada 7 días).
                logger.warning("No se pudo renovar la sesión de Google (%s).", err)
            # TransportError (sin internet) se propaga: es temporal y se reintenta.

        if not self.config.ruta_credentials.exists():
            raise ConfigError(
                f"No se encontró '{self.config.ruta_credentials.name}' en {self.config.script_dir}. "
                "Descárgalo desde Google Cloud Console (ID de cliente OAuth de escritorio)."
            )
        if not self.interactive:
            raise AuthError("La sesión de Google expiró. Ejecuta el monitor para volver a iniciar sesión.")

        if self.on_login_required:
            self.on_login_required()
        logger.info("🔐 Abriendo el navegador para iniciar sesión en Google...")
        flow = InstalledAppFlow.from_client_secrets_file(str(self.config.ruta_credentials), scopes)
        try:
            creds = flow.run_local_server(
                port=0,
                open_browser=True,
                authorization_prompt_message="",
                success_message="Sesión iniciada correctamente. Ya puedes cerrar esta pestaña.",
                timeout_seconds=300,
            )
        except Exception as err:
            raise AuthError(f"No se completó el inicio de sesión en Google ({type(err).__name__}).") from err
        if not creds:
            raise AuthError("No se completó el inicio de sesión en Google.")
        self._save_token(creds)
        logger.info("✅ Sesión de Google guardada en %s", token.name)
        return creds

    def get_service(self) -> Any:
        with self._lock:
            if self._service is None:
                self._creds = self._load_credentials()
                self._service = self.service_builder("sheets", "v4", credentials=self._creds, cache_discovery=False)
            return self._service

    def reset(self) -> None:
        with self._lock:
            self._service = None
            self._creds = None

    def refresh_if_needed(self) -> bool:
        """Renueva la sesión por adelantado (se llama cada minuto en segundo plano).

        El acceso de Google dura 1 hora. Renovarlo antes de que caduque, en cuanto la
        red lo permita, evita depender de oauth2.googleapis.com (que el firewall
        bloquea a ratos) justo cuando llega un CSV. Devuelve True si queda vigente.
        """
        with self._lock:
            creds = self._creds
            if creds is None:
                return False
            expiry = getattr(creds, "expiry", None)
            now = datetime.now(timezone.utc).replace(tzinfo=None)  # google-auth usa UTC sin tz
            if creds.valid and expiry is not None and expiry - now > timedelta(minutes=self.config.minutos_margen_token):
                return True
            try:
                creds.refresh(Request())
                self._save_token(creds)
                logger.debug("🔑 Sesión de Google renovada por adelantado.")
                return True
            except RefreshError:
                raise
            except Exception as err:
                logger.debug("Renovación anticipada falló (%s); se reintentará.", describe_error(err).titulo)
                return False

    # ------------------------------ utilidades ------------------------------
    def _execute(self, request: Any) -> Any:
        return request.execute(num_retries=self.config.reintentos)

    def get_sheet_info(self, rule: FileRule) -> SheetInfo:
        service = self.get_service()
        meta = self._execute(
            service.spreadsheets().get(
                spreadsheetId=rule.spreadsheet_id,
                fields="sheets.properties(title,sheetId,gridProperties(rowCount,columnCount))",
            )
        )
        sheets = [s["properties"] for s in meta.get("sheets", [])]
        wanted = rule.nombre_hoja.strip().casefold()
        for props in sheets:
            if props["title"].strip().casefold() == wanted:
                grid = props.get("gridProperties", {})
                return SheetInfo(props["title"], props["sheetId"], grid.get("rowCount", 0), grid.get("columnCount", 0))
        disponibles = ", ".join(p["title"] for p in sheets[:15])
        raise SheetNotFoundError(
            f"La pestaña '{rule.nombre_hoja}' no existe en el libro. Disponibles: {disponibles}"
        )

    # ------------------------------ escritura -------------------------------
    def update_sheet_with_csv(self, csv_path: Path, rule: FileRule, dry_run: bool = False) -> SyncResult:
        """Sube el CSV a la pestaña indicada por la regla.

        Orden seguro: primero se escriben los datos nuevos encima de los viejos y
        luego se limpian las filas sobrantes. Así la hoja nunca queda vacía si algo
        falla a mitad de camino, y las columnas a la derecha del CSV (fórmulas) no
        se tocan.
        """
        inicio = time.monotonic()
        con_encabezado = rule.incluir_encabezados
        values = read_csv_values(csv_path)
        # Número de columnas según el CSV completo (incluido el encabezado), para
        # limpiar siempre el mismo bloque de columnas aunque se omita la fila 1.
        n_cols = max((len(r) for r in values), default=0)
        if not con_encabezado:
            values = values[1:]  # la hoja ya tiene su encabezado en la fila 1
        if not values:
            logger.warning("El archivo %s no tiene datos.", csv_path.name)
            return SyncResult(csv_path.name, rule.nombre_hoja, 0, 0, "A", rule.url, 0.0, vacio=True,
                              simulacion=dry_run, con_encabezado=con_encabezado)

        start_row = 1 if con_encabezado else 2  # fila de la hoja donde empiezan los datos
        n_rows = len(values)
        last_row = start_row + n_rows - 1
        # Rellenar filas cortas con "" para sobrescribir restos de la carga anterior.
        values = [row + [""] * (n_cols - len(row)) for row in values]
        last_col = col_to_letter(n_cols)

        with self._lock:
            service = self.get_service()
            info = self.get_sheet_info(rule)
            url = f"{rule.url}#gid={info.sheet_id}"
            quoted = quote_sheet_name(info.title)

            if dry_run:
                return SyncResult(csv_path.name, info.title, n_rows, n_cols, last_col, url,
                                  time.monotonic() - inicio, simulacion=True, con_encabezado=con_encabezado)

            # 1. Ampliar la cuadrícula si el CSV no cabe.
            grow = []
            if last_row > info.row_count:
                grow.append({"appendDimension": {"sheetId": info.sheet_id, "dimension": "ROWS",
                                                 "length": last_row - info.row_count}})
            if n_cols > info.column_count:
                grow.append({"appendDimension": {"sheetId": info.sheet_id, "dimension": "COLUMNS",
                                                 "length": n_cols - info.column_count}})
            if grow:
                self._execute(service.spreadsheets().batchUpdate(
                    spreadsheetId=rule.spreadsheet_id, body={"requests": grow}))

            # 2. Escribir por lotes (RAW = se guarda como texto, sin reinterpretar
            #    números largos, ceros a la izquierda, fechas ni fórmulas).
            lote = max(1, self.config.filas_por_lote)
            for start in range(0, n_rows, lote):
                chunk = values[start:start + lote]
                self._execute(service.spreadsheets().values().update(
                    spreadsheetId=rule.spreadsheet_id,
                    range=f"{quoted}!A{start_row + start}",
                    valueInputOption="RAW",
                    body={"values": chunk},
                ))

            # 3. Limpiar filas sobrantes de la carga anterior (solo columnas del CSV).
            if info.row_count > last_row:
                self._execute(service.spreadsheets().values().clear(
                    spreadsheetId=rule.spreadsheet_id,
                    range=f"{quoted}!A{last_row + 1}:{last_col}",
                    body={},
                ))

            # 4. Nota de auditoría en A1.
            if rule.poner_nota_a1:
                now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                filas_datos = n_rows - 1 if con_encabezado else n_rows
                note = (f"Última actualización: {now_str}\nArchivo: {csv_path.name}\n"
                        f"Filas de datos: {filas_datos} (hasta la fila {last_row}) · Columnas: A–{last_col}")
                self._execute(service.spreadsheets().batchUpdate(
                    spreadsheetId=rule.spreadsheet_id,
                    body={"requests": [{"updateCells": {
                        "range": {"sheetId": info.sheet_id, "startRowIndex": 0, "endRowIndex": 1,
                                  "startColumnIndex": 0, "endColumnIndex": 1},
                        "rows": [{"values": [{"note": note}]}],
                        "fields": "note",
                    }}]},
                ))

        result = SyncResult(csv_path.name, info.title, n_rows, n_cols, last_col, url, time.monotonic() - inicio,
                            con_encabezado=con_encabezado)
        logger.info("✅ %s (%.1fs) %s", result.resumen, result.segundos, url)
        return result


# ==============================================================================
# --- PROCESAMIENTO DE ARCHIVOS ---
# ==============================================================================
@dataclass
class SyncStatus:
    """Estado compartido para el icono de la bandeja."""

    ultimo_resultado: Optional[SyncResult] = None
    ultimo_error: Optional[ErrorInfo] = None
    ultimo_archivo: Optional[Path] = None
    ultima_regla: Optional[FileRule] = None
    hora: Optional[datetime] = None
    procesando: Optional[str] = None
    pausado: bool = False
    listeners: List[Callable[[], None]] = field(default_factory=list)

    def changed(self) -> None:
        for cb in list(self.listeners):
            try:
                cb()
            except Exception:
                logger.debug("Listener de estado falló", exc_info=True)


def move_to_processed(path: Path, dest_dir: Path, retries: int = 5, wait: float = 0.5) -> Path:
    """Mueve el archivo a dest_dir sin pisar otro con el mismo nombre ('x (2).csv')."""
    dest_dir.mkdir(parents=True, exist_ok=True)
    target = dest_dir / path.name
    n = 2
    while target.exists():
        target = dest_dir / f"{path.stem} ({n}){path.suffix}"
        n += 1
    for attempt in range(retries):
        try:
            return Path(shutil.move(str(path), str(target)))
        except PermissionError:
            if attempt == retries - 1:
                raise
            time.sleep(wait)  # p. ej. Excel o el antivirus lo tienen abierto
    return target  # pragma: no cover


def in_folder(path: Path, folder: Path) -> bool:
    """True si path está directamente en folder (no en una subcarpeta)."""
    return os.path.normcase(os.path.abspath(path.parent)) == os.path.normcase(os.path.abspath(folder))


class LastSyncStore:
    """Fecha (mtime) del último CSV subido por regla, en logs/ultimas_subidas.json.

    Protege contra subir un CSV más viejo que el último que ya se subió, aunque
    ese último ya no esté en Descargas (se movió a 'procesados') o el monitor se
    haya reiniciado.
    """

    def __init__(self, path: Path):
        self.path = path
        self._lock = threading.Lock()

    def _load(self) -> Dict[str, Dict[str, Any]]:
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
            return data if isinstance(data, dict) else {}
        except (OSError, ValueError):
            return {}

    def last_mtime(self, prefijo: str) -> Optional[int]:
        with self._lock:
            entry = self._load().get(prefijo)
        return int(entry["mtime_ns"]) if entry and "mtime_ns" in entry else None

    def record(self, prefijo: str, path: Path, mtime_ns: int) -> None:
        with self._lock:
            data = self._load()
            prev = data.get(prefijo)
            if prev and int(prev.get("mtime_ns", 0)) > mtime_ns:
                return  # nunca retroceder (p. ej. tras una subida manual de uno viejo)
            data[prefijo] = {"archivo": path.name, "mtime_ns": mtime_ns,
                             "subido": datetime.now().isoformat(timespec="seconds")}
            self.path.parent.mkdir(parents=True, exist_ok=True)
            tmp = self.path.with_suffix(".json.tmp")
            tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
            os.replace(tmp, self.path)


class PendingStore:
    """CSV que no se pudieron subir por un problema temporal (logs/pendientes.json).

    Se guarda UN archivo por regla (el más reciente): cada CSV reemplaza la hoja
    completa, así que reintentar uno viejo después de uno nuevo pisaría datos
    actuales. Cada operación relee el archivo, para compartirlo entre procesos
    (monitor y --archivo).
    """

    def __init__(self, path: Path):
        self.path = path
        self._lock = threading.Lock()

    def _load(self) -> Dict[str, Dict[str, Any]]:
        try:
            data = json.loads(self.path.read_text(encoding="utf-8"))
            return data if isinstance(data, dict) else {}
        except FileNotFoundError:
            return {}
        except (OSError, ValueError) as err:
            logger.warning("No se pudo leer %s (%s); se empieza vacío.", self.path.name, err)
            return {}

    def _save(self, items: Dict[str, Dict[str, Any]]) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        tmp = self.path.with_suffix(".json.tmp")
        tmp.write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(tmp, self.path)

    def set(self, prefijo: str, path: Path, motivo: str) -> bool:
        """Registra el archivo de la regla. True si es nuevo (o reemplaza a otro)."""
        with self._lock:
            items = self._load()
            prev = items.get(prefijo)
            nuevo = prev is None or prev.get("archivo") != str(path)
            if nuevo:
                if prev:
                    logger.info("♻️ '%s' reemplaza a '%s' en pendientes (es más reciente).",
                                path.name, Path(prev["archivo"]).name)
                prev = {"archivo": str(path), "desde": datetime.now().isoformat(timespec="seconds"), "intentos": 0}
            prev["intentos"] = int(prev.get("intentos", 0)) + 1
            prev["motivo"] = motivo
            prev["ultimo_intento"] = datetime.now().isoformat(timespec="seconds")
            items[prefijo] = prev
            self._save(items)
            return nuevo

    def get(self, prefijo: str) -> Optional[Dict[str, Any]]:
        with self._lock:
            return self._load().get(prefijo)

    def remove(self, prefijo: str, only_if: Optional[Path] = None) -> bool:
        with self._lock:
            items = self._load()
            entry = items.get(prefijo)
            if entry is None or (only_if is not None and entry.get("archivo") != str(only_if)):
                return False
            del items[prefijo]
            self._save(items)
            return True

    def resolve_after_success(self, prefijo: str, synced: Path) -> bool:
        """Tras subir `synced`, quita el pendiente de la regla si es ese mismo archivo
        o uno más antiguo (ya no tiene sentido subirlo). Devuelve True si quitó algo."""
        with self._lock:
            items = self._load()
            entry = items.get(prefijo)
            if entry is None:
                return False
            pending_path = Path(entry["archivo"])
            if pending_path != synced:
                try:
                    if pending_path.exists() and pending_path.stat().st_mtime > synced.stat().st_mtime:
                        return False  # el pendiente es más nuevo: se conserva
                except OSError:
                    pass
            del items[prefijo]
            self._save(items)
            return True

    def items(self) -> List[Tuple[str, Dict[str, Any]]]:
        with self._lock:
            return list(self._load().items())

    def __len__(self) -> int:
        with self._lock:
            return len(self._load())


class FileProcessor:
    """Espera a que la descarga termine, sube el CSV y notifica el resultado."""

    def __init__(
        self,
        config: AppConfig,
        rules: RulesStore,
        sheets: GoogleSheetsService,
        notifier: NotifierProtocol,
        status: Optional[SyncStatus] = None,
        sleep: Callable[[float], None] = time.sleep,
        clock: Callable[[], float] = time.monotonic,
        pending: Optional[PendingStore] = None,
    ):
        self.config = config
        self.rules = rules
        self.sheets = sheets
        self.notifier = notifier
        self.pending = pending or PendingStore(config.ruta_pendientes)
        self.last_sync = LastSyncStore(config.ruta_ultimas_subidas)
        self.status = status or SyncStatus()
        self.sleep = sleep
        self.clock = clock
        # Hora (reloj de pared) en que arrancó el monitor: los archivos que ya
        # estaban en Descargas antes de esta hora NUNCA se suben automáticamente.
        self.started_at = time.time()
        self._seen: Dict[str, Tuple[int, str]] = {}  # ruta -> (tamaño, sha1) ya procesados
        self._lock = threading.Lock()

    # ------------------------------------------------------------------ filtro
    def match(self, path: Path) -> Optional[FileRule]:
        if path.suffix.lower() in TEMP_SUFFIXES or path.name.startswith((".", "~$")):
            return None
        if path.suffix.lower() != self.config.extension_requerida.lower():
            return None
        return self.rules.match(path.name)

    MARGEN_ARRANQUE_SEG = 10.0

    def is_new_download(self, path: Path) -> bool:
        """True solo si el contenido del archivo cambió desde que arrancó el monitor.

        Windows (antivirus, indexador, vista previa del Explorador) lee los archivos
        de Descargas y eso dispara eventos de "modificado" aunque el contenido no
        cambie. Un CSV viejo NUNCA debe volver a subirse por eso: se exige que su
        fecha de modificación sea posterior al arranque del monitor.
        """
        try:
            mtime = path.stat().st_mtime
        except OSError:
            return False
        return mtime >= self.started_at - self.MARGEN_ARRANQUE_SEG

    def newer_file_for_rule(self, path: Path, rule: FileRule) -> Optional[Path]:
        """Otro CSV de la misma regla más reciente en la carpeta (si existe)."""
        try:
            mtime = path.stat().st_mtime_ns
        except OSError:
            return None
        newest: Optional[Tuple[int, Path]] = None
        try:
            candidates = list(path.parent.iterdir())
        except OSError:
            return None
        for other in candidates:
            if other == path or other.suffix.lower() != self.config.extension_requerida.lower():
                continue
            if not other.name.lower().startswith(rule.prefijo.lower()):
                continue
            other_rule = self.match(other)
            if other_rule is None or other_rule.prefijo != rule.prefijo:
                continue  # pertenece a una regla más específica
            try:
                other_mtime = other.stat().st_mtime_ns
            except OSError:
                continue
            if other_mtime > mtime and (newest is None or other_mtime > newest[0]):
                newest = (other_mtime, other)
        return newest[1] if newest else None

    # ----------------------------------------------------- espera de descarga
    def wait_until_ready(self, path: Path) -> bool:
        """True cuando el archivo existe, su tamaño no cambia y se puede abrir."""
        deadline = self.clock() + self.config.segundos_espera_maxima
        last_sig: Optional[Tuple[int, int]] = None
        stable_since: Optional[float] = None
        while self.clock() < deadline:
            try:
                st = path.stat()
                sig = (st.st_size, st.st_mtime_ns)
            except FileNotFoundError:
                return False
            except OSError:
                sig = None

            now = self.clock()
            if sig is not None and sig == last_sig and sig[0] > 0:
                if stable_since is not None and now - stable_since >= self.config.segundos_estabilidad:
                    try:
                        with open(path, "rb") as fh:
                            fh.read(1)
                        return True
                    except PermissionError:
                        pass  # el navegador aún lo tiene bloqueado
            else:
                last_sig, stable_since = sig, now
            self.sleep(0.25)
        logger.warning("⏱️ %s no terminó de descargarse en %ss.", path.name, self.config.segundos_espera_maxima)
        return False

    def _is_duplicate(self, path: Path) -> bool:
        """True si ya se procesó este archivo con exactamente el mismo contenido.

        Se compara una huella (SHA-1) del contenido, no la fecha: así ninguna
        lectura o cambio de atributos de Windows provoca otra subida, sin ventana
        de tiempo. Si el contenido cambia de verdad, se procesa de nuevo.
        """
        try:
            data = path.read_bytes()
        except OSError:
            return False
        key, sig = str(path).lower(), (len(data), hashlib.sha1(data).hexdigest())
        with self._lock:
            if self._seen.get(key) == sig:
                return True
            self._seen[key] = sig
            return False

    # -------------------------------------------------------------- proceso
    def process(self, path: Path, force: bool = False, dry_run: bool = False,
                from_pending: bool = False) -> Optional[SyncResult]:
        """Procesa el archivo. Nunca lanza excepciones (las notifica y registra)."""
        try:
            rule = self.match(path)
        except ConfigError as err:
            self._report_error(path, err)
            return None
        if rule is None:
            return None
        if self.status.pausado and not force:
            logger.info("⏸️ Monitoreo en pausa: se ignora %s", path.name)
            return None
        if from_pending:
            entry = self.pending.get(rule.prefijo)
            if entry is None or entry.get("archivo") != str(path):
                logger.debug("%s ya no es el pendiente de '%s'; se omite.", path.name, rule.prefijo)
                return None

        if not force and not self.is_new_download(path):
            logger.debug("Ignorado (ya existía antes de arrancar el monitor): %s", path.name)
            return None

        if not self.wait_until_ready(path):
            if from_pending and not path.exists():
                self.pending.remove(rule.prefijo, only_if=path)
                logger.info("🗑️ %s ya no existe; se quita de pendientes.", path.name)
            return None
        if not force and self._is_duplicate(path):
            logger.debug("Evento duplicado ignorado: %s", path.name)
            return None
        if not force or from_pending:
            # Cada CSV reemplaza la hoja entera: nunca subir uno más antiguo que otro
            # de la misma regla que ya está en la carpeta.
            newer = self.newer_file_for_rule(path, rule)
            try:
                mtime_ns = path.stat().st_mtime_ns
            except OSError:
                mtime_ns = 0
            last = self.last_sync.last_mtime(rule.prefijo)
            motivo = None
            if newer is not None:
                motivo = f"hay uno más reciente en la carpeta ({newer.name})"
            elif last is not None and mtime_ns < last:
                motivo = "ya se subió uno más reciente"
            if motivo:
                logger.info("⏭️ Se omite %s (regla '%s'): %s.", path.name, rule.prefijo, motivo)
                if from_pending:
                    self.pending.remove(rule.prefijo, only_if=path)
                return None

        logger.info("📄 %s: %s → regla '%s' (pestaña '%s')", "Reintentando pendiente" if from_pending else "Archivo detectado",
                    path.name, rule.prefijo, rule.nombre_hoja)
        self.status.procesando = path.name
        self.status.ultimo_archivo, self.status.ultima_regla = path, rule
        self.status.changed()
        try:
            result = self._sync_with_retries(path, rule, dry_run, quick=from_pending)
        except Exception as err:
            self._handle_failure(path, rule, err, dry_run)
            return None
        finally:
            self.status.procesando = None

        was_pending = False if dry_run else self.pending.resolve_after_success(rule.prefijo, path)
        if not dry_run and not result.vacio:
            try:
                self.last_sync.record(rule.prefijo, path, path.stat().st_mtime_ns)
            except OSError:
                pass
        if rule.mover_a_procesados and not dry_run and not result.vacio:
            try:
                destino = move_to_processed(path, self.config.ruta_procesados)
                logger.info("📦 %s movido a %s", path.name, destino.parent)
                self.status.ultimo_archivo = destino  # "Reintentar último archivo" sigue funcionando
            except Exception as err:  # no es grave: la hoja ya está actualizada
                logger.warning("No se pudo mover %s a procesados: %s", path.name, err)
        self.status.ultimo_resultado, self.status.ultimo_error = result, None
        self.status.hora = datetime.now()
        self.status.changed()
        level = "warning" if result.vacio else "info"
        if result.vacio:
            titulo = "⚠️ CSV sin datos"
        elif result.simulacion:
            titulo = f"🔎 Simulación: {result.hoja}"
        elif was_pending and from_pending:
            titulo = f"✅ {result.hoja} actualizada tras reintentar"
        else:
            titulo = f"✅ {result.hoja} actualizada"
        self.notifier.notify(
            titulo,
            f"{result.resumen}\n{result.archivo}",
            level=level,
            buttons=(NotificationButton("Abrir hoja", result.url),),
            tag=f"ok-{rule.prefijo}",
        )
        return result

    def _sync_with_retries(self, path: Path, rule: FileRule, dry_run: bool, quick: bool = False) -> SyncResult:
        esperas = () if quick else tuple(self.config.esperas_reintento)
        relogin_done = False
        intento = 0
        while True:
            try:
                return self.sheets.update_sheet_with_csv(path, rule, dry_run=dry_run)
            except RefreshError as err:
                # Sesión revocada: se borra el token y se reintenta una vez (pedirá
                # iniciar sesión si el monitor es interactivo).
                self.sheets.reset()
                try:
                    self.sheets.config.ruta_token.unlink(missing_ok=True)
                except OSError:
                    pass
                if relogin_done:
                    raise AuthError("La sesión de Google expiró. Vuelve a iniciar sesión.") from err
                relogin_done = True
                continue
            except Exception as err:
                info = describe_error(err)
                if not info.transitorio or intento >= len(esperas):
                    raise
                logger.warning("🔁 %s: %s. Reintento %d/%d en %.0fs...",
                               path.name, info.titulo, intento + 1, len(esperas), esperas[intento])
                self.sleep(esperas[intento])
                intento += 1

    def _handle_failure(self, path: Path, rule: FileRule, err: BaseException, dry_run: bool) -> None:
        info = describe_error(err)
        if info.transitorio and not dry_run:
            nuevo = self.pending.set(rule.prefijo, path, info.titulo)
            logger.warning("⏳ %s queda pendiente (%s). Se reintentará cada %g min.",
                           path.name, info.titulo, self.config.minutos_reintento_pendientes)
            log_details(err, logging.DEBUG)
            self.status.ultimo_error, self.status.hora = info, datetime.now()
            self.status.changed()
            if nuevo:
                self.notifier.notify(
                    f"⏳ {info.titulo}",
                    f"{path.name}\nSe reintentará automáticamente cada {self.config.minutos_reintento_pendientes:g} min.",
                    level="warning", buttons=(NotificationButton("Ver registro", str(LOG_FILE)),), tag="pendiente",
                )
            return
        self._report_error(path, err)

    def _report_error(self, path: Path, err: BaseException) -> None:
        info = describe_error(err)
        logger.error("❌ Fallo al procesar %s: %s — %s", path.name, info.titulo, info.mensaje)
        log_details(err, logging.ERROR)
        self.status.ultimo_error, self.status.hora = info, datetime.now()
        self.status.changed()
        buttons = [NotificationButton("Ver registro", str(LOG_FILE))]
        if self.status.ultima_regla is not None and isinstance(err, (SheetNotFoundError, HttpError)):
            buttons.insert(0, NotificationButton("Abrir libro", self.status.ultima_regla.url))
        self.notifier.notify(
            f"❌ {info.titulo}",
            f"{info.mensaje}\nArchivo: {path.name}",
            level="error",
            buttons=buttons,
            tag="error",
        )


class SyncWorker:
    """Cola de trabajo: los eventos de watchdog solo encolan; un hilo procesa.

    Así el observador nunca se bloquea ni muere por una excepción, y los
    archivos se procesan de uno en uno (la API de Google no es thread-safe).
    """

    def __init__(self, processor: FileProcessor):
        self.processor = processor
        self.queue: "queue.Queue[Tuple[Path, bool]]" = queue.Queue()
        self._pending: set = set()
        self._lock = threading.Lock()
        self._stop = threading.Event()
        self._thread = threading.Thread(target=self._run, name="sync-worker", daemon=True)

    def start(self) -> None:
        self._thread.start()

    def stop(self, timeout: float = 5.0) -> None:
        self._stop.set()
        self._thread.join(timeout)

    def submit(self, path: Path, force: bool = False, from_pending: bool = False) -> bool:
        key = str(path).lower()
        with self._lock:
            if key in self._pending and (not force or from_pending):
                return False
            self._pending.add(key)
        self.queue.put((path, force, from_pending))
        return True

    def _run(self) -> None:
        while not self._stop.is_set():
            try:
                path, force, from_pending = self.queue.get(timeout=0.5)
            except queue.Empty:
                continue
            try:
                self.processor.process(path, force=force, from_pending=from_pending)
            except Exception:  # defensa extra: el hilo nunca debe morir
                logger.exception("Error inesperado en el procesador")
            finally:
                with self._lock:
                    self._pending.discard(str(path).lower())
                self.queue.task_done()


class BackgroundJobs:
    """Tareas periódicas: renovar la sesión antes de que caduque y reintentar pendientes."""

    def __init__(self, config: AppConfig, sheets: GoogleSheetsService, worker: SyncWorker,
                 pending: PendingStore, notifier: NotifierProtocol):
        self.config, self.sheets, self.worker = config, sheets, worker
        self.pending, self.notifier = pending, notifier
        self._stop = threading.Event()
        self._threads = [
            threading.Thread(target=self._token_loop, name="token-keeper", daemon=True),
            threading.Thread(target=self._pending_loop, name="pending-retry", daemon=True),
        ]

    def start(self) -> None:
        for t in self._threads:
            t.start()

    def stop(self) -> None:
        self._stop.set()

    def _token_loop(self) -> None:
        while not self._stop.wait(60):
            try:
                self.sheets.refresh_if_needed()
            except RefreshError:
                logger.warning("La sesión de Google fue revocada; se pedirá iniciar sesión en la próxima subida.")
                self.sheets.reset()
            except Exception:
                logger.debug("token-keeper", exc_info=True)

    def _pending_loop(self) -> None:
        if self._stop.wait(20):  # primer intento poco después de arrancar
            return
        while True:
            try:
                self.retry_pending()
            except Exception:
                logger.exception("Error al reintentar pendientes")
            if self._stop.wait(self.config.minutos_reintento_pendientes * 60):
                return

    def retry_pending(self) -> int:
        """Encola los pendientes; descarta los que ya no existen o son demasiado antiguos."""
        enviados = 0
        limite = datetime.now() - timedelta(hours=self.config.horas_max_pendiente)
        for prefijo, entry in self.pending.items():
            path = Path(entry.get("archivo", ""))
            if not path.exists():
                self.pending.remove(prefijo, only_if=path)
                logger.info("🗑️ %s ya no existe; se quita de pendientes.", path.name)
                continue
            try:
                desde = datetime.fromisoformat(entry.get("desde", ""))
            except ValueError:
                desde = datetime.now()
            if desde < limite:
                self.pending.remove(prefijo, only_if=path)
                logger.error("❌ %s lleva más de %gh sin poder subirse; se descarta.", path.name,
                             self.config.horas_max_pendiente)
                self.notifier.notify("❌ No se pudo subir", f"{path.name}\nVuelve a descargarlo o usa 'Reintentar' en la bandeja.",
                                     level="error", buttons=(NotificationButton("Ver registro", str(LOG_FILE)),))
                continue
            if self.worker.submit(path, force=True, from_pending=True):
                enviados += 1
        return enviados


class CSVDownloadHandler(FileSystemEventHandler):
    """Traduce eventos del sistema de archivos a trabajos en la cola."""

    def __init__(self, worker: SyncWorker):
        self.worker = worker

    def _handle(self, raw_path: Any) -> None:
        if isinstance(raw_path, bytes):
            raw_path = os.fsdecode(raw_path)
        path = Path(raw_path)
        try:
            processor = self.worker.processor
            if (processor.match(path) is not None and in_folder(path, processor.config.carpeta_origen)
                    and processor.is_new_download(path)):  # ignora lo movido a 'procesados'
                self.worker.submit(path)
        except ConfigError as err:
            self.worker.processor._report_error(path, err)
        except Exception:
            logger.exception("Error al evaluar el evento de %s", path)

    def on_created(self, event: FileSystemEvent) -> None:
        if not event.is_directory:
            self._handle(event.src_path)

    def on_moved(self, event: FileSystemEvent) -> None:  # Chrome/Edge: .crdownload -> .csv
        if not event.is_directory:
            self._handle(event.dest_path)

    def on_modified(self, event: FileSystemEvent) -> None:  # Firefox / sobrescrituras
        if not event.is_directory:
            self._handle(event.src_path)


# ==============================================================================
# --- INSTANCIA ÚNICA ---
# ==============================================================================
class SingleInstance:
    """Evita que el monitor se ejecute dos veces (subidas duplicadas)."""

    def __init__(self, lock_path: Path):
        self.lock_path = lock_path
        self._fh: Any = None

    def acquire(self) -> bool:
        self.lock_path.parent.mkdir(parents=True, exist_ok=True)
        fh = open(self.lock_path, "a+")
        try:
            if sys.platform == "win32":
                import msvcrt

                fh.seek(0)
                msvcrt.locking(fh.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl

                fcntl.flock(fh.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError:
            fh.close()
            return False
        fh.seek(0)
        fh.truncate()
        fh.write(str(os.getpid()))
        fh.flush()
        self._fh = fh
        return True

    def release(self) -> None:
        if self._fh is None:
            return
        try:
            if sys.platform == "win32":
                import msvcrt

                self._fh.seek(0)
                msvcrt.locking(self._fh.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                import fcntl

                fcntl.flock(self._fh.fileno(), fcntl.LOCK_UN)
        except OSError:
            pass
        self._fh.close()
        self._fh = None


# ==============================================================================
# --- ICONO EN LA BANDEJA (opcional, requiere pystray + Pillow) ---
# ==============================================================================
def open_path(target: str) -> None:
    """Abre una URL, carpeta o archivo con la aplicación predeterminada."""
    try:
        if re.match(r"^https?://", target):
            import webbrowser

            webbrowser.open(target)
        elif sys.platform == "win32":
            os.startfile(target)  # type: ignore[attr-defined]
        else:
            subprocess.Popen(["xdg-open", target])
    except Exception as err:
        logger.warning("No se pudo abrir %s: %s", target, err)


class TrayApp:
    """Icono en la bandeja del sistema con estado y accesos rápidos."""

    def __init__(self, config: AppConfig, worker: SyncWorker, status: SyncStatus, on_quit: Callable[[], None],
                 jobs: Optional[BackgroundJobs] = None):
        import pystray  # noqa: F401  (lanza ImportError si no está instalado)
        from PIL import Image

        self.pystray = pystray
        self.config = config
        self.worker = worker
        self.status = status
        self.on_quit = on_quit
        self.jobs = jobs
        self.pending = worker.processor.pending
        base = Image.open(ICON_FILE).convert("RGBA") if ICON_FILE.exists() else Image.new("RGBA", (64, 64), (15, 157, 88, 255))
        self.images = {
            "ok": base,
            "error": self._badge(base, (219, 68, 55, 255)),
            "busy": self._badge(base, (251, 188, 5, 255)),
            "pending": self._badge(base, (255, 138, 0, 255)),
            "paused": self._badge(base, (128, 128, 128, 255)),
        }
        self.icon = pystray.Icon("SheetsAutoSync", base, APP_NAME, menu=self._menu())
        status.listeners.append(self.refresh)

    @staticmethod
    def _badge(base: Any, color: Tuple[int, int, int, int]) -> Any:
        from PIL import ImageDraw

        img = base.copy()
        w, h = img.size
        r = int(w * 0.22)
        draw = ImageDraw.Draw(img)
        draw.ellipse((w - 2 * r - 4, 4, w - 4, 2 * r + 4), fill=color, outline=(255, 255, 255, 255), width=max(2, w // 40))
        return img

    def _tooltip(self) -> str:
        s = self.status
        if s.pausado:
            return f"{APP_NAME} — en pausa"
        if s.procesando:
            return _truncate(f"{APP_NAME} — subiendo {s.procesando}", 120)
        n = len(self.pending)
        if n:
            return f"{APP_NAME} — {n} pendiente(s) por reintentar"
        if s.ultimo_error and s.hora:
            return _truncate(f"{APP_NAME} — error {s.hora:%H:%M}: {s.ultimo_error.titulo}", 120)
        if s.ultimo_resultado and s.hora:
            return _truncate(f"{APP_NAME} — {s.ultimo_resultado.hoja} OK a las {s.hora:%H:%M}", 120)
        return f"{APP_NAME} — vigilando Descargas"

    def _menu(self) -> Any:
        item, menu = self.pystray.MenuItem, self.pystray.Menu
        s = self.status
        return menu(
            item(lambda _: self._tooltip().replace(f"{APP_NAME} — ", "Estado: "), None, enabled=False),
            menu.SEPARATOR,
            item("Abrir última hoja", lambda: open_path(s.ultimo_resultado.url),
                 enabled=lambda _: bool(s.ultimo_resultado), default=True),
            item(lambda _: f"Reintentar pendientes ({len(self.pending)})", self._retry_pending,
                 enabled=lambda _: self.jobs is not None and len(self.pending) > 0),
            item("Reintentar último archivo", self._retry,
                 enabled=lambda _: bool(s.ultimo_archivo and s.ultimo_archivo.exists())),
            item("Sincronizar un archivo…", self._pick_file),
            item("Pausar monitoreo", self._toggle_pause, checked=lambda _: s.pausado),
            menu.SEPARATOR,
            item("Abrir carpeta Descargas", lambda: open_path(str(self.config.carpeta_origen))),
            item("Abrir carpeta de procesados", lambda: open_path(str(self.config.ruta_procesados)),
                 enabled=lambda _: self.config.ruta_procesados.is_dir()),
            item("Editar reglas (rules.yaml)", lambda: open_path(str(self.config.ruta_rules))),
            item("Ver registro", lambda: open_path(str(LOG_FILE))),
            menu.SEPARATOR,
            item("Salir", self._quit),
        )

    def refresh(self) -> None:
        s = self.status
        key = ("paused" if s.pausado else "busy" if s.procesando else "pending" if len(self.pending)
               else "error" if s.ultimo_error else "ok")
        try:
            self.icon.icon = self.images[key]
            self.icon.title = self._tooltip()[:127]
            self.icon.update_menu()
        except Exception:
            logger.debug("No se pudo refrescar el icono", exc_info=True)

    def _retry(self) -> None:
        if self.status.ultimo_archivo:
            self.worker.submit(self.status.ultimo_archivo, force=True)

    def _retry_pending(self) -> None:
        if self.jobs is not None:
            logger.info("🔁 Reintento manual de %d pendiente(s).", self.jobs.retry_pending())

    def _toggle_pause(self) -> None:
        self.status.pausado = not self.status.pausado
        logger.info("⏯️ Monitoreo %s", "en pausa" if self.status.pausado else "reanudado")
        self.status.changed()

    def _pick_file(self) -> None:
        def ask() -> None:
            try:
                import tkinter as tk
                from tkinter import filedialog

                root = tk.Tk()
                root.withdraw()
                root.attributes("-topmost", True)
                chosen = filedialog.askopenfilename(
                    title="Elige el CSV a sincronizar",
                    initialdir=str(self.config.carpeta_origen),
                    filetypes=[("CSV", "*.csv")],
                )
                root.destroy()
                if chosen:
                    path = Path(chosen)
                    if self.worker.processor.match(path) is None:
                        self.worker.processor.notifier.notify(
                            "Sin regla para este archivo",
                            f"'{path.name}' no empieza por ningún prefijo de rules.yaml.",
                            level="warning",
                            buttons=(NotificationButton("Editar reglas", str(self.config.ruta_rules)),),
                        )
                    else:
                        self.worker.submit(path, force=True)
            except Exception:
                logger.exception("No se pudo abrir el selector de archivos")

        threading.Thread(target=ask, name="file-picker", daemon=True).start()

    def _quit(self) -> None:
        self.on_quit()
        self.icon.stop()

    def run(self) -> None:
        self.refresh()
        self.icon.run()


# ==============================================================================
# --- COMANDOS ---
# ==============================================================================
def build_app(config: AppConfig, notifier: NotifierProtocol, interactive: bool = True):
    def rules_error(err: Exception) -> None:
        notifier.notify("⚠️ Error en rules.yaml", f"{err}\nSe siguen usando las reglas anteriores.",
                        level="warning", buttons=(NotificationButton("Editar reglas", str(config.ruta_rules)),))

    def login_required() -> None:
        notifier.notify("🔐 Inicia sesión en Google",
                        "Se abrió el navegador para renovar el acceso a Google Sheets.", level="warning")

    rules = RulesStore(config.ruta_rules, on_error=rules_error)
    sheets = GoogleSheetsService(config, interactive=interactive, on_login_required=login_required)
    status = SyncStatus()
    processor = FileProcessor(config, rules, sheets, notifier, status)
    return rules, sheets, processor, status


def check_host(host: str) -> Tuple[bool, str]:
    """¿La red deja llegar al servidor real de Google o muestra la página de bloqueo?"""
    try:
        ip = socket.gethostbyname(host)
    except OSError as err:
        return False, f"no se pudo resolver ({err})"
    try:
        ctx = ssl.create_default_context()
        with socket.create_connection((host, 443), timeout=8) as sock, ctx.wrap_socket(sock, server_hostname=host):
            return True, ip
    except ssl.SSLCertVerificationError:
        return False, f"{ip} → certificado no válido (bloqueo del firewall)"
    except OSError as err:
        return False, f"{ip} → {err}"


def call_with_network_retry(fn: Callable[[], Any], intentos: int = 7, espera: float = 10.0,
                            sleep: Callable[[float], None] = time.sleep) -> Any:
    """Para comandos interactivos: reintenta mientras la red bloquee a Google."""
    for i in range(intentos):
        try:
            return fn()
        except Exception as err:
            info = describe_error(err)
            if not info.transitorio or i == intentos - 1:
                raise
            print(f"   ⏳ {info.titulo}. Reintentando en {espera:.0f}s ({i + 1}/{intentos - 1})…", flush=True)
            sleep(espera)


def cmd_verify(config: AppConfig) -> int:
    """Comprueba reglas, red, sesión de Google, pestañas y pendientes."""
    ok = True
    print(f"📁 Carpeta vigilada: {config.carpeta_origen} {'✔' if config.carpeta_origen.is_dir() else '✘ NO EXISTE'}")
    ok &= config.carpeta_origen.is_dir()
    try:
        rules = RulesStore(config.ruta_rules).get()
        print(f"📋 rules.yaml: {len(rules)} regla(s) ✔")
    except ConfigError as err:
        print(f"📋 rules.yaml ✘ {err}")
        return 1
    for host in ("oauth2.googleapis.com", "sheets.googleapis.com"):
        good, detail = check_host(host)
        print(f"🌐 {host}: {'✔ ' + detail if good else '✘ ' + detail}")
    try:
        sheets = GoogleSheetsService(config, interactive=True)
        call_with_network_retry(sheets.get_service)
        print("🔐 Sesión de Google ✔")
    except Exception as err:
        e = describe_error(err)
        print(f"🔐 Sesión de Google ✘ {e.titulo}: {e.mensaje}")
        if is_network_block(err):
            print("   La red sigue bloqueando oauth2.googleapis.com. Inténtalo en unos minutos o pide a TI\n"
                  "   que lo agregue a la lista de permitidos del filtro Fortinet.")
        log_details(err, logging.WARNING)
        return 1
    for rule in rules:
        try:
            info = sheets.get_sheet_info(rule)
            print(f"   ✔ '{rule.prefijo}*' → '{info.title}' ({info.row_count} filas × {info.column_count} cols)")
        except Exception as err:
            ok = False
            e = describe_error(err)
            print(f"   ✘ '{rule.prefijo}*' → {e.titulo}: {e.mensaje}")
    pend = PendingStore(config.ruta_pendientes).items()
    print(f"⏳ Pendientes: {len(pend)}")
    for prefijo, entry in pend:
        print(f"   • {prefijo}: {Path(entry['archivo']).name} (desde {entry.get('desde')}, "
              f"{entry.get('intentos')} intento(s), {entry.get('motivo')})")
    print("Todo correcto ✔" if ok else "Hay problemas que revisar ✘")
    return 0 if ok else 1


def cmd_single_file(config: AppConfig, path: Path, dry_run: bool, notifier: NotifierProtocol) -> int:
    _, _, processor, _ = build_app(config, notifier)
    if not path.exists():
        print(f"No existe: {path}")
        return 1
    if processor.match(path) is None:
        print(f"'{path.name}' no coincide con ninguna regla de rules.yaml.")
        return 1
    result = processor.process(path, force=True, dry_run=dry_run)
    if result:
        print(result.resumen)
        print(result.url)
        return 0
    rule = processor.match(path)
    entry = processor.pending.get(rule.prefijo) if rule else None
    if entry and entry.get("archivo") == str(path):
        print("⏳ No se pudo subir ahora (problema temporal). Quedó pendiente: el monitor lo reintentará solo.")
    return 1


def run_monitor(config: AppConfig, use_tray: bool = True) -> int:
    notifier = build_default_notifier()
    instance = SingleInstance(config.ruta_lock)
    if not instance.acquire():
        logger.warning("Ya hay otro monitor en ejecución. Se cierra esta instancia.")
        notifier.notify("Auto-Sync ya está en ejecución", "No se inició una segunda copia del monitor.", level="warning")
        time.sleep(1)  # dar tiempo a mostrar la notificación
        return 0

    observer = worker = jobs = None
    try:
        rules, sheets, processor, status = build_app(config, notifier)
        reglas = rules.get()  # valida rules.yaml al arrancar
        if not config.carpeta_origen.is_dir():
            raise ConfigError(f"La carpeta a vigilar no existe: {config.carpeta_origen}")

        # Verifica la sesión al arrancar (si hace falta login, se pide ahora y no
        # en medio de una descarga). Sin red no bloquea: se reintenta en segundo plano.
        try:
            sheets.get_service()
        except (TransportError, OSError) as err:
            logger.warning("Sin acceso a Google al iniciar (%s). Se reintentará en segundo plano.",
                           describe_error(err).titulo)

        worker = SyncWorker(processor)
        worker.start()
        jobs = BackgroundJobs(config, sheets, worker, processor.pending, notifier)
        jobs.start()
        observer = Observer()
        observer.schedule(CSVDownloadHandler(worker), str(config.carpeta_origen), recursive=False)
        observer.start()

        logger.info("🚀 Monitor activo. Carpeta: %s", config.carpeta_origen)
        for regla in reglas:
            logger.info("   • '%s*' → hoja '%s' (%s)", regla.prefijo, regla.nombre_hoja, regla.spreadsheet_id)
        n_pend = len(processor.pending)
        notifier.notify(
            "Auto-Sync activo",
            f"Vigilando Descargas · {len(reglas)} regla(s): " + ", ".join(r.prefijo for r in reglas)
            + (f" · {n_pend} pendiente(s)" if n_pend else ""),
            level="silent",
            buttons=(NotificationButton("Ver registro", str(LOG_FILE)),),
            tag="startup",
        )

        stop = threading.Event()
        tray = None
        if use_tray:
            try:
                tray = TrayApp(config, worker, status, on_quit=stop.set, jobs=jobs)
            except ImportError:
                logger.info("pystray no instalado: se ejecuta sin icono en la bandeja.")
            except Exception:
                logger.exception("No se pudo crear el icono de la bandeja")

        try:
            if tray is not None:
                tray.run()  # bloquea hasta "Salir"
            else:
                while not stop.wait(1):
                    if not observer.is_alive():
                        raise RuntimeError("El observador de archivos se detuvo inesperadamente.")
        except KeyboardInterrupt:
            pass
        logger.info("🛑 Deteniendo el monitor...")
        return 0
    except Exception as err:
        info = describe_error(err)
        logger.critical("❌ No se pudo iniciar el monitor: %s — %s", info.titulo, info.mensaje)
        log_details(err, logging.CRITICAL)
        notifier.notify(f"❌ Auto-Sync no pudo iniciar: {info.titulo}", info.mensaje, level="error",
                        buttons=(NotificationButton("Ver registro", str(LOG_FILE)),
                                 NotificationButton("Editar reglas", str(config.ruta_rules))))
        time.sleep(2)  # que la notificación salga antes de cerrar
        return 1
    finally:
        if observer is not None:
            observer.stop()
            observer.join(5)
        if jobs is not None:
            jobs.stop()
        if worker is not None:
            worker.stop()
        instance.release()


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Sincroniza CSV descargados con Google Sheets.")
    parser.add_argument("--verificar", action="store_true", help="Valida reglas, sesión de Google y pestañas.")
    parser.add_argument("--archivo", type=Path, help="Sincroniza este CSV y termina.")
    parser.add_argument("--simulacion", action="store_true", help="Con --archivo: no escribe nada en Sheets.")
    parser.add_argument("--carpeta", type=Path, help="Carpeta a vigilar (por defecto, Descargas).")
    parser.add_argument("--sin-bandeja", action="store_true", help="No mostrar icono en la bandeja.")
    parser.add_argument("--debug", action="store_true", help="Registro detallado.")
    return parser.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)
    setup_logging(level=logging.DEBUG if args.debug else logging.INFO)
    config = AppConfig(carpeta_origen=args.carpeta) if args.carpeta else AppConfig()

    if args.verificar:
        return cmd_verify(config)
    if args.archivo:
        notifier = SafeNotifier(build_default_notifier().backends, asynchronous=False)
        config = dataclasses.replace(config, esperas_reintento=(5.0, 15.0, 30.0))  # manual: esperas más cortas
        return cmd_single_file(config, args.archivo.expanduser().resolve(), args.simulacion, notifier)
    return run_monitor(config, use_tray=not args.sin_bandeja)


if __name__ == "__main__":
    sys.exit(main())
