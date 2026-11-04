"""Incidente 2026-11-04 y cola de pendientes.

Windows (antivirus / indexador) leyó todos los CSV de Descargas; esas lecturas
disparan eventos de "modificado" y el monitor volvió a subir los CSV viejos,
pisando la hoja con datos antiguos. Estas pruebas garantizan que no se repita.
"""
from __future__ import annotations

import json
import os
import ssl
import time
from datetime import datetime, timedelta
from pathlib import Path

from google.auth.exceptions import TransportError

from test_monitor import (  # noqa: F401  (fixtures reutilizadas)
    RecordingNotifier, api, config, m, processor, sheets, workdir, write_csv,
)


def old_file(folder: Path, name: str, content: str, hours_ago: float) -> Path:
    p = write_csv(folder / name, content)
    t = time.time() - hours_ago * 3600
    os.utime(p, (t, t))
    return p


def simulate_windows_scan(p: Path) -> None:
    """Lo que hace un antivirus/indexador: leer y tocar acceso/atributos, sin cambiar contenido."""
    with open(p, "rb") as fh:
        fh.read()
    st = p.stat()
    os.utime(p, (time.time(), st.st_mtime))
    os.chmod(p, 0o444)
    os.chmod(p, 0o666)


def fortinet_error() -> TransportError:
    try:
        try:
            raise ssl.SSLCertVerificationError(1, "[SSL: CERTIFICATE_VERIFY_FAILED] self-signed certificate")
        except ssl.SSLError as e:
            raise TransportError("HTTPSConnectionPool(host='oauth2.googleapis.com')") from e
    except TransportError as e:
        return e


# ------------------------------------------------------------------- incidente
def test_incident_reads_of_old_files_never_upload(processor, config, api):
    d = config.carpeta_origen
    old = [old_file(d, f"Enrolled Cursos ({i}).csv", f"a\n{i}\n", hours_ago=48 - i) for i in range(1, 6)]
    worker = m.SyncWorker(processor)
    worker.start()
    observer = m.Observer()
    observer.schedule(m.CSVDownloadHandler(worker), str(d), recursive=False)
    observer.start()
    try:
        for p in old:
            simulate_windows_scan(p)
        time.sleep(1.5)
        assert api.calls == [] and processor.notifier.items == []

        # Una descarga nueva de verdad sí se sube, una sola vez…
        tmp = d / "Enrolled Cursos (6).csv.crdownload"
        write_csv(tmp, "a\n6\n")
        tmp.rename(d / "Enrolled Cursos (6).csv")
        deadline = time.time() + 10
        while time.time() < deadline and not processor.notifier.items:
            time.sleep(0.05)
        time.sleep(0.8)
        # …y volver a leerla después tampoco la sube otra vez.
        simulate_windows_scan(d / "Enrolled Cursos (6).csv")
        time.sleep(1.5)
    finally:
        observer.stop()
        observer.join()
        worker.stop()
    assert api.kinds().count("get") == 1, api.kinds()
    assert len(processor.notifier.items) == 1


def test_old_file_ignored_but_can_be_forced(processor, config, api):
    p = old_file(config.carpeta_origen, "Enrolled (1).csv", "a\n1\n", hours_ago=5)
    assert processor.process(p) is None and api.calls == []
    assert processor.process(p, force=True) is not None  # --archivo / bandeja


def test_never_uploads_older_csv_when_newer_exists(processor, config, api):
    d = config.carpeta_origen
    older = write_csv(d / "Enrolled (92).csv", "a\nviejo\n")
    t = time.time()
    os.utime(older, (t - 5, t - 5))
    newer = write_csv(d / "Enrolled (93).csv", "a\nnuevo\n")
    assert processor.process(older) is None and api.calls == []
    assert processor.process(newer) is not None


def test_newer_check_respects_more_specific_rule(processor, config, api):
    d = config.carpeta_origen
    p = write_csv(d / "Enrolled (1).csv", "a\n1\n")
    t = time.time()
    os.utime(p, (t - 5, t - 5))
    write_csv(d / "Enrolled Especial (1).csv", "a\n1\n")  # pertenece a otra regla
    assert processor.process(p) is not None


def test_same_content_never_reuploaded_however_late(processor, config, api):
    p = write_csv(config.carpeta_origen / "Enrolled.csv", "a\n1\n")
    processor.process(p)
    processor.clock = lambda: time.monotonic() + 10_000
    processor.process(p)
    assert api.kinds().count("get") == 1


# ------------------------------------------------------------------ pendientes
def test_fortinet_block_is_transient():
    info = m.describe_error(fortinet_error())
    assert info.transitorio and "bloqueó" in info.titulo


def test_pending_flow_and_auto_retry(processor, config, api):
    p = write_csv(config.carpeta_origen / "Enrolled (1).csv", "a\n1\n")
    api.errors = [fortinet_error()] * 3
    assert processor.process(p) is None
    assert processor.pending.get("Enrolled")["archivo"] == str(p)
    assert len(processor.notifier.items) == 1 and processor.notifier.items[0].level == "warning"

    worker = m.SyncWorker(processor)
    jobs = m.BackgroundJobs(config, processor.sheets, worker, processor.pending, RecordingNotifier())
    worker.start()
    api.errors = [fortinet_error()]
    jobs.retry_pending()
    worker.queue.join()
    assert len(processor.notifier.items) == 1  # sigue bloqueado: no repite aviso
    assert processor.pending.get("Enrolled")["intentos"] == 2
    jobs.retry_pending()  # la red ya responde
    worker.queue.join()
    worker.stop()
    assert processor.pending.get("Enrolled") is None
    assert "tras reintentar" in processor.notifier.items[-1].title


def test_pending_keeps_only_latest_and_newer_download_supersedes(processor, config, api):
    d = config.carpeta_origen
    p1 = write_csv(d / "Enrolled (1).csv", "a\n1\n")
    api.errors = [fortinet_error()] * 3
    processor.process(p1)
    time.sleep(0.02)
    p2 = write_csv(d / "Enrolled (2).csv", "a\n2\n")
    api.errors = [fortinet_error()] * 3
    processor.process(p2)
    assert processor.pending.get("Enrolled")["archivo"] == str(p2)  # reemplazó al (1)
    time.sleep(0.02)
    p3 = write_csv(d / "Enrolled (3).csv", "a\n3\n")
    assert processor.process(p3) is not None
    assert processor.pending.get("Enrolled") is None  # el (2) es más viejo: se descarta

    processor.pending.set("Enrolled", p2, "Red")  # aunque reapareciera…
    before = api.kinds().count("get")
    assert processor.process(p2, force=True, from_pending=True) is None  # …no pisa al (3)
    assert api.kinds().count("get") == before and processor.pending.get("Enrolled") is None


def test_pending_shared_between_processes(config):
    a, b = m.PendingStore(config.ruta_pendientes), m.PendingStore(config.ruta_pendientes)
    a.set("Enrolled", Path("C:/x.csv"), "Red")
    assert len(b) == 1
    b.remove("Enrolled")
    assert len(a) == 0


def test_expired_pending_discarded(processor, config):
    p = write_csv(config.carpeta_origen / "Enrolled.csv", "a\n1\n")
    processor.pending.set("Enrolled", p, "Red")
    data = json.loads(config.ruta_pendientes.read_text(encoding="utf-8"))
    data["Enrolled"]["desde"] = (datetime.now() - timedelta(hours=100)).isoformat()
    config.ruta_pendientes.write_text(json.dumps(data), encoding="utf-8")
    notifier = RecordingNotifier()
    jobs = m.BackgroundJobs(config, processor.sheets, m.SyncWorker(processor), processor.pending, notifier)
    assert jobs.retry_pending() == 0 and len(processor.pending) == 0
    assert "No se pudo subir" in notifier.items[0].title


class FakeCreds:
    def __init__(self, minutes, exc=None):
        self.valid, self.exc, self.refreshes = True, exc, 0
        self.expiry = datetime.utcnow() + timedelta(minutes=minutes)

    def refresh(self, req):
        self.refreshes += 1
        if self.exc:
            raise self.exc
        self.expiry = datetime.utcnow() + timedelta(minutes=60)

    def to_json(self):
        return "{}"


def test_refresh_ahead(sheets, monkeypatch):
    monkeypatch.setattr(m, "Request", lambda: None)
    sheets._creds = FakeCreds(5)
    assert sheets.refresh_if_needed() and sheets._creds.refreshes == 1
    assert sheets.refresh_if_needed() and sheets._creds.refreshes == 1  # ya vigente
    sheets._creds = FakeCreds(5, fortinet_error())
    assert sheets.refresh_if_needed() is False  # bloqueado: no lanza, reintenta luego
