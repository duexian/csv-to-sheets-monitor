"""Opción mover_a_procesados: los CSV subidos con éxito salen de Descargas."""
from __future__ import annotations

import ssl
import time
from pathlib import Path

from google.auth.exceptions import TransportError

from test_monitor import (  # noqa: F401  (fixtures reutilizadas)
    RecordingNotifier, api, config, m, processor, sheets, workdir, write_csv,
)


def procesados(config) -> Path:
    return config.carpeta_origen / "procesados" / "Sheets"


def test_rule_option_parsed_default_true():
    r = m.parse_rules({"A": {"spreadsheet_id": "x", "nombre_hoja": "h"},
                       "B": {"spreadsheet_id": "x", "nombre_hoja": "h", "mover_a_procesados": False}})
    by = {x.prefijo: x for x in r}
    assert by["A"].mover_a_procesados is True and by["B"].mover_a_procesados is False


def test_real_rules_enable_it():
    assert all(r.mover_a_procesados for r in m.RulesStore(Path(m.SCRIPT_DIR) / "rules.yaml").get())


def test_moved_after_successful_upload(processor, config, api):
    p = write_csv(config.carpeta_origen / "Enrolled (5).csv", "a\n1\n")
    assert processor.process(p) is not None
    assert not p.exists() and (procesados(config) / "Enrolled (5).csv").exists()
    assert processor.status.ultimo_archivo == procesados(config) / "Enrolled (5).csv"


def test_not_moved_when_pending(processor, config, api):
    err = TransportError("x")
    err.__cause__ = ssl.SSLCertVerificationError(1, "CERTIFICATE_VERIFY_FAILED")
    api.errors = [err] * 3
    p = write_csv(config.carpeta_origen / "Enrolled.csv", "a\n1\n")
    assert processor.process(p) is None
    assert p.exists() and processor.pending.get("Enrolled")["archivo"] == str(p)


def test_not_moved_on_permanent_error_dry_run_or_empty(processor, config, api):
    from test_monitor import http_error
    api.errors = [http_error(403)]
    p = write_csv(config.carpeta_origen / "Enrolled (1).csv", "a\n1\n")
    processor.process(p)
    assert p.exists()
    p2 = write_csv(config.carpeta_origen / "Enrolled (2).csv", "a\n1\n")
    processor.process(p2, dry_run=True)
    assert p2.exists()
    p3 = write_csv(config.carpeta_origen / "Enrolled (3).csv", "\n")
    processor.process(p3)
    assert p3.exists()


def test_disabled_per_rule(processor, config, api):
    config.ruta_rules.write_text(
        'Enrolled:\n  spreadsheet_id: "S"\n  nombre_hoja: "Enrolled"\n  mover_a_procesados: false\n', encoding="utf-8")
    time.sleep(0.02)
    p = write_csv(config.carpeta_origen / "Enrolled.csv", "a\n1\n")
    assert processor.process(p) is not None and p.exists()


def test_pending_retry_success_moves_file(processor, config, api):
    err = TransportError("x")
    err.__cause__ = ssl.SSLCertVerificationError(1, "CERTIFICATE_VERIFY_FAILED")
    api.errors = [err] * 3
    p = write_csv(config.carpeta_origen / "Enrolled.csv", "a\n1\n")
    processor.process(p)
    assert processor.process(p, force=True, from_pending=True) is not None
    assert not p.exists() and processor.pending.get("Enrolled") is None


def test_events_from_processed_folder_ignored(processor, config):
    worker = m.SyncWorker(processor)
    handler = m.CSVDownloadHandler(worker)
    sub = procesados(config)
    sub.mkdir(parents=True)
    inside = write_csv(sub / "Enrolled (9).csv", "a\n1\n")
    handler._handle(str(inside))
    assert worker.queue.qsize() == 0
    handler._handle(str(write_csv(config.carpeta_origen / "Enrolled (9).csv", "a\n1\n")))
    assert worker.queue.qsize() == 1


def test_collision_in_processed_folder(processor, config, api):
    p = write_csv(config.carpeta_origen / "Enrolled.csv", "a\n1\n")
    processor.process(p)
    time.sleep(0.02)
    p = write_csv(config.carpeta_origen / "Enrolled.csv", "a\n2\n")
    processor.process(p)
    assert (procesados(config) / "Enrolled.csv").read_text(encoding="utf-8-sig") == "a\n1\n"
    assert (procesados(config) / "Enrolled (2).csv").read_text(encoding="utf-8-sig") == "a\n2\n"


def test_older_csv_never_uploaded_after_newer_was_moved_away(processor, config, api):
    """El (93) se subió y se movió a procesados; un (92) más viejo no debe pisarlo,
    ni siquiera tras reiniciar el monitor (la fecha se guarda en disco)."""
    import os
    d = config.carpeta_origen
    old = write_csv(d / "Enrolled (92).csv", "a\nviejo\n")
    t = time.time()
    os.utime(old, (t - 30, t - 30))
    new = write_csv(d / "Enrolled (93).csv", "a\nnuevo\n")
    assert processor.process(new) is not None and not new.exists()

    restarted = m.FileProcessor(config, processor.rules, processor.sheets, RecordingNotifier(),
                                sleep=lambda s: None)
    restarted.started_at = 0  # que el (92) cuente como "nuevo" para forzar la comprobación
    before = api.kinds().count("get")
    assert restarted.process(old) is None
    processor.pending.set("Enrolled", old, "Red")
    assert restarted.process(old, force=True, from_pending=True) is None
    assert api.kinds().count("get") == before and processor.pending.get("Enrolled") is None
    assert restarted.process(old, force=True) is not None  # subida manual explícita: permitida


def test_end_to_end_move_does_not_retrigger(processor, config, api):
    worker = m.SyncWorker(processor)
    worker.start()
    obs = m.Observer()
    obs.schedule(m.CSVDownloadHandler(worker), str(config.carpeta_origen), recursive=False)
    obs.start()
    try:
        tmp = config.carpeta_origen / "Enrolled (7).csv.crdownload"
        write_csv(tmp, "a\n7\n")
        tmp.rename(config.carpeta_origen / "Enrolled (7).csv")
        deadline = time.time() + 10
        while time.time() < deadline and not (procesados(config) / "Enrolled (7).csv").exists():
            time.sleep(0.05)
        time.sleep(1.0)
    finally:
        obs.stop()
        obs.join()
        worker.stop()
    assert api.kinds().count("get") == 1 and len(processor.notifier.items) == 1
