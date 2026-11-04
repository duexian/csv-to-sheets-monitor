"""Pruebas del monitor CSV → Google Sheets.

Ejecutar:  python -m pytest -q
No usan internet ni tocan Google: la API se reemplaza por un doble en memoria.
"""
from __future__ import annotations

import base64
import importlib.util
import logging
import re
import sys
import threading
import time
from pathlib import Path
from types import SimpleNamespace

import httplib2
import pytest
from google.auth.exceptions import RefreshError, TransportError
from googleapiclient.errors import HttpError

ROOT = Path(__file__).resolve().parents[1]
_spec = importlib.util.spec_from_file_location("monitor", ROOT / "csv-to-sheets-monitor.py")
m = importlib.util.module_from_spec(_spec)
sys.modules["monitor"] = m
_spec.loader.exec_module(m)


# ==============================================================================
# Dobles de prueba
# ==============================================================================
class FakeRequest:
    def __init__(self, api, kind, kwargs):
        self.api, self.kind, self.kwargs = api, kind, kwargs

    def execute(self, num_retries=0):
        self.api.calls.append((self.kind, self.kwargs))
        if self.api.errors:
            err = self.api.errors.pop(0)
            if err is not None:
                raise err
        if self.kind == "get":
            return {"sheets": [{"properties": p} for p in self.api.sheets]}
        return {}


class FakeSheetsAPI:
    def __init__(self, sheets=None):
        self.calls = []
        self.errors = []  # excepciones a lanzar en orden (None = éxito)
        self.sheets = sheets or [
            {"title": "Otra", "sheetId": 1, "gridProperties": {"rowCount": 100, "columnCount": 5}},
            {"title": "Enrolled", "sheetId": 42, "gridProperties": {"rowCount": 10, "columnCount": 4}},
        ]

    # API encadenada: service.spreadsheets().values().update(...)
    def spreadsheets(self):
        return self

    def values(self):
        return SimpleNamespace(
            update=lambda **kw: FakeRequest(self, "update", kw),
            clear=lambda **kw: FakeRequest(self, "clear", kw),
        )

    def get(self, **kw):
        return FakeRequest(self, "get", kw)

    def batchUpdate(self, **kw):
        return FakeRequest(self, "batchUpdate", kw)

    def kinds(self):
        return [k for k, _ in self.calls]


class RecordingNotifier:
    def __init__(self):
        self.items = []

    def notify(self, title, message, *, level="info", buttons=(), tag=None):
        self.items.append(SimpleNamespace(title=title, message=message, level=level, buttons=list(buttons), tag=tag))


def http_error(status: int, msg: str = "boom") -> HttpError:
    resp = httplib2.Response({"status": status})
    resp.reason = msg
    return HttpError(resp, f'{{"error": {{"message": "{msg}"}}}}'.encode())


@pytest.fixture
def workdir(tmp_path):
    (tmp_path / "rules.yaml").write_text(
        'Enrolled:\n  spreadsheet_id: "SHEET1"\n  nombre_hoja: "enrolled"\n  poner_nota_a1: true\n'
        'Enrolled Especial:\n  spreadsheet_id: "SHEET2"\n  nombre_hoja: "Otra"\n  poner_nota_a1: false\n',
        encoding="utf-8",
    )
    (tmp_path / "Descargas").mkdir()
    return tmp_path


@pytest.fixture
def config(workdir):
    return m.AppConfig(
        script_dir=workdir,
        carpeta_origen=workdir / "Descargas",
        segundos_estabilidad=0.05,
        segundos_espera_maxima=3,
        reintentos=3,
        esperas_reintento=(0.01, 0.01),
        filas_por_lote=2,
    )


@pytest.fixture
def api():
    return FakeSheetsAPI()


@pytest.fixture
def sheets(config, api):
    svc = m.GoogleSheetsService(config, interactive=False)
    svc._service = api  # sin autenticación real
    return svc


@pytest.fixture
def processor(config, sheets):
    notifier = RecordingNotifier()
    proc = m.FileProcessor(config, m.RulesStore(config.ruta_rules), sheets, notifier, sleep=lambda s: time.sleep(0.01))
    return proc


def write_csv(path: Path, text: str, encoding="utf-8-sig") -> Path:
    path.write_bytes(text.encode(encoding))
    return path


# ==============================================================================
# Reglas
# ==============================================================================
class TestRules:
    def test_valid_rules_sorted_longest_prefix_first(self):
        rules = m.parse_rules({"A": {"spreadsheet_id": "x", "nombre_hoja": "h"},
                               "ABC": {"spreadsheet_id": "y", "nombre_hoja": "h"}})
        assert [r.prefijo for r in rules] == ["ABC", "A"]
        assert rules[0].poner_nota_a1 is True
        assert rules[0].incluir_encabezados is True  # por defecto, compatible con reglas anteriores

    def test_incluir_encabezados_option(self):
        rules = m.parse_rules({"A": {"spreadsheet_id": "x", "nombre_hoja": "h", "incluir_encabezados": False}})
        assert rules[0].incluir_encabezados is False
        with pytest.raises(m.ConfigError, match="incluir_encabezados debe ser true o false"):
            m.parse_rules({"A": {"spreadsheet_id": "x", "nombre_hoja": "h", "incluir_encabezados": "no"}})

    def test_accepts_full_url_as_id(self):
        rules = m.parse_rules({"P": {"spreadsheet_id": "https://docs.google.com/spreadsheets/d/AbC_-123/edit#gid=0",
                                     "nombre_hoja": "h"}})
        assert rules[0].spreadsheet_id == "AbC_-123"

    @pytest.mark.parametrize("data, fragment", [
        (None, "vacío"),
        ([], "vacío"),
        ({"P": "texto"}, "debe tener campos"),
        ({"P": {"nombre_hoja": "h"}}, "falta: spreadsheet_id"),
        ({"P": {"spreadsheet_id": "x", "nombre_hoja": "h", "nombre_hjoa": 1}}, "desconocidos: nombre_hjoa"),
        ({"P": {"spreadsheet_id": "", "nombre_hoja": "h"}}, "spreadsheet_id vacío"),
        ({"P": {"spreadsheet_id": "x", "nombre_hoja": " "}}, "nombre_hoja vacío"),
        ({"P": {"spreadsheet_id": "x", "nombre_hoja": "h", "poner_nota_a1": "si"}}, "true o false"),
    ])
    def test_invalid_rules_give_clear_message(self, data, fragment):
        with pytest.raises(m.ConfigError, match=fragment):
            m.parse_rules(data)

    def test_real_rules_file_is_valid(self):
        rules = m.RulesStore(ROOT / "rules.yaml").get()
        assert {r.prefijo for r in rules} >= {"Enrolled", "Participantes"}

    def test_match_case_insensitive_and_most_specific(self, config):
        store = m.RulesStore(config.ruta_rules)
        assert store.match("enrolled cursos (3).csv").spreadsheet_id == "SHEET1"
        assert store.match("Enrolled Especial.csv").spreadsheet_id == "SHEET2"
        assert store.match("Otro.csv") is None

    def test_hot_reload_and_keeps_last_good_rules_on_error(self, config):
        errors = []
        store = m.RulesStore(config.ruta_rules, on_error=errors.append)
        assert len(store.get()) == 2

        time.sleep(0.02)
        config.ruta_rules.write_text('Nuevo:\n  spreadsheet_id: "Z"\n  nombre_hoja: "h"\n', encoding="utf-8")
        assert [r.prefijo for r in store.get()] == ["Nuevo"]

        time.sleep(0.02)
        config.ruta_rules.write_text("Nuevo: [esto: no: es: valido", encoding="utf-8")
        assert [r.prefijo for r in store.get()] == ["Nuevo"]  # conserva las anteriores
        assert len(errors) == 1
        store.get()
        assert len(errors) == 1  # no avisa dos veces del mismo error

    def test_missing_rules_file(self, tmp_path):
        with pytest.raises(m.ConfigError, match="No se encontró"):
            m.RulesStore(tmp_path / "no.yaml").get()


# ==============================================================================
# CSV y utilidades
# ==============================================================================
class TestCsv:
    def test_utf8_bom_emojis_and_quotes(self, tmp_path):
        p = write_csv(tmp_path / "a.csv", 'Grupo,Participantes,"Inscrito en"\n"2023 cohorte",28,"🆕 Disciplina, Laboral"\n')
        assert m.read_csv_values(p) == [["Grupo", "Participantes", "Inscrito en"],
                                        ["2023 cohorte", "28", "🆕 Disciplina, Laboral"]]

    def test_cp1252_fallback(self, tmp_path):
        p = write_csv(tmp_path / "a.csv", "Nombre,Ciudad\nJosé,Bogotá\n", encoding="cp1252")
        assert m.read_csv_values(p)[1] == ["José", "Bogotá"]

    def test_semicolon_delimiter(self, tmp_path):
        p = write_csv(tmp_path / "a.csv", "a;b;c\n1;2;3\n")
        assert m.read_csv_values(p) == [["a", "b", "c"], ["1", "2", "3"]]

    def test_comma_wins_when_header_has_commas(self, tmp_path):
        p = write_csv(tmp_path / "a.csv", "a,b;x,c\n1,2,3\n")
        assert m.read_csv_values(p)[0] == ["a", "b;x", "c"]

    def test_trailing_empty_rows_removed_and_empty_file(self, tmp_path):
        p = write_csv(tmp_path / "a.csv", "a,b\n1,2\n,\n\n")
        assert m.read_csv_values(p) == [["a", "b"], ["1", "2"]]
        assert m.read_csv_values(write_csv(tmp_path / "e.csv", "﻿\n  \n")) == []

    def test_multiline_cells(self, tmp_path):
        p = write_csv(tmp_path / "a.csv", 'a,b\n"línea1\nlínea2",2\n')
        assert m.read_csv_values(p) == [["a", "b"], ["línea1\nlínea2", "2"]]

    @pytest.mark.parametrize("n, letter", [(1, "A"), (26, "Z"), (27, "AA"), (52, "AZ"), (703, "AAA")])
    def test_col_to_letter(self, n, letter):
        assert m.col_to_letter(n) == letter

    def test_col_to_letter_invalid(self):
        with pytest.raises(ValueError):
            m.col_to_letter(0)

    def test_quote_sheet_name_escapes_apostrophe(self):
        assert m.quote_sheet_name("Lista d'Alumnos") == "'Lista d''Alumnos'"

    def test_real_downloads_sample_if_present(self):
        downloads = m.get_downloads_folder()
        files = sorted(downloads.glob("Enrolled*.csv"), key=lambda p: p.stat().st_mtime)[-1:] if downloads.is_dir() else []
        if not files:
            pytest.skip("No hay CSV reales en Descargas")
        rows = m.read_csv_values(files[0])
        assert rows and rows[0][0] == "Grupo"
        assert all(len(r) == len(rows[0]) for r in rows)


# ==============================================================================
# Google Sheets (con API falsa)
# ==============================================================================
class TestSheetsService:
    def rule(self, **kw):
        base = dict(prefijo="Enrolled", spreadsheet_id="SHEET1", nombre_hoja="ENROLLED", poner_nota_a1=True)
        base.update(kw)
        return m.FileRule(**base)

    def test_full_upload_flow(self, tmp_path, sheets, api):
        p = write_csv(tmp_path / "Enrolled (1).csv", "a,b,c\n1,2\n007,=SUMA(1),3\n")
        res = sheets.update_sheet_with_csv(p, self.rule())

        assert res.hoja == "Enrolled" and res.filas == 3 and res.columnas == 3 and res.ultima_columna == "C"
        assert res.url.endswith("#gid=42")
        assert api.kinds() == ["get", "update", "update", "clear", "batchUpdate"]

        updates = [kw for k, kw in api.calls if k == "update"]
        assert updates[0]["range"] == "'Enrolled'!A1" and updates[1]["range"] == "'Enrolled'!A3"
        assert all(u["valueInputOption"] == "RAW" for u in updates)
        assert updates[0]["body"]["values"][1] == ["1", "2", ""]  # fila corta rellenada
        assert updates[1]["body"]["values"][0] == ["007", "=SUMA(1)", "3"]  # se guarda como texto

        clear = dict(api.calls)["clear"]
        assert clear["range"] == "'Enrolled'!A4:C"  # solo filas sobrantes y columnas del CSV

        note = api.calls[-1][1]["body"]["requests"][0]["updateCells"]
        assert note["range"]["sheetId"] == 42
        assert "Enrolled (1).csv" in note["rows"][0]["values"][0]["note"]

    def test_no_note_and_grid_expansion(self, tmp_path, sheets, api):
        rows = "\n".join(",".join(str(c) for c in range(7)) for _ in range(12))
        p = write_csv(tmp_path / "x.csv", rows + "\n")
        sheets.update_sheet_with_csv(p, self.rule(poner_nota_a1=False))

        grow = api.calls[1]
        assert grow[0] == "batchUpdate"
        reqs = grow[1]["body"]["requests"]
        assert {"appendDimension": {"sheetId": 42, "dimension": "ROWS", "length": 2}} in reqs
        assert {"appendDimension": {"sheetId": 42, "dimension": "COLUMNS", "length": 3}} in reqs
        assert "clear" not in api.kinds()  # la hoja quedó exactamente del tamaño del CSV
        assert api.kinds().count("batchUpdate") == 1  # sin nota

    def test_empty_csv_makes_no_api_calls(self, tmp_path, sheets, api):
        res = sheets.update_sheet_with_csv(write_csv(tmp_path / "e.csv", ""), self.rule())
        assert res.vacio and api.calls == []

    def test_dry_run_does_not_write(self, tmp_path, sheets, api):
        res = sheets.update_sheet_with_csv(write_csv(tmp_path / "x.csv", "a\n1\n"), self.rule(), dry_run=True)
        assert res.simulacion and api.kinds() == ["get"]
        assert "Simulación" in res.resumen

    def test_sheet_not_found_lists_available(self, tmp_path, sheets):
        with pytest.raises(m.SheetNotFoundError, match="Disponibles: Otra, Enrolled"):
            sheets.update_sheet_with_csv(write_csv(tmp_path / "x.csv", "a\n"), self.rule(nombre_hoja="Nope"))

    def test_resumen_format(self):
        r = m.SyncResult("f.csv", "Hoja", 7130, 3, "C", "u", 1.0)
        assert r.resumen == "Se cargaron 7.129 filas (+ encabezado) en 'Hoja', columnas A–C."
        r = m.SyncResult("f.csv", "Hoja", 7129, 3, "C", "u", 1.0, con_encabezado=False)
        assert r.resumen == "Se cargaron 7.129 filas desde la fila 2 (encabezado de la hoja intacto) en 'Hoja', columnas A–C."

    def test_without_headers_writes_from_row_2(self, tmp_path, sheets, api):
        p = write_csv(tmp_path / "Enrolled.csv", "a,b,c\n1,2,3\n4,5\n6,7,8\n")
        res = sheets.update_sheet_with_csv(p, self.rule(incluir_encabezados=False))

        assert res.filas == 3 and res.con_encabezado is False and res.ultima_columna == "C"
        updates = [kw for k, kw in api.calls if k == "update"]
        # filas_por_lote=2 -> dos lotes: filas 2-3 y fila 4 de la hoja
        assert [u["range"] for u in updates] == ["'Enrolled'!A2", "'Enrolled'!A4"]
        written = [row for u in updates for row in u["body"]["values"]]
        assert written == [["1", "2", "3"], ["4", "5", ""], ["6", "7", "8"]]  # sin la fila "a,b,c"
        assert dict(api.calls)["clear"]["range"] == "'Enrolled'!A5:C"  # el encabezado (fila 1) no se limpia
        note = api.calls[-1][1]["body"]["requests"][0]["updateCells"]["rows"][0]["values"][0]["note"]
        assert "Filas de datos: 3 (hasta la fila 4)" in note

    def test_without_headers_grid_expansion_counts_header_row(self, tmp_path, sheets, api):
        # Hoja de 10 filas: 10 filas de datos + encabezado de la hoja = 11 -> hay que añadir 1
        rows = "h1,h2\n" + "".join(f"{i},{i}\n" for i in range(10))
        sheets.update_sheet_with_csv(write_csv(tmp_path / "x.csv", rows), self.rule(incluir_encabezados=False))
        grow = api.calls[1][1]["body"]["requests"]
        assert grow == [{"appendDimension": {"sheetId": 42, "dimension": "ROWS", "length": 1}}]
        assert "clear" not in api.kinds()

    def test_without_headers_only_header_row_is_empty(self, tmp_path, sheets, api):
        res = sheets.update_sheet_with_csv(write_csv(tmp_path / "x.csv", "a,b\n"), self.rule(incluir_encabezados=False))
        assert res.vacio and api.calls == []
        assert "no tiene datos" in res.resumen

    def test_without_headers_dry_run(self, tmp_path, sheets, api):
        res = sheets.update_sheet_with_csv(write_csv(tmp_path / "x.csv", "a\n1\n2\n"),
                                           self.rule(incluir_encabezados=False), dry_run=True)
        assert res.filas == 2 and "Simulación" in res.resumen and "fila 2" in res.resumen
        assert api.kinds() == ["get"]


class FakeCreds:
    def __init__(self, valid=False, refresh_token="r", refresh_exc=None):
        self.valid, self.refresh_token, self.refresh_exc = valid, refresh_token, refresh_exc
        self.expired = not valid
        self.refreshed = False

    def refresh(self, request):
        if self.refresh_exc:
            raise self.refresh_exc
        self.refreshed, self.valid = True, True

    def to_json(self):
        return '{"token": "nuevo"}'


class TestAuth:
    def make(self, config, monkeypatch, creds, interactive=False):
        config.ruta_token.write_text("{}", encoding="utf-8")
        config.ruta_credentials.write_text("{}", encoding="utf-8")
        monkeypatch.setattr(m.Credentials, "from_authorized_user_file", staticmethod(lambda *a, **k: creds))
        return m.GoogleSheetsService(config, interactive=interactive, service_builder=lambda *a, **k: "SERVICE")

    def test_refresh_ok_saves_token(self, config, monkeypatch):
        creds = FakeCreds()
        svc = self.make(config, monkeypatch, creds)
        assert svc.get_service() == "SERVICE"
        assert creds.refreshed
        assert config.ruta_token.read_text(encoding="utf-8") == '{"token": "nuevo"}'

    def test_revoked_token_non_interactive_raises_auth_error(self, config, monkeypatch):
        svc = self.make(config, monkeypatch, FakeCreds(refresh_exc=RefreshError("invalid_grant")))
        with pytest.raises(m.AuthError, match="expiró"):
            svc.get_service()

    def test_revoked_token_interactive_runs_login(self, config, monkeypatch):
        flagged = []
        new_creds = FakeCreds(valid=True)
        flow = SimpleNamespace(run_local_server=lambda **kw: new_creds)
        monkeypatch.setattr(m.InstalledAppFlow, "from_client_secrets_file", staticmethod(lambda *a, **k: flow))
        svc = self.make(config, monkeypatch, FakeCreds(refresh_exc=RefreshError("invalid_grant")), interactive=True)
        svc.on_login_required = lambda: flagged.append(True)
        assert svc.get_service() == "SERVICE"
        assert flagged == [True]

    def test_no_internet_propagates_transport_error(self, config, monkeypatch):
        svc = self.make(config, monkeypatch, FakeCreds(refresh_exc=TransportError("sin red")))
        with pytest.raises(TransportError):
            svc.get_service()
        assert config.ruta_token.read_text(encoding="utf-8") == "{}"  # token intacto

    def test_corrupt_token_and_missing_credentials(self, config, monkeypatch):
        def broken(*a, **k):
            raise ValueError("bad json")
        monkeypatch.setattr(m.Credentials, "from_authorized_user_file", staticmethod(broken))
        config.ruta_token.write_text("{corrupto", encoding="utf-8")
        svc = m.GoogleSheetsService(config, interactive=False)
        with pytest.raises(m.ConfigError, match="credentials.json"):
            svc.get_service()


# ==============================================================================
# Mensajes de error
# ==============================================================================
class TestDescribeError:
    @pytest.mark.parametrize("err, titulo, transitorio", [
        (http_error(403), "Sin permiso", False),
        (http_error(404), "Libro no encontrado", False),
        (http_error(429), "Límite", True),
        (http_error(503), "no responde", True),
        (TransportError("x"), "Sin conexión", True),
        (ConnectionResetError(), "Sin conexión", True),
        (TimeoutError(), "Sin conexión", True),
        (PermissionError(), "Archivo en uso", True),
        (FileNotFoundError(), "Archivo no encontrado", False),
        (m.SheetNotFoundError("x"), "Pestaña no encontrada", False),
        (m.AuthError("x"), "Inicia sesión", False),
        (m.ConfigError("x"), "configuración", False),
        (UnicodeDecodeError("utf-8", b"\xff", 0, 1, "bad"), "CSV ilegible", False),
        (KeyError("k"), "Error inesperado (KeyError)", False),
    ])
    def test_mapping(self, err, titulo, transitorio):
        info = m.describe_error(err)
        assert titulo in info.titulo
        assert info.transitorio is transitorio

    def test_format_error_details_outside_except(self):
        try:
            raise ValueError("detalle")
        except ValueError as e:
            err = e
        text = m.format_error_details(err)  # fuera del except: antes daba "NoneType: None"
        assert "ValueError" in text and "detalle" in text and "Traceback" in text


# ==============================================================================
# Procesador
# ==============================================================================
class TestProcessor:
    def test_match_filters(self, processor, config):
        d = config.carpeta_origen
        assert processor.match(d / "Enrolled (1).csv")
        assert processor.match(d / "Enrolled (1).CSV")
        assert processor.match(d / "Enrolled (1).csv.crdownload") is None
        assert processor.match(d / "Enrolled.xlsx") is None
        assert processor.match(d / "~$Enrolled.csv") is None
        assert processor.match(d / "Otro.csv") is None

    def test_success_notifies_with_open_sheet_button(self, processor, config, api):
        p = write_csv(config.carpeta_origen / "Enrolled (5).csv", "a,b\n1,2\n")
        res = processor.process(p)
        assert res and res.filas == 2
        n = processor.notifier.items[-1]
        assert n.level == "info" and "Enrolled actualizada" in n.title
        assert n.buttons[0].etiqueta == "Abrir hoja" and "#gid=42" in n.buttons[0].destino
        assert processor.status.ultimo_resultado is res and processor.status.ultimo_error is None

    def test_duplicate_events_processed_once(self, processor, config, api):
        p = write_csv(config.carpeta_origen / "Enrolled.csv", "a\n1\n")
        processor.process(p)
        processor.process(p)
        assert api.kinds().count("get") == 1
        processor.process(p, force=True)  # "Reintentar" desde la bandeja sí vuelve a subir
        assert api.kinds().count("get") == 2

    def test_modified_file_is_processed_again(self, processor, config, api):
        p = write_csv(config.carpeta_origen / "Enrolled.csv", "a\n1\n")
        processor.process(p)
        time.sleep(0.02)
        write_csv(p, "a\n1\n2\n")
        processor.process(p)
        assert api.kinds().count("get") == 2

    def test_transient_errors_are_retried(self, processor, config, api):
        api.errors = [http_error(503), TransportError("red")]
        p = write_csv(config.carpeta_origen / "Enrolled.csv", "a\n1\n")
        assert processor.process(p) is not None
        assert api.kinds().count("get") == 3
        assert processor.notifier.items[-1].level == "info"

    def test_permanent_error_notifies_and_never_raises(self, processor, config, api):
        api.errors = [http_error(403)]
        p = write_csv(config.carpeta_origen / "Enrolled.csv", "a\n1\n")
        assert processor.process(p) is None
        n = processor.notifier.items[-1]
        assert n.level == "error" and "Sin permiso" in n.title and "Enrolled.csv" in n.message
        assert [b.etiqueta for b in n.buttons] == ["Abrir libro", "Ver registro"]
        assert api.kinds().count("get") == 1  # sin reintentos inútiles
        assert processor.status.ultimo_error.titulo == "Sin permiso sobre la hoja"

    def test_retries_exhausted_goes_to_pending(self, processor, config, api):
        api.errors = [http_error(500)] * 5
        p = write_csv(config.carpeta_origen / "Enrolled.csv", "a\n1\n")
        assert processor.process(p) is None
        assert api.kinds().count("get") == 3
        n = processor.notifier.items[-1]
        assert n.level == "warning" and "no responde" in n.title and "reintentará" in n.message
        assert processor.pending.get("Enrolled")["archivo"] == str(p)

    def test_refresh_error_during_run_resets_session(self, processor, config, api):
        config.ruta_token.write_text("{}", encoding="utf-8")
        config.ruta_credentials.write_text("{}", encoding="utf-8")
        api.errors = [RefreshError("invalid_grant")]
        p = write_csv(config.carpeta_origen / "Enrolled.csv", "a\n1\n")
        # tras el reset el servicio intenta re-autenticar (no interactivo) -> AuthError notificado
        assert processor.process(p) is None
        assert not config.ruta_token.exists()
        assert "Inicia sesión" in processor.notifier.items[-1].title

    def test_waits_for_growing_file(self, processor, config, api):
        p = config.carpeta_origen / "Enrolled.csv"
        write_csv(p, "a,b\n")

        def grow():
            for i in range(5):
                time.sleep(0.03)
                with open(p, "a", encoding="utf-8") as fh:
                    fh.write(f"{i},{i}\n")
        t = threading.Thread(target=grow)
        t.start()
        res = processor.process(p)
        t.join()
        assert res.filas == 6  # leyó el archivo completo

    def test_vanished_file_is_ignored(self, processor, config, api):
        assert processor.process(config.carpeta_origen / "Enrolled-borrado.csv") is None
        assert api.calls == [] and processor.notifier.items == []

    def test_paused(self, processor, config, api):
        processor.status.pausado = True
        p = write_csv(config.carpeta_origen / "Enrolled.csv", "a\n")
        assert processor.process(p) is None and api.calls == []

    def test_empty_csv_warns(self, processor, config, api):
        p = write_csv(config.carpeta_origen / "Enrolled.csv", "\n")
        processor.segundos = 0
        # archivo de 4 bytes (BOM + \n) -> válido pero sin datos
        res = processor.process(p)
        assert res.vacio and processor.notifier.items[-1].level == "warning"


# ==============================================================================
# Integración: watchdog real + simulación de descarga de Chrome
# ==============================================================================
def test_end_to_end_chrome_download(processor, config, api):
    worker = m.SyncWorker(processor)
    worker.start()
    observer = m.Observer()
    observer.schedule(m.CSVDownloadHandler(worker), str(config.carpeta_origen), recursive=False)
    observer.start()
    try:
        tmp = config.carpeta_origen / "Enrolled Cursos (7).csv.crdownload"
        write_csv(tmp, "Grupo,Participantes\n")
        time.sleep(0.1)
        with open(tmp, "a", encoding="utf-8") as fh:
            fh.write("g1,28\n")
        tmp.rename(config.carpeta_origen / "Enrolled Cursos (7).csv")
        write_csv(config.carpeta_origen / "No relacionado.csv", "x\n")

        deadline = time.time() + 10
        while time.time() < deadline and not processor.notifier.items:
            time.sleep(0.05)
        time.sleep(0.5)  # por si llegan eventos duplicados
    finally:
        observer.stop()
        observer.join()
        worker.stop()

    assert api.kinds().count("get") == 1, api.kinds()  # subido exactamente una vez
    assert len(processor.notifier.items) == 1 and processor.notifier.items[0].level == "info"


def test_worker_survives_unexpected_exceptions(config):
    class Boom:
        notifier = RecordingNotifier()

        def __init__(self):
            self.n = 0

        def process(self, path, force=False, **kw):
            self.n += 1
            if self.n == 1:
                raise RuntimeError("inesperado")

    proc = Boom()
    worker = m.SyncWorker(proc)
    worker.start()
    worker.submit(Path("a.csv"))
    worker.submit(Path("b.csv"))
    worker.queue.join()
    worker.stop()
    assert proc.n == 2


# ==============================================================================
# Notificaciones
# ==============================================================================
class TestNotifications:
    def test_toast_xml_escapes_and_has_buttons(self, tmp_path):
        icon = tmp_path / "i.png"
        icon.write_bytes(b"x")
        n = m.WindowsToastNotifier(icon=icon)
        xml = n.build_xml('A & <B> "$x"', "línea 1 & $env:USERNAME\nlínea 2",
                          level="error",
                          buttons=[m.NotificationButton("Abrir hoja", "https://x/y?a=1&b=2"),
                                   m.NotificationButton("Ver registro", str(tmp_path / "mi log.txt"))])
        import xml.dom.minidom as minidom
        doc = minidom.parseString(xml)  # XML bien formado
        texts = [t.firstChild.data for t in doc.getElementsByTagName("text")]
        assert texts[0] == 'A & <B> "$x"' and texts[1] == "línea 1 & $env:USERNAME" and texts[2] == "línea 2"
        actions = doc.getElementsByTagName("action")
        assert actions[0].getAttribute("arguments") == "https://x/y?a=1&b=2"
        assert actions[1].getAttribute("arguments").startswith("file:///") and "mi%20log.txt" in actions[1].getAttribute("arguments")
        assert doc.documentElement.getAttribute("duration") == "long"

    def test_long_text_truncated(self):
        xml = m.WindowsToastNotifier(icon=None).build_xml("T" * 200, "M" * 1000)
        assert "T" * 79 + "…" in xml and "M" * 1000 not in xml

    def test_script_is_encoded_and_runner_called(self):
        calls = []
        n = m.WindowsToastNotifier(icon=None, runner=lambda cmd, **kw: calls.append((cmd, kw)) or SimpleNamespace(returncode=0))
        n.notify("Título 🚀", "Mensaje con 'comillas' y $dólar", tag="ok-Enrolled Cursos")
        cmd, kw = calls[0]
        assert cmd[0] == "powershell.exe" and cmd[-2] == "-EncodedCommand"
        script = base64.b64decode(cmd[-1]).decode("utf-16-le")
        xml_b64 = re.search(r"FromBase64String\('([^']+)'\)", script).group(1)
        assert "Título 🚀" in base64.b64decode(xml_b64).decode("utf-8")
        assert "$t.Tag = 'ok-Enrolled_Cursos'" in script

    def test_toast_failure_raises_for_fallback(self):
        n = m.WindowsToastNotifier(icon=None, runner=lambda *a, **k: SimpleNamespace(returncode=1, stderr=b"fallo"))
        with pytest.raises(RuntimeError, match="fallo"):
            n.notify("a", "b")

    def test_safe_notifier_falls_back_and_never_raises(self, caplog):
        class Broken:
            def notify(self, *a, **k):
                raise OSError("sin toasts")
        good = RecordingNotifier()
        safe = m.SafeNotifier([Broken(), good], asynchronous=False)
        safe.notify("t", "m")
        assert good.items[0].title == "t"

        all_broken = m.SafeNotifier([Broken()], asynchronous=False)
        with caplog.at_level(logging.INFO, logger="SheetsAutoSync"):
            all_broken.notify("t2", "m2")  # no lanza
        assert "Notificación no mostrada" in caplog.text

    def test_repeated_errors_are_throttled(self):
        now = [0.0]
        rec = RecordingNotifier()
        safe = m.SafeNotifier([rec], throttle_seconds=60, asynchronous=False, clock=lambda: now[0])
        safe.notify("Error", "x", level="error")
        safe.notify("Error", "x", level="error")
        safe.notify("OK", "y")
        safe.notify("OK", "y")  # los éxitos no se agrupan
        now[0] = 61
        safe.notify("Error", "x", level="error")
        assert [i.title for i in rec.items] == ["Error", "OK", "OK", "Error"]


# ==============================================================================
# Varios
# ==============================================================================
def test_single_instance_lock(tmp_path):
    a = m.SingleInstance(tmp_path / "x.lock")
    b = m.SingleInstance(tmp_path / "x.lock")
    assert a.acquire()
    assert not b.acquire()
    a.release()
    assert b.acquire()
    b.release()


def test_logging_survives_pythonw_and_cp1252(tmp_path, monkeypatch):
    import io
    monkeypatch.setattr(sys, "stderr", io.TextIOWrapper(io.BytesIO(), encoding="cp1252"))
    log = tmp_path / "logs" / "m.log"
    m.setup_logging(log)
    m.logger.info("🚀 emoji ✅")
    sys.stderr.flush()
    assert "🚀 emoji ✅" in log.read_text(encoding="utf-8")

    monkeypatch.setattr(sys, "stderr", None)  # pythonw.exe
    m.setup_logging(log)
    m.logger.info("sin consola")
    assert "sin consola" in log.read_text(encoding="utf-8")
    m.setup_logging(None)


def test_downloads_folder_exists():
    assert m.get_downloads_folder().is_dir()


def test_cli_single_file_dry_run(config, api, monkeypatch, capsys):
    p = write_csv(config.carpeta_origen / "Enrolled.csv", "a,b\n1,2\n")
    monkeypatch.setattr(m.GoogleSheetsService, "get_service", lambda self: api)
    code = m.cmd_single_file(config, p, dry_run=True, notifier=RecordingNotifier())
    assert code == 0
    assert "Simulación" in capsys.readouterr().out
    assert api.kinds() == ["get"]
    assert m.cmd_single_file(config, config.carpeta_origen / "Otro.csv", True, RecordingNotifier()) == 1
