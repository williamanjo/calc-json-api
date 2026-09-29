# -*- coding: utf-8 -*-
"""JSON API Importer: imports the arrays of a JSON API response as Calc tables.

Each array in the response becomes a table: title (array path), header (object
keys, with nested objects flattened as "a.b") and one row per item. Table
positions are saved in the document so "Refresh" can rewrite them in place.

Registered as a UNO component (XJobExecutor); the menu entries in Addons.xcu
call it with "service:<SERVICE_NAME>?<command>".
"""

import json
import locale
import os
import re
import ssl
import traceback
import urllib.error
import urllib.request

try:
    import uno
    import unohelper
    from com.sun.star.awt import XActionListener, XItemListener
    from com.sun.star.beans import PropertyValue
    from com.sun.star.task import XJobExecutor
except ImportError:  # outside LibreOffice (tests)
    uno = None

IMPLEMENTATION_NAME = "io.github.williamanjo.CalcJsonApi.Impl"
SERVICE_NAME = "io.github.williamanjo.CalcJsonApi"
CONFIG_PROP = "JsonApiImportConfig"
METHODS = ("GET", "POST")
ROOT = "$"  # path of an array at the root of the response
BOLD = 150.0
# VALUE | DATETIME | STRING | FORMULA: clears content, keeps the user's formatting
CONTENT_FLAGS = 1 | 2 | 4 | 16
MAX_COL_WIDTH = 8000  # 1/100 mm
CHUNK = 64 * 1024
WRITE_CHUNK = 2000  # rows per setDataArray, so the progress bar moves
HEADER_NAME_RE = re.compile(r"[!#$%&'*+.^_`|~0-9A-Za-z-]+")

# Values of headers marked as sensitive live only in memory, never in the file.
# Key: (url, lowercase header name).
_SECRETS = {}


# ---------------------------------------------------------------- i18n

_STRINGS = {
    "title": ("API Data", "Dados da API"),
    "connect_failed": ("Connection failed: {0}", "Falha ao conectar: {0}"),
    "request_failed": ("Request failed: {0}", "Falha na requisição: {0}"),
    "bad_json": ("The response is not valid JSON: {0}", "A resposta não é JSON válido: {0}"),
    "header_no_name": ("A header has a value but no name.", "Há um cabeçalho com valor e sem nome."),
    "header_bad_name": ("Invalid header name: {0}\nUse only letters, digits and - (e.g. X-Tenant).",
                        "Nome de cabeçalho inválido: {0}\n"
                        "Use só letras, números e - (ex.: X-Tenant)."),
    "bad_body": ("The body is not valid JSON: {0}", "O body não é JSON válido: {0}"),
    "bad_url": ("Enter a URL starting with http:// or https://",
                "Informe uma URL começando com http:// ou https://"),
    "root": ("(root)", "(raiz)"),
    "value": ("value", "valor"),
    "empty": ("(empty)", "(vazio)"),
    "sheet": ("data", "dados"),
    "downloading_of": ("Downloading JSON: {0} of {1}", "Baixando JSON: {0} de {1}"),
    "downloading": ("Downloading JSON: {0}", "Baixando JSON: {0}"),
    "connecting": ("Downloading JSON: connecting...", "Baixando JSON: conectando..."),
    "importing": ("Importing tables...", "Importando tabelas..."),
    "importing_rows": ("Importing {0}: {1} of {2} rows", "Importando {0}: {1} de {2} linhas"),
    "dlg_import": ("Import from API", "Importar da API"),
    "dlg_refresh": ("Refresh: enter the sensitive headers",
                    "Atualizar: informe os cabeçalhos sensíveis"),
    "method": ("Method:", "Método:"),
    "url": ("API URL:", "URL da API:"),
    "headers": ("Headers:", "Cabeçalhos:"),
    "add_header": ("+ Header", "+ Cabeçalho"),
    "name": ("Name", "Nome"),
    "header_value": ("Value", "Valor"),
    "sensitive": ("Sensitive", "Sensível"),
    "remove_header": ("Remove header", "Remover cabeçalho"),
    "body": ("Body (JSON):", "Body (JSON):"),
    "hint": ("Mark Sensitive on headers that must not be saved in the document (token, "
             "password). The value is kept in memory only, until LibreOffice is closed.",
             "Marque Sensível nos cabeçalhos que não devem ser salvos no documento "
             "(token, senha). O valor fica só na memória até fechar o LibreOffice."),
    "fetch": ("Fetch", "Buscar"),
    "cancel": ("Cancel", "Cancelar"),
    "dlg_arrays": ("Choose arrays", "Escolher arrays"),
    "arrays_found": ("Arrays found (Ctrl+click to pick several):",
                     "Arrays encontrados (Ctrl+clique para vários):"),
    "items": ("{0}   ({1} items)", "{0}   ({1} itens)"),
    "where": ("Where to place:", "Onde colocar:"),
    "side": ("Side by side, starting at {0}", "Lado a lado, a partir de {0}"),
    "sheets": ("One sheet per array (at A1)", "Uma aba para cada array (em A1)"),
    "import": ("Import", "Importar"),
    "no_fit": ("The table ({0} rows x {1} columns) does not fit in the sheet from this position.",
               "A tabela ({0} linhas x {1} colunas) não cabe na planilha a partir dessa posição."),
    "no_arrays": ("No arrays found in the response.", "Nenhum array encontrado na resposta."),
    "overwrite": ("There is already data where the tables will be placed. Overwrite?",
                  "Já existem dados onde as tabelas serão coladas. Sobrescrever?"),
    "no_tables": ("No tables were imported in this document. Use \"Import from API...\".",
                  "Nenhuma tabela importada neste documento. Use \"Importar da API...\"."),
    "missing_array": ("array not in the response", "array não veio na resposta"),
    "missing_sheet": ("sheet no longer exists", "aba não existe mais"),
    "updated": ("{0} table(s) refreshed.", "{0} tabela(s) atualizada(s)."),
    "problems": ("Problems:", "Problemas:"),
    "no_links": ("This document has no API links.", "Este documento não tem vínculos com API."),
    "remove_links": ("Remove the API links? The data stays in the cells, but \"Refresh\" "
                     "will no longer update it.",
                     "Remover os vínculos com a API? Os dados nas células continuam, "
                     "mas \"Atualizar\" deixa de funcionar para eles."),
    "unexpected": ("Unexpected error:\n{0}", "Erro inesperado:\n{0}"),
}

_LANG = 0  # 0 = English, 1 = Portuguese


def _(key, *args):
    return _STRINGS[key][_LANG].format(*args)


def set_language(tag):
    global _LANG
    _LANG = 1 if tag.lower().replace("_", "-").startswith(("pt", "portuguese")) else 0


def _ui_language(ctx):
    """UI language of LibreOffice; falls back to the OS locale when it is "Default"."""
    try:
        provider = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.configuration.ConfigurationProvider", ctx)
        node = PropertyValue(Name="nodepath", Value="/org.openoffice.Setup/L10N")
        access = provider.createInstanceWithArguments(
            "com.sun.star.configuration.ConfigurationAccess", (node,))
        tag = access.getByName("ooLocale")
        if tag:
            return tag
    except Exception:
        pass
    return locale.getlocale()[0] or os.environ.get("LANG", "")


class ApiError(Exception):
    pass


# ---------------------------------------------------------------- core
# Headers are [name, value, sensitive] lists; in the document, the value of a
# sensitive header is saved as None.

def fetch_json(url, method="GET", headers=(), body="", timeout=60, progress=None):
    """Sends the request and returns the JSON. progress(read, total_or_None) per chunk."""
    data = body.encode("utf-8") if method == "POST" else None
    request = urllib.request.Request(url, data=data, method=method)
    request.add_header("Accept", "application/json")
    request.add_header("User-Agent", "calc-json-api/1.3")
    if data is not None:
        request.add_header("Content-Type", "application/json")
    for name, value, *_rest in headers:
        if value:
            request.add_header(name, value)  # overrides the defaults above
    try:
        with urllib.request.urlopen(request, timeout=timeout,
                                    context=ssl.create_default_context()) as resp:
            length = resp.headers.get("Content-Length", "")
            total = int(length) if length.isdigit() else None
            charset = resp.headers.get_content_charset() or "utf-8"
            chunks, received = [], 0
            while True:
                chunk = resp.read(CHUNK)
                if not chunk:
                    break
                chunks.append(chunk)
                received += len(chunk)
                if progress:
                    progress(received, total)
            raw = b"".join(chunks)
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", "replace")[:300]
        raise ApiError(f"HTTP {e.code} {e.reason}\n{detail}")
    except urllib.error.URLError as e:
        raise ApiError(_("connect_failed", e.reason))
    except (ValueError, OSError) as e:
        raise ApiError(_("request_failed", e))
    try:
        return json.loads(raw.decode(charset, "replace"))
    except ValueError as e:
        raise ApiError(_("bad_json", e))


def clean_headers(rows):
    """Validates the dialog rows. Fully empty rows are dropped."""
    headers = []
    for name, value, sensitive in rows:
        name, value = name.strip(), value.strip()
        if not name and not value:
            continue
        if not name:
            raise ApiError(_("header_no_name"))
        if not HEADER_NAME_RE.fullmatch(name):
            raise ApiError(_("header_bad_name", name))
        headers.append([name, value, bool(sensitive)])
    return headers


def validate_body(method, headers, body):
    if method != "POST" or not body.strip():
        return
    content_type = next((v for n, v, *_rest in headers if n.lower() == "content-type"),
                        "application/json")
    if "json" in content_type.lower():
        try:
            json.loads(body)
        except ValueError as e:
            raise ApiError(_("bad_body", e))


def storable(request):
    """Copy of the request without sensitive values, to save in the document."""
    headers = [[n, None if s else v, s] for n, v, s in request["headers"]]
    return {**request, "headers": headers}


def remember_secrets(request):
    for name, value, sensitive in request["headers"]:
        if sensitive and value:
            _SECRETS[(request["url"], name.lower())] = value


def fill_secrets(stored):
    """Returns (request with secrets from memory, whether any is missing)."""
    headers, missing = [], False
    for name, value, sensitive in stored["headers"]:
        if sensitive and value is None:
            value = _SECRETS.get((stored["url"], name.lower()), "")
            missing = missing or not value
        headers.append([name, value or "", sensitive])
    return {**stored, "headers": headers}, missing


def find_arrays(data, path=""):
    """Lists (path, array) for every array in the JSON. Does not descend into arrays."""
    if isinstance(data, list):
        return [(path or ROOT, data)]
    found = []
    if isinstance(data, dict):
        for key, value in data.items():
            found.extend(find_arrays(value, f"{path}.{key}" if path else str(key)))
    return found


def path_label(path):
    return _("root") if path == ROOT else path


def _flatten(obj, prefix=""):
    out = {}
    for key, value in obj.items():
        name = f"{prefix}.{key}" if prefix else str(key)
        if isinstance(value, dict) and value:
            out.update(_flatten(value, name))
        else:
            out[name] = value
    return out


def cell_value(value):
    if value is None:
        return ""
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, int):
        # above 2^53 Calc's double loses digits (e.g. big IDs)
        return float(value) if abs(value) <= 2 ** 53 else str(value)
    if isinstance(value, float):
        return value
    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False) if value else ""
    return str(value)


def build_table(items):
    """Turns a JSON array into (columns, rows)."""
    if items and all(isinstance(i, dict) for i in items):
        flat = [_flatten(i) for i in items]
        columns = list(dict.fromkeys(k for row in flat for k in row))
        if not columns:
            return [_("empty")], [[""] for _row in flat]
        return columns, [[cell_value(row.get(c)) for c in columns] for row in flat]
    return [_("value")], [[cell_value(i)] for i in items]


def sheet_name_for(path):
    name = re.sub(r"[\[\]*?:/\\']", "_", path_label(path)).strip()[:31]
    return name or _("sheet")


def _size(n):
    return f"{n / 1024:.0f} KB" if n < 1024 * 1024 else f"{n / 1024 / 1024:.1f} MB"


# ---------------------------------------------------------------- UNO

if uno is not None:
    class _Action(unohelper.Base, XActionListener):
        def __init__(self, fn):
            self.fn = fn

        def actionPerformed(self, event):
            self.fn()

        def disposing(self, event):
            pass

    class _Item(unohelper.Base, XItemListener):
        def __init__(self, fn):
            self.fn = fn

        def itemStateChanged(self, event):
            self.fn()

        def disposing(self, event):
            pass


def _msg(doc, text, kind="INFOBOX", buttons=1):
    window = doc.CurrentController.Frame.ContainerWindow
    box = window.getToolkit().createMessageBox(
        window, uno.Enum("com.sun.star.awt.MessageBoxType", kind), buttons, _("title"), text)
    try:
        return box.execute()
    finally:
        box.dispose()


def _confirm(doc, text):
    # BUTTONS_YES_NO = 3, result YES = 2
    return _msg(doc, text, "QUERYBOX", 3) == 2


class _Progress:
    """Progress bar in the document's status bar."""

    def __init__(self, doc, text):
        self.indicator = doc.CurrentController.Frame.createStatusIndicator()
        self.indicator.start(text, 100)

    def download(self, received, total):
        if total:
            self.indicator.setValue(min(100, received * 100 // total))
            self.indicator.setText(_("downloading_of", _size(received), _size(total)))
        else:
            # no Content-Length: the bar cycles just to show activity
            self.indicator.setValue(received // CHUNK % 100)
            self.indicator.setText(_("downloading", _size(received)))

    def update(self, done, total, text):
        self.indicator.setValue(min(100, done * 100 // max(total, 1)))
        self.indicator.setText(text)

    def end(self):
        self.indicator.end()


class _Dialog:
    def __init__(self, ctx, title, width, height):
        self.ctx = ctx
        smgr = ctx.ServiceManager
        self.model = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
        self.model.Title = title
        self.model.Width = width
        self.model.Height = height
        self.dialog = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
        self.dialog.setModel(self.model)

    def add(self, kind, name, x, y, width, height, **props):
        model = self.model.createInstance(f"com.sun.star.awt.UnoControl{kind}Model")
        model.PositionX, model.PositionY = x, y
        model.Width, model.Height = width, height
        for key, value in props.items():
            setattr(model, key, value)
        self.model.insertByName(name, model)
        return model

    def buttons(self, y, ok_label):
        width = self.model.Width
        self.add("Button", "ok", width - 116, y, 52, 14, Label=ok_label,
                 PushButtonType=1, DefaultButton=True)
        self.add("Button", "cancel", width - 58, y, 52, 14, Label=_("cancel"), PushButtonType=2)

    def run(self, on_peer=None):
        toolkit = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.awt.Toolkit", self.ctx)
        self.dialog.createPeer(toolkit, None)
        if on_peer:
            on_peer()  # visibility set before the peer exists is lost when it is created
        return self.dialog.execute() == 1

    def field(self, name):
        return self.model.getByName(name)

    def control(self, name):
        return self.dialog.getControl(name)

    def dispose(self):
        self.dialog.dispose()


def _select(listbox_model, positions):
    uno.invoke(listbox_model, "setPropertyValue",
               ("SelectedItems", uno.Any("[]short", tuple(positions))))


class _RequestDialog:
    """Method, URL, headers (rows added with "+") and body (POST only)."""

    WIDTH = 314
    ROW_HEIGHT = 16
    MAX_HEADERS = 15
    ROW_PARTS = ("hn", "hv", "hs", "hr")  # name, value, sensitive, remove

    def __init__(self, ctx, title, request):
        d = self.d = _Dialog(ctx, title, self.WIDTH, 100)
        d.add("FixedText", "l_method", 6, 6, 50, 10, Label=_("method"))
        method = d.add("ListBox", "method", 6, 16, 50, 12, Dropdown=True, StringItemList=METHODS)
        _select(method, [METHODS.index(request["method"]) if request["method"] in METHODS else 0])
        d.add("FixedText", "l_url", 62, 6, 246, 10, Label=_("url"))
        d.add("Edit", "url", 62, 16, 246, 12, Text=request["url"])
        d.add("FixedText", "l_hdr", 6, 36, 150, 10, Label=_("headers"))
        d.add("Button", "add", 244, 33, 64, 13, Label=_("add_header"))
        d.add("FixedText", "l_name", 6, 49, 90, 9, Label=_("name"))
        d.add("FixedText", "l_value", 100, 49, 136, 9, Label=_("header_value"))
        d.add("FixedText", "l_body", 6, 0, 302, 10, Label=_("body"))
        d.add("Edit", "body", 6, 0, 302, 62, MultiLine=True, VScroll=True, Text=request["body"])
        d.add("FixedText", "hint", 6, 0, 302, 18, MultiLine=True, Label=_("hint"))
        d.buttons(0, _("fetch"))

        self.rows = []
        self._next = 0
        for name, value, sensitive in request["headers"]:
            self.add_row(name, value or "", sensitive)
        if not self.rows:
            self.add_row()
        d.control("add").addActionListener(_Action(self._on_add))
        d.control("method").addItemListener(_Item(self.layout))
        self.layout()

    def method(self):
        selected = self.d.field("method").SelectedItems
        return METHODS[selected[0]] if selected else "GET"

    def add_row(self, name="", value="", sensitive=False):
        d, i = self.d, self._next
        self._next += 1
        d.add("Edit", f"hn{i}", 6, 0, 90, 12, Text=name)
        d.add("Edit", f"hv{i}", 100, 0, 136, 12, Text=value,
              EchoChar=ord("*") if sensitive else 0)
        d.add("CheckBox", f"hs{i}", 240, 0, 50, 12, Label=_("sensitive"),
              State=int(bool(sensitive)))
        d.add("Button", f"hr{i}", 294, 0, 14, 12, Label="−", HelpText=_("remove_header"))
        d.control(f"hs{i}").addItemListener(_Item(lambda: self._on_sensitive(i)))
        d.control(f"hr{i}").addActionListener(_Action(lambda: self._on_remove(i)))
        self.rows.append(i)
        return i

    def _on_add(self):
        if len(self.rows) >= self.MAX_HEADERS:
            return
        i = self.add_row()
        self.layout()
        self.d.control(f"hn{i}").setFocus()

    def _on_sensitive(self, i):
        sensitive = self.d.field(f"hs{i}").State == 1
        self.d.field(f"hv{i}").EchoChar = ord("*") if sensitive else 0

    def _on_remove(self, i):
        # hide instead of removing from the model: the clicked button is one of the row's controls
        self.rows.remove(i)
        for part in self.ROW_PARTS:
            self.d.control(f"{part}{i}").setVisible(False)
        if not self.rows:
            self.add_row()
        self.layout()

    def layout(self):
        d = self.d
        y = 59
        for i in self.rows:
            for part in self.ROW_PARTS:
                d.field(f"{part}{i}").PositionY = y
            y += self.ROW_HEIGHT
        y += 4
        post = self.method() == "POST"
        d.control("l_body").setVisible(post)
        d.control("body").setVisible(post)
        if post:
            d.field("l_body").PositionY = y
            d.field("body").PositionY = y + 10
            y += 78
        d.field("hint").PositionY = y
        y += 22
        d.field("ok").PositionY = d.field("cancel").PositionY = y
        d.field("add").Enabled = len(self.rows) < self.MAX_HEADERS
        d.model.Height = y + 20

    def values(self):
        d = self.d
        return {
            "url": d.field("url").Text.strip(),
            "method": self.method(),
            "headers": [[d.field(f"hn{i}").Text, d.field(f"hv{i}").Text,
                         d.field(f"hs{i}").State == 1] for i in self.rows],
            "body": d.field("body").Text,
        }


def _ask_request(ctx, doc, request=None, title=None):
    """Shows the request dialog. Reopens with what was typed while there are errors."""
    state = request or {"url": "", "method": "GET", "headers": [], "body": ""}
    while True:
        dlg = _RequestDialog(ctx, title or _("dlg_import"), state)
        try:
            if not dlg.d.run(on_peer=dlg.layout):
                return None
            state = dlg.values()
        finally:
            dlg.d.dispose()
        try:
            if not re.match(r"https?://", state["url"], re.I):
                raise ApiError(_("bad_url"))
            headers = clean_headers(state["headers"])
            validate_body(state["method"], headers, state["body"])
        except ApiError as e:
            _msg(doc, str(e), "ERRORBOX")
            continue
        post = state["method"] == "POST"
        return {**state, "headers": headers, "body": state["body"] if post else ""}


def _ask_arrays(ctx, arrays, anchor_name):
    items = tuple(_("items", path_label(path), len(items)) for path, items in arrays)
    dlg = _Dialog(ctx, _("dlg_arrays"), 270, 176)
    dlg.add("FixedText", "l_lst", 6, 6, 258, 10, Label=_("arrays_found"))
    lst = dlg.add("ListBox", "lst", 6, 16, 258, 96, MultiSelection=True, StringItemList=items)
    _select(lst, range(len(items)))
    dlg.add("FixedText", "l_dst", 6, 118, 258, 10, Label=_("where"))
    dlg.add("RadioButton", "side", 12, 128, 252, 10, State=1, Label=_("side", anchor_name))
    dlg.add("RadioButton", "sheets", 12, 140, 252, 10, Label=_("sheets"))
    dlg.buttons(156, _("import"))
    try:
        if not dlg.run():
            return None
        selected = list(dlg.control("lst").getSelectedItemsPos())
        mode = "sheets" if dlg.field("sheets").State == 1 else "side"
    finally:
        dlg.dispose()
    return (selected, mode) if selected else None


def _fetch(doc, request):
    progress = _Progress(doc, _("connecting"))
    try:
        return fetch_json(request["url"], request["method"], request["headers"],
                          request["body"], progress=progress.download)
    finally:
        progress.end()


def _normalize(entry):
    """Upgrades links saved by earlier versions to the current format."""
    if "request" not in entry:  # 1.0: url + one header whose value lived in memory
        header = entry.pop("header", "")
        entry["request"] = {"url": entry.pop("url", ""), "method": "GET",
                            "headers": [[header, None, True]] if header else [], "body": ""}
    # 1.1: [name, value], value None when it was sensitive
    entry["request"]["headers"] = [h if len(h) == 3 else [h[0], h[1], h[1] is None]
                                   for h in entry["request"]["headers"]]
    if entry.get("path") == "(raiz)":  # up to 1.2 the root array was saved translated
        entry["path"] = ROOT
    return entry


def _load_config(doc):
    props = doc.DocumentProperties.UserDefinedProperties
    if not props.getPropertySetInfo().hasPropertyByName(CONFIG_PROP):
        return []
    try:
        return [_normalize(e) for e in json.loads(props.getPropertyValue(CONFIG_PROP))]
    except ValueError:
        return []


def _save_config(doc, config):
    props = doc.DocumentProperties.UserDefinedProperties
    text = json.dumps(config, ensure_ascii=False)
    if props.getPropertySetInfo().hasPropertyByName(CONFIG_PROP):
        props.setPropertyValue(CONFIG_PROP, text)
    else:
        props.addProperty(CONFIG_PROP, 128, text)  # PropertyAttribute.REMOVEABLE
    doc.setModified(True)


def _anchor(doc):
    try:
        addr = doc.CurrentSelection.getRangeAddress()
        return doc.Sheets.getByIndex(addr.Sheet), addr.StartColumn, addr.StartRow
    except AttributeError:
        return doc.CurrentController.ActiveSheet, 0, 0


def _check_fits(sheet, col, row, ncols, nrows):
    if col + ncols > sheet.Columns.Count or row + nrows > sheet.Rows.Count:
        raise ApiError(_("no_fit", nrows, ncols))


def _has_content(sheet, col, row, ncols, nrows):
    rng = sheet.getCellRangeByPosition(col, row, col + ncols - 1, row + nrows - 1)
    return rng.queryContentCells(CONTENT_FLAGS).getCount() > 0


def _write_table(sheet, col, row, title, columns, rows, old=None, fit_columns=False,
                 progress=None):
    """Writes title + header + rows. Returns (height, width).

    progress(rows_written) is called after each block of WRITE_CHUNK rows.
    """
    ncols, height = len(columns), len(rows) + 2
    _check_fits(sheet, col, row, ncols, height)
    if old:
        old_h, old_w = old
        sheet.getCellRangeByPosition(col, row, col + old_w - 1, row + old_h - 1) \
            .clearContents(CONTENT_FLAGS)
    title_cell = sheet.getCellByPosition(col, row)
    title_cell.String = title
    title_cell.CharWeight = BOLD
    block = (tuple(columns),) + tuple(tuple(r) for r in rows)
    for start in range(0, len(block), WRITE_CHUNK):
        part = block[start:start + WRITE_CHUNK]
        top = row + 1 + start
        sheet.getCellRangeByPosition(col, top, col + ncols - 1, top + len(part) - 1) \
            .setDataArray(part)
        if progress:
            progress(start + len(part))
    sheet.getCellRangeByPosition(col, row + 1, col + ncols - 1, row + 1).CharWeight = BOLD
    if fit_columns:
        for i in range(col, col + ncols):
            column = sheet.Columns.getByIndex(i)
            column.OptimalWidth = True
            if column.Width > MAX_COL_WIDTH:
                column.Width = MAX_COL_WIDTH
    return height, ncols


class _WriteProgress:
    """Overall progress, in rows, of writing several tables."""

    def __init__(self, doc, total_rows):
        self.progress = _Progress(doc, _("importing"))
        self.total = total_rows
        self.done = 0

    def table(self, path, nrows):
        base = self.done
        self.done += nrows
        label = path_label(path)

        def report(written):
            self.progress.update(base + written, self.total,
                                 _("importing_rows", label, written, nrows))
        return report

    def end(self):
        self.progress.end()


def _target_sheet(doc, name):
    sheets = doc.Sheets
    if not sheets.hasByName(name):
        sheets.insertNewByName(name, sheets.getCount())
    return sheets.getByName(name)


def importar(ctx, doc):
    config = _load_config(doc)
    last = fill_secrets(config[-1]["request"])[0] if config else None
    request = _ask_request(ctx, doc, last)
    if not request:
        return
    remember_secrets(request)
    try:
        data = _fetch(doc, request)
    except ApiError as e:
        _msg(doc, str(e), "ERRORBOX")
        return

    arrays = find_arrays(data)
    if not arrays:
        _msg(doc, _("no_arrays"), "WARNINGBOX")
        return
    sheet, col, row = _anchor(doc)
    choice = _ask_arrays(ctx, arrays, sheet.getCellByPosition(col, row).AbsoluteName)
    if not choice:
        return
    selected, mode = choice

    # plan every position before writing, to confirm overwriting only once
    plan = []
    for index in selected:
        path, items = arrays[index]
        columns, rows = build_table(items)
        if mode == "sheets":
            target, c, r = _target_sheet(doc, sheet_name_for(path)), 0, 0
        else:
            target, c, r = sheet, col, row
            col += len(columns) + 1
        plan.append((path, columns, rows, target, c, r))

    try:
        for _path, columns, rows, target, c, r in plan:
            _check_fits(target, c, r, len(columns), len(rows) + 2)
    except ApiError as e:
        _msg(doc, str(e), "ERRORBOX")
        return
    if any(_has_content(t, c, r, len(cols), len(rows) + 2) for _p, cols, rows, t, c, r in plan):
        if not _confirm(doc, _("overwrite")):
            return

    saved = storable(request)
    progress = _WriteProgress(doc, sum(len(p[2]) + 1 for p in plan))
    try:
        for path, columns, rows, target, c, r in plan:
            height, width = _write_table(target, c, r, path_label(path), columns, rows,
                                         fit_columns=True,
                                         progress=progress.table(path, len(rows) + 1))
            config = [e for e in config
                      if (e["sheet"], e["col"], e["row"]) != (target.Name, c, r)]
            config.append({"request": saved, "path": path, "sheet": target.Name,
                           "col": c, "row": r, "nrows": height, "ncols": width})
    finally:
        progress.end()
    _save_config(doc, config)
    if mode == "sheets":
        doc.CurrentController.setActiveSheet(plan[0][3])


def atualizar(ctx, doc):
    config = _load_config(doc)
    if not config:
        _msg(doc, _("no_tables"))
        return

    # one request per distinct method/URL/headers/body combination
    responses = {}
    for entry in config:
        key = json.dumps(entry["request"], sort_keys=True)
        if key in responses:
            continue
        request, missing = fill_secrets(entry["request"])
        if missing:
            request = _ask_request(ctx, doc, request, _("dlg_refresh"))
            if not request:
                return
            remember_secrets(request)
        try:
            responses[key] = (storable(request), dict(find_arrays(_fetch(doc, request))))
        except ApiError as e:
            responses[key] = (storable(request), e)

    problems = []
    tables = []
    for entry in config:
        entry["request"], arrays = responses[json.dumps(entry["request"], sort_keys=True)]
        where = f"{entry['sheet']} / {path_label(entry['path'])}"
        if isinstance(arrays, ApiError):
            problems.append(f"{where}: {arrays}")
        elif entry["path"] not in arrays:
            problems.append(f"{where}: {_('missing_array')}")
        elif not doc.Sheets.hasByName(entry["sheet"]):
            problems.append(f"{where}: {_('missing_sheet')}")
        else:
            tables.append((entry, *build_table(arrays[entry["path"]])))

    progress = _WriteProgress(doc, sum(len(t[2]) + 1 for t in tables))
    updated = 0
    try:
        for entry, columns, rows in tables:
            try:
                entry["nrows"], entry["ncols"] = _write_table(
                    doc.Sheets.getByName(entry["sheet"]), entry["col"], entry["row"],
                    path_label(entry["path"]), columns, rows,
                    old=(entry["nrows"], entry["ncols"]),
                    progress=progress.table(entry["path"], len(rows) + 1))
                updated += 1
            except ApiError as e:
                problems.append(f"{entry['sheet']} / {path_label(entry['path'])}: {e}")
    finally:
        progress.end()
    _save_config(doc, config)

    text = _("updated", updated)
    if problems:
        text += f"\n\n{_('problems')}\n" + "\n".join(problems)
    _msg(doc, text, "WARNINGBOX" if problems else "INFOBOX")


def remover_vinculos(ctx, doc):
    props = doc.DocumentProperties.UserDefinedProperties
    if not props.getPropertySetInfo().hasPropertyByName(CONFIG_PROP):
        _msg(doc, _("no_links"))
        return
    if _confirm(doc, _("remove_links")):
        props.removeProperty(CONFIG_PROP)
        doc.setModified(True)


COMMANDS = {"importar": importar, "atualizar": atualizar, "remover_vinculos": remover_vinculos}


if uno is not None:
    class CalcJsonApi(unohelper.Base, XJobExecutor):
        """Entry point for the menu: service:<SERVICE_NAME>?<command>."""

        def __init__(self, ctx):
            self.ctx = ctx

        def trigger(self, command):
            action = COMMANDS.get(command)
            desktop = self.ctx.ServiceManager.createInstanceWithContext(
                "com.sun.star.frame.Desktop", self.ctx)
            doc = desktop.getCurrentComponent()
            if action is None or doc is None or \
                    not doc.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
                return
            set_language(_ui_language(self.ctx))
            try:
                action(self.ctx, doc)
            except Exception:
                _msg(doc, _("unexpected", traceback.format_exc(limit=3)), "ERRORBOX")

    g_ImplementationHelper = unohelper.ImplementationHelper()
    g_ImplementationHelper.addImplementation(CalcJsonApi, IMPLEMENTATION_NAME, (SERVICE_NAME,))
