# ============================================================
# Control de Medidas y Pesos - PySide6 (DESKTOP)
# ============================================================
import os
import sys
import glob
import json
import shutil
import logging
import traceback
from datetime import datetime

import openpyxl
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QLineEdit, QPushButton, QScrollArea, QComboBox, QStackedWidget, QFrame,
    QScroller, QFileDialog, QStackedLayout
)
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QPixmap, QPainter, QColor, QIcon

# ============ CONFIGURACIÓN ============
EXCEL_NAME = "Control Medidas y Pesos - FREDGLIN SALINAS.xlsx"
SHEET_NAME = "MEDIDAS Y PESOS"
OPCIONES_ESTADO = ('registrado', 'no registrado', 'sin existencia')

TAGS_FIN_TABLA = [
    "FREDGLIN SALINAS CONTROL",
    "FILTRADO GENERAL",
    "FILTRADO POR MARCA",
    "TOTAL GENERAL",
    "CÓMO USAR",
]
# Match exacto (uppercase) para evitar cortar productos que CONTENGAN estos textos
TAGS_FIN_UPPER = frozenset(t.upper() for t in TAGS_FIN_TABLA)

PAGE_SIZE = 20

# Desktop: usar diálogos nativos del SO (mejor integración). Cambiar a False
# si en alguna distro Linux el diálogo nativo falla.
USE_NATIVE_DIALOGS = True


# ============ DIRECTORIO DE DATOS DE USUARIO (cross-platform) ============
def _user_data_dir():
    if os.name == "nt":
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
    elif sys.platform == "darwin":
        base = os.path.expanduser("~/Library/Application Support")
    else:
        base = os.environ.get("XDG_CONFIG_HOME") or os.path.expanduser("~/.config")
    d = os.path.join(base, "ControlMedidas")
    try:
        os.makedirs(d, exist_ok=True)
    except Exception:
        d = os.getcwd()
    return d


DATA_DIR = _user_data_dir()
LOG_FILE = os.path.join(DATA_DIR, "app_pyside.log")
CONFIG_FILE = os.path.join(DATA_DIR, "app_config.json")

try:
    logging.basicConfig(
        filename=LOG_FILE,
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
    )
    logging.info("=== INICIO APP PYSIDE6 (desktop) ===")
except Exception:
    pass

# ============ PALETA ============
C_BG      = "#F4F9F4"
C_PRIMARY = "#0F8A2A"
C_CARD    = "#FFFFFF"
C_TEXT    = "#1C1B1F"
C_TEXT2   = "#49454F"
C_CHIP    = "#E8F5E9"
C_SUCCESS = "#43A047"
C_ERROR   = "#B3261E"
C_TINT    = "#E3F2FD"
C_ACCENT  = "#1073B8"


# ============ HELPERS ============
_CARPETAS_CACHE = None

def _carpetas_busqueda():
    """Rutas de búsqueda cross-platform, cacheadas por sesión."""
    global _CARPETAS_CACHE
    if _CARPETAS_CACHE is not None:
        return _CARPETAS_CACHE

    base = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
    home = os.path.expanduser("~")

    carpetas = [
        base,
        os.path.join(home, "Downloads"),
        os.path.join(home, "Documents"),
        os.path.join(home, "Desktop"),
        home,
        os.getcwd(),
    ]
    # En Windows, Downloads/Documents son subcarpetas del perfil (ya cubierto),
    # pero APPDATA puede tener también atajos con archivos copiados.
    if os.name == "nt":
        appdata = os.environ.get("APPDATA")
        if appdata:
            carpetas.append(os.path.join(appdata, "ControlMedidas"))

    vistos = set()
    resultado = []
    for c in carpetas:
        if c and c not in vistos and os.path.isdir(c):
            vistos.add(c)
            resultado.append(c)
    _CARPETAS_CACHE = resultado
    return resultado


def load_config():
    try:
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    except Exception:
        return {}


def save_config(cfg):
    try:
        with open(CONFIG_FILE, "w") as f:
            json.dump(cfg, f, indent=2)
    except Exception:
        pass


def find_excel():
    vistos = set()
    carpetas = []
    for c in _carpetas_busqueda():
        if c and c not in vistos and os.path.isdir(c):
            vistos.add(c)
            carpetas.append(c)

    # 1) Nombre exacto
    for carp in carpetas:
        p = os.path.join(carp, EXCEL_NAME)
        if os.path.isfile(p):
            return p
    # 2) Preferir "FREDGLIN ... CORREGIDO"
    for carp in carpetas:
        for f in glob.glob(os.path.join(carp, "*.xlsx")):
            b = os.path.basename(f)
            if b.startswith("~$"):
                continue
            up = b.upper()
            if "FREDGLIN" in up and "CORREGIDO" in up:
                return f
    # 3) Cualquier FREDGLIN
    for carp in carpetas:
        for f in glob.glob(os.path.join(carp, "*.xlsx")):
            b = os.path.basename(f)
            if b.startswith("~$"):
                continue
            if "FREDGLIN" in b.upper():
                return f
    return None


def carpetas_revisadas_str():
    return "\n".join(f"  • {c}" for c in _carpetas_busqueda()) or "  (ninguna accesible)"


def find_logo():
    """Busca logo de forma flexible (case-insensitive)."""
    nombres_base = ["logo", "insuaminca", "insuamina", "marca"]
    extensiones = [".png", ".jpg", ".jpeg"]
    for carp in _carpetas_busqueda():
        try:
            for archivo in os.listdir(carp):
                if archivo.startswith("~$"):
                    continue
                bajo = archivo.lower()
                nombre_sin_ext, ext = os.path.splitext(bajo)
                if ext in extensiones and any(nb in nombre_sin_ext for nb in nombres_base):
                    return os.path.join(carp, archivo)
        except Exception:
            continue
    return None


def calcular_tipo_y_validacion(unidad, caja):
    u = (unidad or "").strip().lower()
    c = (caja or "").strip().lower()
    if u == "registrado" and c == "registrado":
        return "completado", "OK"
    elif u == "sin existencia" and c == "sin existencia":
        return "sin completar", "sin existencia"
    elif "registrado" in (u, c):
        return "parcial", "revisar"
    else:
        return "sin completar", "revisar"


# ============ WIDGET DE MARCA DE AGUA ============
class WatermarkWidget(QWidget):
    """Capa de fondo con el logo INSUAMINCA (o texto) como marca de agua.

    El pixmap escalado se cachea y solo se recalcula al cambiar tamaño.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.bg_pixmap = None
        self._scaled_cache = None

    def set_watermark(self, pixmap):
        self.bg_pixmap = pixmap
        self._scaled_cache = None
        self.update()

    def resizeEvent(self, event):
        # Invalidar caché al cambiar tamaño
        self._scaled_cache = None
        super().resizeEvent(event)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setRenderHint(QPainter.SmoothPixmapTransform)
        painter.fillRect(self.rect(), QColor(C_BG))

        if self.bg_pixmap and not self.bg_pixmap.isNull():
            if self._scaled_cache is None:
                self._scaled_cache = self.bg_pixmap.scaled(
                    int(self.width() * 0.75), int(self.height() * 0.75),
                    Qt.KeepAspectRatio, Qt.SmoothTransformation
                )
            scaled = self._scaled_cache
            painter.setOpacity(0.15)
            x = (self.width() - scaled.width()) // 2
            y = (self.height() - scaled.height()) // 2 - int(self.height() * 0.05)
            painter.drawPixmap(x, y, scaled)

            painter.setOpacity(0.06)
            font = painter.font()
            font.setPointSize(40)
            font.setBold(True)
            painter.setFont(font)
            painter.setPen(QColor(C_PRIMARY))
            painter.drawText(
                self.rect().adjusted(0, int(self.height() * 0.62), 0, 0),
                Qt.AlignHCenter | Qt.AlignTop,
                "INSUAMINCA"
            )
        else:
            painter.setOpacity(0.08)
            font = painter.font()
            font.setPointSize(52)
            font.setBold(True)
            painter.setFont(font)
            painter.setPen(QColor(C_PRIMARY))
            painter.drawText(self.rect(), Qt.AlignCenter, "INSUAMINCA")


# ============ WIDGETS PERSONALIZADOS ============
class ProductCard(QFrame):
    def __init__(self, prod: str, details: str, parent=None):
        super().__init__(parent)
        self.setObjectName("productCard")
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setAttribute(Qt.WA_Hover, True)
        self.setCursor(Qt.PointingHandCursor)
        self._apply_style(C_CARD)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(12, 8, 12, 8)
        lay.setSpacing(2)

        lbl1 = QLabel(prod)
        lbl1.setWordWrap(True)
        lbl1.setStyleSheet(
            f"color: {C_TEXT}; font-weight: bold; font-size: 13px; background: transparent;"
        )
        lbl2 = QLabel(details)
        lbl2.setWordWrap(True)
        lbl2.setStyleSheet(
            f"color: {C_TEXT2}; font-size: 11px; background: transparent;"
        )
        lay.addWidget(lbl1)
        lay.addWidget(lbl2)
        self.row = None

    def _apply_style(self, bg):
        self.setStyleSheet(
            f"QFrame#productCard {{ background-color: {bg}; border-radius: 10px; }}"
        )

    def enterEvent(self, event):
        self._apply_style(C_TINT)
        super().enterEvent(event)

    def leaveEvent(self, event):
        self._apply_style(C_CARD)
        super().leaveEvent(event)


# ============ VENTANA PRINCIPAL ============
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Control Medidas y Pesos")
        self.resize(900, 850)
        self.setMinimumSize(400, 480)

        # --- Configuración persistente ---
        self.config = load_config()

        # --- Logo ---
        logo_path = self.config.get("logo_path")
        if not logo_path or not os.path.isfile(logo_path):
            logo_path = find_logo()
        self.logo_path = logo_path
        self.logo_pixmap = None
        if logo_path:
            try:
                icon = QIcon(logo_path)
                self.setWindowIcon(icon)
                self.logo_pixmap = QPixmap(logo_path)
                self.config["logo_path"] = logo_path
                save_config(self.config)
            except Exception:
                pass

        # --- Excel ---
        excel_path = self.config.get("excel_path")
        if not excel_path or not os.path.isfile(excel_path):
            excel_path = find_excel()
        self.excel_path = excel_path

        # --- Estado ---
        self.rows = []
        self.filtered = []
        self.selected_row = None
        self.wb = None
        self.ws = None
        self.cols = {}
        self.load_err = None
        self.schema_warning = None
        self._visible_count = 0
        self._last_query = ""
        self._warning_widget = None
        self._more_btn = None
        self._loading_more = False
        self._viewport_press_pos = None

        # --- Contenedor con capa de marca de agua + capa de contenido ---
        central_container = QWidget()
        stack_layout = QStackedLayout(central_container)
        stack_layout.setStackingMode(QStackedLayout.StackAll)

        self.watermark = WatermarkWidget()
        if self.logo_pixmap and not self.logo_pixmap.isNull():
            self.watermark.set_watermark(self.logo_pixmap)
        stack_layout.addWidget(self.watermark)

        content = QWidget()
        content.setStyleSheet("background: transparent;")
        content.setAttribute(Qt.WA_NoSystemBackground, True)
        main_layout = QVBoxLayout(content)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(8)

        self.stacked = QStackedWidget()
        self.stacked.setStyleSheet("QStackedWidget { background: transparent; }")
        self.stacked.setAttribute(Qt.WA_NoSystemBackground, True)
        main_layout.addWidget(self.stacked)

        stack_layout.addWidget(content)
        stack_layout.setCurrentWidget(content)

        self.setCentralWidget(central_container)

        self.search_timer = QTimer(self)
        self.search_timer.setSingleShot(True)
        self.search_timer.setInterval(180)
        self.search_timer.timeout.connect(self._do_search_from_input)

        self._build_list_view()
        self.stacked.addWidget(self.list_widget)
        self.stacked.setCurrentWidget(self.list_widget)

        QTimer.singleShot(20, self._load_data)

    # ---------- VISTA LISTA ----------
    def _build_list_view(self):
        self.list_widget = QWidget()
        self.list_widget.setStyleSheet("background: transparent;")
        self.list_widget.setAttribute(Qt.WA_NoSystemBackground, True)
        v = QVBoxLayout(self.list_widget)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(8)

        # Header
        header = QFrame()
        header.setStyleSheet(f"background-color: {C_PRIMARY}; border-radius: 12px;")
        hv = QVBoxLayout(header)
        hv.setContentsMargins(14, 10, 14, 10)
        hv.setSpacing(2)
        title = QLabel("Control Medidas y Pesos")
        title.setStyleSheet("color: white; font-size: 17px; font-weight: bold; background: transparent;")
        self.count_label = QLabel("Cargando…")
        self.count_label.setStyleSheet(f"color: {C_CHIP}; font-size: 11px; background: transparent;")
        hv.addWidget(title)
        hv.addWidget(self.count_label)
        v.addWidget(header)

        # Fila de botones
        btn_row = QHBoxLayout()
        btn_row.setSpacing(6)

        open_btn = QPushButton("📂 Abrir Excel")
        open_btn.setCursor(Qt.PointingHandCursor)
        open_btn.setStyleSheet(
            f"background-color: {C_ACCENT}; color: white; border-radius: 8px; "
            f"padding: 8px; font-weight: bold; font-size: 11px;"
        )
        open_btn.clicked.connect(self._open_excel_dialog)
        btn_row.addWidget(open_btn)

        reload_btn = QPushButton("🔄 Recargar")
        reload_btn.setCursor(Qt.PointingHandCursor)
        reload_btn.setStyleSheet(
            f"background-color: {C_SUCCESS}; color: white; border-radius: 8px; "
            f"padding: 8px; font-weight: bold; font-size: 11px;"
        )
        reload_btn.clicked.connect(self._reload_excel)
        btn_row.addWidget(reload_btn)

        logo_btn = QPushButton("🖼️ Logo")
        logo_btn.setCursor(Qt.PointingHandCursor)
        logo_btn.setStyleSheet(
            f"background-color: {C_PRIMARY}; color: white; border-radius: 8px; "
            f"padding: 8px; font-weight: bold; font-size: 11px;"
        )
        logo_btn.clicked.connect(self._open_logo_dialog)
        btn_row.addWidget(logo_btn)

        v.addLayout(btn_row)

        # Etiqueta del archivo actual
        self.file_label = QLabel("")
        self.file_label.setWordWrap(True)
        self.file_label.setStyleSheet(
            f"color: {C_TEXT2}; font-size: 10px; background: transparent; padding: 2px;"
        )
        v.addWidget(self.file_label)
        self._update_file_label()

        # Buscador
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Buscar por producto, marca, SKU o código de barras…")
        self.search_input.setStyleSheet(
            f"background-color: {C_CARD}; padding: 10px; border-radius: 8px; "
            f"color: {C_TEXT}; border: none;"
        )
        self.search_input.textChanged.connect(self._on_search)
        v.addWidget(self.search_input)

        # Salir
        exit_btn = QPushButton("SALIR")
        exit_btn.setCursor(Qt.PointingHandCursor)
        exit_btn.setStyleSheet(
            f"background-color: {C_ERROR}; color: white; border-radius: 8px; "
            f"padding: 8px; font-weight: bold;"
        )
        exit_btn.clicked.connect(self.close)
        v.addWidget(exit_btn)

        # Scroll con scroller táctil (solo TouchGesture en desktop;
        # el wheel del ratón ya funciona de forma nativa)
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setStyleSheet(
            "QScrollArea { border: none; background: transparent; }"
            "QScrollArea > QWidget > QWidget { background: transparent; }"
        )
        self.scroll.viewport().setStyleSheet("background: transparent;")
        self.scroll.viewport().setAutoFillBackground(False)
        self.scroll.viewport().setAttribute(Qt.WA_NoSystemBackground, True)
        self.scroll.viewport().setAttribute(Qt.WA_AcceptTouchEvents, True)

        QScroller.grabGesture(self.scroll.viewport(), QScroller.TouchGesture)
        self.scroll.viewport().installEventFilter(self)

        self.list_container = QWidget()
        self.list_container.setStyleSheet("background: transparent;")
        self.list_container.setAttribute(Qt.WA_NoSystemBackground, True)
        self.list_layout = QVBoxLayout(self.list_container)
        self.list_layout.setContentsMargins(0, 0, 0, 0)
        self.list_layout.setSpacing(6)
        self.list_layout.addStretch()
        self.scroll.setWidget(self.list_container)
        v.addWidget(self.scroll)

        self.scroll.verticalScrollBar().valueChanged.connect(self._on_scroll)

    # ---------- FILTRO DE EVENTOS (CLICS EN TARJETAS) ----------
    def eventFilter(self, obj, event):
        if obj is self.scroll.viewport():
            if event.type() == event.Type.MouseButtonPress and event.button() == Qt.LeftButton:
                self._viewport_press_pos = event.position().toPoint()
            elif event.type() == event.Type.MouseButtonRelease and event.button() == Qt.LeftButton:
                if self._viewport_press_pos is not None:
                    release_pos = event.position().toPoint()
                    if (release_pos - self._viewport_press_pos).manhattanLength() < 15:
                        pos_in_container = self.list_container.mapFrom(self.scroll.viewport(), release_pos)
                        widget = self.list_container.childAt(pos_in_container)
                        while widget is not None and not isinstance(widget, ProductCard):
                            widget = widget.parentWidget()
                        if isinstance(widget, ProductCard) and widget.row is not None:
                            self._show_detail(widget.row)
                    self._viewport_press_pos = None
        return super().eventFilter(obj, event)

    # ---------- SELECCIÓN DE ARCHIVOS ----------
    def _update_file_label(self):
        if self.excel_path and os.path.isfile(self.excel_path):
            self.file_label.setText(f"📄 {os.path.basename(self.excel_path)}")
        else:
            self.file_label.setText("⚠️ Sin archivo Excel seleccionado")

    def _start_dir(self):
        if self.excel_path and os.path.isfile(self.excel_path):
            return os.path.dirname(self.excel_path)
        for c in (
            os.path.join(os.path.expanduser("~"), "Downloads"),
            os.path.join(os.path.expanduser("~"), "Documents"),
            os.path.expanduser("~"),
        ):
            if os.path.isdir(c):
                return c
        return os.getcwd()

    def _make_dialog(self, title, name_filter):
        dialog = QFileDialog(self, title)
        dialog.setOption(QFileDialog.DontUseNativeDialog, not USE_NATIVE_DIALOGS)
        dialog.setFileMode(QFileDialog.ExistingFile)
        dialog.setNameFilter(name_filter)
        dialog.setDirectory(self._start_dir())
        return dialog

    def _open_excel_dialog(self):
        dialog = self._make_dialog(
            "Seleccionar archivo Excel",
            "Excel (*.xlsx *.xlsm *.xls);;Todos los archivos (*)"
        )
        if dialog.exec():
            files = dialog.selectedFiles()
            if files:
                self.excel_path = files[0]
                self.config["excel_path"] = self.excel_path
                save_config(self.config)
                self._update_file_label()
                self._reload_excel()

    def _open_logo_dialog(self):
        dialog = self._make_dialog(
            "Seleccionar logo",
            "Imágenes (*.png *.jpg *.jpeg);;Todos los archivos (*)"
        )
        if dialog.exec():
            files = dialog.selectedFiles()
            if files:
                self.logo_path = files[0]
                self.config["logo_path"] = self.logo_path
                save_config(self.config)
                pix = QPixmap(self.logo_path)
                if not pix.isNull():
                    self.logo_pixmap = pix
                    self.watermark.set_watermark(pix)
                    try:
                        self.setWindowIcon(QIcon(self.logo_path))
                    except Exception:
                        pass

    def _reload_excel(self):
        # Detener búsqueda pendiente y limpiar caja de texto
        self.search_timer.stop()
        self.search_input.blockSignals(True)
        self.search_input.clear()
        self.search_input.blockSignals(False)

        self.load_err = None
        self.schema_warning = None
        self._visible_count = 0
        self._last_query = ""
        self.selected_row = None
        self.rows = []
        self.filtered = []
        self.wb = None
        self.ws = None
        self.cols = {}
        self._update_file_label()
        self._load_data()

    # ---------- CARGA ----------
    def _load_data(self):
        try:
            if self.excel_path is None or not os.path.isfile(self.excel_path):
                self.load_err = (
                    "No hay ningún archivo Excel cargado.\n\n"
                    "Usa el botón '📂 Abrir Excel' para seleccionar uno, o\n"
                    "colócalo en cualquiera de estas carpetas:\n"
                    f"{carpetas_revisadas_str()}\n\n"
                    "El archivo debe tener extensión .xlsx (o .xlsm).\n"
                    "La app detecta automáticamente las columnas Producto,\n"
                    "Marca, SKU, Código de barras, Unidad y Caja."
                )
            else:
                try:
                    wb = openpyxl.load_workbook(self.excel_path, data_only=False)
                except PermissionError:
                    raise RuntimeError(
                        "El archivo Excel está abierto en otra aplicación.\n"
                        "Ciérralo e intenta de nuevo."
                    )
                if SHEET_NAME in wb.sheetnames:
                    ws = wb[SHEET_NAME]
                else:
                    ws = wb.active
                cols = self._map_columns(ws)
                if not cols.get("sku") and not cols.get("barras"):
                    self.schema_warning = (
                        "Aviso: no se detectaron columnas SKU ni Código de barras.\n"
                        "Agrega esos headers en la fila 1 del Excel para editarlos."
                    )
                rows = self._build_rows(ws, cols)
                self.wb = wb
                self.ws = ws
                self.cols = cols
                self.rows = rows
                self.filtered = rows[:]
                self.config["excel_path"] = self.excel_path
                save_config(self.config)
        except Exception as e:
            self.load_err = f"Error al leer Excel:\n{e}"
            logging.exception("Error al leer Excel")

        self.count_label.setText(f"{len(self.rows)} productos disponibles")

        if self.load_err:
            self._clear_all_widgets()
            err_lbl = QLabel(self.load_err)
            err_lbl.setWordWrap(True)
            err_lbl.setStyleSheet(
                f"background-color: #F9DEDC; color: {C_ERROR}; "
                f"padding: 12px; border-radius: 8px;"
            )
            self.list_layout.addWidget(err_lbl)
            self.list_layout.addStretch()
            return

        self._render_list(reset=True)

    def _map_columns(self, ws):
        m = {}
        for col in range(1, 25):
            v = ws.cell(row=1, column=col).value
            if not v:
                continue
            h = str(v).strip().lower()
            if ("producto" in h or h == "nombre" or "descrip" in h) and "producto" not in m:
                m["producto"] = col
            elif ("marca" in h or "laboratorio" in h) and "marca" not in m:
                m["marca"] = col
            elif h == "sku" and "sku" not in m:
                m["sku"] = col
            elif ("barras" in h or "barcode" in h or "ean" in h) and "barras" not in m:
                m["barras"] = col
            elif "unidad" in h and "unidad" not in m:
                m["unidad"] = col
            elif "caja" in h and "caja" not in m:
                m["caja"] = col
            elif "tipo" in h and "tipo" not in m:
                m["tipo"] = col
            elif "valid" in h and "validacion" not in m:
                m["validacion"] = col
        m.setdefault("producto", 1)
        m.setdefault("marca", 2)
        m.setdefault("unidad", 5)
        m.setdefault("caja", 6)
        return m

    def _build_rows(self, ws, cols):
        rows = []
        cp = cols["producto"]
        cm = cols.get("marca")
        cs = cols.get("sku")
        cb = cols.get("barras")
        for idx, row_vals in enumerate(ws.iter_rows(values_only=True), start=1):
            if idx == 1:
                continue
            v = row_vals[cp - 1] if cp - 1 < len(row_vals) else None
            if v is None:
                continue
            nombre = str(v).strip()
            if not nombre:
                continue
            # Match exacto (uppercase) para evitar falsos positivos
            if nombre.upper() in TAGS_FIN_UPPER:
                break
            marca = row_vals[cm - 1] if cm and cm - 1 < len(row_vals) else None
            sku = row_vals[cs - 1] if cs and cs - 1 < len(row_vals) else None
            barras = row_vals[cb - 1] if cb and cb - 1 < len(row_vals) else None
            rows.append(
                (idx, nombre, str(marca or "").strip(),
                 str(sku or "").strip(), str(barras or "").strip())
            )
        return rows

    # ---------- RENDER INCREMENTAL ----------
    def _clear_all_widgets(self):
        while self.list_layout.count():
            item = self.list_layout.takeAt(0)
            w = item.widget()
            if w is not None:
                w.deleteLater()
        self._warning_widget = None
        self._more_btn = None
        self._visible_count = 0

    def _render_list(self, reset=True):
        if reset:
            self._rebuild_list()
        else:
            self._on_load_more()

    def _rebuild_list(self):
        self._clear_all_widgets()
        if self.schema_warning:
            self._warning_widget = QLabel(self.schema_warning)
            self._warning_widget.setWordWrap(True)
            self._warning_widget.setStyleSheet(
                "background-color: #FFF3CD; color: #7A5D00; padding: 8px; "
                "border-radius: 8px; font-size: 11px;"
            )
            self.list_layout.addWidget(self._warning_widget)
        self._append_page()

    def _append_page(self):
        if self._more_btn is not None:
            self.list_layout.removeWidget(self._more_btn)
            self._more_btn.deleteLater()
            self._more_btn = None
        if self.list_layout.count() > 0:
            last = self.list_layout.itemAt(self.list_layout.count() - 1)
            if last is not None and last.spacerItem() is not None:
                self.list_layout.takeAt(self.list_layout.count() - 1)

        total = len(self.filtered)
        if total == 0:
            already = any(
                isinstance(self.list_layout.itemAt(i).widget(), QLabel)
                and self.list_layout.itemAt(i).widget().text() == "Sin coincidencias"
                for i in range(self.list_layout.count())
            )
            if not already:
                lbl = QLabel("Sin coincidencias")
                lbl.setAlignment(Qt.AlignCenter)
                lbl.setStyleSheet(f"color: {C_TEXT2}; padding: 20px; font-size: 13px;")
                self.list_layout.addWidget(lbl)
            self.list_layout.addStretch()
            return

        start = self._visible_count
        end = min(start + PAGE_SIZE, total)
        for r, prod, marca, sku, barras in self.filtered[start:end]:
            marca_txt = marca if marca else "S/M"
            extras = []
            if sku:
                extras.append(f"SKU: {sku}")
            if barras:
                extras.append(f"Barras: {barras}")
            linea2 = f"Marca: {marca_txt}"
            if extras:
                linea2 += "  |  " + "  |  ".join(extras)
            card = ProductCard(prod, linea2)
            card.row = r
            self.list_layout.addWidget(card)

        self._visible_count = end

        if end < total:
            restantes = total - end
            self._more_btn = QPushButton(f"Cargar más ({restantes} restantes)")
            self._more_btn.setCursor(Qt.PointingHandCursor)
            self._more_btn.setStyleSheet(
                f"background-color: {C_ACCENT}; color: white; border-radius: 8px; "
                f"padding: 10px; font-weight: bold;"
            )
            self._more_btn.clicked.connect(self._on_load_more)
            self.list_layout.addWidget(self._more_btn)

        self.list_layout.addStretch()

    def _on_load_more(self):
        if self._loading_more:
            return
        if self._visible_count >= len(self.filtered):
            return
        self._loading_more = True
        try:
            self._append_page()
        finally:
            self._loading_more = False

    def _on_scroll(self, value):
        if self.stacked.currentWidget() is not self.list_widget:
            return
        if self._loading_more:
            return
        bar = self.scroll.verticalScrollBar()
        if bar.maximum() > 0 and value >= bar.maximum() - 30:
            if self._visible_count < len(self.filtered):
                QTimer.singleShot(100, self._on_load_more)

    # ---------- BÚSQUEDA ----------
    def _on_search(self, _text):
        self.search_timer.stop()
        self.search_timer.start()

    def _do_search_from_input(self):
        self._refresh_filter()
        self._render_list(reset=True)
        self.scroll.verticalScrollBar().setValue(0)

    def _refresh_filter(self):
        """Único punto que sincroniza _last_query con el input y refiltra."""
        self._last_query = self.search_input.text()
        self._apply_filter()

    def _apply_filter(self):
        q = self._last_query.strip().lower()
        if not q:
            self.filtered = self.rows[:]
        else:
            self.filtered = [
                r for r in self.rows
                if q in r[1].lower() or q in r[2].lower()
                or q in r[3].lower() or q in r[4].lower()
            ]

    # ---------- DETALLE ----------
    def _show_list(self):
        self.selected_row = None
        self.stacked.setCurrentWidget(self.list_widget)

    def _show_detail(self, row):
        self.selected_row = row
        while self.stacked.count() > 1:
            w = self.stacked.widget(1)
            self.stacked.removeWidget(w)
            w.deleteLater()
        detail = self._build_detail_widget(row)
        self.stacked.addWidget(detail)
        self.stacked.setCurrentWidget(detail)

    def _build_detail_widget(self, row):
        c = self.cols
        prod = self.ws.cell(row=row, column=c["producto"]).value or ""
        marca = (self.ws.cell(row=row, column=c["marca"]).value or "") if c.get("marca") else ""
        sku = (self.ws.cell(row=row, column=c["sku"]).value or "") if c.get("sku") else ""
        barras = (self.ws.cell(row=row, column=c["barras"]).value or "") if c.get("barras") else ""
        u = self.ws.cell(row=row, column=c["unidad"]).value
        ca = self.ws.cell(row=row, column=c["caja"]).value

        w = QWidget()
        w.setStyleSheet("background: transparent;")
        v = QVBoxLayout(w)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(8)

        header = QFrame()
        header.setStyleSheet(f"background-color: {C_PRIMARY}; border-radius: 12px;")
        hv = QHBoxLayout(header)
        hv.setContentsMargins(8, 8, 8, 8)
        back = QPushButton("ATRÁS")
        back.setCursor(Qt.PointingHandCursor)
        back.setStyleSheet(
            f"background-color: {C_ACCENT}; color: white; border-radius: 8px; "
            f"padding: 8px 12px; font-weight: bold;"
        )
        back.clicked.connect(self._show_list)
        hv.addWidget(back)
        title = QLabel("Editar Producto")
        title.setStyleSheet("color: white; font-size: 15px; font-weight: bold; background: transparent;")
        hv.addWidget(title)
        hv.addStretch()
        v.addWidget(header)

        info = QFrame()
        info.setStyleSheet(f"background-color: {C_TINT}; border-radius: 10px;")
        iv = QVBoxLayout(info)
        iv.setContentsMargins(12, 8, 12, 8)
        i1 = QLabel(f"Fila {row} | Marca: {marca if marca else 'S/M'}")
        i1.setStyleSheet(f"color: {C_ACCENT}; font-weight: bold; font-size: 11px; background: transparent;")
        i2 = QLabel(str(prod))
        i2.setWordWrap(True)
        i2.setStyleSheet(f"color: {C_TEXT}; font-size: 13px; background: transparent;")
        iv.addWidget(i1)
        iv.addWidget(i2)
        v.addWidget(info)

        self.sku_input = None
        if c.get("sku"):
            v.addWidget(self._section_label("SKU"))
            self.sku_input = QLineEdit(str(sku))
            self.sku_input.setStyleSheet(
                f"background-color: {C_CHIP}; padding: 8px; border-radius: 8px; "
                f"color: {C_TEXT}; border: none;"
            )
            v.addWidget(self.sku_input)

        self.barras_input = None
        if c.get("barras"):
            v.addWidget(self._section_label("CÓDIGO DE BARRAS"))
            self.barras_input = QLineEdit(str(barras))
            self.barras_input.setStyleSheet(
                f"background-color: {C_CHIP}; padding: 8px; border-radius: 8px; "
                f"color: {C_TEXT}; border: none;"
            )
            v.addWidget(self.barras_input)

        v.addWidget(self._section_label("UNIDAD"))
        self.unidad_combo = self._mk_combo(u)
        v.addWidget(self.unidad_combo)

        v.addWidget(self._section_label("CAJA"))
        self.caja_combo = self._mk_combo(ca)
        v.addWidget(self.caja_combo)

        save = QPushButton("GUARDAR CAMBIOS")
        save.setCursor(Qt.PointingHandCursor)
        save.setStyleSheet(
            f"background-color: {C_SUCCESS}; color: white; border-radius: 8px; "
            f"padding: 12px; font-weight: bold; font-size: 13px;"
        )
        save.clicked.connect(self._save_detail)
        v.addWidget(save)

        self.msg = QLabel("")
        self.msg.setWordWrap(True)
        self.msg.setStyleSheet(f"color: {C_SUCCESS}; font-size: 12px; background: transparent;")
        v.addWidget(self.msg)
        v.addStretch()

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("QScrollArea { border: none; background: transparent; }")
        scroll.viewport().setStyleSheet("background: transparent;")
        scroll.viewport().setAutoFillBackground(False)
        scroll.viewport().setAttribute(Qt.WA_NoSystemBackground, True)
        scroll.viewport().setAttribute(Qt.WA_AcceptTouchEvents, True)
        QScroller.grabGesture(scroll.viewport(), QScroller.TouchGesture)
        scroll.setWidget(w)

        outer = QWidget()
        outer.setStyleSheet("background: transparent;")
        ov = QVBoxLayout(outer)
        ov.setContentsMargins(0, 0, 0, 0)
        ov.addWidget(scroll)
        return outer

    def _section_label(self, txt):
        lbl = QLabel(txt)
        lbl.setStyleSheet(f"color: {C_TEXT2}; font-size: 11px; font-weight: bold; background: transparent;")
        return lbl

    def _mk_combo(self, current):
        combo = QComboBox()
        combo.addItems(OPCIONES_ESTADO)
        val = str(current).strip().lower() if current else OPCIONES_ESTADO[2]
        if val not in OPCIONES_ESTADO:
            val = OPCIONES_ESTADO[2]
        combo.setCurrentIndex(OPCIONES_ESTADO.index(val))
        combo.setStyleSheet(
            f"QComboBox {{ background-color: {C_CHIP}; padding: 8px; "
            f"border-radius: 8px; color: {C_TEXT}; border: none; }}"
            f"QComboBox::drop-down {{ border: none; }}"
            f"QComboBox QAbstractItemView {{ background-color: {C_CARD}; "
            f"color: {C_TEXT}; selection-background-color: {C_TINT}; }}"
        )
        return combo

    # ---------- GUARDADO ----------
    def _backup_if_needed(self):
        """Crea backup diario (una vez por día por archivo, sin bloquear
        backups de días posteriores en la misma sesión)."""
        if not self.excel_path:
            return
        try:
            base, ext = os.path.splitext(self.excel_path)
            hoy = datetime.now().strftime("%Y%m%d")
            backup = f"{base}_backup_{hoy}{ext}"
            if not os.path.exists(backup):
                shutil.copy2(self.excel_path, backup)
                logging.info(f"Backup creado: {backup}")
        except Exception:
            logging.exception("No se pudo crear el backup")

    def _save_detail(self):
        if self.selected_row is None:
            return
        try:
            self._backup_if_needed()
            r = self.selected_row
            c = self.cols
            u_val = self.unidad_combo.currentText()
            c_val = self.caja_combo.currentText()
            tipo, validacion = calcular_tipo_y_validacion(u_val, c_val)

            if c.get("sku") and self.sku_input is not None:
                self.ws.cell(row=r, column=c["sku"]).value = self.sku_input.text().strip()
            if c.get("barras") and self.barras_input is not None:
                self.ws.cell(row=r, column=c["barras"]).value = self.barras_input.text().strip()

            self.ws.cell(row=r, column=c["unidad"]).value = u_val
            self.ws.cell(row=r, column=c["caja"]).value = c_val
            if c.get("tipo"):
                self.ws.cell(row=r, column=c["tipo"]).value = tipo
            if c.get("validacion"):
                self.ws.cell(row=r, column=c["validacion"]).value = validacion

            self.wb.save(self.excel_path)

            nuevo_sku = self.sku_input.text().strip() if self.sku_input else ""
            nuevo_barras = self.barras_input.text().strip() if self.barras_input else ""
            for i, row_tuple in enumerate(self.rows):
                if row_tuple[0] == r:
                    self.rows[i] = (r, row_tuple[1], row_tuple[2], nuevo_sku, nuevo_barras)
                    break

            self._refresh_filter()

            self.msg.setText(f"Fila {r} actualizada")
            self.msg.setStyleSheet(f"color: {C_SUCCESS}; font-size: 12px; background: transparent;")
            QTimer.singleShot(400, self._return_to_list_after_save)

        except PermissionError:
            self.msg.setText("Cierra el Excel en otra app antes de guardar.")
            self.msg.setStyleSheet(f"color: {C_ERROR}; font-size: 12px; background: transparent;")
        except Exception as e:
            self.msg.setText(f"Error: {e}")
            self.msg.setStyleSheet(f"color: {C_ERROR}; font-size: 12px; background: transparent;")
            logging.exception("Error al guardar")

    def _return_to_list_after_save(self):
        self._show_list()
        self._render_list(reset=True)

    # ---------- TECLA ATRÁS ----------
    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape and self.stacked.currentIndex() == 1:
            self._show_list()
            return
        super().keyPressEvent(event)


# ============ ENTRADA ============
if __name__ == "__main__":
    try:
        app = QApplication(sys.argv)
        app.setStyle("Fusion")
        window = MainWindow()
        window.show()
        sys.exit(app.exec())
    except Exception:
        err = traceback.format_exc()
        logging.error(err)
        try:
            with open(os.path.join(DATA_DIR, "error_crash.txt"), "w") as f:
                f.write(err)
        except Exception:
            pass
        raise