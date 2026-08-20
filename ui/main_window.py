"""
DataHub v2 — Interface PyQt6 (tema dark).

Substitui a camada tkinter mantendo EXATAMENTE a mesma interface pública usada
por app.py:
    MainWindow(cfg, on_executar, on_parar, on_tratar)
    .log(msg) / .log_tratamento(msg) / .set_rodando(bool)
    .set_tratando(bool) / .mostrar_concluido(msg) / .run()

Todo o fluxo (core/) fica inalterado.
"""
import sys
import calendar
from datetime import datetime, timedelta
from pathlib import Path
from typing import Callable

from PyQt6.QtCore import Qt, pyqtSignal, QDate, QSize, QRectF, QPointF
from PyQt6.QtGui import (
    QFontDatabase, QIcon, QColor, QTextCharFormat,
    QPixmap, QPainter, QPolygonF, QPen, QPalette,
)
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QComboBox, QRadioButton, QButtonGroup,
    QLineEdit, QCheckBox, QPlainTextEdit, QFileDialog, QMessageBox,
    QDateEdit, QFrame, QStackedWidget, QCalendarWidget, QDialog, QScrollArea,
    QSystemTrayIcon, QMenu,
)

from config import AppConfig, save_config

try:
    from core.updater import (
        local_version, check_for_update, REPO,
        prepare_update, launch_updater, is_installed,
    )
except Exception:
    local_version = lambda: "?"
    check_for_update = lambda repo=None: None
    REPO = ""
    prepare_update = lambda u, p=None: ""
    launch_updater = lambda b: None
    is_installed = lambda: False


POLOS = [
    "Todos os polos", "UTN", "UTS", "Campos", "Lagos", "Macaé",
    "Noroeste", "Magé", "Niterói", "São Gonçalo", "Serrana", "Sul",
]


def _resolve(name: str) -> str:
    """Resolve um recurso (ico) tanto no código quanto no bundle PyInstaller."""
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        return str(Path(meipass) / name)
    return name


# ── Temas (QSS) ───────────────────────────────────────────────────────────────

def _dark_theme():
    return {
        "bg": "#0D1013", "panel": "#12161B", "panel2": "#171C22",
        "border": "#262E37", "text": "#EDEFF2", "muted": "#9BA5B0",
        "muted2": "#5C6672", "nav": "#9BA5B0",
        "accent": "#F0A742", "accent_text": "#2A1E0C", "accent_hover": "#f5b45c",
        "btn_bg": "#1E242C", "btn_border": "#262E37", "btn_hover": "#171C22",
        "input_bg": "#0D1013", "sel": "#F0A742", "sel_text": "#2A1E0C",
        "alt": "#0D1013",
        "danger_bg": "#402019", "danger_border": "#E0654F", "danger_text": "#E0654F",
        "disabled_text": "#5C6672", "brand": "#F0A742", "brand_sub": "#5C6672",
        "page_title": "#EDEFF2", "page_sub": "#9BA5B0", "section": "#9BA5B0",
        "ok": "#3FB88A", "plain_bg": "#0D1013", "plain_text": "#9BA5B0",
        "radio_text": "#EDEFF2", "check_text": "#EDEFF2",
        "scroll_handle": "#262E37", "scroll_handle_hover": "#3A4452",
        "nav_active_bg": "#4A3A22",
    }


def _light_theme():
    return {
        "bg": "#ECEFF3", "panel": "#FFFFFF", "panel2": "#F4F6F9",
        "border": "#D2D8E0", "text": "#1A1F26", "muted": "#5A636F",
        "muted2": "#8A93A2", "nav": "#4A5562",
        "accent": "#2D6CDF", "accent_text": "#FFFFFF", "accent_hover": "#4A82E8",
        "btn_bg": "#E4E9F0", "btn_border": "#C9D2DD", "btn_hover": "#D6DDE6",
        "input_bg": "#E4E9F0", "sel": "#2D6CDF", "sel_text": "#FFFFFF",
        "alt": "#E4E9F0",
        "danger_bg": "#FBEAEA", "danger_border": "#E0A4A4", "danger_text": "#C0392B",
        "disabled_text": "#9AA4B2", "brand": "#2D6CDF", "brand_sub": "#8A93A2",
        "page_title": "#1A1F26", "page_sub": "#5A636F", "section": "#5A636F",
        "ok": "#27AE60", "plain_bg": "#FFFFFF", "plain_text": "#1A1F26",
        "radio_text": "#1A1F26", "check_text": "#1A1F26",
        "scroll_handle": "#C2CAD6", "scroll_handle_hover": "#AEB8C6",
        "nav_active_bg": "#E4E9F0",
    }


def _build_style(t: dict) -> str:
    return f"""
QWidget {{
    background-color: {t['panel']};
    color: {t['text']};
    font-family: "Manrope", "Segoe UI", "Inter", sans-serif;
    font-size: 13px;
}}
QWidget#Rail {{ background-color: {t['panel2']}; border: none; }}
QWidget#Content {{ background-color: {t['panel']}; }}
QScrollArea {{ background-color: transparent; border: none; }}
QLabel#Brand {{ font-size: 18px; font-weight: 700; color: {t['brand']}; }}
QLabel#BrandSub {{ font-size: 10px; color: {t['muted2']}; letter-spacing: 2px; }}
QLabel#PageTitle {{ font-size: 20px; font-weight: 700; color: {t['page_title']}; }}
QLabel#PageSub {{ font-size: 12px; color: {t['page_sub']}; }}
QLabel#SectionLabel {{ font-size: 11px; font-weight: 700; color: {t['section']};
    letter-spacing: 1px; }}

QFrame#Card {{
    background-color: {t['panel']};
    border: 1px solid {t['border']};
    border-radius: 12px;
}}
QFrame#Sep {{ background-color: {t['border']}; max-height: 1px; border: none; }}

QPushButton {{
    background-color: {t['btn_bg']}; border: 1px solid {t['btn_border']}; border-radius: 9px;
    padding: 8px 14px; color: {t['text']};
}}
QPushButton:hover {{ background-color: {t['btn_hover']}; border-color: {t['border']}; }}
QPushButton:pressed {{ background-color: {t['panel2']}; }}
QPushButton:disabled {{ color: {t['disabled_text']}; background-color: {t['panel2']}; border-color: {t['border']}; }}

QPushButton#Primary {{ background-color: {t['accent']}; color: {t['accent_text']}; border: none; font-weight: 700; }}
QPushButton#Primary:hover {{ background-color: {t['accent_hover']}; }}
QPushButton#Primary:disabled {{ background-color: {t['btn_border']}; color: {t['disabled_text']}; }}

QPushButton#Danger {{ background-color: {t['danger_bg']}; border: 1px solid {t['danger_border']}; color: {t['danger_text']}; }}
QPushButton#Danger:hover {{ background-color: {t['danger_border']}; }}

QPushButton#Nav {{
    background-color: transparent; border: none; border-radius: 9px;
    text-align: left; padding: 10px 14px; color: {t['nav']}; font-weight: 600;
}}
QPushButton#Nav:hover {{ background-color: {t['panel2']}; color: {t['text']}; }}
QPushButton#Nav[active="true"] {{
    background-color: {t['nav_active_bg']}; color: {t['accent']}; border-left: 3px solid {t['accent']};
}}

QLineEdit, QComboBox {{
    background-color: {t['input_bg']}; border: 1px solid {t['btn_border']}; border-radius: 8px;
    padding: 6px 10px; color: {t['text']};
    selection-background-color: {t['sel']}; selection-color: {t['sel_text']};
}}
QLineEdit:focus, QComboBox:focus {{ border-color: {t['accent']}; }}
QComboBox QAbstractItemView {{
    background-color: {t['input_bg']}; color: {t['text']}; selection-background-color: {t['btn_border']};
    border: 1px solid {t['btn_border']};
}}

QRadioButton {{ color: {t['radio_text']}; spacing: 8px; }}
QRadioButton::indicator {{ width: 16px; height: 16px; border-radius: 8px;
    border: 2px solid {t['btn_border']}; background-color: {t['input_bg']}; }}
QRadioButton::indicator:checked {{ background-color: {t['accent']}; border-color: {t['accent']}; }}

QCheckBox {{ color: {t['check_text']}; spacing: 8px; }}
QCheckBox::indicator {{ width: 16px; height: 16px; border-radius: 4px;
    border: 2px solid {t['btn_border']}; background-color: {t['input_bg']}; }}
QCheckBox::indicator:checked {{ background-color: {t['ok']}; border-color: {t['ok']}; }}

QPlainTextEdit {{
    background-color: {t['plain_bg']}; border: 1px solid {t['border']}; border-radius: 8px;
    color: {t['plain_text']}; font-family: "JetBrains Mono", "Consolas", "Courier New", monospace;
    font-size: 11px; padding: 8px;
}}

QScrollBar:vertical {{ background: transparent; width: 7px; border-radius: 4px;
    margin: 4px 3px 4px 6px; }}
QScrollBar::handle:vertical {{ background: {t['scroll_handle']}; border-radius: 4px; min-height: 28px; }}
QScrollBar::handle:vertical:hover {{ background: {t['scroll_handle_hover']}; }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{ background: transparent; }}

QToolTip {{ background-color: {t['panel']}; color: {t['text']}; border: 1px solid {t['btn_border']}; padding: 4px 8px; }}

QDialog, QMessageBox {{ background-color: {t['panel']}; color: {t['text']}; }}
QMessageBox QLabel {{ color: {t['text']}; background-color: transparent; }}
QMessageBox QAbstractButton {{ min-width: 90px; }}

QFileDialog {{ background-color: {t['bg']}; color: {t['text']}; }}
QFileDialog QAbstractItemView {{ background-color: {t['input_bg']}; color: {t['text']}; }}

QCalendarWidget {{ background-color: {t['bg']}; color: {t['text']}; border: none; }}
QCalendarWidget QWidget#qt_calendar_navigationbar {{ background-color: {t['panel']}; }}
QCalendarWidget QAbstractItemView {{
    background-color: {t['panel']}; color: {t['text']};
    selection-background-color: {t['accent']}; selection-color: {t['accent_text']};
    alternate-background-color: {t['panel']};
}}
QCalendarWidget QMenu {{ background-color: {t['panel']}; color: {t['text']}; }}
QCalendarWidget QSpinBox {{ background-color: {t['input_bg']}; color: {t['text']}; padding: 2px; }}
QCalendarWidget QToolButton {{
    background-color: {t['panel']}; color: {t['text']}; padding: 6px; border-radius: 6px;
}}
QCalendarWidget QToolButton:hover {{ background-color: {t['btn_hover']}; }}
"""


def _apply_palette(app: QApplication, t: dict):
    """Paleta global: garante que widgets nativos (modais, diálogos do SO,
    áreas de texto internas) usem exatamente as cores do tema — não só os
    widgets cobertos pelo QSS. Evita fundos diferentes dentro dos modais."""
    p = QPalette()
    bg = QColor(t["bg"])
    text = QColor(t["text"])
    p.setColor(QPalette.ColorRole.Window, bg)
    p.setColor(QPalette.ColorRole.WindowText, text)
    p.setColor(QPalette.ColorRole.Base, QColor(t["panel"]))
    p.setColor(QPalette.ColorRole.Text, text)
    p.setColor(QPalette.ColorRole.Button, QColor(t["btn_bg"]))
    p.setColor(QPalette.ColorRole.ButtonText, text)
    p.setColor(QPalette.ColorRole.Highlight, QColor(t["sel"]))
    p.setColor(QPalette.ColorRole.HighlightedText, QColor(t["sel_text"]))
    p.setColor(QPalette.ColorRole.Link, QColor(t["accent"]))
    p.setColor(QPalette.ColorRole.ToolTipBase, QColor(t["panel"]))
    p.setColor(QPalette.ColorRole.ToolTipText, text)
    p.setColor(QPalette.ColorRole.PlaceholderText, QColor(t["muted"]))
    app.setPalette(p)


def _symbol_icon(kind: str, color: str = "#1A1206") -> QIcon:
    """Ícone pequeno na cor da fonte do botão (▶ ou ■) para Executar/Parar."""
    pm = QPixmap(16, 16)
    pm.fill(Qt.GlobalColor.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.RenderHint.Antialiasing)
    p.setPen(Qt.PenStyle.NoPen)
    p.setBrush(QColor(color))
    if kind == "play":
        p.drawPolygon(QPolygonF([QPointF(4, 3), QPointF(4, 13), QPointF(13, 8)]))
    else:
        p.drawRect(QRectF(4, 4, 8, 8))
    p.end()
    return QIcon(pm)


def _set_fonts():
    """Tenta registrar Manrope/JetBrains Mono se presentes no bundle; silencioso se não."""
    for fam in ("Manrope", "JetBrains Mono"):
        QFontDatabase.addApplicationFont(fam)


class _DateField(QWidget):
    """Campo de data com seletor estilo calendário (botão 📅 abre popup)."""
    _MESES = [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro",
    ]

    def __init__(self, valor=None, get_accent=lambda: "#F4B740"):
        super().__init__()
        self._date = (valor or datetime.now()).date()
        self._get_accent = get_accent
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(8)
        self._entry = QLineEdit(self._date.strftime("%d/%m/%Y"))
        self._entry.setReadOnly(True)
        lay.addWidget(self._entry, stretch=1)
        btn = QPushButton("📅")
        btn.setFixedWidth(38)
        btn.setFixedHeight(34)
        btn.clicked.connect(self._open)
        lay.addWidget(btn)

    def _open(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("Selecionar data")
        v = QVBoxLayout(dlg)

        hoje = datetime.now()
        cab = QHBoxLayout()
        cab.setSpacing(8)
        prev = QPushButton("‹")
        prev.setFixedWidth(34)
        prev.setFixedHeight(32)
        prox = QPushButton("›")
        prox.setFixedWidth(34)
        prox.setFixedHeight(32)
        mes = QComboBox()
        mes.addItems(self._MESES)
        ano = QComboBox()
        for y in range(hoje.year - 10, hoje.year + 3):
            ano.addItem(str(y))
        cab.addWidget(prev)
        cab.addWidget(mes, stretch=1)
        cab.addWidget(ano, stretch=1)
        cab.addWidget(prox)
        v.addLayout(cab)

        cal = QCalendarWidget()
        cal.setNavigationBarVisible(False)
        cal.setVerticalHeaderFormat(QCalendarWidget.VerticalHeaderFormat.NoVerticalHeader)
        fmt = QTextCharFormat()
        fmt.setForeground(QColor(self._get_accent()))
        cal.setWeekdayTextFormat(Qt.DayOfWeek.Saturday, fmt)
        cal.setWeekdayTextFormat(Qt.DayOfWeek.Sunday, fmt)
        cal.setSelectedDate(QDate(self._date.year, self._date.month, self._date.day))
        v.addWidget(cal)

        def _navegar():
            y = int(ano.currentText())
            m = mes.currentIndex() + 1
            cal.setCurrentPage(y, m)
            ultimo = calendar.monthrange(y, m)[1]
            cal.setSelectedDate(QDate(y, m, min(self._date.day, ultimo)))

        def _sync(y, m):
            mes.blockSignals(True)
            ano.blockSignals(True)
            mes.setCurrentIndex(m - 1)
            ano.setCurrentText(str(y))
            mes.blockSignals(False)
            ano.blockSignals(False)

        mes.currentIndexChanged.connect(_navegar)
        ano.currentTextChanged.connect(_navegar)
        cal.currentPageChanged.connect(_sync)
        prev.clicked.connect(cal.showPreviousMonth)
        prox.clicked.connect(cal.showNextMonth)

        mes.setCurrentIndex(self._date.month - 1)
        ano.setCurrentText(str(self._date.year))

        h = QHBoxLayout()
        h.addStretch(1)
        ok = QPushButton("OK", objectName="Primary")
        ok.setFixedWidth(90)
        ok.clicked.connect(lambda: self._pick(cal.selectedDate(), dlg))
        h.addWidget(ok)
        v.addLayout(h)
        dlg.exec()

    def _pick(self, qd, dlg):
        self._date = qd.toPyDate()
        self._entry.setText(self._date.strftime("%d/%m/%Y"))
        dlg.accept()

    def get_date(self) -> datetime:
        return datetime(self._date.year, self._date.month, self._date.day)


# ── Janela principal ──────────────────────────────────────────────────────────

class MainWindow(QMainWindow):
    log_signal = pyqtSignal(str)
    log_trat_signal = pyqtSignal(str)
    rodando_signal = pyqtSignal(bool)
    tratando_signal = pyqtSignal(bool)
    concluido_signal = pyqtSignal(str)
    formato_signal = pyqtSignal(str)
    operview_login_signal = pyqtSignal(object)
    quit_signal = pyqtSignal()

    def __init__(self, cfg: AppConfig, on_executar: Callable,
                 on_parar: Callable, on_tratar: Callable):
        if QApplication.instance() is None:
            self._app = QApplication(sys.argv)
        else:
            self._app = QApplication.instance()

        super().__init__()
        self.setWindowTitle("DataHub")
        self.setWindowIcon(QIcon(_resolve("appicon.ico")))
        self._init_tray()
        self.setGeometry(100, 60, 620, 620)
        self.setMinimumSize(620, 620)
        self._center()

        self._cfg = cfg
        self._on_executar = on_executar
        self._on_parar = on_parar
        self._on_tratar = on_tratar

        self._tema = _dark_theme()
        self._accent = self._tema["accent"]
        self._claro = False

        self.log_signal.connect(self._append_log_exec)
        self.log_trat_signal.connect(self._append_log_trat)
        self.rodando_signal.connect(self._set_rodando)
        self.tratando_signal.connect(self._set_tratando)
        self.formato_signal.connect(self._set_formato)
        self.operview_login_signal.connect(self._mostrar_aviso_login)
        self.concluido_signal.connect(self._mostrar_concluido)

        self._app.setStyleSheet(_build_style(self._tema))
        _apply_palette(self._app, self._tema)
        _set_fonts()
        self._build()

    # ── Montagem ──────────────────────────────────────────────────────────────

    def _center(self):
        frame = self.frameGeometry()
        screen = QApplication.primaryScreen()
        if screen is not None:
            center = screen.availableGeometry().center()
            frame.moveCenter(center)
            self.move(frame.topLeft())

    def _build(self):
        root = QWidget()
        self.setCentralWidget(root)
        h = QHBoxLayout(root)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(0)

        h.addWidget(self._build_rail())
        h.addWidget(self._build_content(), stretch=1)
        self._goto(0)

        # altura uniforme e legível para TODOS os campos (line/edit/combo/date)
        for campo in self.findChildren((QLineEdit, QComboBox, QDateEdit)):
            campo.setFixedHeight(34)

    def _build_rail(self):
        rail = QWidget(objectName="Rail")
        rail.setFixedWidth(170)
        v = QVBoxLayout(rail)
        v.setContentsMargins(12, 16, 12, 16)
        v.setSpacing(6)

        v.addWidget(QLabel("DataHub", objectName="Brand"))
        v.addWidget(QLabel("EXTRATOR DE DADOS", objectName="BrandSub"))
        v.addSpacing(18)

        self._nav_btns = []
        for icon, text, idx in (
            ("⚡", "Execução", 0),
            ("🛠", "Tratamento", 1),
            ("⚙", "Configurações", 2),
        ):
            b = QPushButton(f"{icon}  {text}")
            b.setObjectName("Nav")
            b.setCheckable(True)
            b.setMinimumHeight(40)
            b.clicked.connect(lambda _=False, i=idx: self._goto(i))
            self._nav_btns.append(b)
            v.addWidget(b)

        v.addStretch(1)

        ver = QLabel(f"v{local_version()}")
        ver.setStyleSheet("color:#5A636F; font-size:11px;")
        v.addWidget(ver)

        return rail

    def _goto(self, idx: int):
        self._stack.setCurrentIndex(idx)
        for i, b in enumerate(self._nav_btns):
            b.setProperty("active", i == idx)
            b.style().polish(b)
        titles = ["Execução", "Tratamento", "Configurações"]
        subs = [
            "Extração e tratamento de incidências",
            "Ferramentas de arquivo",
            "Credenciais e integrações",
        ]
        self._page_title.setText(titles[idx])
        self._page_sub.setText(subs[idx])

    def _build_content(self):
        content = QWidget(objectName="Content")
        v = QVBoxLayout(content)
        v.setContentsMargins(26, 22, 26, 22)
        v.setSpacing(16)

        self._page_title = QLabel(objectName="PageTitle")
        self._page_sub = QLabel(objectName="PageSub")
        v.addWidget(self._page_title)
        v.addWidget(self._page_sub)

        sep = QFrame(objectName="Sep")
        sep.setFrameShape(QFrame.Shape.HLine)
        v.addWidget(sep)

        self._stack = QStackedWidget()
        area = QScrollArea()
        area.setWidgetResizable(True)
        area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        area.setWidget(self._stack)
        v.addWidget(area, stretch=1)

        self._stack.addWidget(self._build_execucao())
        self._stack.addWidget(self._build_tratamento())
        self._stack.addWidget(self._build_configuracoes())
        return content

    @staticmethod
    def _card(widget: QWidget) -> QFrame:
        card = QFrame(objectName="Card")
        card.setFrameShape(QFrame.Shape.StyledPanel)
        lay = QVBoxLayout(card)
        lay.setContentsMargins(16, 16, 16, 16)
        lay.setSpacing(12)
        lay.addWidget(widget)
        return card

    # ── Execução ────────────────────────────────────────────────────────────────

    def _build_execucao(self):
        page = QWidget()
        v = QVBoxLayout(page)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(14)

        peri = QWidget()
        pv = QVBoxLayout(peri)
        pv.setContentsMargins(0, 0, 0, 0)
        pv.setSpacing(10)
        pv.addWidget(QLabel("Período", objectName="SectionLabel"))
        row = QHBoxLayout()
        row.setSpacing(10)
        row.addWidget(QLabel("De"))
        self._data_ini = _DateField(datetime.now() - timedelta(days=6),
                                    get_accent=lambda: self._accent)
        row.addWidget(self._data_ini)
        row.addWidget(QLabel("Até"))
        self._data_fim = _DateField(datetime.now(),
                                    get_accent=lambda: self._accent)
        row.addWidget(self._data_fim)
        row.addStretch(1)
        pv.addLayout(row)
        v.addWidget(self._card(peri))

        polow = QWidget()
        pw = QVBoxLayout(polow)
        pw.setContentsMargins(0, 0, 0, 0)
        pw.setSpacing(10)
        pw.addWidget(QLabel("Polo", objectName="SectionLabel"))
        self._polo_exec = QComboBox()
        self._polo_exec.addItems(POLOS)
        self._polo_exec.setFixedWidth(260)
        pw.addWidget(self._polo_exec)
        v.addWidget(self._card(polow))

        modow = QWidget()
        mw = QVBoxLayout(modow)
        mw.setContentsMargins(0, 0, 0, 0)
        mw.setSpacing(10)
        mw.addWidget(QLabel("Execução", objectName="SectionLabel"))
        self._modo = QButtonGroup(self)
        modo_row = QHBoxLayout()
        modo_row.setSpacing(18)
        for txt, val in (("Uma vez", "uma_vez"),
                         ("A cada 30 min", "30min"),
                         ("A cada 60 min", "60min")):
            rb = QRadioButton(txt)
            self._modo.addButton(rb, {"uma_vez": 0, "30min": 1, "60min": 2}[val])
            if val == "uma_vez":
                rb.setChecked(True)
            modo_row.addWidget(rb)
        modo_row.addStretch(1)
        mw.addLayout(modo_row)
        v.addWidget(self._card(modow))

        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)
        self._btn_executar = QPushButton("▶  Executar", objectName="Primary")
        self._btn_executar.clicked.connect(self._executar)
        self._btn_parar = QPushButton("■  Parar", objectName="Danger")
        self._btn_parar.clicked.connect(self._parar)
        self._btn_parar.setEnabled(False)
        btn_row.addWidget(self._btn_executar)
        btn_row.addWidget(self._btn_parar)
        btn_row.addStretch(1)
        v.addLayout(btn_row)

        log_w = QWidget()
        lw = QVBoxLayout(log_w)
        lw.setContentsMargins(0, 0, 0, 0)
        lw.setSpacing(8)
        lw.addWidget(QLabel("Log", objectName="SectionLabel"))
        self._log_exec = QPlainTextEdit()
        self._log_exec.setReadOnly(True)
        self._log_exec.setMinimumHeight(220)
        lw.addWidget(self._log_exec)
        v.addWidget(self._card(log_w))
        return page

    # ── Tratamento ──────────────────────────────────────────────────────────────

    def _build_tratamento(self):
        page = QWidget()
        v = QVBoxLayout(page)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(14)

        tools = QWidget()
        tv = QVBoxLayout(tools)
        tv.setContentsMargins(0, 0, 0, 0)
        tv.setSpacing(10)
        tv.addWidget(QLabel("Ferramentas", objectName="SectionLabel"))
        self._tool_var = QButtonGroup(self)
        for txt, val in (("Tratar Incidências", "tratar"),
                         ("XLSX → CSV", "xlsx_csv"),
                         ("CSV → XLSX", "csv_xlsx"),
                         ("Merge Arquivos", "merge")):
            rb = QRadioButton(txt)
            rb.toggled.connect(lambda _=False: self._toggle_tool_ui())
            self._tool_var.addButton(rb, {"tratar": 0, "xlsx_csv": 1, "csv_xlsx": 2, "merge": 3}[val])
            tv.addWidget(rb)
        v.addWidget(self._card(tools))

        self._frm_tratar = QWidget()
        ft = QVBoxLayout(self._frm_tratar)
        ft.setContentsMargins(0, 0, 0, 0)
        ft.setSpacing(8)
        ft.addWidget(QLabel("Arquivo de entrada"))
        arq1 = QHBoxLayout()
        self._arquivo_var = QLineEdit()
        self._arquivo_var.setReadOnly(True)
        arq1.addWidget(self._arquivo_var, stretch=1)
        arq1.addWidget(self._browse_btn(self._browse_arquivo))
        ft.addLayout(arq1)
        self._fmt_var = QLabel("Formato detectado: —", objectName="PageSub")
        ft.addWidget(self._fmt_var)
        v.addWidget(self._card(self._frm_tratar))

        self._frm_conversor = QWidget()
        fc = QVBoxLayout(self._frm_conversor)
        fc.setContentsMargins(0, 0, 0, 0)
        fc.setSpacing(8)
        fc.addWidget(QLabel("Arquivo de entrada"))
        arq2 = QHBoxLayout()
        self._conv_arquivo_var = QLineEdit()
        self._conv_arquivo_var.setReadOnly(True)
        arq2.addWidget(self._conv_arquivo_var, stretch=1)
        arq2.addWidget(self._browse_btn(self._browse_conv_arquivo))
        fc.addLayout(arq2)
        fc.addWidget(QLabel("Separador (CSV)"))
        self._conv_sep_var = QLineEdit(";")
        self._conv_sep_var.setFixedWidth(60)
        fc.addWidget(self._conv_sep_var)
        v.addWidget(self._card(self._frm_conversor))

        self._frm_merge = QWidget()
        fm = QVBoxLayout(self._frm_merge)
        fm.setContentsMargins(0, 0, 0, 0)
        fm.setSpacing(8)
        fm.addWidget(QLabel("Arquivos de entrada (selecione múltiplos)"))
        m1 = QHBoxLayout()
        self._merge_arquivos_var = QLineEdit()
        self._merge_arquivos_var.setReadOnly(True)
        m1.addWidget(self._merge_arquivos_var, stretch=1)
        m1.addWidget(self._browse_btn(self._browse_merge))
        fm.addLayout(m1)
        fm.addWidget(QLabel("Arquivo de saída"))
        m2 = QHBoxLayout()
        self._merge_saida_var = QLineEdit()
        m2.addWidget(self._merge_saida_var, stretch=1)
        m2.addWidget(self._browse_btn(self._browse_merge_saida))
        fm.addLayout(m2)
        fm.addWidget(QLabel("Coluna chave (opcional, remove duplicatas)"))
        self._merge_chave_var = QLineEdit()
        fm.addWidget(self._merge_chave_var)
        v.addWidget(self._card(self._frm_merge))

        btn_row = QHBoxLayout()
        self._btn_tratar = QPushButton("▶  Executar", objectName="Primary")
        self._btn_tratar.clicked.connect(self._executar_tool)
        btn_row.addWidget(self._btn_tratar)
        btn_row.addStretch(1)
        v.addLayout(btn_row)

        log_w = QWidget()
        lw = QVBoxLayout(log_w)
        lw.setContentsMargins(0, 0, 0, 0)
        lw.setSpacing(8)
        lw.addWidget(QLabel("Log", objectName="SectionLabel"))
        self._log_trat = QPlainTextEdit()
        self._log_trat.setReadOnly(True)
        self._log_trat.setMinimumHeight(220)
        lw.addWidget(self._log_trat)
        v.addWidget(self._card(log_w))

        self._tool_var.button(0).setChecked(True)
        self._toggle_tool_ui()
        return page

    @staticmethod
    def _browse_btn(slot: Callable) -> QPushButton:
        b = QPushButton("...")
        b.setFixedWidth(40)
        b.clicked.connect(slot)
        return b

    # ── Configurações ──────────────────────────────────────────────────────────

    def _build_configuracoes(self):
        page = QWidget()
        v = QVBoxLayout(page)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(14)
        self._vars_cfg = {}

        orig = QWidget()
        ov = QVBoxLayout(orig)
        ov.setContentsMargins(0, 0, 0, 0)
        ov.setSpacing(10)
        ov.addWidget(QLabel("Origem do relatório", objectName="SectionLabel"))
        self._origem_var = QButtonGroup(self)
        orig_row = QHBoxLayout()
        orig_row.setSpacing(18)
        for txt, val in (("GeoOnline", "geonline"), ("Operview", "operview")):
            rb = QRadioButton(txt)
            self._origem_var.addButton(rb, 0 if val == "geonline" else 1)
            if val == "geonline":
                rb.setChecked(True)
            orig_row.addWidget(rb)
        ov.addLayout(orig_row)
        v.addWidget(self._card(orig))

        geo = QWidget()
        g = QVBoxLayout(geo)
        g.setContentsMargins(0, 0, 0, 0)
        g.setSpacing(10)
        g.addWidget(QLabel("GeoOnline", objectName="SectionLabel"))
        for label, key, secret in (
            ("URL", "geo_url", False),
            ("Conta", "geo_conta", False),
            ("Login", "geo_login", False),
            ("Senha", "geo_senha", True),
        ):
            row = QHBoxLayout()
            row.setSpacing(10)
            row.addWidget(QLabel(label))
            var = QLineEdit()
            if secret:
                var.setEchoMode(QLineEdit.EchoMode.Password)
            self._vars_cfg[key] = var
            row.addWidget(var, stretch=1)
            g.addLayout(row)
        v.addWidget(self._card(geo))

        opv = QWidget()
        op = QVBoxLayout(opv)
        op.setContentsMargins(0, 0, 0, 0)
        op.setSpacing(10)
        op.addWidget(QLabel("Operview", objectName="SectionLabel"))
        for label, key, secret in (
            ("URL", "opv_url", False),
            ("Login", "opv_login", False),
            ("Senha", "opv_senha", True),
        ):
            row = QHBoxLayout()
            row.setSpacing(10)
            row.addWidget(QLabel(label))
            var = QLineEdit()
            if secret:
                var.setEchoMode(QLineEdit.EchoMode.Password)
            self._vars_cfg[key] = var
            row.addWidget(var, stretch=1)
            op.addLayout(row)
        v.addWidget(self._card(opv))

        apar = QWidget()
        ap = QVBoxLayout(apar)
        ap.setContentsMargins(0, 0, 0, 0)
        ap.setSpacing(10)
        ap.addWidget(QLabel("Aparência", objectName="SectionLabel"))
        self._chk_claro = QCheckBox("Modo claro")
        self._chk_claro.toggled.connect(self._aplicar_tema)
        ap.addWidget(self._chk_claro)
        v.addWidget(self._card(apar))

        gen = QWidget()
        gn = QVBoxLayout(gen)
        gn.setContentsMargins(0, 0, 0, 0)
        gn.setSpacing(10)
        gn.addWidget(QLabel("Geral", objectName="SectionLabel"))

        pasta_row = QHBoxLayout()
        pasta_row.setSpacing(10)
        pasta_row.addWidget(QLabel("Pasta local"))
        v_pasta = QLineEdit()
        self._vars_cfg["pasta_local"] = v_pasta
        pasta_row.addWidget(v_pasta, stretch=1)
        b = QPushButton("...")
        b.setFixedWidth(40)
        b.clicked.connect(lambda: v_pasta.setText(
            QFileDialog.getExistingDirectory(
                self, "Selecionar pasta", options=QFileDialog.Option.DontUseNativeDialog)
            or v_pasta.text()))
        pasta_row.addWidget(b)
        gn.addLayout(pasta_row)

        self._vars_cfg["sp_enabled"] = QCheckBox("Enviar para SharePoint")
        self._vars_cfg["sp_enabled"].toggled.connect(self._toggle_sp)
        gn.addWidget(self._vars_cfg["sp_enabled"])

        sp_row = QHBoxLayout()
        sp_row.setSpacing(10)
        sp_row.addWidget(QLabel("Pasta SharePoint"))
        v_sp_pasta = QLineEdit()
        self._vars_cfg["sp_pasta"] = v_sp_pasta
        self._entry_sp = v_sp_pasta
        sp_row.addWidget(v_sp_pasta, stretch=1)
        bsp = QPushButton("...")
        bsp.setFixedWidth(40)
        bsp.clicked.connect(lambda: v_sp_pasta.setText(
            QFileDialog.getExistingDirectory(
                self, "Selecionar pasta", options=QFileDialog.Option.DontUseNativeDialog)
            or v_sp_pasta.text()))
        self._btn_sp = bsp
        sp_row.addWidget(bsp)
        gn.addLayout(sp_row)

        self._vars_cfg["n8n_enabled"] = QCheckBox("Enviar análises para n8n")
        self._vars_cfg["n8n_enabled"].toggled.connect(self._toggle_n8n)
        gn.addWidget(self._vars_cfg["n8n_enabled"])

        n8n_row = QHBoxLayout()
        n8n_row.setSpacing(10)
        n8n_row.addWidget(QLabel("Webhook n8n"))
        v_n8n = QLineEdit()
        self._vars_cfg["n8n_webhook"] = v_n8n
        self._entry_n8n = v_n8n
        n8n_row.addWidget(v_n8n, stretch=1)
        gn.addLayout(n8n_row)

        self._vars_cfg["base_mensal_enabled"] = QCheckBox("Atualizar base mensal (mensal)")
        gn.addWidget(self._vars_cfg["base_mensal_enabled"])
        v.addWidget(self._card(gen))

        upd = QWidget()
        u = QVBoxLayout(upd)
        u.setContentsMargins(0, 0, 0, 0)
        u.setSpacing(10)
        u.addWidget(QLabel("Atualização (GitHub)", objectName="SectionLabel"))
        u.addWidget(QLabel(f"Versão atual: {local_version()}"))
        self._update_status = QLabel("")
        self._update_status.setStyleSheet("color:#3DDC97;")
        u.addWidget(self._update_status)
        bupd = QPushButton("Verificar atualização")
        bupd.clicked.connect(self._verificar_update)
        u.addWidget(bupd, alignment=Qt.AlignmentFlag.AlignLeft)
        if REPO:
            u.addWidget(QLabel(f"Repositório: {REPO}"))
        v.addWidget(self._card(upd))

        save = QPushButton("💾  Salvar configurações", objectName="Primary")
        save.clicked.connect(self._salvar_config)
        v.addWidget(save, alignment=Qt.AlignmentFlag.AlignRight)

        self._carregar_config()
        return page

    def _aplicar_tema(self, claro: bool = False):
        self._claro = bool(claro)
        self._tema = _light_theme() if self._claro else _dark_theme()
        self._accent = self._tema["accent"]
        self._app.setStyleSheet(_build_style(self._tema))
        _apply_palette(self._app, self._tema)

    # ── Callbacks Execução ──────────────────────────────────────────────────────

    def _executar(self):
        data_ini = self._data_ini.get_date()
        data_fim = self._data_fim.get_date()
        if data_ini > data_fim:
            QMessageBox.critical(self, "Erro de data",
                                 "Data inicial não pode ser maior que a final.")
            return
        self._btn_executar.setEnabled(False)
        self._btn_parar.setEnabled(True)
        modo = {0: "uma_vez", 1: "30min", 2: "60min"}[self._modo.checkedId()]
        self._on_executar(
            selecionados={"geonline": True},
            data_ini=data_ini,
            data_fim=data_fim,
            modo=modo,
            polo=self._polo_exec.currentText(),
        )

    def _parar(self):
        self._on_parar()

    # ── Callbacks Tratamento ───────────────────────────────────────────────────

    def _toggle_tool_ui(self):
        tool = {0: "tratar", 1: "xlsx_csv", 2: "csv_xlsx", 3: "merge"}[self._tool_var.checkedId()]
        self._frm_tratar.setVisible(tool == "tratar")
        self._frm_conversor.setVisible(tool in ("xlsx_csv", "csv_xlsx"))
        self._frm_merge.setVisible(tool == "merge")
        labels = {"tratar": "▶  Tratar", "xlsx_csv": "▶  Converter",
                  "csv_xlsx": "▶  Converter", "merge": "▶  Merge"}
        self._btn_tratar.setText(labels.get(tool, "▶  Executar"))

    def _browse_arquivo(self):
        p, _ = QFileDialog.getOpenFileName(
            self, "Selecionar arquivo", "",
            "Excel / CSV (*.xlsx *.xls *.csv);;Todos (*.*)",
            options=QFileDialog.Option.DontUseNativeDialog)
        if p:
            self._arquivo_var.setText(p)

    def _browse_conv_arquivo(self):
        tool = {1: "xlsx_csv", 2: "csv_xlsx"}[self._tool_var.checkedId()]
        ext = "*.xlsx" if tool == "xlsx_csv" else "*.csv"
        p, _ = QFileDialog.getOpenFileName(
            self, "Selecionar arquivo", "", f"Arquivos ({ext});;Todos (*.*)",
            options=QFileDialog.Option.DontUseNativeDialog)
        if p:
            self._conv_arquivo_var.setText(p)

    def _browse_merge(self):
        ps, _ = QFileDialog.getOpenFileNames(
            self, "Selecionar arquivos", "",
            "Excel / CSV (*.xlsx *.xls *.csv);;Todos (*.*)",
            options=QFileDialog.Option.DontUseNativeDialog)
        if ps:
            self._merge_arquivos_var.setText("; ".join(ps))

    def _browse_merge_saida(self):
        p, _ = QFileDialog.getSaveFileName(
            self, "Salvar como", "", "Excel (*.xlsx);;CSV (*.csv)",
            options=QFileDialog.Option.DontUseNativeDialog)
        if p:
            self._merge_saida_var.setText(p)

    def _executar_tool(self):
        tool = {0: "tratar", 1: "xlsx_csv", 2: "csv_xlsx", 3: "merge"}[self._tool_var.checkedId()]
        self._btn_tratar.setEnabled(False)

        if tool == "tratar":
            arq = self._arquivo_var.text().strip()
            if not arq:
                QMessageBox.warning(self, "Atenção", "Selecione um arquivo de entrada.")
                self._btn_tratar.setEnabled(True)
                return
            p = Path(arq)
            if not p.exists():
                QMessageBox.critical(self, "Erro", f"Arquivo não encontrado:\n{arq}")
                self._btn_tratar.setEnabled(True)
                return
            self._on_tratar(arquivo=p, polo="Todos os polos", log_fn=self.log_tratamento)
        elif tool in ("xlsx_csv", "csv_xlsx"):
            import threading
            arq = self._conv_arquivo_var.text().strip()
            sep = self._conv_sep_var.text()
            if not arq:
                QMessageBox.warning(self, "Atenção", "Selecione um arquivo de entrada.")
                self._btn_tratar.setEnabled(True)
                return

            def _run():
                try:
                    from core.processors.utilidades import xlsx_para_csv, csv_para_xlsx
                    if tool == "xlsx_csv":
                        saida = xlsx_para_csv(arq, sep=sep, log_fn=self.log_tratamento)
                    else:
                        saida = csv_para_xlsx(arq, sep=sep, log_fn=self.log_tratamento)
                    self.log_tratamento(f"✅ Convertido: {saida}")
                except Exception as e:
                    self.log_tratamento(f"❌ Erro: {e}")
                finally:
                    self.tratando_signal.emit(False)
            threading.Thread(target=_run, daemon=True).start()
        elif tool == "merge":
            import threading
            arquivos_str = self._merge_arquivos_var.text().strip()
            saida = self._merge_saida_var.text().strip()
            chave = self._merge_chave_var.text().strip() or None
            if not arquivos_str or not saida:
                QMessageBox.warning(self, "Atenção", "Informe arquivos de entrada e saída.")
                self._btn_tratar.setEnabled(True)
                return
            arquivos = [a.strip() for a in arquivos_str.split(";") if a.strip()]
            if len(arquivos) < 2:
                QMessageBox.warning(self, "Atenção", "Informe pelo menos 2 arquivos.")
                self._btn_tratar.setEnabled(True)
                return

            def _run():
                try:
                    from core.processors.utilidades import merge_arquivos
                    saida_final = merge_arquivos(
                        arquivos, saida, col_chave=chave, log_fn=self.log_tratamento)
                    self.log_tratamento(f"✅ Merge concluído: {saida_final}")
                except Exception as e:
                    self.log_tratamento(f"❌ Erro: {e}")
                finally:
                    self.tratando_signal.emit(False)
            threading.Thread(target=_run, daemon=True).start()

    # ── Callbacks Configurações ────────────────────────────────────────────────

    def _toggle_sp(self):
        estado = self._vars_cfg["sp_enabled"].isChecked()
        self._entry_sp.setEnabled(estado)
        self._btn_sp.setEnabled(estado)

    def _toggle_n8n(self):
        self._entry_n8n.setEnabled(self._vars_cfg["n8n_enabled"].isChecked())

    def _carregar_config(self):
        cfg = self._cfg
        self._vars_cfg["geo_url"].setText(cfg.geonline.url)
        self._vars_cfg["geo_conta"].setText(cfg.geonline.conta)
        self._vars_cfg["geo_login"].setText(cfg.geonline.login)
        self._vars_cfg["geo_senha"].setText(cfg.geonline.senha)
        self._vars_cfg["opv_url"].setText(cfg.operview.url)
        self._vars_cfg["opv_login"].setText(cfg.operview.login)
        self._vars_cfg["opv_senha"].setText(cfg.operview.senha)
        for rb in self._origem_var.buttons():
            if self._origem_var.id(rb) == (1 if cfg.origem_relatorio == "operview" else 0):
                rb.setChecked(True)
        self._vars_cfg["pasta_local"].setText(cfg.pasta_local)
        self._vars_cfg["sp_enabled"].setChecked(cfg.sharepoint.enabled)
        self._vars_cfg["sp_pasta"].setText(cfg.sharepoint.pasta)
        self._vars_cfg["n8n_enabled"].setChecked(cfg.n8n.enabled)
        self._vars_cfg["n8n_webhook"].setText(cfg.n8n.webhook)
        self._vars_cfg["base_mensal_enabled"].setChecked(cfg.base_mensal_enabled)
        self._chk_claro.setChecked(bool(cfg.tema_claro))
        self._toggle_sp()
        self._toggle_n8n()

    def _salvar_config(self):
        cfg = self._cfg
        cfg.geonline.url = self._vars_cfg["geo_url"].text().strip()
        cfg.geonline.conta = self._vars_cfg["geo_conta"].text().strip()
        cfg.geonline.login = self._vars_cfg["geo_login"].text().strip()
        cfg.geonline.senha = self._vars_cfg["geo_senha"].text().strip()
        cfg.operview.url = self._vars_cfg["opv_url"].text().strip()
        cfg.operview.login = self._vars_cfg["opv_login"].text().strip()
        cfg.operview.senha = self._vars_cfg["opv_senha"].text().strip()
        cfg.origem_relatorio = (
            "operview" if self._origem_var.id(self._origem_var.checkedButton()) == 1
            else "geonline"
        )
        cfg.pasta_local = self._vars_cfg["pasta_local"].text().strip()
        cfg.sharepoint.enabled = self._vars_cfg["sp_enabled"].isChecked()
        cfg.sharepoint.pasta = self._vars_cfg["sp_pasta"].text().strip()
        cfg.n8n.enabled = self._vars_cfg["n8n_enabled"].isChecked()
        cfg.n8n.webhook = self._vars_cfg["n8n_webhook"].text().strip()
        cfg.webhook_n8n = cfg.n8n.webhook
        cfg.base_mensal_enabled = self._vars_cfg["base_mensal_enabled"].isChecked()
        cfg.tema_claro = self._chk_claro.isChecked()

        if not cfg.pasta_local:
            QMessageBox.warning(self, "Atenção", "Informe a pasta local.")
            return

        save_config(cfg)
        QMessageBox.information(self, "Salvo", "Configurações salvas com sucesso!")

    def _verificar_update(self):
        if not is_installed():
            self._update_status.setText(
                "Executando a partir do código-fonte. Atualize via git/pull do repositório.")
            return
        self._update_status.setText("Verificando...")
        self._app.processEvents()
        try:
            res = check_for_update(REPO)
        except Exception as e:
            self._update_status.setText(f"Erro ao verificar: {e}")
            return
        if not res:
            self._update_status.setText(
                "DataHub está atualizado (ou não foi possível verificar a conexão).")
            return
        tag, url = res
        resp = QMessageBox.question(
            self, "Atualização disponível",
            f"Nova versão {tag} disponível.\n"
            "Deseja baixar e atualizar agora?\n"
            "O aplicativo será reiniciado.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if resp != QMessageBox.StandardButton.Yes:
            self._update_status.setText("Atualização adiada.")
            return
        self._update_status.setText(f"Baixando {tag}...")
        self._app.processEvents()
        try:
            bat = prepare_update(url, progress=None)
        except Exception as e:
            self._update_status.setText(f"Falha no download: {e}")
            return
        launch_updater(bat)

    # ── Log público (thread-safe) ──────────────────────────────────────────────

    def log(self, mensagem: str):
        self.log_signal.emit(mensagem)

    def log_tratamento(self, mensagem: str):
        print(f"[DEBUG TRAT] {mensagem}")
        self.log_trat_signal.emit(mensagem)

    def _append_log_exec(self, msg: str):
        self._log_exec.appendPlainText(f"{datetime.now():%H:%M:%S}  {msg}")

    def _append_log_trat(self, msg: str):
        self._log_trat.appendPlainText(f"{datetime.now():%H:%M:%S}  {msg}")

    def set_rodando(self, rodando: bool):
        self.rodando_signal.emit(rodando)

    def _set_rodando(self, rodando: bool):
        self._btn_executar.setEnabled(not rodando)
        self._btn_parar.setEnabled(rodando)

    def set_tratando(self, tratando: bool):
        self.tratando_signal.emit(tratando)

    def _set_tratando(self, tratando: bool):
        self._btn_tratar.setEnabled(not tratando)

    def _set_formato(self, fmt: str):
        self._fmt_var.setText(f"Formato detectado: {fmt or '—'}")

    # ── Ícone na bandeja do sistema ────────────────────────────────────────────
    def _init_tray(self):
        icon = QIcon(_resolve("appicon.ico"))
        self._tray = QSystemTrayIcon(icon, self)
        self._tray.setToolTip("DataHub")
        menu = QMenu()
        act_abrir = menu.addAction("Abrir")
        act_abrir.triggered.connect(self._mostrar_e_focar)
        act_sair = menu.addAction("Sair")
        act_sair.triggered.connect(self._sair)
        self._tray.setContextMenu(menu)
        self._tray.activated.connect(
            lambda r: self._mostrar_e_focar()
            if r == QSystemTrayIcon.ActivationReason.Trigger else None)
        self._tray.show()

    def _mostrar_e_focar(self):
        self.showNormal()
        self.show()
        self.raise_()
        self.activateWindow()

    def _sair(self):
        try:
            self.quit_signal.emit()
        except Exception:
            pass

    def closeEvent(self, event):
        event.ignore()
        self.hide()
        try:
            self._tray.showMessage(
                "DataHub",
                "O app continua na bandeja do sistema (clique no ícone para abrir).",
                QSystemTrayIcon.MessageIcon.Information, 2500)
        except Exception:
            pass

    def mostrar_execucao(self):
        self._stack.setCurrentIndex(0)

    def mostrar_concluido(self, mensagem: str = "Fluxo concluído com sucesso!"):
        self.concluido_signal.emit(mensagem)

    def _mostrar_concluido(self, mensagem: str):
        QMessageBox.information(self, "DataHub", mensagem)

    def _mostrar_aviso_login(self, login_event):
        QMessageBox.information(
            self, "Operview — Login necessário",
            "O download do Operview precisa de login.\n\n"
            "Faça o login no navegador que foi aberto e clique OK para continuar.",
        )
        if login_event is not None:
            login_event.set()

    # ── Loop ───────────────────────────────────────────────────────────────────

    def mainloop(self):
        self.show()
        self._app.exec()

    def run(self):
        self.mainloop()


if __name__ == "__main__":
    from config import AppConfig
    MainWindow(AppConfig(), lambda **k: None, lambda: None, lambda **k: None).run()
