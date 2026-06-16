"""
Modüller arası ortak UI yardımcıları.

Tema bütünlüğü (koyu zemin + cam panel + sarı/lacivert aksan) burada toplanır;
her modül aynı görünümü paylaşsın. Salt sunum amaçlı, iş mantığı içermez.
"""

from __future__ import annotations

from PyQt6.QtWidgets import (
    QWidget, QLabel, QFrame, QVBoxLayout, QHBoxLayout, QPushButton,
)
from PyQt6.QtCore import Qt


# Fenerbahçe aksanı + koyu tema renkleri (tek yerden yönetilsin)
LACIVERT = "#00295C"
SARI = "#FFED00"
ZEMIN = "#11151c"
PANEL = "rgba(255,255,255,0.04)"
KENAR = "rgba(255,255,255,0.08)"
SOLUK = "#8b98a5"

# Modül panellerinde kullanılacak ortak stil (ana_pencere global stiline ek)
MODUL_STIL = f"""
QLabel#modul_baslik {{ font-size: 20px; font-weight: 600; color: #e6edf3; }}
QLabel#bolum_baslik {{ font-size: 14px; font-weight: 600; color: {SARI}; }}
QLabel#ipucu {{ color: {SOLUK}; font-size: 12px; }}

QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox, QPlainTextEdit, QTextEdit {{
    background-color: rgba(255,255,255,0.05);
    border: 1px solid {KENAR};
    border-radius: 8px;
    padding: 7px 9px;
    color: #e6edf3;
    selection-background-color: rgba(88,166,255,0.4);
}}
QLineEdit:focus, QComboBox:focus, QSpinBox:focus, QDoubleSpinBox:focus,
QPlainTextEdit:focus, QTextEdit:focus {{
    border: 1px solid rgba(88,166,255,0.6);
}}
QComboBox::drop-down {{ border: none; width: 22px; }}
QComboBox QAbstractItemView {{
    background-color: #1a2029;
    border: 1px solid {KENAR};
    selection-background-color: rgba(88,166,255,0.3);
    color: #e6edf3;
}}

QPushButton {{
    background-color: rgba(255,255,255,0.07);
    border: 1px solid {KENAR};
    border-radius: 8px;
    padding: 8px 14px;
    color: #e6edf3;
}}
QPushButton:hover {{ background-color: rgba(255,255,255,0.12); }}
QPushButton#birincil {{
    background-color: {SARI};
    color: {LACIVERT};
    font-weight: 600;
    border: none;
}}
QPushButton#birincil:hover {{ background-color: #ffe11a; }}
QPushButton#tehlike {{ color: #ff9a9a; }}

QTableWidget {{
    background-color: rgba(255,255,255,0.03);
    border: 1px solid {KENAR};
    border-radius: 8px;
    gridline-color: {KENAR};
    color: #e6edf3;
}}
QHeaderView::section {{
    background-color: rgba(255,255,255,0.06);
    color: {SOLUK};
    border: none;
    border-bottom: 1px solid {KENAR};
    padding: 6px;
    font-weight: 600;
}}
QTableWidget::item:selected {{ background-color: rgba(88,166,255,0.25); }}
/* Tablo içi düzenleme editörü: OPAK zemin + sıfır kenar.
   Aksi halde altdaki item metni editörün arkasından görünüp 'üst üste binme'
   izlenimi verir. */
QTableWidget QLineEdit, QTableWidget QAbstractItemView {{
    background-color: #1a2029;
    color: #ffffff;
    border: 1px solid rgba(88,166,255,0.7);
    border-radius: 0px;
    padding: 2px 4px;
    margin: 0px;
    selection-background-color: rgba(88,166,255,0.5);
}}

QCheckBox {{ color: #e6edf3; spacing: 7px; }}
QCheckBox::indicator {{
    width: 18px; height: 18px;
    border: 1px solid {KENAR};
    border-radius: 4px;
    background: rgba(255,255,255,0.05);
}}
QCheckBox::indicator:checked {{
    background: {SARI};
    border: 1px solid {SARI};
}}
QListWidget {{
    background-color: rgba(255,255,255,0.03);
    border: 1px solid {KENAR};
    border-radius: 8px;
    color: #e6edf3;
}}
QListWidget::item {{ padding: 8px 10px; border-radius: 6px; }}
QListWidget::item:selected {{ background-color: rgba(88,166,255,0.25); }}
"""


def baslik_etiketi(metin: str) -> QLabel:
    e = QLabel(metin)
    e.setObjectName("modul_baslik")
    return e


def bolum_etiketi(metin: str) -> QLabel:
    e = QLabel(metin)
    e.setObjectName("bolum_baslik")
    return e


def ipucu_etiketi(metin: str) -> QLabel:
    e = QLabel(metin)
    e.setObjectName("ipucu")
    e.setWordWrap(True)
    return e


def ayirici() -> QFrame:
    c = QFrame()
    c.setFrameShape(QFrame.Shape.HLine)
    c.setStyleSheet(f"color: {KENAR}; background: {KENAR}; max-height: 1px;")
    return c


def kart_paneli() -> QWidget:
    """Cam görünümlü içerik paneli (objectName='panel', ana_pencere stiliyle uyumlu)."""
    w = QWidget()
    w.setObjectName("panel")
    return w
