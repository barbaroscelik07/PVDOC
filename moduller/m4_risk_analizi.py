"""
Risk Analizi modülü (Bölüm 6 / Tablo 3 + 6.1 / Tablo 4).

Her ikisi de tamamen manuel (kullanıcı kararı):
- Tablo 3: operasyon bazında Kritik(E)/Kritik değil(H) değerlendirmesi.
- Tablo 4: öngörülen proses parametreleri.
Üretim yöntemi farklı olursa kullanıcı yeni satırlar ekler.
"""

from __future__ import annotations

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QTableWidget,
    QTableWidgetItem, QHeaderView, QAbstractItemView, QComboBox,
)
from PyQt6.QtCore import Qt

from core.models import ProjeVerisi, RiskSatiri, ProsesParametresi
from moduller.widget_yardimcilari import (
    MODUL_STIL, baslik_etiketi, bolum_etiketi, ipucu_etiketi, ayirici,
)


class RiskModulu(QWidget):
    def __init__(self, proje: ProjeVerisi):
        super().__init__()
        self.proje = proje
        self.setObjectName("panel")
        self.setStyleSheet(MODUL_STIL)
        self._arayuzu_kur()
        self._risk_doldur()
        self._param_doldur()

    def _arayuzu_kur(self) -> None:
        kok = QVBoxLayout(self)
        kok.setContentsMargins(22, 22, 22, 22)
        kok.setSpacing(12)

        kok.addWidget(baslik_etiketi("Proses Parametreleri Değerlendirmesi (Risk Analizi)"))

        # --- Tablo 3 ---
        u1 = QHBoxLayout()
        u1.addWidget(bolum_etiketi("Tablo 3 — Kritik / Kritik Olmayan Parametreler"))
        u1.addStretch(1)
        b1 = QPushButton("+ Satır"); b1.clicked.connect(self._risk_ekle)
        b1s = QPushButton("− Sil"); b1s.setObjectName("tehlike"); b1s.clicked.connect(self._risk_sil)
        u1.addWidget(b1); u1.addWidget(b1s)
        kok.addLayout(u1)

        t1 = QTableWidget(0, 5)
        t1.setHorizontalHeaderLabels(["Op. No", "Operasyon", "Kritik (E/H)", "Testler", "Yorumlar"])
        t1.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        t1.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        t1.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        t1.verticalHeader().setVisible(False)
        t1.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        t1.cellChanged.connect(self._risk_hucre)
        self.t_risk = t1
        kok.addWidget(t1, 1)

        kok.addWidget(ayirici())

        # --- Tablo 4 ---
        u2 = QHBoxLayout()
        u2.addWidget(bolum_etiketi("Tablo 4 — Öngörülen Proses Parametreleri"))
        u2.addStretch(1)
        b2 = QPushButton("+ Satır"); b2.clicked.connect(self._param_ekle)
        b2s = QPushButton("− Sil"); b2s.setObjectName("tehlike"); b2s.clicked.connect(self._param_sil)
        u2.addWidget(b2); u2.addWidget(b2s)
        kok.addLayout(u2)

        t2 = QTableWidget(0, 3)
        t2.setHorizontalHeaderLabels(["Açıklama (örn. Operasyon 2: Aşama 8)", "Parametre", "Değer"])
        t2.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        t2.verticalHeader().setVisible(False)
        t2.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        t2.cellChanged.connect(self._param_hucre)
        self.t_param = t2
        kok.addWidget(t2, 1)

    # =============================================================== Tablo 3
    def _risk_ekle(self) -> None:
        self.proje.risk_satirlari.append(RiskSatiri(operasyon_no=0, operasyon=""))
        self._risk_doldur()

    def _risk_sil(self) -> None:
        r = self.t_risk.currentRow()
        if 0 <= r < len(self.proje.risk_satirlari):
            del self.proje.risk_satirlari[r]
            self._risk_doldur()

    def _risk_hucre(self, row: int, col: int) -> None:
        if not (0 <= row < len(self.proje.risk_satirlari)):
            return
        rs = self.proje.risk_satirlari[row]
        if col == 2:
            # combobox ile yönetiliyor; metinden okumayalım
            return
        it = self.t_risk.item(row, col)
        v = it.text() if it else ""
        if col == 0:
            try:
                rs.operasyon_no = int(v) if v.strip() else 0
            except ValueError:
                rs.operasyon_no = 0
        elif col == 1:
            rs.operasyon = v
        elif col == 3:
            rs.testler = v
        elif col == 4:
            rs.yorumlar = v

    def _risk_doldur(self) -> None:
        self.t_risk.blockSignals(True)
        self.t_risk.setRowCount(0)
        for i, rs in enumerate(self.proje.risk_satirlari):
            self.t_risk.insertRow(i)
            self.t_risk.setItem(i, 0, QTableWidgetItem(str(rs.operasyon_no or "")))
            self.t_risk.setItem(i, 1, QTableWidgetItem(rs.operasyon))
            self.t_risk.setItem(i, 3, QTableWidgetItem(rs.testler))
            self.t_risk.setItem(i, 4, QTableWidgetItem(rs.yorumlar))
            # Kritik E/H combobox
            cmb = QComboBox()
            cmb.addItems(["H", "E"])
            cmb.setCurrentText("E" if rs.kritik else "H")
            cmb.currentTextChanged.connect(lambda t, idx=i: self._kritik_degisti(idx, t))
            self.t_risk.setCellWidget(i, 2, cmb)
        self.t_risk.blockSignals(False)

    def _kritik_degisti(self, idx: int, metin: str) -> None:
        if 0 <= idx < len(self.proje.risk_satirlari):
            self.proje.risk_satirlari[idx].kritik = (metin == "E")

    # =============================================================== Tablo 4
    def _param_ekle(self) -> None:
        self.proje.proses_parametreleri.append(ProsesParametresi(aciklama=""))
        self._param_doldur()

    def _param_sil(self) -> None:
        r = self.t_param.currentRow()
        if 0 <= r < len(self.proje.proses_parametreleri):
            del self.proje.proses_parametreleri[r]
            self._param_doldur()

    def _param_hucre(self, row: int, col: int) -> None:
        if not (0 <= row < len(self.proje.proses_parametreleri)):
            return
        it = self.t_param.item(row, col)
        v = it.text() if it else ""
        pp = self.proje.proses_parametreleri[row]
        if col == 0:
            pp.aciklama = v
        elif col == 1:
            pp.parametre = v
        elif col == 2:
            pp.deger = v

    def _param_doldur(self) -> None:
        self.t_param.blockSignals(True)
        self.t_param.setRowCount(0)
        for i, pp in enumerate(self.proje.proses_parametreleri):
            self.t_param.insertRow(i)
            for c, v in enumerate([pp.aciklama, pp.parametre, pp.deger]):
                self.t_param.setItem(i, c, QTableWidgetItem(v))
        self.t_param.blockSignals(False)
