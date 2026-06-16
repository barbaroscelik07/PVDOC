"""
Ekipman modülü (Bölüm 7 / Tablo 5).

Kullanıcı ekipman listesini serbestçe düzenler: satır ekler, siler,
hücreleri doğrudan yazar. Şablondaki örnek satırlar başlangıç şablonu
olarak tek tıkla eklenebilir.
"""

from __future__ import annotations

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QTableWidget,
    QTableWidgetItem, QHeaderView, QAbstractItemView,
)

from core.models import ProjeVerisi, Ekipman
from moduller.widget_yardimcilari import (
    MODUL_STIL, baslik_etiketi, ipucu_etiketi,
)


# Şablondaki örnek ekipman listesi (hızlı başlangıç için)
_ORNEK = [
    (1, "Tartım", "TERAZİ-KANTAR", "20,0 kg / 600,0 kg"),
    (2, "Eleme", "QUADRO-FREWİTT", "100 rpm / 2800 rpm"),
    (2, "Karıştırma", "SERVOLIFT MC", "3,0 kg / 1200,0 Kg"),
    (3, "Tablet Baskı", "FETTE", "30000 tab.-saat / 470400 tab.-saat"),
    (4, "Film Kaplama", "GLATT", "350 lt"),
    (5, "Blisterleme", "UHLMANN", "60 blister-dk / 600 blister-dk"),
    (5, "Kutulama", "UHLMANN", "25 kutu-dk / 250 kutu-dk"),
]


class EkipmanModulu(QWidget):
    def __init__(self, proje: ProjeVerisi):
        super().__init__()
        self.proje = proje
        self.setObjectName("panel")
        self.setStyleSheet(MODUL_STIL)
        self._arayuzu_kur()
        self._verilerden_doldur()

    def _arayuzu_kur(self) -> None:
        kok = QVBoxLayout(self)
        kok.setContentsMargins(22, 22, 22, 22)
        kok.setSpacing(12)

        ust = QHBoxLayout()
        ust.addWidget(baslik_etiketi("Ekipman Listesi (Tablo 5)"))
        ust.addStretch(1)
        b_ekle = QPushButton("+ Satır")
        b_ekle.clicked.connect(self._satir_ekle)
        b_sil = QPushButton("− Seçili Satır")
        b_sil.setObjectName("tehlike")
        b_sil.clicked.connect(self._satir_sil)
        b_ornek = QPushButton("Örnek Listeyi Yükle")
        b_ornek.clicked.connect(self._ornek_yukle)
        ust.addWidget(b_ornek)
        ust.addWidget(b_ekle)
        ust.addWidget(b_sil)
        kok.addLayout(ust)

        kok.addWidget(ipucu_etiketi(
            "Hücrelere doğrudan yazabilir, satır ekleyip silebilirsiniz."
        ))

        t = QTableWidget(0, 4)
        t.setHorizontalHeaderLabels(["Operasyon No", "Operasyon", "Ekipman Adı", "Ekipman Kapasitesi"])
        t.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        t.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        t.verticalHeader().setVisible(False)
        t.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        t.cellChanged.connect(self._hucre_degisti)
        self.tablo = t
        kok.addWidget(t, 1)

    # ----------------------------------------------------------- veri ↔ UI
    def _satir_ekle(self) -> None:
        self.proje.ekipmanlar.append(Ekipman(operasyon_no=0, operasyon="", ekipman_adi=""))
        self._verilerden_doldur()

    def _satir_sil(self) -> None:
        r = self.tablo.currentRow()
        if 0 <= r < len(self.proje.ekipmanlar):
            del self.proje.ekipmanlar[r]
            self._verilerden_doldur()

    def _ornek_yukle(self) -> None:
        self.proje.ekipmanlar = [
            Ekipman(operasyon_no=o, operasyon=op, ekipman_adi=ad, kapasite=kap)
            for (o, op, ad, kap) in _ORNEK
        ]
        self._verilerden_doldur()

    def _hucre_degisti(self, row: int, col: int) -> None:
        if not (0 <= row < len(self.proje.ekipmanlar)):
            return
        it = self.tablo.item(row, col)
        v = it.text() if it else ""
        e = self.proje.ekipmanlar[row]
        if col == 0:
            try:
                e.operasyon_no = int(v) if v.strip() else 0
            except ValueError:
                e.operasyon_no = 0
        elif col == 1:
            e.operasyon = v
        elif col == 2:
            e.ekipman_adi = v
        elif col == 3:
            e.kapasite = v

    def _verilerden_doldur(self) -> None:
        self.tablo.blockSignals(True)
        self.tablo.setRowCount(0)
        for e in self.proje.ekipmanlar:
            r = self.tablo.rowCount()
            self.tablo.insertRow(r)
            degerler = [str(e.operasyon_no or ""), e.operasyon, e.ekipman_adi, e.kapasite]
            for c, v in enumerate(degerler):
                self.tablo.setItem(r, c, QTableWidgetItem(v))
        self.tablo.blockSignals(False)
