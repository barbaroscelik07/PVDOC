"""
Numune Alma Planı modülü (Bölüm 9 / Tablo 10).

Bölüm başlığı/metni sabittir; Tablo 10 ürün formuna göre düzenlenir:
- Tablet     : Karıştırma, Tablet Baskı, (Film Kaplama YOK), Blisterleme
- Film Tablet: Karıştırma, Tablet Baskı, Film Kaplama, Blisterleme
- Kapsül     : Karıştırma, Dolum, (Tablet Baskı / Film Kaplama YOK), Blisterleme

Kullanıcı "Forma Göre Doldur" ile şablon satırları üretebilir, sonra
tabloyu serbestçe düzenleyebilir (manuel giriş de mümkün).
"""

from __future__ import annotations

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QTableWidget,
    QTableWidgetItem, QHeaderView, QAbstractItemView,
)

from core.models import ProjeVerisi, NumuneAlmaSatiri, UrunFormu
from moduller.widget_yardimcilari import (
    MODUL_STIL, baslik_etiketi, ipucu_etiketi,
)


# Forma göre numune alma noktaları (şablon Tablo 10 temelli)
_KARISTIRMA = ("Karıştırma", "1,2,3,4,5,6,7,8,9,10",
               "Her nokta ~10 g")
_TABLET = ("Tablet Baskı", "Baş, Orta, Son", "300 tb")
_FILM = ("Film Kaplama", "Baş, Orta, Son", "300 ftb")
_DOLUM = ("Dolum", "Baş, Orta, Son", "300 kapsül")
_BLISTER = ("Blisterleme", "Baş, Orta, Son", "150 ftb")


def _forma_gore_satirlar(form: UrunFormu) -> list[NumuneAlmaSatiri]:
    """Ürün formuna göre Tablo 10 başlangıç satırlarını üretir."""
    plan: list[tuple[str, str, str]] = [_KARISTIRMA]
    if form is UrunFormu.TABLET:
        plan += [_TABLET, _BLISTER]
        op_no = {"Karıştırma": 2, "Tablet Baskı": 3, "Blisterleme": 5}
    elif form is UrunFormu.FILM_TABLET:
        plan += [_TABLET, _FILM, _BLISTER]
        op_no = {"Karıştırma": 2, "Tablet Baskı": 3, "Film Kaplama": 4, "Blisterleme": 5}
    elif form is UrunFormu.KAPSUL:
        plan += [_DOLUM, _BLISTER]
        op_no = {"Karıştırma": 2, "Dolum": 3, "Blisterleme": 5}
    else:
        op_no = {}

    return [
        NumuneAlmaSatiri(
            operasyon_no=op_no.get(op, 0),
            operasyon=op,
            numune_noktasi=nokta,
            toplam_miktar=miktar,
        )
        for (op, nokta, miktar) in plan
    ]


class NumuneModulu(QWidget):
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
        ust.addWidget(baslik_etiketi("Numune Alma Planı (Tablo 10)"))
        ust.addStretch(1)
        b_form = QPushButton("Forma Göre Doldur")
        b_form.setObjectName("birincil")
        b_form.clicked.connect(self._forma_gore_doldur)
        b_ekle = QPushButton("+ Satır"); b_ekle.clicked.connect(self._satir_ekle)
        b_sil = QPushButton("− Sil"); b_sil.setObjectName("tehlike"); b_sil.clicked.connect(self._satir_sil)
        ust.addWidget(b_form); ust.addWidget(b_ekle); ust.addWidget(b_sil)
        kok.addLayout(ust)

        kok.addWidget(ipucu_etiketi(
            "‘Forma Göre Doldur’ ürün formuna uygun satırları üretir "
            "(ör. kapsülde tablet baskı/film kaplama olmaz, dolum olur). "
            "Sonrasında tabloyu serbestçe düzenleyebilirsiniz."
        ))

        t = QTableWidget(0, 4)
        t.setHorizontalHeaderLabels(["Op. No", "Operasyon", "Numune Alma Noktası", "Toplam Numune Miktarı"])
        t.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        t.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        t.verticalHeader().setVisible(False)
        t.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        t.cellChanged.connect(self._hucre_degisti)
        self.tablo = t
        kok.addWidget(t, 1)

    # ----------------------------------------------------------- veri ↔ UI
    def _forma_gore_doldur(self) -> None:
        self.proje.numune_plani = _forma_gore_satirlar(self.proje.urun_formu)
        self._verilerden_doldur()

    def _satir_ekle(self) -> None:
        self.proje.numune_plani.append(NumuneAlmaSatiri(operasyon_no=0, operasyon=""))
        self._verilerden_doldur()

    def _satir_sil(self) -> None:
        r = self.tablo.currentRow()
        if 0 <= r < len(self.proje.numune_plani):
            del self.proje.numune_plani[r]
            self._verilerden_doldur()

    def _hucre_degisti(self, row: int, col: int) -> None:
        if not (0 <= row < len(self.proje.numune_plani)):
            return
        it = self.tablo.item(row, col)
        v = it.text() if it else ""
        n = self.proje.numune_plani[row]
        if col == 0:
            try:
                n.operasyon_no = int(v) if v.strip() else 0
            except ValueError:
                n.operasyon_no = 0
        elif col == 1:
            n.operasyon = v
        elif col == 2:
            n.numune_noktasi = v
        elif col == 3:
            n.toplam_miktar = v

    def _verilerden_doldur(self) -> None:
        self.tablo.blockSignals(True)
        self.tablo.setRowCount(0)
        for n in self.proje.numune_plani:
            r = self.tablo.rowCount()
            self.tablo.insertRow(r)
            for c, v in enumerate([str(n.operasyon_no or ""), n.operasyon,
                                   n.numune_noktasi, n.toplam_miktar]):
                self.tablo.setItem(r, c, QTableWidgetItem(v))
        self.tablo.blockSignals(False)
