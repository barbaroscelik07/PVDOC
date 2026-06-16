"""
Birim/Seri Formül modülü (Bölüm 5 / Tablo 1).

Hammadde listesini düzenler. Çift katmanlı tabletler için her satır bir
katmana (0=yok, 1=Katman I, 2=Katman II) atanabilir; "ara toplam" satırları
(Katman I Ağırlık vb.) ayrıca işaretlenir. Sayısal sütunlar serbest metin
olarak değil, biçimli giriş olarak tutulur ama doğrulama gevşektir (k.m., U.Y.
gibi değerler de yazılabildiği için ham metin saklanır).
"""

from __future__ import annotations

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QTableWidget,
    QTableWidgetItem, QHeaderView, QAbstractItemView, QComboBox, QApplication,
)
from PyQt6.QtGui import QKeySequence, QShortcut

from core.models import ProjeVerisi, Hammadde
from moduller.widget_yardimcilari import (
    MODUL_STIL, baslik_etiketi, ipucu_etiketi,
)


_KATMAN_SECENEK = ["—", "Katman I", "Katman II"]


def _f(deger) -> str:
    """float|None → düzenlenebilir metin."""
    return "" if deger is None else (f"{deger:g}")


def _pf(metin: str):
    """metin → float|None (boş/geçersizse None)."""
    metin = metin.strip().replace(",", ".")
    if not metin:
        return None
    try:
        return float(metin)
    except ValueError:
        return None


class FormulModulu(QWidget):
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
        ust.addWidget(baslik_etiketi("Birim ve Seri Formül (Tablo 1)"))
        ust.addStretch(1)
        b_ekle = QPushButton("+ Hammadde"); b_ekle.clicked.connect(self._satir_ekle)
        b_top = QPushButton("+ Ara Toplam"); b_top.clicked.connect(self._toplam_ekle)
        b_yapistir = QPushButton("Excel/Word Yapıştır"); b_yapistir.clicked.connect(self._panodan_yapistir)
        b_sil = QPushButton("− Sil"); b_sil.setObjectName("tehlike"); b_sil.clicked.connect(self._satir_sil)
        ust.addWidget(b_yapistir); ust.addWidget(b_ekle); ust.addWidget(b_top); ust.addWidget(b_sil)
        kok.addLayout(ust)

        kok.addWidget(ipucu_etiketi(
            "Çift katmanlı tablette her satırı bir katmana atayın. "
            "Excel/Word'den kopyaladığınız tabloyu ‘Excel/Word Yapıştır’ ile "
            "(veya tabloya tıklayıp Ctrl+V) mevcut satırların altına ekleyebilirsiniz. "
            "Sütun sırası: Hammadde · Fonksiyon · Birim Formül · % İçerik · kg/seri."
        ))

        t = QTableWidget(0, 6)
        t.setHorizontalHeaderLabels(
            ["Hammadde / Yardımcı Madde", "Fonksiyon", "Birim Formül (mg/tb)",
             "% İçerik", "kg / seri", "Katman"]
        )
        t.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        t.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)
        t.verticalHeader().setVisible(False)
        t.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        t.cellChanged.connect(self._hucre_degisti)
        self.tablo = t
        kok.addWidget(t, 1)

        # Ctrl+V: tablo odaktayken panodan yapıştır
        ksy = QShortcut(QKeySequence.StandardKey.Paste, t)
        ksy.activated.connect(self._panodan_yapistir)

    # ----------------------------------------------------------- veri ↔ UI
    def _satir_ekle(self) -> None:
        self.proje.hammaddeler.append(Hammadde(ad=""))
        self._verilerden_doldur()

    def _panodan_yapistir(self) -> None:
        """
        Excel/Word'den kopyalanan tabloyu mevcut satırların ALTINA ekler.
        Satırlar yeni satır (\\n), sütunlar sekme (\\t) ile ayrılır — Excel/Word
        kopyalamasının standart biçimi. Sütun sırası:
        Hammadde · Fonksiyon · Birim Formül · % İçerik · kg/seri.
        """
        metin = QApplication.clipboard().text()
        if not metin.strip():
            return
        eklenen = 0
        for satir in metin.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
            if not satir.strip():
                continue
            h = satir.split("\t")
            # eksik sütunları boş kabul et
            h += [""] * (5 - len(h))
            self.proje.hammaddeler.append(Hammadde(
                ad=h[0].strip(),
                fonksiyon=h[1].strip(),
                birim_formul=_pf(h[2]),
                yuzde_icerik=_pf(h[3]),
                seri_miktar=_pf(h[4]),
            ))
            eklenen += 1
        if eklenen:
            self._verilerden_doldur()

    def _toplam_ekle(self) -> None:
        self.proje.hammaddeler.append(Hammadde(ad="Katman Ağırlık", ara_toplam=True))
        self._verilerden_doldur()

    def _satir_sil(self) -> None:
        r = self.tablo.currentRow()
        if 0 <= r < len(self.proje.hammaddeler):
            del self.proje.hammaddeler[r]
            self._verilerden_doldur()

    def _hucre_degisti(self, row: int, col: int) -> None:
        if not (0 <= row < len(self.proje.hammaddeler)):
            return
        if col == 5:
            return  # katman combobox ile yönetiliyor
        it = self.tablo.item(row, col)
        v = it.text() if it else ""
        h = self.proje.hammaddeler[row]
        if col == 0:
            h.ad = v
        elif col == 1:
            h.fonksiyon = v
        elif col == 2:
            h.birim_formul = _pf(v)
        elif col == 3:
            h.yuzde_icerik = _pf(v)
        elif col == 4:
            h.seri_miktar = _pf(v)

    def _katman_degisti(self, idx: int, secim_idx: int) -> None:
        if 0 <= idx < len(self.proje.hammaddeler):
            self.proje.hammaddeler[idx].katman = secim_idx  # 0,1,2

    def _verilerden_doldur(self) -> None:
        self.tablo.blockSignals(True)
        self.tablo.setRowCount(0)
        for i, h in enumerate(self.proje.hammaddeler):
            self.tablo.insertRow(i)
            self.tablo.setItem(i, 0, QTableWidgetItem(h.ad))
            self.tablo.setItem(i, 1, QTableWidgetItem(h.fonksiyon))
            self.tablo.setItem(i, 2, QTableWidgetItem(_f(h.birim_formul)))
            self.tablo.setItem(i, 3, QTableWidgetItem(_f(h.yuzde_icerik)))
            self.tablo.setItem(i, 4, QTableWidgetItem(_f(h.seri_miktar)))
            cmb = QComboBox()
            cmb.addItems(_KATMAN_SECENEK)
            cmb.setCurrentIndex(h.katman if 0 <= h.katman <= 2 else 0)
            cmb.currentIndexChanged.connect(lambda s, idx=i: self._katman_degisti(idx, s))
            self.tablo.setCellWidget(i, 5, cmb)
        self.tablo.blockSignals(False)
