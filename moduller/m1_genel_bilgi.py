"""
Genel Bilgi modülü (şablon başlığı + Bölüm 5.2 Kapsanan Ürünler / Tablo 2).

Buradaki alanlar şablonun her sayfasındaki başlık/altbilgiyi ve kapsanan
ürün serilerini besler. Doküman no formatı firmaya özgü → kullanıcı elle girer.
Seri sayısı sabit 3 (tasarım kararı); seri no / boyut alanları düzenlenebilir.
"""

from __future__ import annotations

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QLineEdit,
    QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView,
)

from core.models import ProjeVerisi, SERI_SAYISI
from moduller.widget_yardimcilari import (
    MODUL_STIL, baslik_etiketi, bolum_etiketi, ipucu_etiketi, ayirici,
)


class GenelBilgiModulu(QWidget):
    def __init__(self, proje: ProjeVerisi):
        super().__init__()
        self.proje = proje
        self.setObjectName("panel")
        self.setStyleSheet(MODUL_STIL)
        self._arayuzu_kur()
        self._verilerden_doldur()

    # ------------------------------------------------------------------ UI
    def _arayuzu_kur(self) -> None:
        kok = QVBoxLayout(self)
        kok.setContentsMargins(22, 22, 22, 22)
        kok.setSpacing(14)

        kok.addWidget(baslik_etiketi("Genel Bilgi"))
        kok.addWidget(ipucu_etiketi(
            "Doküman numarası formatı firmanıza özgüdür; her projede kendiniz girersiniz."
        ))

        izgara = QGridLayout()
        izgara.setHorizontalSpacing(12)
        izgara.setVerticalSpacing(10)
        s = 0

        def satir(etiket: str, placeholder: str = "") -> QLineEdit:
            nonlocal s
            izgara.addWidget(QLabel(etiket), s, 0)
            le = QLineEdit()
            if placeholder:
                le.setPlaceholderText(placeholder)
            izgara.addWidget(le, s, 1)
            s += 1
            return le

        self.in_firma = satir("Firma İsmi:", "örn. {Firma ismi}")
        self.in_urun = satir("Ürün Adı:", "örn. Xxx Film Kaplı Tablet")
        self.in_pvp_no = satir("PVP Doküman No:", "örn. AG-PV-xxx")
        self.in_pvr_no = satir("PVR Doküman No:", "örn. AG-PV-xxx-R")
        self.in_rev_no = satir("Revizyon No:", "00")
        self.in_rev_tarih = satir("Revizyon Tarihi:", "U.Y.")
        self.in_form_no = satir("Form No:", "örn. N-15-0506")

        # Alanlar değişince anında veriye yaz
        self.in_firma.textChanged.connect(lambda t: self._yaz("firma_ismi", t))
        self.in_urun.textChanged.connect(self._urun_degisti)
        self.in_pvp_no.textChanged.connect(lambda t: self._yaz("pvp_dokuman_no", t))
        self.in_pvr_no.textChanged.connect(lambda t: self._yaz("pvr_dokuman_no", t))
        self.in_rev_no.textChanged.connect(lambda t: self._yaz("revizyon_no", t))
        self.in_rev_tarih.textChanged.connect(lambda t: self._yaz("revizyon_tarihi", t))
        self.in_form_no.textChanged.connect(lambda t: self._yaz("form_no", t))

        kok.addLayout(izgara)
        kok.addWidget(ayirici())

        # Tablo 2: Kapsanan ürünler / seriler (sabit 3)
        kok.addWidget(bolum_etiketi("Kapsanan Ürünler — Seriler (Tablo 2)"))
        kok.addWidget(ipucu_etiketi(
            "Pilot üretim proses validasyonu sabit 3 seriye uygulanır."
        ))
        kok.addWidget(self._seri_tablosu())
        kok.addStretch(1)

    def _seri_tablosu(self) -> QTableWidget:
        t = QTableWidget(SERI_SAYISI, 4)
        t.setHorizontalHeaderLabels(["Ürün İsmi", "Seri No", "Seri Boyutu (adet)", "Seri Boyutu (kg)"])
        t.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        t.verticalHeader().setVisible(False)
        t.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        t.setMaximumHeight(170)
        t.cellChanged.connect(self._seri_hucre_degisti)
        self.seri_tablo = t
        return t

    # -------------------------------------------------------------- veri ↔ UI
    def _yaz(self, alan: str, deger: str) -> None:
        setattr(self.proje.dokuman, alan, deger)

    def _urun_degisti(self, t: str) -> None:
        self.proje.dokuman.urun_adi = t
        # Seri tablosundaki ürün ismi sütununu boşsa otomatik doldurmak için
        # kullanıcıya bırakıyoruz; sadece veriyi güncelliyoruz.

    def _seri_hucre_degisti(self, row: int, col: int) -> None:
        if not (0 <= row < len(self.proje.seriler)):
            return
        it = self.seri_tablo.item(row, col)
        v = it.text() if it else ""
        seri = self.proje.seriler[row]
        if col == 0:
            seri.urun_ismi = v
        elif col == 1:
            seri.seri_no = v
        elif col == 2:
            seri.seri_boyutu_adet = v
        elif col == 3:
            seri.seri_boyutu_kg = v

    def _verilerden_doldur(self) -> None:
        d = self.proje.dokuman
        self.in_firma.setText(d.firma_ismi)
        self.in_urun.setText(d.urun_adi)
        self.in_pvp_no.setText(d.pvp_dokuman_no)
        self.in_pvr_no.setText(d.pvr_dokuman_no)
        self.in_rev_no.setText(d.revizyon_no)
        self.in_rev_tarih.setText(d.revizyon_tarihi)
        self.in_form_no.setText(d.form_no)

        self.seri_tablo.blockSignals(True)
        for r in range(SERI_SAYISI):
            seri = self.proje.seriler[r]
            for c, v in enumerate([seri.urun_ismi, seri.seri_no,
                                   seri.seri_boyutu_adet, seri.seri_boyutu_kg]):
                self.seri_tablo.setItem(r, c, QTableWidgetItem(v))
        self.seri_tablo.blockSignals(False)
