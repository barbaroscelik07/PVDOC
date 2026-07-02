"""
Genel Bilgi modülü (şablon başlığı + Bölüm 5.2 Kapsanan Ürünler / Tablo 2).

Buradaki alanlar şablonun her sayfasındaki başlık/altbilgiyi ve kapsanan
ürün serilerini besler. Doküman no formatı firmaya özgü → kullanıcı elle girer.
Seri sayısı sabit 3 (tasarım kararı); seri no / boyut alanları düzenlenebilir.
"""

from __future__ import annotations

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QLineEdit,
    QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView, QCheckBox,
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

        self.in_firma = satir("Firma İsmi:", "NEUTEC İLAÇ")
        self.in_urun = satir("Ürün Adı:", "örn. Xxx Film Kaplı Tablet")
        self.in_pvp_no = satir("PVP Doküman No:", "örn. AG-PV-xxx")
        self.in_pvr_no = satir("PVR Doküman No:", "örn. AG-PV-xxx-R")
        self.in_rev_no = satir("Revizyon No:", "03")
        self.in_rev_tarih = satir("Revizyon Tarihi:", "16.03.2022")
        self.in_pvp_form_no = satir("PVP Form No:", "N-15-506")
        self.in_pvr_form_no = satir("PVR Form No:", "N-15-507")

        # Alanlar değişince anında veriye yaz
        self.in_firma.textChanged.connect(lambda t: self._yaz("firma_ismi", t))
        self.in_urun.textChanged.connect(self._urun_degisti)
        self.in_pvp_no.textChanged.connect(self._pvp_no_degisti)
        self.in_pvr_no.textChanged.connect(lambda t: self._yaz("pvr_dokuman_no", t))
        self.in_rev_no.textChanged.connect(lambda t: self._yaz("revizyon_no", t))
        self.in_rev_tarih.textChanged.connect(lambda t: self._yaz("revizyon_tarihi", t))
        self.in_pvp_form_no.textChanged.connect(lambda t: self._yaz("pvp_form_no", t))
        self.in_pvr_form_no.textChanged.connect(lambda t: self._yaz("pvr_form_no", t))

        kok.addLayout(izgara)

        # Çift katman seçimi (Tablo 6 türetmesini etkiler)
        katman = QHBoxLayout()
        katman.addWidget(QLabel("Tablet Yapısı:"))
        self.chk_cift = QCheckBox("Çift katmanlı tablet")
        self.chk_cift.setChecked(getattr(self.proje.spek_karti, "cift_katman", False))
        self.chk_cift.toggled.connect(lambda v: setattr(self.proje.spek_karti, "cift_katman", v))
        katman.addWidget(self.chk_cift)
        katman.addWidget(ipucu_etiketi(
            "Çift katmanda Karışım aşamasında Görünüş/Elek/Bulk-Tap her etken için ayrı olur."))
        katman.addStretch(1)
        kok.addLayout(katman)
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
    def _seri_no_otomatik_artir(self, ilk: str) -> None:
        """
        İlk seri no'nun SONUNDAKİ sayıyı +1 yaparak alt 2 satırı doldurur.
        'NI00101-2606-P01' → P02, P03. Sayı bulunamazsa dokunmaz.
        """
        import re
        m = re.search(r"(\d+)(\D*)$", ilk)
        if not m:
            return
        sayi = m.group(1)
        son_ek = m.group(2)
        bas = ilk[:m.start(1)]
        genislik = len(sayi)
        try:
            taban = int(sayi)
        except ValueError:
            return
        self.seri_tablo.blockSignals(True)
        for r in range(1, SERI_SAYISI):
            yeni_sayi = str(taban + r).zfill(genislik)
            yeni = f"{bas}{yeni_sayi}{son_ek}"
            self.proje.seriler[r].seri_no = yeni
            self._hucreyi_ayarla(r, 1, yeni)
        self.seri_tablo.blockSignals(False)

    def _yaz(self, alan: str, deger: str) -> None:
        setattr(self.proje.dokuman, alan, deger)

    def _pvp_no_degisti(self, t: str) -> None:
        self.proje.dokuman.pvp_dokuman_no = t
        # PVR doküman no otomatik: PVP + "-R" (kullanıcı elle değiştirmediyse)
        mevcut_pvr = self.in_pvr_no.text().strip()
        otomatik_onceki = getattr(self, "_otomatik_pvr", "")
        if not mevcut_pvr or mevcut_pvr == otomatik_onceki:
            yeni_pvr = f"{t}-R" if t.strip() else ""
            self._otomatik_pvr = yeni_pvr
            self.in_pvr_no.blockSignals(True)
            self.in_pvr_no.setText(yeni_pvr)
            self.in_pvr_no.blockSignals(False)
            self.proje.dokuman.pvr_dokuman_no = yeni_pvr

    def _urun_degisti(self, t: str) -> None:
        self.proje.dokuman.urun_adi = t
        # Ürün adı yazıldıkça Kapsanan Ürünler tablosundaki 3 satırın
        # "Ürün İsmi" sütununu otomatik güncelle (üzerine yaz).
        for seri in self.proje.seriler:
            seri.urun_ismi = t
        self.seri_tablo.blockSignals(True)
        for r in range(SERI_SAYISI):
            self._hucreyi_ayarla(r, 0, t)
        self.seri_tablo.blockSignals(False)

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
            if row == 0:
                self._seri_no_otomatik_artir(v)
        elif col == 2:
            seri.seri_boyutu_adet = v
            if row == 0:
                self._alt_satirlara_kopyala(2, v, "seri_boyutu_adet")
        elif col == 3:
            seri.seri_boyutu_kg = v
            if row == 0:
                self._alt_satirlara_kopyala(3, v, "seri_boyutu_kg")

    def _alt_satirlara_kopyala(self, col: int, deger: str, alan: str) -> None:
        """İlk satırdaki seri boyutu değerini alttaki 2 satıra kopyalar (üzerine yazar)."""
        self.seri_tablo.blockSignals(True)
        for r in range(1, SERI_SAYISI):
            setattr(self.proje.seriler[r], alan, deger)
            self._hucreyi_ayarla(r, col, deger)
        self.seri_tablo.blockSignals(False)

    def _hucreyi_ayarla(self, row: int, col: int, metin: str) -> None:
        """Var olan item'ı yeniden kullanır; yoksa oluşturur (eski item birikmesini önler)."""
        it = self.seri_tablo.item(row, col)
        if it is None:
            it = QTableWidgetItem(metin)
            self.seri_tablo.setItem(row, col, it)
        else:
            it.setText(metin)

    def _verilerden_doldur(self) -> None:
        d = self.proje.dokuman
        for le, deger in (
            (self.in_firma, d.firma_ismi), (self.in_urun, d.urun_adi),
            (self.in_pvp_no, d.pvp_dokuman_no), (self.in_pvr_no, d.pvr_dokuman_no),
            (self.in_rev_no, d.revizyon_no), (self.in_rev_tarih, d.revizyon_tarihi),
            (self.in_pvp_form_no, d.pvp_form_no), (self.in_pvr_form_no, d.pvr_form_no),
        ):
            le.blockSignals(True)
            le.setText(deger)
            le.blockSignals(False)

        self.seri_tablo.blockSignals(True)
        for r in range(SERI_SAYISI):
            seri = self.proje.seriler[r]
            for c, v in enumerate([seri.urun_ismi, seri.seri_no,
                                   seri.seri_boyutu_adet, seri.seri_boyutu_kg]):
                self._hucreyi_ayarla(r, c, v)
        self.seri_tablo.blockSignals(False)
