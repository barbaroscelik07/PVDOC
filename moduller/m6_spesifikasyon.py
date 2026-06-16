"""
Spesifikasyon modülü (Bölüm 8 / Tablo 6-7).

Kapsam (Faz 1, Görev 1.3–1.6):
- 1.3 Tek test girişi: limit türüne göre dinamik alanlar (aralık/min/maks/metin/bilgi).
- 1.4 Dinamik etkin madde + impurite ekle/çıkar.
- 1.5 Ürün formu seçimi → operasyon listesini süzme (tablet/film/kapsül).
- 1.6 Test başına "IPK mi?" ve "* (validasyon serilerinde)" işaretleme;
      spek kartını kütüphaneye kaydet / kütüphaneden yükle.

Bu modül ProjeVerisi.spek_karti üzerinde çalışır; değişiklikler doğrudan
merkezi veriye yansır. Tablo tipi test adından otomatik atanır (core.test_tipi).
"""

from __future__ import annotations

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QLineEdit,
    QComboBox, QDoubleSpinBox, QPushButton, QTableWidget, QTableWidgetItem,
    QHeaderView, QListWidget, QListWidgetItem, QInputDialog, QMessageBox,
    QCheckBox, QAbstractItemView, QScrollArea,
)
from PyQt6.QtCore import Qt

from core.models import (
    ProjeVerisi, SpekKarti, UrunFormu, EtkinMadde, Impurite, Test,
    Spesifikasyon, LimitTuru, TabloTipi,
)
from core.test_tipi import tablo_tipini_belirle, ipk_testi_mi
from core import spek_kutuphanesi as lib
from moduller.widget_yardimcilari import (
    MODUL_STIL, baslik_etiketi, bolum_etiketi, ipucu_etiketi, ayirici,
)


# Limit türü görünen ad <-> enum
_LIMIT_SECENEKLERI = [
    ("Aralık (alt–üst)", LimitTuru.ARALIK),
    ("Minimum", LimitTuru.MINIMUM),
    ("Maksimum", LimitTuru.MAKSIMUM),
    ("Metin (örn. Pozitif)", LimitTuru.METIN),
    ("Bilgi amaçlıdır", LimitTuru.BILGI),
]


class SpekModulu(QWidget):
    """Spesifikasyon giriş paneli."""

    def __init__(self, proje: ProjeVerisi):
        super().__init__()
        self.proje = proje
        self.setObjectName("panel")
        self.setStyleSheet(MODUL_STIL)
        self._arayuzu_kur()
        self._verilerden_doldur()

    # =====================================================================
    # Arayüz
    # =====================================================================
    def _arayuzu_kur(self) -> None:
        kok = QVBoxLayout(self)
        kok.setContentsMargins(22, 22, 22, 22)
        kok.setSpacing(14)

        # Başlık + kart işlemleri
        ust = QHBoxLayout()
        ust.addWidget(baslik_etiketi("Spesifikasyonlar (Tablo 6 / 7)"))
        ust.addStretch(1)
        btn_yukle = QPushButton("Karttan Yükle")
        btn_yukle.clicked.connect(self._karttan_yukle)
        btn_kaydet = QPushButton("Kart Olarak Kaydet")
        btn_kaydet.setObjectName("birincil")
        btn_kaydet.clicked.connect(self._kart_kaydet)
        ust.addWidget(btn_yukle)
        ust.addWidget(btn_kaydet)
        kok.addLayout(ust)

        # Ürün formu (1.5)
        form_satir = QHBoxLayout()
        form_satir.addWidget(QLabel("Ürün Formu:"))
        self.cmb_form = QComboBox()
        for f in UrunFormu:
            self.cmb_form.addItem(f.value, f)
        self.cmb_form.currentIndexChanged.connect(self._form_degisti)
        form_satir.addWidget(self.cmb_form)
        form_satir.addStretch(1)
        kok.addLayout(form_satir)

        kok.addWidget(ayirici())

        # Orta alan: solda etkin madde/impurite, sağda test ekleme formu
        orta = QHBoxLayout()
        orta.setSpacing(16)
        orta.addLayout(self._etkin_madde_bolumu(), 1)
        orta.addLayout(self._test_ekleme_bolumu(), 1)
        kok.addLayout(orta)

        kok.addWidget(ayirici())

        # Alt: test tablosu
        kok.addWidget(bolum_etiketi("Tanımlı Testler"))
        kok.addWidget(self._test_tablosu_olustur(), 1)

    # ---- sol: etkin madde + impurite (1.4) -------------------------------
    def _etkin_madde_bolumu(self) -> QVBoxLayout:
        d = QVBoxLayout()
        d.setSpacing(8)
        d.addWidget(bolum_etiketi("Etkin Maddeler"))

        self.liste_em = QListWidget()
        self.liste_em.currentRowChanged.connect(self._em_secildi)
        d.addWidget(self.liste_em)

        em_btn = QHBoxLayout()
        b_ekle = QPushButton("+ Etkin Madde")
        b_ekle.clicked.connect(self._em_ekle)
        b_sil = QPushButton("− Sil")
        b_sil.setObjectName("tehlike")
        b_sil.clicked.connect(self._em_sil)
        em_btn.addWidget(b_ekle)
        em_btn.addWidget(b_sil)
        d.addLayout(em_btn)

        d.addWidget(ipucu_etiketi("Seçili etkin maddenin impuriteleri:"))
        self.liste_imp = QListWidget()
        self.liste_imp.setMaximumHeight(120)
        d.addWidget(self.liste_imp)

        imp_btn = QHBoxLayout()
        bi_ekle = QPushButton("+ İmpurite")
        bi_ekle.clicked.connect(self._imp_ekle)
        bi_sil = QPushButton("− Sil")
        bi_sil.setObjectName("tehlike")
        bi_sil.clicked.connect(self._imp_sil)
        imp_btn.addWidget(bi_ekle)
        imp_btn.addWidget(bi_sil)
        d.addLayout(imp_btn)
        return d

    # ---- sağ: test ekleme formu (1.3) ------------------------------------
    def _test_ekleme_bolumu(self) -> QVBoxLayout:
        d = QVBoxLayout()
        d.setSpacing(8)
        d.addWidget(bolum_etiketi("Yeni Test Ekle"))

        izgara = QGridLayout()
        izgara.setHorizontalSpacing(10)
        izgara.setVerticalSpacing(8)
        s = 0

        izgara.addWidget(QLabel("Test Adı:"), s, 0)
        self.in_ad = QLineEdit()
        self.in_ad.setPlaceholderText("örn. Etkin madde 1 Miktar Tayini")
        izgara.addWidget(self.in_ad, s, 1); s += 1

        izgara.addWidget(QLabel("Operasyon:"), s, 0)
        self.cmb_op = QComboBox()  # form filtresine göre dolar (1.5)
        izgara.addWidget(self.cmb_op, s, 1); s += 1

        izgara.addWidget(QLabel("Etkin Madde:"), s, 0)
        self.cmb_em = QComboBox()  # -1 = ürüne ait genel
        izgara.addWidget(self.cmb_em, s, 1); s += 1

        izgara.addWidget(QLabel("Limit Türü:"), s, 0)
        self.cmb_limit = QComboBox()
        for ad, _ in _LIMIT_SECENEKLERI:
            self.cmb_limit.addItem(ad)
        self.cmb_limit.currentIndexChanged.connect(self._limit_turu_degisti)
        izgara.addWidget(self.cmb_limit, s, 1); s += 1

        # Dinamik sayısal alanlar
        self.in_hedef = self._spin()
        self.in_alt = self._spin()
        self.in_ust = self._spin()
        self.in_min = self._spin()
        self.in_maks = self._spin()
        self.in_metin = QLineEdit()
        self.in_metin.setPlaceholderText("örn. Beyaz renkli toz / Pozitif / Uygun")
        self.in_birim = QLineEdit()
        self.in_birim.setPlaceholderText("mg/f.tab, %, kP, mm, dakika …")

        self.lbl_hedef = QLabel("Hedef:"); izgara.addWidget(self.lbl_hedef, s, 0); izgara.addWidget(self.in_hedef, s, 1); s += 1
        self.lbl_alt = QLabel("Alt Limit:"); izgara.addWidget(self.lbl_alt, s, 0); izgara.addWidget(self.in_alt, s, 1); s += 1
        self.lbl_ust = QLabel("Üst Limit:"); izgara.addWidget(self.lbl_ust, s, 0); izgara.addWidget(self.in_ust, s, 1); s += 1
        self.lbl_min = QLabel("Minimum:"); izgara.addWidget(self.lbl_min, s, 0); izgara.addWidget(self.in_min, s, 1); s += 1
        self.lbl_maks = QLabel("Maksimum:"); izgara.addWidget(self.lbl_maks, s, 0); izgara.addWidget(self.in_maks, s, 1); s += 1
        self.lbl_metin = QLabel("Sabit Sonuç:"); izgara.addWidget(self.lbl_metin, s, 0); izgara.addWidget(self.in_metin, s, 1); s += 1
        self.lbl_birim = QLabel("Birim:"); izgara.addWidget(self.lbl_birim, s, 0); izgara.addWidget(self.in_birim, s, 1); s += 1

        d.addLayout(izgara)

        # İşaretler (1.6)
        isaret = QHBoxLayout()
        self.chk_ipk = QCheckBox("IPK testi (Tablo 7)")
        self.chk_yildiz = QCheckBox("* Validasyon serilerinde")
        isaret.addWidget(self.chk_ipk)
        isaret.addWidget(self.chk_yildiz)
        isaret.addStretch(1)
        d.addLayout(isaret)

        b_test_ekle = QPushButton("Testi Ekle")
        b_test_ekle.setObjectName("birincil")
        b_test_ekle.clicked.connect(self._test_ekle)
        d.addWidget(b_test_ekle)
        d.addStretch(1)

        self._limit_turu_degisti()  # ilk durumda alanları ayarla
        return d

    def _spin(self) -> QDoubleSpinBox:
        sp = QDoubleSpinBox()
        sp.setDecimals(3)
        sp.setRange(-1_000_000, 1_000_000)
        sp.setButtonSymbols(QDoubleSpinBox.ButtonSymbols.NoButtons)
        return sp

    def _test_tablosu_olustur(self) -> QTableWidget:
        t = QTableWidget(0, 7)
        t.setHorizontalHeaderLabels(
            ["Op.", "Operasyon", "Test Adı", "Spesifikasyon", "Tablo Tipi", "IPK", "*"]
        )
        t.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        t.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        t.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        t.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        t.verticalHeader().setVisible(False)
        self.tablo = t

        # Satır silme için kısayol: çift tık → sor
        t.cellDoubleClicked.connect(self._test_sil_sor)
        return t

    # =====================================================================
    # Veri ↔ UI
    # =====================================================================
    @property
    def kart(self) -> SpekKarti:
        return self.proje.spek_karti

    def _verilerden_doldur(self) -> None:
        """ProjeVerisi'nden tüm widget'ları doldurur."""
        # form
        idx = self.cmb_form.findData(self.kart.urun_formu)
        if idx >= 0:
            self.cmb_form.setCurrentIndex(idx)
        self._operasyonlari_yenile()
        self._em_listesini_yenile()
        self._em_combo_yenile()
        self._tabloyu_yenile()

    # ---- ürün formu (1.5) ----
    def _form_degisti(self) -> None:
        f = self.cmb_form.currentData()
        if isinstance(f, UrunFormu):
            self.kart.urun_formu = f
            self.proje.urun_formu = f
        self._operasyonlari_yenile()

    def _operasyonlari_yenile(self) -> None:
        f = self.cmb_form.currentData()
        self.cmb_op.clear()
        if isinstance(f, UrunFormu):
            self.cmb_op.addItems(f.operasyonlar)

    # ---- etkin madde (1.4) ----
    def _em_listesini_yenile(self) -> None:
        self.liste_em.clear()
        for em in self.kart.etkin_maddeler:
            self.liste_em.addItem(em.ad)
        if self.kart.etkin_maddeler:
            self.liste_em.setCurrentRow(0)
        self._imp_listesini_yenile()

    def _em_combo_yenile(self) -> None:
        self.cmb_em.clear()
        self.cmb_em.addItem("(Ürüne ait / genel)", -1)
        for i, em in enumerate(self.kart.etkin_maddeler):
            self.cmb_em.addItem(em.ad, i)

    def _em_ekle(self) -> None:
        ad, ok = QInputDialog.getText(self, "Etkin Madde", "Etkin madde adı:")
        if ok and ad.strip():
            self.kart.etkin_maddeler.append(EtkinMadde(ad=ad.strip()))
            self._em_listesini_yenile()
            self._em_combo_yenile()

    def _em_sil(self) -> None:
        r = self.liste_em.currentRow()
        if 0 <= r < len(self.kart.etkin_maddeler):
            del self.kart.etkin_maddeler[r]
            self._em_listesini_yenile()
            self._em_combo_yenile()

    def _em_secildi(self, _row: int) -> None:
        self._imp_listesini_yenile()

    def _secili_em(self) -> EtkinMadde | None:
        r = self.liste_em.currentRow()
        if 0 <= r < len(self.kart.etkin_maddeler):
            return self.kart.etkin_maddeler[r]
        return None

    # ---- impurite (1.4) ----
    def _imp_listesini_yenile(self) -> None:
        self.liste_imp.clear()
        em = self._secili_em()
        if em:
            for imp in em.impuriteler:
                etiket = f"{imp.ad}  —  {imp.limit_metni or ('Maksimum %'+str(imp.maksimum_deger) if imp.maksimum_deger is not None else '')}"
                self.liste_imp.addItem(etiket.strip(" —"))

    def _imp_ekle(self) -> None:
        em = self._secili_em()
        if not em:
            QMessageBox.information(self, "İmpurite", "Önce bir etkin madde seçin.")
            return
        ad, ok = QInputDialog.getText(self, "İmpurite", "İmpurite adı (örn. imp. a):")
        if not (ok and ad.strip()):
            return
        limit, ok2 = QInputDialog.getText(self, "İmpurite Limiti", "Limit (örn. Maksimum %1.0):")
        em.impuriteler.append(Impurite(ad=ad.strip(), limit_metni=limit.strip() if ok2 else ""))
        self._imp_listesini_yenile()

    def _imp_sil(self) -> None:
        em = self._secili_em()
        r = self.liste_imp.currentRow()
        if em and 0 <= r < len(em.impuriteler):
            del em.impuriteler[r]
            self._imp_listesini_yenile()

    # ---- limit türü (1.3) ----
    def _limit_turu_degisti(self) -> None:
        tur = _LIMIT_SECENEKLERI[self.cmb_limit.currentIndex()][1]
        goster = {
            "hedef": tur is LimitTuru.ARALIK,
            "alt": tur is LimitTuru.ARALIK,
            "ust": tur is LimitTuru.ARALIK,
            "min": tur is LimitTuru.MINIMUM,
            "maks": tur is LimitTuru.MAKSIMUM,
            "metin": tur is LimitTuru.METIN,
            "birim": tur in (LimitTuru.ARALIK, LimitTuru.MINIMUM, LimitTuru.MAKSIMUM),
        }
        for ad, vis in goster.items():
            getattr(self, f"lbl_{ad}").setVisible(vis)
            getattr(self, f"in_{ad}").setVisible(vis)

    # ---- test ekle (1.3 + 1.6) ----
    def _test_ekle(self) -> None:
        ad = self.in_ad.text().strip()
        if not ad:
            QMessageBox.information(self, "Test", "Test adı boş olamaz.")
            return

        tur = _LIMIT_SECENEKLERI[self.cmb_limit.currentIndex()][1]
        spek = Spesifikasyon(limit_turu=tur, birim=self.in_birim.text().strip())
        if tur is LimitTuru.ARALIK:
            spek.hedef_deger = self.in_hedef.value()
            spek.alt_limit = self.in_alt.value()
            spek.ust_limit = self.in_ust.value()
        elif tur is LimitTuru.MINIMUM:
            spek.minimum_deger = self.in_min.value()
        elif tur is LimitTuru.MAKSIMUM:
            spek.maksimum_deger = self.in_maks.value()
        elif tur is LimitTuru.METIN:
            spek.sabit_sonuc = self.in_metin.text().strip()

        test = Test(
            ad=ad,
            operasyon=self.cmb_op.currentText(),
            spesifikasyon=spek,
            tablo_tipi=tablo_tipini_belirle(ad),     # otomatik (1.1)
            etkin_madde_index=self.cmb_em.currentData(),
            ipk=self.chk_ipk.isChecked(),
            yildizli=self.chk_yildiz.isChecked(),
        )
        self.kart.testler.append(test)
        self._tabloyu_yenile()

        # formu sıfırla (operasyon/form korunur, hızlı ardışık giriş için)
        self.in_ad.clear()
        self.in_metin.clear()
        self.chk_ipk.setChecked(False)
        self.chk_yildiz.setChecked(False)

    def _test_sil_sor(self, row: int, _col: int) -> None:
        if not (0 <= row < len(self.kart.testler)):
            return
        ad = self.kart.testler[row].ad
        c = QMessageBox.question(self, "Test Sil", f"'{ad}' silinsin mi?")
        if c == QMessageBox.StandardButton.Yes:
            del self.kart.testler[row]
            self._tabloyu_yenile()

    def _tabloyu_yenile(self) -> None:
        self.tablo.setRowCount(0)
        for test in self.kart.testler:
            r = self.tablo.rowCount()
            self.tablo.insertRow(r)
            degerler = [
                str(test.operasyon_no or ""),
                test.operasyon,
                test.ad,
                test.spesifikasyon.metni_olustur(),
                test.tablo_tipi.value,
                "E" if test.ipk else "",
                "*" if test.yildizli else "",
            ]
            for c, v in enumerate(degerler):
                it = QTableWidgetItem(v)
                if c in (0, 5, 6):
                    it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.tablo.setItem(r, c, it)

    # =====================================================================
    # Spek kartı kaydet / yükle (1.6)
    # =====================================================================
    def _kart_kaydet(self) -> None:
        varsayilan = self.kart.kart_adi or self.proje.dokuman.urun_adi or "Yeni Kart"
        ad, ok = QInputDialog.getText(self, "Kart Olarak Kaydet", "Kart adı:", text=varsayilan)
        if not (ok and ad.strip()):
            return
        self.kart.kart_adi = ad.strip()
        try:
            if lib.kart_var_mi(self.kart.kart_adi):
                c = QMessageBox.question(self, "Üzerine Yaz",
                                         "Aynı adlı kart var. Üzerine yazılsın mı?")
                if c != QMessageBox.StandardButton.Yes:
                    return
            lib.karti_kaydet(self.kart)
            QMessageBox.information(self, "Kaydedildi", f"'{self.kart.kart_adi}' kütüphaneye kaydedildi.")
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Kart kaydedilemedi:\n{e}")

    def _karttan_yukle(self) -> None:
        kartlar = lib.kartlari_listele()
        if not kartlar:
            QMessageBox.information(self, "Karttan Yükle", "Kütüphanede kayıtlı kart yok.")
            return
        adlar = [f"{k['kart_adi']}  ({k['test_sayisi']} test)" for k in kartlar]
        sec, ok = QInputDialog.getItem(self, "Karttan Yükle", "Kart seç:", adlar, 0, False)
        if not ok:
            return
        secilen = kartlar[adlar.index(sec)]
        try:
            yeni = lib.karti_yukle(secilen["kart_adi"])
            self.proje.spek_karti = yeni
            self.proje.urun_formu = yeni.urun_formu
            self._verilerden_doldur()
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Kart yüklenemedi:\n{e}")
