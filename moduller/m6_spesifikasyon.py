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
    QComboBox, QDoubleSpinBox, QSpinBox, QPushButton, QTableWidget, QTableWidgetItem,
    QHeaderView, QListWidget, QListWidgetItem, QInputDialog, QMessageBox,
    QCheckBox, QAbstractItemView, QScrollArea, QDialog, QDialogButtonBox,
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
        self.chk_t89 = QCheckBox("Tablo 8/9'u (Serbest Bırakma / Raf Ömrü) otomatik üret")
        self.chk_t89.toggled.connect(lambda v: setattr(self.kart, "tablo89_ekle", v))
        form_satir.addWidget(self.chk_t89)
        kok.addLayout(form_satir)

        # Tablo 8/9 tolerans alanları
        tol_satir = QHBoxLayout()
        tol_satir.addWidget(QLabel("Serbest Bırakma toleransı:"))
        self.in_t8_tol = QLineEdit("±%5")
        self.in_t8_tol.setMaximumWidth(90)
        self.in_t8_tol.textChanged.connect(lambda t: setattr(self.kart, "serbest_birakma_tolerans", t.strip()))
        tol_satir.addWidget(self.in_t8_tol)
        tol_satir.addWidget(QLabel("Raf Ömrü toleransı:"))
        self.in_t9_tol = QLineEdit("±%7.5")
        self.in_t9_tol.setMaximumWidth(90)
        self.in_t9_tol.textChanged.connect(lambda t: setattr(self.kart, "raf_omru_tolerans", t.strip()))
        tol_satir.addWidget(self.in_t9_tol)
        tol_satir.addStretch(1)
        kok.addLayout(tol_satir)

        kok.addWidget(ayirici())

        # Orta alan: solda etkin madde/impurite, sağda test ekleme formu
        orta = QHBoxLayout()
        orta.setSpacing(16)
        orta.addLayout(self._etkin_madde_bolumu(), 2)
        orta.addLayout(self._test_ekleme_bolumu(), 3)
        kok.addLayout(orta)

        kok.addWidget(ayirici())

        # Alt: test tablosu
        kok.addWidget(bolum_etiketi("Tanımlı Testler"))
        kok.addWidget(self._test_tablosu_olustur(), 1)
        kok.addLayout(self._tasima_butonlari())

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

        izgara.addWidget(QLabel("Operasyon No:"), s, 0)
        self.sp_op_no = QSpinBox()
        self.sp_op_no.setRange(0, 99)
        self.sp_op_no.setButtonSymbols(QSpinBox.ButtonSymbols.UpDownArrows)
        izgara.addWidget(self.sp_op_no, s, 1); s += 1

        izgara.addWidget(QLabel("Etkin Madde:"), s, 0)
        self.cmb_em = QComboBox()  # -1 = ürüne ait genel
        izgara.addWidget(self.cmb_em, s, 1); s += 1

        izgara.addWidget(QLabel("Limit Türü:"), s, 0)
        self.cmb_limit = QComboBox()
        for ad, _ in _LIMIT_SECENEKLERI:
            self.cmb_limit.addItem(ad)
        self.cmb_limit.currentIndexChanged.connect(self._limit_turu_degisti)
        izgara.addWidget(self.cmb_limit, s, 1); s += 1

        # Dinamik giriş alanları — QLineEdit ile ondalık BİREBİR korunur.
        self.in_hedef = QLineEdit()
        self.in_alt = QLineEdit()
        self.in_ust = QLineEdit()
        self.in_min = QLineEdit()
        self.in_maks = QLineEdit()
        self.in_tol = QLineEdit()
        self.in_tol.setPlaceholderText("örn. ±%5  (boş bırakılabilir)")
        for le in (self.in_hedef, self.in_alt, self.in_ust, self.in_min, self.in_maks):
            le.setPlaceholderText("sayı (örn. 5,0)")
        # Sabit sonuç: spesifikasyon hücresinde GÖRÜNECEK metin (elle yazılır).
        self.in_metin = QLineEdit()
        self.in_metin.setPlaceholderText("örn. 10.0 mg/f.tab ±%5 (9.5 – 10.5 mg/f.tab)")
        self.in_birim = QLineEdit()
        self.in_birim.setPlaceholderText("mg/f.tab, %, kP, mm …")

        self.lbl_hedef = QLabel("Hedef:"); izgara.addWidget(self.lbl_hedef, s, 0); izgara.addWidget(self.in_hedef, s, 1); s += 1
        self.lbl_tol = QLabel("Tolerans:"); izgara.addWidget(self.lbl_tol, s, 0); izgara.addWidget(self.in_tol, s, 1); s += 1
        self.lbl_alt = QLabel("Alt Limit (veri üretimi):"); izgara.addWidget(self.lbl_alt, s, 0); izgara.addWidget(self.in_alt, s, 1); s += 1
        self.lbl_ust = QLabel("Üst Limit (veri üretimi):"); izgara.addWidget(self.lbl_ust, s, 0); izgara.addWidget(self.in_ust, s, 1); s += 1
        self.lbl_min = QLabel("Minimum:"); izgara.addWidget(self.lbl_min, s, 0); izgara.addWidget(self.in_min, s, 1); s += 1
        self.lbl_maks = QLabel("Maksimum:"); izgara.addWidget(self.lbl_maks, s, 0); izgara.addWidget(self.in_maks, s, 1); s += 1
        self.lbl_metin = QLabel("Sabit Sonuç metni:"); izgara.addWidget(self.lbl_metin, s, 0); izgara.addWidget(self.in_metin, s, 1); s += 1
        self.lbl_birim = QLabel("Birim:"); izgara.addWidget(self.lbl_birim, s, 0); izgara.addWidget(self.in_birim, s, 1); s += 1

        d.addLayout(izgara)
        d.addWidget(ipucu_etiketi(
            "‘Sabit Sonuç metni’ spesifikasyon hücresinde aynen görünür. "
            "Alt/Üst Limit ise PVR sonuç değerlerinin bu aralıkta üretilmesini sağlar."
        ))

        # İşaretler
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

        # Hazır kalıplar + kopyalama
        kalip = QHBoxLayout()
        b_mikro = QPushButton("+ Mikrobiyolojik Kontrol")
        b_mikro.clicked.connect(self._mikrobiyolojik_ekle)
        b_agirlik = QPushButton("+ Ağırlık Tekdüzeliği")
        b_agirlik.clicked.connect(self._agirlik_ekle)
        kalip.addWidget(b_mikro)
        kalip.addWidget(b_agirlik)
        d.addLayout(kalip)

        b_kopyala = QPushButton("Seçili testi başka operasyona kopyala")
        b_kopyala.clicked.connect(self._bolume_kopyala)
        d.addWidget(b_kopyala)
        d.addStretch(1)

        self._limit_turu_degisti()  # ilk durumda alanları ayarla
        return d

    def _test_tablosu_olustur(self) -> QTableWidget:
        t = QTableWidget(0, 7)
        t.setHorizontalHeaderLabels(
            ["Op.", "Operasyon", "Test Adı", "Spesifikasyon", "Tablo Tipi", "IPK", "*"]
        )
        # Kullanıcı sütun genişliklerini elle ayarlayabilsin (Interactive)
        t.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        t.horizontalHeader().setStretchLastSection(True)
        t.setColumnWidth(0, 40)
        t.setColumnWidth(1, 110)
        t.setColumnWidth(2, 220)
        t.setColumnWidth(3, 240)
        t.setColumnWidth(4, 130)
        t.setColumnWidth(5, 45)
        t.setColumnWidth(6, 35)
        t.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        t.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        t.verticalHeader().setVisible(False)
        t.setMinimumHeight(260)
        t.setWordWrap(True)
        self.tablo = t

        # Satır silme için kısayol: çift tık → sor
        t.cellDoubleClicked.connect(self._test_sil_sor)
        return t

    def _tasima_butonlari(self) -> QHBoxLayout:
        """Test sırasını elle değiştirme + operasyona göre otomatik sıralama + düzenle/sil."""
        h = QHBoxLayout()
        b_yukari = QPushButton("▲ Yukarı"); b_yukari.clicked.connect(lambda: self._tasi(-1))
        b_asagi = QPushButton("▼ Aşağı"); b_asagi.clicked.connect(lambda: self._tasi(1))
        b_sirala = QPushButton("Operasyona Göre Sırala")
        b_sirala.clicked.connect(self._operasyona_gore_sirala)
        b_duzenle = QPushButton("Seçiliyi Düzenle")
        b_duzenle.clicked.connect(self._secili_duzenle)
        b_sil = QPushButton("Seçiliyi Sil")
        b_sil.clicked.connect(self._secili_sil)
        h.addWidget(b_yukari); h.addWidget(b_asagi); h.addWidget(b_sirala)
        h.addWidget(b_duzenle); h.addWidget(b_sil); h.addStretch(1)
        return h

    def _secili_sil(self) -> None:
        r = self.tablo.currentRow()
        if 0 <= r < len(self.kart.testler):
            c = QMessageBox.question(self, "Sil", f"'{self.kart.testler[r].ad}' silinsin mi?")
            if c == QMessageBox.StandardButton.Yes:
                del self.kart.testler[r]
                self._tabloyu_yenile()

    def _secili_duzenle(self) -> None:
        """Seçili testin değerlerini forma yükler; tekrar 'Testi Ekle' ile güncellenir."""
        r = self.tablo.currentRow()
        if not (0 <= r < len(self.kart.testler)):
            QMessageBox.information(self, "Düzenle", "Önce tablodan bir test seçin.")
            return
        t = self.kart.testler[r]
        sp = t.spesifikasyon
        self.in_ad.setText(t.ad.rstrip("*"))
        self.cmb_op.setCurrentText(t.operasyon)
        self.sp_op_no.setValue(t.operasyon_no or 0)
        # limit türü
        for i, (_, tur) in enumerate(_LIMIT_SECENEKLERI):
            if tur is sp.limit_turu:
                self.cmb_limit.setCurrentIndex(i); break
        self.in_hedef.setText(sp.hedef_metin)
        self.in_tol.setText(sp.tolerans)
        self.in_alt.setText(sp.alt_metin)
        self.in_ust.setText(sp.ust_metin)
        self.in_min.setText(sp.minimum_metin)
        self.in_maks.setText(sp.maksimum_metin)
        self.in_metin.setText(sp.spesifikasyon_metni or sp.sabit_sonuc)
        self.in_birim.setText(sp.birim)
        self.chk_ipk.setChecked(t.ipk)
        self.chk_yildiz.setChecked(t.yildizli)
        # eskisini sil ki 'Testi Ekle' güncellenmiş halini eklesin
        del self.kart.testler[r]
        self._tabloyu_yenile()
        QMessageBox.information(self, "Düzenle",
                               "Test bilgileri forma yüklendi. Değiştirip 'Testi Ekle' ile tekrar ekleyin.")

    def _tasi(self, yon: int) -> None:
        r = self.tablo.currentRow()
        yeni = r + yon
        if 0 <= r < len(self.kart.testler) and 0 <= yeni < len(self.kart.testler):
            self.kart.testler[r], self.kart.testler[yeni] = self.kart.testler[yeni], self.kart.testler[r]
            self._tabloyu_yenile()
            self.tablo.selectRow(yeni)

    def _operasyona_gore_sirala(self) -> None:
        """Testleri operasyon numarasına göre stabil sıralar (giriş sırası korunur)."""
        self.kart.testler.sort(key=lambda t: (t.operasyon_no or 99))
        self._tabloyu_yenile()

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
        # Tablo 8/9 ayarları
        self.chk_t89.blockSignals(True); self.chk_t89.setChecked(self.kart.tablo89_ekle); self.chk_t89.blockSignals(False)
        self.in_t8_tol.blockSignals(True); self.in_t8_tol.setText(self.kart.serbest_birakma_tolerans); self.in_t8_tol.blockSignals(False)
        self.in_t9_tol.blockSignals(True); self.in_t9_tol.setText(self.kart.raf_omru_tolerans); self.in_t9_tol.blockSignals(False)
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
        # operasyon seçilince no öner
        try:
            self.cmb_op.currentTextChanged.disconnect(self._op_no_oner)
        except TypeError:
            pass
        self.cmb_op.currentTextChanged.connect(self._op_no_oner)
        self._op_no_oner(self.cmb_op.currentText())

    # Operasyon adı → tipik operasyon numarası (kullanıcı değiştirebilir)
    _OP_NO = {"Karıştırma": 2, "Tablet Baskı": 3, "Film Kaplama": 4,
              "Dolum": 3, "Blisterleme": 5}

    def _op_no_oner(self, op_adi: str) -> None:
        if hasattr(self, "sp_op_no"):
            self.sp_op_no.setValue(self._OP_NO.get(op_adi, 0))

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

        dlg = ImpuriteDialog(self, operasyonlar=[self.cmb_op.itemText(i)
                                                 for i in range(self.cmb_op.count())])
        if not dlg.exec():
            return
        v = dlg.degerler()
        if not v["ad"]:
            return

        # İmpurite, etkin maddeye eklenir. İlgili Bileşikler grubu çıktıda
        # (Tablo 6/8/9) bu listeden GRUPLU üretilir — ayrı test OLUŞTURULMAZ.
        maks = (v["maks"] or "").strip()
        te = maks.upper().replace(" ", "") in ("T.E.", "T.E", "TE")
        em.impuriteler.append(Impurite(
            ad=v["ad"],
            limit_metni=v["limit"] or (f"Maksimum %{maks}" if maks and not te else ("Maksimum T.E." if te else "")),
            maksimum_deger=self._sayi(maks),
            operasyon=v["operasyon"],
            operasyon_no=self._OP_NO.get(v["operasyon"], 0),
            yildizli=v["yildiz"],
            te=te,
        ))
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
            "tol": tur is LimitTuru.ARALIK,
            "alt": tur in (LimitTuru.ARALIK, LimitTuru.METIN),
            "ust": tur in (LimitTuru.ARALIK, LimitTuru.METIN),
            "min": tur is LimitTuru.MINIMUM,
            "maks": tur is LimitTuru.MAKSIMUM,
            "metin": True,   # Sabit sonuç metni HER ZAMAN görünür (spek hücresinde gösterilir)
            "birim": tur in (LimitTuru.ARALIK, LimitTuru.MINIMUM, LimitTuru.MAKSIMUM),
        }
        for ad, vis in goster.items():
            getattr(self, f"lbl_{ad}").setVisible(vis)
            getattr(self, f"in_{ad}").setVisible(vis)

    @staticmethod
    def _sayi(metin: str):
        """Ham metni float'a çevirir (veri üretimi için). Boş/geçersizse None."""
        metin = metin.strip().replace(",", ".")
        if not metin:
            return None
        try:
            return float(metin)
        except ValueError:
            return None

    # ---- test ekle (1.3 + 1.6) ----
    def _test_ekle(self) -> None:
        ad = self.in_ad.text().strip()
        if not ad:
            QMessageBox.information(self, "Test", "Test adı boş olamaz.")
            return

        tur = _LIMIT_SECENEKLERI[self.cmb_limit.currentIndex()][1]
        spek = Spesifikasyon(limit_turu=tur, birim=self.in_birim.text().strip())
        # Alt/Üst limit her tipte veri üretimi için saklanabilir
        spek.alt_metin = self.in_alt.text().strip()
        spek.ust_metin = self.in_ust.text().strip()
        spek.alt_limit = self._sayi(spek.alt_metin)
        spek.ust_limit = self._sayi(spek.ust_metin)
        if tur is LimitTuru.ARALIK:
            spek.hedef_metin = self.in_hedef.text().strip()
            spek.tolerans = self.in_tol.text().strip()
            spek.hedef_deger = self._sayi(spek.hedef_metin)
        elif tur is LimitTuru.MINIMUM:
            spek.minimum_metin = self.in_min.text().strip()
            spek.minimum_deger = self._sayi(spek.minimum_metin)
        elif tur is LimitTuru.MAKSIMUM:
            spek.maksimum_metin = self.in_maks.text().strip()
            spek.maksimum_deger = self._sayi(spek.maksimum_metin)

        # Sabit sonuç metni doluysa spesifikasyon hücresinde AYNEN gösterilir.
        sabit = self.in_metin.text().strip()
        if sabit:
            spek.spesifikasyon_metni = sabit
            spek.sabit_sonuc = sabit

        test = Test(
            ad=ad,
            operasyon=self.cmb_op.currentText(),
            operasyon_no=self.sp_op_no.value(),
            spesifikasyon=spek,
            tablo_tipi=tablo_tipini_belirle(ad),     # otomatik (1.1)
            etkin_madde_index=self.cmb_em.currentData(),
            ipk=self.chk_ipk.isChecked(),
            yildizli=self.chk_yildiz.isChecked(),
        )
        self.kart.testler.append(test)
        self._tabloyu_yenile()

        # formu sıfırla (operasyon/form korunur, hızlı ardışık giriş için)
        for le in (self.in_ad, self.in_metin, self.in_hedef, self.in_alt,
                   self.in_ust, self.in_min, self.in_maks, self.in_tol):
            le.clear()
        self.chk_ipk.setChecked(False)
        self.chk_yildiz.setChecked(False)

    # ---- hazır kalıplar ----
    def _mikrobiyolojik_ekle(self) -> None:
        """Mikrobiyolojik Kontrol: sabit 3 alt satır, sonuç hep 'Uygun'."""
        op = self.cmb_op.currentText()
        test = Test(
            ad="Mikrobiyolojik Kontrol",
            operasyon=op,
            operasyon_no=self.sp_op_no.value(),
            tablo_tipi=TabloTipi.MATRIS,
            mikrobiyolojik=True,
            ipk=self.chk_ipk.isChecked(),
            yildizli=self.chk_yildiz.isChecked(),
            spesifikasyon=Spesifikasyon(limit_turu=LimitTuru.METIN, sabit_sonuc="Uygun"),
            alt_satirlar=[
                ("-Toplam Aerobik Mikroorganizma Sayısı", "≤10³ cfu/g"),
                ("-Küf ve Maya Sayısı", "≤10² cfu/g"),
                ("-E. coli", "0 cfu/g"),
            ],
        )
        self.kart.testler.append(test)
        self._tabloyu_yenile()

    def _agirlik_ekle(self) -> None:
        """
        Ağırlık Tekdüzeliği (resimdeki yapı):
          satır 1: 'Ağırlık Tekdüzeliği' — sağı BOŞ
          satır 2: '—20 tablette ... maksimum 2 tanesi ... sapabilir.' — sağında 1. limit
          satır 3: '—Hiçbir tablet bu limitten sapmamalıdır' — sağında 2. limit
        Sonuç verisi için ana alt/üst limit (Ortalama Ağırlık ile eşleşir).
        """
        dlg = AgirlikDialog(self, operasyon=self.cmb_op.currentText())
        if not dlg.exec():
            return
        v = dlg.degerler()
        if not (v["alt"] and v["ust"]):
            QMessageBox.information(self, "Ağırlık Tekdüzeliği",
                                    "Alt ve üst limit (örn. 270.75 / 299.25) zorunludur.")
            return
        test = Test(
            ad="Ağırlık Tekdüzeliği",
            operasyon=self.cmb_op.currentText(),
            operasyon_no=self.sp_op_no.value(),
            tablo_tipi=TabloTipi.AGIRLIK_TEKDUZELIGI,
            ipk=self.chk_ipk.isChecked(),
            yildizli=self.chk_yildiz.isChecked(),
            spesifikasyon=Spesifikasyon(limit_turu=LimitTuru.ARALIK,
                                        alt_metin=v["alt"], ust_metin=v["ust"],
                                        alt_limit=self._sayi(v["alt"]), ust_limit=self._sayi(v["ust"]),
                                        birim="mg"),
            aciklama_etiketi="—20 tablette tek tek tabletlerden maksimum 2 tanesi bu limitten sapabilir.",
            aciklama_spek=v["limit1"],
            aciklama2_etiketi="—Hiçbir tablet bu limitten sapmamalıdır",
            aciklama2_spek=v["limit2"],
        )
        self.kart.testler.append(test)
        self._tabloyu_yenile()

    def _bolume_kopyala(self) -> None:
        """Seçili testi başka bir operasyona kopyalar (yıldız HARİÇ)."""
        r = self.tablo.currentRow()
        if not (0 <= r < len(self.kart.testler)):
            QMessageBox.information(self, "Kopyala", "Önce tablodan bir test seçin.")
            return
        kaynak = self.kart.testler[r]
        operasyonlar = [self.cmb_op.itemText(i) for i in range(self.cmb_op.count())]
        hedef, ok = QInputDialog.getItem(self, "Başka Operasyona Kopyala",
                                         "Hedef operasyon:", operasyonlar, 0, False)
        if not ok:
            return
        import copy as _copy
        yeni = _copy.deepcopy(kaynak)
        yeni.operasyon = hedef
        yeni.operasyon_no = self._OP_NO.get(hedef, kaynak.operasyon_no)
        # Yıldız (* validasyon serilerinde) durumunu her kopyada sor
        c = QMessageBox.question(
            self, "Validasyon (*)",
            f"'{hedef}' bölümünde bu test '*' (validasyon serilerinde) olarak işaretlensin mi?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        yeni.yildizli = (c == QMessageBox.StandardButton.Yes)
        yeni.sonuc_verisi = {}
        self.kart.testler.append(yeni)
        self._operasyona_gore_sirala()  # kopya sonrası otomatik düzenle

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
                test.ad + ("*" if test.yildizli else ""),
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


class ImpuriteDialog(QDialog):
    """İmpurite eklerken ad, limit, operasyon ve yıldız bilgisini tek seferde toplar."""

    def __init__(self, parent, operasyonlar: list[str]):
        super().__init__(parent)
        self.setWindowTitle("İmpurite Ekle")
        self.setStyleSheet(MODUL_STIL)
        self.setMinimumWidth(360)

        izg = QGridLayout(self)
        izg.addWidget(QLabel("İmpurite adı:"), 0, 0)
        self.in_ad = QLineEdit(); self.in_ad.setPlaceholderText("örn. imp. a / Toplam imp.")
        izg.addWidget(self.in_ad, 0, 1)

        izg.addWidget(QLabel("Limit metni:"), 1, 0)
        self.in_limit = QLineEdit(); self.in_limit.setPlaceholderText("örn. Maksimum %1.0 (boşsa otomatik)")
        izg.addWidget(self.in_limit, 1, 1)

        izg.addWidget(QLabel("Maksimum (sayı):"), 2, 0)
        self.in_maks = QLineEdit(); self.in_maks.setPlaceholderText("örn. 1,0")
        izg.addWidget(self.in_maks, 2, 1)

        izg.addWidget(QLabel("Operasyon:"), 3, 0)
        self.cmb_op = QComboBox(); self.cmb_op.addItems(operasyonlar)
        izg.addWidget(self.cmb_op, 3, 1)

        self.chk_yildiz = QCheckBox("* Validasyon serilerinde uygulanır")
        izg.addWidget(self.chk_yildiz, 4, 0, 1, 2)

        btn = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btn.accepted.connect(self.accept)
        btn.rejected.connect(self.reject)
        izg.addWidget(btn, 5, 0, 1, 2)

    def degerler(self) -> dict:
        return {
            "ad": self.in_ad.text().strip(),
            "limit": self.in_limit.text().strip(),
            "maks": self.in_maks.text().strip(),
            "operasyon": self.cmb_op.currentText(),
            "yildiz": self.chk_yildiz.isChecked(),
        }


class AgirlikDialog(QDialog):
    """Ağırlık Tekdüzeliği: 2 limit çifti + sonuç üretimi için alt/üst limit."""

    def __init__(self, parent, operasyon: str):
        super().__init__(parent)
        self.setWindowTitle("Ağırlık Tekdüzeliği Ekle")
        self.setStyleSheet(MODUL_STIL)
        self.setMinimumWidth(460)

        izg = QGridLayout(self)
        s = 0
        izg.addWidget(QLabel("Sonuç üretimi için (Ortalama Ağırlık ile eşleşir):"), s, 0, 1, 2); s += 1
        izg.addWidget(QLabel("Alt Limit:"), s, 0)
        self.in_alt = QLineEdit(); self.in_alt.setPlaceholderText("örn. 270.75")
        izg.addWidget(self.in_alt, s, 1); s += 1
        izg.addWidget(QLabel("Üst Limit:"), s, 0)
        self.in_ust = QLineEdit(); self.in_ust.setPlaceholderText("örn. 299.25")
        izg.addWidget(self.in_ust, s, 1); s += 1

        izg.addWidget(QLabel("— maksimum 2 tanesi sapabilir → sağ değer:"), s, 0)
        self.in_l1 = QLineEdit(); self.in_l1.setPlaceholderText("örn. ≤ 270.75 veya ≥ 299.25 mg")
        izg.addWidget(self.in_l1, s, 1); s += 1
        izg.addWidget(QLabel("— hiçbir tablet sapmamalıdır → sağ değer:"), s, 0)
        self.in_l2 = QLineEdit(); self.in_l2.setPlaceholderText("örn. ≤ 256.50 veya ≥ 313.50 mg")
        izg.addWidget(self.in_l2, s, 1); s += 1

        btn = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btn.accepted.connect(self.accept)
        btn.rejected.connect(self.reject)
        izg.addWidget(btn, s, 0, 1, 2)

    def degerler(self) -> dict:
        return {
            "alt": self.in_alt.text().strip(),
            "ust": self.in_ust.text().strip(),
            "limit1": self.in_l1.text().strip(),
            "limit2": self.in_l2.text().strip(),
        }
