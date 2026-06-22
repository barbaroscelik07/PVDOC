"""
Spesifikasyon modülü — Excel benzeri doğrudan tablo girişi.

Kullanıcı kararı: form yerine Tablo 6'yı doğrudan düzenle.
Sütunlar: Op No | Operasyon | Test Adı | Spesifikasyon | Alt Limit | Üst Limit | Sonuç Tipi | *
- Hücrelere doğrudan yazılır (Word'deki gibi).
- "Sonuç Tipi" test adından otomatik atanır (core.test_tipi), dropdown'dan değişir.
- Etkin madde bilgisi test adına gömülü ("Etkin madde 1 Miktar Tayini").
- Özel testler için hazır butonlar: + Mikrobiyolojik, + Ağırlık Tekdüzeliği, + İlgili Bileşikler.
- Program sonuç tablolarını alt/üst limit veya spesifikasyon metninden üretir.
"""

from __future__ import annotations

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QLineEdit,
    QComboBox, QPushButton, QTableWidget, QTableWidgetItem,
    QHeaderView, QInputDialog, QMessageBox, QCheckBox, QAbstractItemView,
    QDialog, QDialogButtonBox,
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor, QBrush

from core.models import (
    ProjeVerisi, SpekKarti, UrunFormu, EtkinMadde, Impurite, Test,
    Spesifikasyon, LimitTuru, TabloTipi,
)
from core.test_tipi import tablo_tipini_belirle, ipk_testi_mi
from moduller.widget_yardimcilari import (
    MODUL_STIL, baslik_etiketi, bolum_etiketi, ipucu_etiketi, ayirici,
)


# Sonuç tipi dropdown: görünen ad <-> enum
_TIP_SECENEKLERI = [
    ("Görünüş/Teşhis (tek sonuç)", TabloTipi.TEK_SONUC),
    ("Miktar/İmpurite (Numune 1-2)", TabloTipi.IKI_NUMUNE),
    ("Karışım Tekdüzeliği (10 numune)", TabloTipi.ON_NUMUNE),
    ("Ağırlık Tekdüzeliği (20 numune)", TabloTipi.AGIRLIK_TEKDUZELIGI),
    ("Sertlik/Çap/Dağılma (nokta)", TabloTipi.BOS_NOKTA),
    ("Mikrobiyolojik (matris)", TabloTipi.MATRIS),
]
_TIP_AD = {t: ad for ad, t in _TIP_SECENEKLERI}

# Operasyon adı → tipik no
_OP_NO = {"Karıştırma": 2, "Tablet Baskı": 3, "Film Kaplama": 4, "Dolum": 3, "Blisterleme": 5}

# Sütun indeksleri
SUT_OPNO, SUT_OP, SUT_AD, SUT_SPEK, SUT_ALT, SUT_UST, SUT_TIP, SUT_YILDIZ = range(8)


class SpekModulu(QWidget):
    """Excel benzeri spesifikasyon tablosu (Tablo 6) + özel test butonları."""

    def __init__(self, proje: ProjeVerisi):
        super().__init__()
        self.proje = proje
        self.kart: SpekKarti = proje.spek_karti
        self._yukleniyor = False
        self.setStyleSheet(MODUL_STIL)
        self._arayuzu_kur()
        self._tabloyu_yenile()

    # ----------------------------------------------------------------- arayüz
    def _arayuzu_kur(self) -> None:
        kok = QVBoxLayout(self)
        kok.setContentsMargins(22, 22, 22, 22)
        kok.setSpacing(12)
        kok.addWidget(baslik_etiketi("Spesifikasyonlar (Tablo 6)"))

        # Üst satır: ürün formu + Tablo 8/9
        ust = QHBoxLayout()
        ust.addWidget(QLabel("Ürün Formu:"))
        self.cmb_form = QComboBox()
        for f in UrunFormu:
            self.cmb_form.addItem(f.value, f)
        self.cmb_form.currentIndexChanged.connect(self._form_degisti)
        ust.addWidget(self.cmb_form)
        ust.addStretch(1)
        self.chk_t89 = QCheckBox("Tablo 8/9 (Serbest Bırakma / Raf Ömrü) otomatik üret")
        self.chk_t89.toggled.connect(lambda v: setattr(self.kart, "tablo89_ekle", v))
        ust.addWidget(self.chk_t89)
        kok.addLayout(ust)

        tol = QHBoxLayout()
        tol.addWidget(QLabel("Serbest Bırakma tol.:"))
        self.in_t8 = QLineEdit("±%5"); self.in_t8.setMaximumWidth(80)
        self.in_t8.textChanged.connect(lambda t: setattr(self.kart, "serbest_birakma_tolerans", t.strip()))
        tol.addWidget(self.in_t8)
        tol.addWidget(QLabel("Raf Ömrü tol.:"))
        self.in_t9 = QLineEdit("±%7.5"); self.in_t9.setMaximumWidth(80)
        self.in_t9.textChanged.connect(lambda t: setattr(self.kart, "raf_omru_tolerans", t.strip()))
        tol.addWidget(self.in_t9)
        tol.addStretch(1)
        kok.addLayout(tol)

        kok.addWidget(ipucu_etiketi(
            "Hücrelere doğrudan yazın. Test adına etkin maddeyi gömün "
            "(örn. ‘Etkin madde 1 Miktar Tayini’). Alt/Üst limit sonuç üretimi içindir; "
            "boş bırakırsanız program spesifikasyon metninden sayıları çıkarır."
        ))

        # Ana tablo
        self.tablo = QTableWidget(0, 8)
        self.tablo.setHorizontalHeaderLabels(
            ["Op No", "Operasyon", "Test Adı", "Spesifikasyon",
             "Alt Limit", "Üst Limit", "Sonuç Tipi", "*"])
        hh = self.tablo.horizontalHeader()
        hh.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.tablo.setColumnWidth(SUT_OPNO, 55)
        self.tablo.setColumnWidth(SUT_OP, 110)
        self.tablo.setColumnWidth(SUT_AD, 240)
        self.tablo.setColumnWidth(SUT_SPEK, 280)
        self.tablo.setColumnWidth(SUT_ALT, 80)
        self.tablo.setColumnWidth(SUT_UST, 80)
        self.tablo.setColumnWidth(SUT_TIP, 200)
        self.tablo.setColumnWidth(SUT_YILDIZ, 35)
        self.tablo.verticalHeader().setVisible(False)
        self.tablo.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.tablo.setMinimumHeight(320)
        self.tablo.cellChanged.connect(self._hucre_degisti)
        kok.addWidget(self.tablo, 1)

        # Satır işlemleri
        sb = QHBoxLayout()
        for metin, fn in [("+ Satır", self._satir_ekle), ("− Sil", self._satir_sil),
                          ("▲ Yukarı", lambda: self._tasi(-1)), ("▼ Aşağı", lambda: self._tasi(1)),
                          ("Operasyona Göre Sırala", self._sirala)]:
            b = QPushButton(metin); b.clicked.connect(fn); sb.addWidget(b)
        sb.addStretch(1)
        kok.addLayout(sb)

        # Özel test butonları
        ozel = QHBoxLayout()
        ozel.addWidget(bolum_etiketi("Hazır Özel Testler:"))
        for metin, fn in [("+ Mikrobiyolojik", self._mikro_ekle),
                          ("+ Ağırlık Tekdüzeliği", self._agirlik_ekle),
                          ("+ İlgili Bileşikler", self._ilgili_ekle)]:
            b = QPushButton(metin); b.clicked.connect(fn); ozel.addWidget(b)
        ozel.addStretch(1)
        kok.addLayout(ozel)
        kok.addWidget(ipucu_etiketi(
            "İlgili Bileşikler: etkin madde + impuriteleri tek seferde ekler "
            "(film kaplama hariç otomatik *). Mikrobiyolojik: blisterleme hariç otomatik *."
        ))

        # form/tablo89 doldur
        self._yukleniyor = True
        idx = self.cmb_form.findData(self.kart.urun_formu)
        if idx >= 0:
            self.cmb_form.setCurrentIndex(idx)
        self.chk_t89.setChecked(self.kart.tablo89_ekle)
        self.in_t8.setText(self.kart.serbest_birakma_tolerans)
        self.in_t9.setText(self.kart.raf_omru_tolerans)
        self._yukleniyor = False

    # ----------------------------------------------------------------- tablo
    def _tabloyu_yenile(self) -> None:
        self._yukleniyor = True
        self.tablo.setRowCount(0)
        for test in self.kart.testler:
            self._satir_ciz(test)
        self._yukleniyor = False

    def _satir_ciz(self, test: Test) -> None:
        r = self.tablo.rowCount()
        self.tablo.insertRow(r)
        sp = test.spesifikasyon
        degerler = {
            SUT_OPNO: str(test.operasyon_no or ""),
            SUT_OP: test.operasyon,
            SUT_AD: test.ad,
            SUT_SPEK: sp.spesifikasyon_metni or sp.sabit_sonuc or sp.metni_olustur(),
            SUT_ALT: sp.alt_metin,
            SUT_UST: sp.ust_metin,
        }
        for c, v in degerler.items():
            it = QTableWidgetItem(v)
            it.setForeground(QBrush(QColor("#1a1f2b")))
            it.setBackground(QBrush(QColor("#ffffff")))
            self.tablo.setItem(r, c, it)
        # Sonuç tipi: dropdown
        cmb = QComboBox()
        for ad, _t in _TIP_SECENEKLERI:
            cmb.addItem(ad)
        for i, (_ad, t) in enumerate(_TIP_SECENEKLERI):
            if t is test.tablo_tipi:
                cmb.setCurrentIndex(i); break
        cmb.currentIndexChanged.connect(lambda _i, rr=r: self._tip_degisti(rr))
        self.tablo.setCellWidget(r, SUT_TIP, cmb)
        # Yıldız: checkbox
        chk = QCheckBox()
        chk.setChecked(test.yildizli)
        chk.toggled.connect(lambda v, rr=r: self._yildiz_degisti(rr, v))
        w = QWidget(); lay = QHBoxLayout(w); lay.addWidget(chk)
        lay.setAlignment(Qt.AlignmentFlag.AlignCenter); lay.setContentsMargins(0, 0, 0, 0)
        self.tablo.setCellWidget(r, SUT_YILDIZ, w)

    def _hucre_degisti(self, row: int, col: int) -> None:
        if self._yukleniyor or not (0 <= row < len(self.kart.testler)):
            return
        it = self.tablo.item(row, col)
        v = it.text().strip() if it else ""
        test = self.kart.testler[row]
        sp = test.spesifikasyon
        if col == SUT_OPNO:
            try: test.operasyon_no = int(v) if v else 0
            except ValueError: pass
        elif col == SUT_OP:
            test.operasyon = v
            if not (self.tablo.item(row, SUT_OPNO) and self.tablo.item(row, SUT_OPNO).text().strip()):
                test.operasyon_no = _OP_NO.get(v, 0)
                self._yukleniyor = True
                self.tablo.setItem(row, SUT_OPNO, QTableWidgetItem(str(test.operasyon_no or "")))
                self._yukleniyor = False
        elif col == SUT_AD:
            test.ad = v
            # tip otomatik güncelle (kullanıcı dropdown'dan değiştirmediyse)
            yeni_tip = tablo_tipini_belirle(v)
            test.tablo_tipi = yeni_tip
            test.ipk = ipk_testi_mi(v)
            cmb = self.tablo.cellWidget(row, SUT_TIP)
            if cmb:
                for i, (_ad, t) in enumerate(_TIP_SECENEKLERI):
                    if t is yeni_tip:
                        cmb.blockSignals(True); cmb.setCurrentIndex(i); cmb.blockSignals(False); break
        elif col == SUT_SPEK:
            sp.spesifikasyon_metni = v
        elif col == SUT_ALT:
            sp.alt_metin = v
            sp.alt_limit = self._sayi(v)
        elif col == SUT_UST:
            sp.ust_metin = v
            sp.ust_limit = self._sayi(v)

    def _tip_degisti(self, row: int) -> None:
        if not (0 <= row < len(self.kart.testler)):
            return
        cmb = self.tablo.cellWidget(row, SUT_TIP)
        if cmb:
            self.kart.testler[row].tablo_tipi = _TIP_SECENEKLERI[cmb.currentIndex()][1]

    def _yildiz_degisti(self, row: int, v: bool) -> None:
        if 0 <= row < len(self.kart.testler):
            self.kart.testler[row].yildizli = v

    # ----------------------------------------------------------------- işlem
    def _yeni_test(self, ad="", op="", tip=None) -> Test:
        opno = _OP_NO.get(op, 0)
        return Test(ad=ad, operasyon=op, operasyon_no=opno,
                    tablo_tipi=tip or TabloTipi.TEK_SONUC,
                    spesifikasyon=Spesifikasyon(limit_turu=LimitTuru.ARALIK))

    def _satir_ekle(self) -> None:
        self.kart.testler.append(self._yeni_test())
        self._tabloyu_yenile()
        self.tablo.selectRow(len(self.kart.testler) - 1)

    def _satir_sil(self) -> None:
        r = self.tablo.currentRow()
        if 0 <= r < len(self.kart.testler):
            del self.kart.testler[r]
            self._tabloyu_yenile()

    def _tasi(self, yon: int) -> None:
        r = self.tablo.currentRow()
        y = r + yon
        if 0 <= r < len(self.kart.testler) and 0 <= y < len(self.kart.testler):
            self.kart.testler[r], self.kart.testler[y] = self.kart.testler[y], self.kart.testler[r]
            self._tabloyu_yenile()
            self.tablo.selectRow(y)

    def _sirala(self) -> None:
        self.kart.testler.sort(key=lambda t: (t.operasyon_no or 99))
        self._tabloyu_yenile()

    # ----------------------------------------------------------------- özel testler
    def _aktif_operasyonlar(self) -> list[str]:
        f = self.cmb_form.currentData()
        if isinstance(f, UrunFormu):
            return list(f.operasyonlar)
        return ["Karıştırma", "Tablet Baskı", "Film Kaplama"]

    def _mikro_ekle(self) -> None:
        ops = self._aktif_operasyonlar()
        op, ok = QInputDialog.getItem(self, "Mikrobiyolojik Kontrol",
                                      "Hangi operasyon?", ops, 0, False)
        if not ok:
            return
        # blisterleme hariç otomatik yıldız
        yildiz = "blister" not in op.lower()
        t = Test(ad="Mikrobiyolojik Kontrol", operasyon=op, operasyon_no=_OP_NO.get(op, 0),
                 tablo_tipi=TabloTipi.MATRIS, mikrobiyolojik=True, yildizli=yildiz,
                 spesifikasyon=Spesifikasyon(limit_turu=LimitTuru.METIN, sabit_sonuc="Uygun"),
                 alt_satirlar=[("-Toplam Aerobik Mikroorganizma Sayısı", "≤10³ cfu/g"),
                               ("-Küf ve Maya Sayısı", "≤10² cfu/g"),
                               ("-E. coli", "0 cfu/g")])
        self.kart.testler.append(t)
        self._tabloyu_yenile()

    def _agirlik_ekle(self) -> None:
        dlg = AgirlikDialog(self, self._aktif_operasyonlar())
        if not dlg.exec():
            return
        v = dlg.degerler()
        if not (v["alt"] and v["ust"]):
            QMessageBox.information(self, "Ağırlık Tekdüzeliği", "Alt ve üst limit zorunludur.")
            return
        t = Test(ad="Ağırlık Tekdüzeliği", operasyon=v["operasyon"],
                 operasyon_no=_OP_NO.get(v["operasyon"], 0),
                 tablo_tipi=TabloTipi.AGIRLIK_TEKDUZELIGI, ipk=True, yildizli=v["yildiz"],
                 spesifikasyon=Spesifikasyon(limit_turu=LimitTuru.ARALIK,
                                             alt_metin=v["alt"], ust_metin=v["ust"],
                                             alt_limit=self._sayi(v["alt"]), ust_limit=self._sayi(v["ust"]),
                                             birim="mg"),
                 aciklama_etiketi="—20 tablette tek tek tabletlerden maksimum 2 tanesi bu limitten sapabilir.",
                 aciklama_spek=v["limit1"],
                 aciklama2_etiketi="—Hiçbir tablet bu limitten sapmamalıdır.",
                 aciklama2_spek=v["limit2"])
        self.kart.testler.append(t)
        self._tabloyu_yenile()

    def _ilgili_ekle(self) -> None:
        dlg = IlgiliBilesiklerDialog(self, self._aktif_operasyonlar())
        if not dlg.exec():
            return
        v = dlg.degerler()
        if not v["em_ad"] or not v["impuriteler"]:
            QMessageBox.information(self, "İlgili Bileşikler",
                                    "Etkin madde adı ve en az bir impurite gerekli.")
            return
        op = v["operasyon"]
        # film kaplama hariç otomatik yıldız
        yildiz = "film" not in op.lower()
        # Etkin maddeyi bul/oluştur
        em = next((e for e in self.kart.etkin_maddeler if e.ad == v["em_ad"]), None)
        if em is None:
            em = EtkinMadde(ad=v["em_ad"])
            self.kart.etkin_maddeler.append(em)
        for imp in v["impuriteler"]:
            te = imp["maks"].upper().replace(" ", "") in ("T.E.", "T.E", "TE")
            em.impuriteler.append(Impurite(
                ad=imp["ad"],
                limit_metni=imp["limit"] or (f"Maksimum %{imp['maks']}" if imp["maks"] and not te else ("Maksimum T.E." if te else "")),
                maksimum_deger=self._sayi(imp["maks"]),
                operasyon=op, operasyon_no=_OP_NO.get(op, 0), yildizli=yildiz, te=te))
        QMessageBox.information(self, "İlgili Bileşikler",
                               f"{v['em_ad']} için {len(v['impuriteler'])} impurite eklendi. "
                               "Çıktıda Tablo 6/8/9'da gruplu görünecek.")

    # ----------------------------------------------------------------- form
    def _form_degisti(self) -> None:
        if self._yukleniyor:
            return
        f = self.cmb_form.currentData()
        if isinstance(f, UrunFormu):
            self.kart.urun_formu = f
            self.proje.urun_formu = f

    @staticmethod
    def _sayi(metin):
        if not metin:
            return None
        try:
            return float(str(metin).strip().replace(",", "."))
        except ValueError:
            return None


# ============================================================================
# Dialoglar
# ============================================================================

class AgirlikDialog(QDialog):
    """Ağırlık Tekdüzeliği: 2 limit çifti + sonuç üretimi için alt/üst limit."""

    def __init__(self, parent, operasyonlar):
        super().__init__(parent)
        self.setWindowTitle("Ağırlık Tekdüzeliği Ekle")
        self.setStyleSheet(MODUL_STIL)
        self.setMinimumWidth(480)
        izg = QGridLayout(self); s = 0
        izg.addWidget(QLabel("Operasyon:"), s, 0)
        self.cmb_op = QComboBox(); self.cmb_op.addItems(operasyonlar)
        izg.addWidget(self.cmb_op, s, 1); s += 1
        izg.addWidget(QLabel("Alt Limit (sonuç üretimi):"), s, 0)
        self.in_alt = QLineEdit(); self.in_alt.setPlaceholderText("örn. 270.75")
        izg.addWidget(self.in_alt, s, 1); s += 1
        izg.addWidget(QLabel("Üst Limit (sonuç üretimi):"), s, 0)
        self.in_ust = QLineEdit(); self.in_ust.setPlaceholderText("örn. 299.25")
        izg.addWidget(self.in_ust, s, 1); s += 1
        izg.addWidget(QLabel("‘sapabilir’ satırı sağ değer:"), s, 0)
        self.in_l1 = QLineEdit(); self.in_l1.setPlaceholderText("örn. ≤ 270.75 veya ≥ 299.25 mg")
        izg.addWidget(self.in_l1, s, 1); s += 1
        izg.addWidget(QLabel("‘sapmamalıdır’ satırı sağ değer:"), s, 0)
        self.in_l2 = QLineEdit(); self.in_l2.setPlaceholderText("örn. ≤ 256.50 veya ≥ 313.50 mg")
        izg.addWidget(self.in_l2, s, 1); s += 1
        self.chk_yildiz = QCheckBox("* Validasyon serilerinde")
        izg.addWidget(self.chk_yildiz, s, 0, 1, 2); s += 1
        btn = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btn.accepted.connect(self.accept); btn.rejected.connect(self.reject)
        izg.addWidget(btn, s, 0, 1, 2)

    def degerler(self):
        return {"operasyon": self.cmb_op.currentText(),
                "alt": self.in_alt.text().strip(), "ust": self.in_ust.text().strip(),
                "limit1": self.in_l1.text().strip(), "limit2": self.in_l2.text().strip(),
                "yildiz": self.chk_yildiz.isChecked()}


class IlgiliBilesiklerDialog(QDialog):
    """Etkin madde + impurite listesini tek seferde toplar."""

    def __init__(self, parent, operasyonlar):
        super().__init__(parent)
        self.setWindowTitle("İlgili Bileşikler Ekle")
        self.setStyleSheet(MODUL_STIL)
        self.setMinimumWidth(560)
        kok = QVBoxLayout(self)
        ust = QHBoxLayout()
        ust.addWidget(QLabel("Etkin madde adı:"))
        self.in_em = QLineEdit(); self.in_em.setPlaceholderText("örn. Etkin madde 1")
        ust.addWidget(self.in_em)
        ust.addWidget(QLabel("Operasyon:"))
        self.cmb_op = QComboBox(); self.cmb_op.addItems(operasyonlar)
        ust.addWidget(self.cmb_op)
        kok.addLayout(ust)
        kok.addWidget(ipucu_etiketi("İmpurite satırları: ad | limit metni | maksimum (sayı veya T.E.)"))
        self.tablo = QTableWidget(0, 3)
        self.tablo.setHorizontalHeaderLabels(["İmpurite Adı", "Limit Metni", "Maksimum"])
        self.tablo.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tablo.verticalHeader().setVisible(False)
        self.tablo.setMinimumHeight(180)
        kok.addWidget(self.tablo)
        sb = QHBoxLayout()
        be = QPushButton("+ Satır"); be.clicked.connect(self._satir_ekle)
        bs = QPushButton("− Sil"); bs.clicked.connect(self._satir_sil)
        sb.addWidget(be); sb.addWidget(bs); sb.addStretch(1)
        kok.addLayout(sb)
        btn = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btn.accepted.connect(self.accept); btn.rejected.connect(self.reject)
        kok.addWidget(btn)
        # 3 örnek satır
        for ad in ("imp. a", "Her bir bilinmeyen imp.", "Toplam imp."):
            self._satir_ekle(ad)

    def _satir_ekle(self, ad=""):
        r = self.tablo.rowCount()
        self.tablo.insertRow(r)
        self.tablo.setItem(r, 0, QTableWidgetItem(ad))
        self.tablo.setItem(r, 1, QTableWidgetItem(""))
        self.tablo.setItem(r, 2, QTableWidgetItem(""))

    def _satir_sil(self):
        r = self.tablo.currentRow()
        if r >= 0:
            self.tablo.removeRow(r)

    def degerler(self):
        imps = []
        for r in range(self.tablo.rowCount()):
            ad = self.tablo.item(r, 0).text().strip() if self.tablo.item(r, 0) else ""
            if not ad:
                continue
            limit = self.tablo.item(r, 1).text().strip() if self.tablo.item(r, 1) else ""
            maks = self.tablo.item(r, 2).text().strip() if self.tablo.item(r, 2) else ""
            imps.append({"ad": ad, "limit": limit, "maks": maks})
        return {"em_ad": self.in_em.text().strip(),
                "operasyon": self.cmb_op.currentText(), "impuriteler": imps}
