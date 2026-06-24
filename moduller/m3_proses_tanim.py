"""
Üretim Yöntemi modülü (Bölüm 5 / Proses Tanımı).

Üretim yöntemi operasyon ve aşamalardan oluşur. Her aşama:
- serbest açıklama metni ("0.450 kg Talk ... 0,8 mm elekten elenerek ..."),
- opsiyonel mini parametre tablosu ("Elek açıklığı | 0,8 mm",
  "Karıştırma süresi | 10 dk" gibi),
- opsiyonel IPK etiketi ("IPK-1", "IPK-2" — kullanıcı kendi girer).

Sol listede aşamalar, sağda seçili aşamanın detay editörü gösterilir.
"""

from __future__ import annotations

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QLineEdit,
    QSpinBox, QPlainTextEdit, QPushButton, QListWidget, QTableWidget,
    QTableWidgetItem, QHeaderView, QAbstractItemView, QSplitter, QScrollArea,
    QFileDialog, QMessageBox,
)
from PyQt6.QtCore import Qt

from core.models import ProjeVerisi, Asama, ParametreSatiri
from moduller.widget_yardimcilari import (
    MODUL_STIL, baslik_etiketi, bolum_etiketi, ipucu_etiketi, ayirici,
)


class ProsesModulu(QWidget):
    def __init__(self, proje: ProjeVerisi):
        super().__init__()
        self.proje = proje
        self.setObjectName("panel")
        self.setStyleSheet(MODUL_STIL)
        self._aktif: int = -1  # seçili aşama index'i
        self._arayuzu_kur()
        self._liste_yenile()

    def _arayuzu_kur(self) -> None:
        kok = QVBoxLayout(self)
        kok.setContentsMargins(22, 22, 22, 22)
        kok.setSpacing(12)
        kok.addWidget(baslik_etiketi("Üretim Yöntemi (Proses Tanımı)"))

        # Word'den üretim yöntemi yükle
        wbar = QHBoxLayout()
        b_word = QPushButton("📄 Word'den Üretim Yöntemi Yükle")
        b_word.setObjectName("birincil")
        b_word.clicked.connect(self._word_yukle)
        wbar.addWidget(b_word)
        wbar.addWidget(ipucu_etiketi(
            "'Operasyon X: Aşama Y' + açıklama çiftlerini okur, çıktıya aynen yazar."))
        wbar.addStretch(1)
        kok.addLayout(wbar)

        # Splitter: dar ekranda sol/sağ panel oranı korunur, paneller ezilmez
        bolucu = QSplitter(Qt.Orientation.Horizontal)

        # --- sol: aşama listesi ---
        sol = QVBoxLayout()
        sol.addWidget(bolum_etiketi("Aşamalar"))
        self.liste = QListWidget()
        self.liste.currentRowChanged.connect(self._asama_secildi)
        sol.addWidget(self.liste, 1)
        sb = QHBoxLayout()
        b_ekle = QPushButton("+ Aşama"); b_ekle.clicked.connect(self._asama_ekle)
        b_sil = QPushButton("− Sil"); b_sil.setObjectName("tehlike"); b_sil.clicked.connect(self._asama_sil)
        sb.addWidget(b_ekle); sb.addWidget(b_sil)
        sol.addLayout(sb)
        sol_w = QWidget(); sol_w.setLayout(sol)
        sol_w.setMinimumWidth(220)
        bolucu.addWidget(sol_w)

        # --- sağ: detay editörü ---
        sag = QVBoxLayout()
        sag.setSpacing(8)
        izg = QGridLayout(); izg.setHorizontalSpacing(10); izg.setVerticalSpacing(8)
        # Operasyon No ve Aşama No ALT ALTA (kullanıcı isteği)
        izg.addWidget(QLabel("Operasyon No:"), 0, 0)
        self.sp_op = QSpinBox(); self.sp_op.setRange(0, 99)
        self.sp_op.valueChanged.connect(self._operasyon_no_degisti)
        izg.addWidget(self.sp_op, 0, 1)
        izg.addWidget(QLabel("Aşama No:"), 1, 0)
        self.sp_as = QSpinBox(); self.sp_as.setRange(0, 99)
        self.sp_as.valueChanged.connect(self._detay_yaz)
        izg.addWidget(self.sp_as, 1, 1)
        izg.addWidget(QLabel("IPK Etiketi:"), 2, 0)
        self.in_ipk = QLineEdit(); self.in_ipk.setPlaceholderText("örn. IPK-1 (boş bırakılabilir)")
        self.in_ipk.textChanged.connect(self._detay_yaz)
        izg.addWidget(self.in_ipk, 2, 1)
        sag.addLayout(izg)

        sag.addWidget(QLabel("Açıklama:"))
        self.txt = QPlainTextEdit()
        self.txt.setPlaceholderText("Aşama açıklaması…")
        self.txt.textChanged.connect(self._detay_yaz)
        self.txt.setMaximumHeight(140)
        sag.addWidget(self.txt)

        sag.addWidget(ayirici())

        pb = QHBoxLayout()
        pb.addWidget(bolum_etiketi("Mini Parametre Tablosu"))
        pb.addStretch(1)
        b_pe = QPushButton("+ Satır"); b_pe.clicked.connect(self._param_ekle)
        b_ps = QPushButton("− Sil"); b_ps.setObjectName("tehlike"); b_ps.clicked.connect(self._param_sil)
        pb.addWidget(b_pe); pb.addWidget(b_ps)
        sag.addLayout(pb)
        sag.addWidget(ipucu_etiketi("örn. ‘Elek açıklığı | 0,8 mm’, ‘Karıştırma süresi | 10 dk’ — ° için: AltGr+, veya kopyala-yapıştır"))

        self.t_param = QTableWidget(0, 2)
        self.t_param.setHorizontalHeaderLabels(["Etiket", "Değer"])
        self.t_param.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.t_param.verticalHeader().setVisible(False)
        self.t_param.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.t_param.cellChanged.connect(self._param_hucre)
        sag.addWidget(self.t_param, 1)

        sag_w = QWidget(); sag_w.setLayout(sag)
        sag_w.setMinimumWidth(300)
        bolucu.addWidget(sag_w)
        bolucu.setStretchFactor(1, 1)
        bolucu.setSizes([260, 600])
        kok.addWidget(bolucu, 1)

        self._detay_aktif(False)

    # ------------------------------------------------------------- aşama listesi
    def _liste_yenile(self) -> None:
        self.liste.blockSignals(True)
        self.liste.clear()
        for a in self.proje.asamalar:
            etiket = f"Operasyon {a.operasyon_no}: Aşama {a.asama_no}"
            if a.ipk_etiketi:
                etiket += f"  ({a.ipk_etiketi})"
            self.liste.addItem(etiket)
        self.liste.blockSignals(False)
        if self.proje.asamalar:
            self.liste.setCurrentRow(min(max(self._aktif, 0), len(self.proje.asamalar) - 1))
        else:
            self._aktif = -1
            self._detay_aktif(False)

    def _word_yukle(self) -> None:
        """Word'den üretim yöntemi adımlarını (Operasyon/Aşama çiftleri) okur."""
        yol, _ = QFileDialog.getOpenFileName(
            self, "Üretim Yöntemi İçeren Word Dosyası Seç", "", "Word (*.docx)")
        if not yol:
            return
        try:
            from core.uretim_okuyucu import uretim_yontemi_coz
            r = uretim_yontemi_coz(yol)
        except Exception as e:
            QMessageBox.warning(self, "Word Okuma Hatası", f"Dosya okunamadı:\n{e}")
            return
        if not r["bulundu"]:
            QMessageBox.warning(self, "Bulunamadı",
                                "Word dosyasında 'Operasyon X: Aşama Y' deseni bulunamadı.\n"
                                "Üretim yöntemi 'Operasyon 1: Aşama 1' ile başlamalı.")
            return
        self.proje.uretim_adimlari = r["adimlar"]
        QMessageBox.information(
            self, "Üretim Yöntemi Yüklendi",
            f"{len(r['adimlar'])} adım yüklendi.\n\n"
            "Çıktı alırken üretim yöntemi bölümüne bu adımlar yazılacak.")

    def _asama_ekle(self) -> None:
        # akıllı varsayılan: son aşamanın operasyonunu sürdür, aşama no +1
        op = self.proje.asamalar[-1].operasyon_no if self.proje.asamalar else 1
        asno = (self.proje.asamalar[-1].asama_no + 1) if self.proje.asamalar else 1
        self.proje.asamalar.append(Asama(operasyon_no=op, asama_no=asno))
        self._aktif = len(self.proje.asamalar) - 1
        self._liste_yenile()

    def _asama_sil(self) -> None:
        if 0 <= self._aktif < len(self.proje.asamalar):
            del self.proje.asamalar[self._aktif]
            self._aktif = min(self._aktif, len(self.proje.asamalar) - 1)
            self._liste_yenile()

    def _asama_secildi(self, row: int) -> None:
        self._aktif = row
        if 0 <= row < len(self.proje.asamalar):
            self._detay_doldur(self.proje.asamalar[row])
            self._detay_aktif(True)
        else:
            self._detay_aktif(False)

    # ------------------------------------------------------------- detay editörü
    def _aktif_asama(self) -> Asama | None:
        if 0 <= self._aktif < len(self.proje.asamalar):
            return self.proje.asamalar[self._aktif]
        return None

    def _detay_aktif(self, durum: bool) -> None:
        for w in (self.sp_op, self.sp_as, self.in_ipk, self.txt, self.t_param):
            w.setEnabled(durum)

    def _detay_doldur(self, a: Asama) -> None:
        for w in (self.sp_op, self.sp_as, self.in_ipk, self.txt):
            w.blockSignals(True)
        self.sp_op.setValue(a.operasyon_no)
        self.sp_as.setValue(a.asama_no)
        self.in_ipk.setText(a.ipk_etiketi)
        self.txt.setPlainText(a.metin)
        for w in (self.sp_op, self.sp_as, self.in_ipk, self.txt):
            w.blockSignals(False)
        self._param_doldur(a)

    def _operasyon_no_degisti(self) -> None:
        """Operasyon no değişince aşama no otomatik 1'den başlar (kullanıcı isteği)."""
        a = self._aktif_asama()
        if a and self.sp_op.value() != a.operasyon_no:
            self.sp_as.blockSignals(True)
            self.sp_as.setValue(1)
            self.sp_as.blockSignals(False)
        self._detay_yaz()

    def _detay_yaz(self) -> None:
        a = self._aktif_asama()
        if not a:
            return
        a.operasyon_no = self.sp_op.value()
        a.asama_no = self.sp_as.value()
        a.ipk_etiketi = self.in_ipk.text().strip()
        a.metin = self.txt.toPlainText()
        # liste etiketini güncelle (sinyal tetiklemeden)
        if 0 <= self._aktif < self.liste.count():
            etiket = f"Operasyon {a.operasyon_no}: Aşama {a.asama_no}"
            if a.ipk_etiketi:
                etiket += f"  ({a.ipk_etiketi})"
            self.liste.blockSignals(True)
            self.liste.item(self._aktif).setText(etiket)
            self.liste.blockSignals(False)

    # ------------------------------------------------------------- mini tablo
    def _param_ekle(self) -> None:
        a = self._aktif_asama()
        if a is None:
            return
        a.parametreler.append(ParametreSatiri(etiket="", deger=""))
        self._param_doldur(a)

    def _param_sil(self) -> None:
        a = self._aktif_asama()
        if a is None:
            return
        r = self.t_param.currentRow()
        if 0 <= r < len(a.parametreler):
            del a.parametreler[r]
            self._param_doldur(a)

    def _param_hucre(self, row: int, col: int) -> None:
        a = self._aktif_asama()
        if a is None or not (0 <= row < len(a.parametreler)):
            return
        it = self.t_param.item(row, col)
        v = it.text() if it else ""
        if col == 0:
            a.parametreler[row].etiket = v
        else:
            a.parametreler[row].deger = v

    def _param_doldur(self, a: Asama) -> None:
        from PyQt6.QtGui import QColor, QBrush
        acik = QBrush(QColor("#1a1f2b"))   # koyu yazı (açık hücre zemini için)
        self.t_param.blockSignals(True)
        self.t_param.setRowCount(0)
        for p in a.parametreler:
            r = self.t_param.rowCount()
            self.t_param.insertRow(r)
            for c, val in ((0, p.etiket), (1, p.deger)):
                it = QTableWidgetItem(val)
                it.setForeground(acik)
                it.setBackground(QBrush(QColor("#ffffff")))
                self.t_param.setItem(r, c, it)
        self.t_param.blockSignals(False)
