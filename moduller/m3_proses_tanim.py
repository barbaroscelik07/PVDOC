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
    QTableWidgetItem, QHeaderView, QAbstractItemView,
)

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

        govde = QHBoxLayout()
        govde.setSpacing(16)

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
        sol_w = QWidget(); sol_w.setLayout(sol); sol_w.setFixedWidth(260)
        govde.addWidget(sol_w)

        # --- sağ: detay editörü ---
        sag = QVBoxLayout()
        sag.setSpacing(8)
        izg = QGridLayout(); izg.setHorizontalSpacing(10); izg.setVerticalSpacing(8)
        izg.addWidget(QLabel("Operasyon No:"), 0, 0)
        self.sp_op = QSpinBox(); self.sp_op.setRange(0, 99)
        self.sp_op.valueChanged.connect(self._detay_yaz)
        izg.addWidget(self.sp_op, 0, 1)
        izg.addWidget(QLabel("Aşama No:"), 0, 2)
        self.sp_as = QSpinBox(); self.sp_as.setRange(0, 99)
        self.sp_as.valueChanged.connect(self._detay_yaz)
        izg.addWidget(self.sp_as, 0, 3)
        izg.addWidget(QLabel("IPK Etiketi:"), 1, 0)
        self.in_ipk = QLineEdit(); self.in_ipk.setPlaceholderText("örn. IPK-1 (boş bırakılabilir)")
        self.in_ipk.textChanged.connect(self._detay_yaz)
        izg.addWidget(self.in_ipk, 1, 1, 1, 3)
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
        sag.addWidget(ipucu_etiketi("örn. ‘Elek açıklığı | 0,8 mm’, ‘Karıştırma süresi | 10 dk’"))

        self.t_param = QTableWidget(0, 2)
        self.t_param.setHorizontalHeaderLabels(["Etiket", "Değer"])
        self.t_param.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.t_param.verticalHeader().setVisible(False)
        self.t_param.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.t_param.cellChanged.connect(self._param_hucre)
        sag.addWidget(self.t_param, 1)

        sag_w = QWidget(); sag_w.setLayout(sag)
        govde.addWidget(sag_w, 1)
        kok.addLayout(govde, 1)

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
        self.t_param.blockSignals(True)
        self.t_param.setRowCount(0)
        for p in a.parametreler:
            r = self.t_param.rowCount()
            self.t_param.insertRow(r)
            self.t_param.setItem(r, 0, QTableWidgetItem(p.etiket))
            self.t_param.setItem(r, 1, QTableWidgetItem(p.deger))
        self.t_param.blockSignals(False)
