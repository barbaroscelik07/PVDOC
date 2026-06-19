"""
Ana pencere: QMainWindow + hub merkezli sekme yapısı.

Sorumluluklar:
- Tek ProjeVerisi örneğini tutar (uygulamanın tek gerçek kaynağı).
- Üst menü: Yeni / Aç / Kaydet / Farklı Kaydet (proje_io üzerinden).
- Sol tarafta modül navigasyonu (hub), sağda aktif modül paneli.
- Modüller henüz yazılmadığı için şimdilik "yer tutucu" paneller gösterilir;
  her modül tamamlandıkça ilgili panel gerçek modülle değiştirilecek.

Tasarım: koyu zemin + cam (glassmorphic) aksan paneller.
"""

from __future__ import annotations

import os
from pathlib import Path

from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QHBoxLayout, QVBoxLayout, QListWidget,
    QListWidgetItem, QStackedWidget, QLabel, QFileDialog, QMessageBox,
)
from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QIcon

from core.models import ProjeVerisi
from core import proje_io


def _ikon_yolu() -> str:
    """
    ikon.ico'nun tam yolunu döndürür.
    PyInstaller ile paketlendiğinde dosya geçici _MEIPASS klasörüne açılır;
    normal çalışmada ise proje kökünden okunur.
    """
    import sys
    import os
    taban = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(taban, "kaynaklar_ikon", "ikon.ico")


# Hub'da görünecek modüller: (anahtar, görünen ad). Sıra = iş akışı sırası.
MODUL_LISTESI = [
    ("genel",        "1 · Genel Bilgi"),
    ("formul",       "2 · Birim/Seri Formül"),
    ("proses",       "3 · Üretim Yöntemi"),
    ("diyagram",     "4 · Akış Diyagramı"),
    ("risk",         "5 · Risk Analizi"),
    ("ekipman",      "6 · Ekipman Listesi"),
    ("spek",         "7 · Spesifikasyonlar"),
    ("numune",       "8 · Numune Alma Planı"),
    ("sonuc",        "9 · Sonuçlar (PVR)"),
    ("cikti",        "10 · Çıktı (PVP / PVR)"),
]


STIL = """
QMainWindow, QWidget { background-color: #11151c; color: #e6edf3; }
QListWidget {
    background-color: rgba(255,255,255,0.04);
    border: 1px solid rgba(255,255,255,0.08);
    border-radius: 12px;
    padding: 8px;
    font-size: 14px;
    outline: none;
}
QListWidget::item {
    padding: 11px 14px;
    border-radius: 8px;
    margin: 2px 0;
}
QListWidget::item:selected {
    background-color: rgba(88,166,255,0.18);
    color: #ffffff;
}
QListWidget::item:hover:!selected {
    background-color: rgba(255,255,255,0.06);
}
QLabel#baslik { font-size: 22px; font-weight: 600; }
QLabel#altbaslik { color: #8b98a5; font-size: 13px; }
QWidget#panel {
    background-color: rgba(255,255,255,0.04);
    border: 1px solid rgba(255,255,255,0.08);
    border-radius: 16px;
}
QLabel#yer_tutucu { color: #8b98a5; font-size: 15px; }
"""


class YerTutucuPanel(QWidget):
    """Modül henüz yazılmadığında gösterilen geçici panel."""

    def __init__(self, baslik: str, aciklama: str = "Bu modül yakında eklenecek."):
        super().__init__()
        self.setObjectName("panel")
        dik = QVBoxLayout(self)
        dik.setContentsMargins(28, 28, 28, 28)
        dik.setSpacing(8)

        b = QLabel(baslik)
        b.setObjectName("baslik")
        a = QLabel(aciklama)
        a.setObjectName("yer_tutucu")
        a.setWordWrap(True)

        dik.addWidget(b)
        dik.addWidget(a)
        dik.addStretch(1)


class AnaPencere(QMainWindow):
    def __init__(self):
        super().__init__()
        self.proje: ProjeVerisi = ProjeVerisi()
        self.aktif_yol: Path | None = None
        self._paneller: dict[str, QWidget] = {}

        self.setWindowTitle("PV-DOC")
        self.resize(1100, 720)
        self.setStyleSheet(STIL)

        ikon = _ikon_yolu()
        if os.path.exists(ikon):
            self.setWindowIcon(QIcon(ikon))

        self._arayuzu_kur()
        self._menuyu_kur()

        # Son oturumu geri yüklemeyi dene
        self._son_oturumu_dene()
        self._baslik_guncelle()

    # ----------------------------------------------------------------- arayüz
    def _arayuzu_kur(self) -> None:
        merkez = QWidget()
        yatay = QHBoxLayout(merkez)
        yatay.setContentsMargins(16, 16, 16, 16)
        yatay.setSpacing(16)

        # Sol: navigasyon (hub)
        sol = QVBoxLayout()
        sol.setSpacing(10)
        baslik = QLabel("PV-DOC")
        baslik.setObjectName("baslik")
        altbaslik = QLabel("Proses Validasyon Protokolü / Raporu")
        altbaslik.setObjectName("altbaslik")
        sol.addWidget(baslik)
        sol.addWidget(altbaslik)

        self.nav = QListWidget()
        self.nav.setFixedWidth(240)
        for anahtar, ad in MODUL_LISTESI:
            it = QListWidgetItem(ad)
            it.setData(Qt.ItemDataRole.UserRole, anahtar)
            it.setSizeHint(QSize(0, 44))
            self.nav.addItem(it)
        self.nav.currentRowChanged.connect(self._modul_degisti)
        sol.addWidget(self.nav, 1)

        sol_kapsayici = QWidget()
        sol_kapsayici.setLayout(sol)
        sol_kapsayici.setFixedWidth(260)

        # Sağ: aktif modül paneli (stack)
        self.stack = QStackedWidget()
        for anahtar, ad in MODUL_LISTESI:
            panel = YerTutucuPanel(ad)
            self._paneller[anahtar] = panel
            self.stack.addWidget(panel)

        yatay.addWidget(sol_kapsayici)
        yatay.addWidget(self.stack, 1)

        self.setCentralWidget(merkez)

        # Gerçek modülleri bağla (tamamlandıkça buraya eklenir)
        self._gercek_modulleri_bagla()

        self.nav.setCurrentRow(0)

    def _gercek_modulleri_bagla(self) -> None:
        """Tamamlanmış modülleri yer tutucularla değiştirir."""
        from moduller.m1_genel_bilgi import GenelBilgiModulu
        from moduller.m2_formul import FormulModulu
        from moduller.m3_proses_tanim import ProsesModulu
        from moduller.m4_risk_analizi import RiskModulu
        from moduller.m5_ekipman import EkipmanModulu
        from moduller.m6_spesifikasyon import SpekModulu
        from moduller.m7_numune_plani import NumuneModulu
        from moduller.m8_cikti import CiktiModulu

        self.modulu_yenile("genel", GenelBilgiModulu(self.proje))
        self.modulu_yenile("formul", FormulModulu(self.proje))
        self.modulu_yenile("proses", ProsesModulu(self.proje))
        self.modulu_yenile("risk", RiskModulu(self.proje))
        self.modulu_yenile("ekipman", EkipmanModulu(self.proje))
        self.modulu_yenile("spek", SpekModulu(self.proje))
        self.modulu_yenile("numune", NumuneModulu(self.proje))
        self.modulu_yenile("cikti", CiktiModulu(self.proje))

    def _menuyu_kur(self) -> None:
        cubuk = self.menuBar()
        dosya = cubuk.addMenu("&Dosya")

        eylem_yeni = dosya.addAction("Yeni Proje")
        eylem_yeni.triggered.connect(self.yeni_proje)
        eylem_ac = dosya.addAction("Aç…")
        eylem_ac.triggered.connect(self.projeyi_ac)
        dosya.addSeparator()
        eylem_kaydet = dosya.addAction("Kaydet")
        eylem_kaydet.triggered.connect(self.kaydet)
        eylem_farkli = dosya.addAction("Farklı Kaydet…")
        eylem_farkli.triggered.connect(self.farkli_kaydet)
        dosya.addSeparator()
        eylem_cikis = dosya.addAction("Çıkış")
        eylem_cikis.triggered.connect(self.close)

    # ----------------------------------------------------------- modül geçişi
    def _modul_degisti(self, satir: int) -> None:
        if 0 <= satir < self.stack.count():
            self.stack.setCurrentIndex(satir)

    def modulu_yenile(self, anahtar: str, panel: QWidget) -> None:
        """
        Gerçek bir modül tamamlandığında yer tutucuyu onunla değiştirir.
        Panel otomatik olarak dikey kaydırılabilir bir alana sarılır; böylece
        pencere küçültüldüğünde (yarım ekran) sayfanın altı her zaman erişilebilir.
        """
        if anahtar not in self._paneller:
            return
        from PyQt6.QtWidgets import QScrollArea
        from PyQt6.QtCore import Qt as _Qt
        kaydir = QScrollArea()
        kaydir.setWidgetResizable(True)
        kaydir.setFrameShape(QScrollArea.Shape.NoFrame)
        kaydir.setHorizontalScrollBarPolicy(_Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        kaydir.setVerticalScrollBarPolicy(_Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        kaydir.setWidget(panel)

        eski = self._paneller[anahtar]
        idx = self.stack.indexOf(eski)
        self.stack.removeWidget(eski)
        eski.deleteLater()
        self.stack.insertWidget(idx, kaydir)
        self._paneller[anahtar] = kaydir

    # --------------------------------------------------------- proje işlemleri
    def yeni_proje(self) -> None:
        self.proje = ProjeVerisi()
        self.aktif_yol = None
        proje_io.oturumu_temizle()
        self._gercek_modulleri_bagla()
        self._baslik_guncelle()

    def projeyi_ac(self) -> None:
        yol, _ = QFileDialog.getOpenFileName(
            self, "Proje Aç", "", "PV-DOC Projesi (*.pvdoc)",
            options=QFileDialog.Option.DontUseNativeDialog,
        )
        if not yol:
            return
        try:
            self.proje = proje_io.projeyi_yukle(yol)
            self.aktif_yol = Path(yol)
            self._gercek_modulleri_bagla()
            self._baslik_guncelle()
        except Exception as e:  # geniş tut: dosya bozuk olabilir
            QMessageBox.critical(self, "Açma Hatası", f"Proje açılamadı:\n{e}")

    def kaydet(self) -> bool:
        """Aktif yola kaydeder; yol yoksa Farklı Kaydet'e düşer."""
        if self.aktif_yol is None:
            return self.farkli_kaydet()
        try:
            proje_io.projeyi_kaydet(self.proje, self.aktif_yol)
            return True
        except Exception as e:
            QMessageBox.critical(self, "Kaydetme Hatası", f"Kaydedilemedi:\n{e}")
            return False

    def farkli_kaydet(self) -> bool:
        onerilen = self.proje.dokuman.urun_adi or "proje"
        yol, _ = QFileDialog.getSaveFileName(
            self, "Farklı Kaydet", onerilen, "PV-DOC Projesi (*.pvdoc)",
            options=QFileDialog.Option.DontUseNativeDialog,
        )
        if not yol:
            return False
        try:
            kaydedilen = proje_io.projeyi_kaydet(self.proje, yol)
            self.aktif_yol = kaydedilen
            self._baslik_guncelle()
            return True
        except Exception as e:
            QMessageBox.critical(self, "Kaydetme Hatası", f"Kaydedilemedi:\n{e}")
            return False

    # ----------------------------------------------------------------- yardımcı
    def _son_oturumu_dene(self) -> None:
        geri = proje_io.son_oturumu_geri_yukle()
        if geri is not None:
            self.proje = geri
            self.aktif_yol = proje_io.son_oturum_yolu()

    def _baslik_guncelle(self) -> None:
        ad = self.aktif_yol.name if self.aktif_yol else "Kaydedilmemiş proje"
        urun = self.proje.dokuman.urun_adi
        parca = f"{urun} — {ad}" if urun else ad
        self.setWindowTitle(f"PV-DOC — {parca}")
