"""
Çıktı modülü (Bölüm 10. sekme).

Kullanıcı buradan PVP ve/veya PVR dosyalarını Word ve (varsa) PDF olarak üretir.
PVR için sonuç verisi otomatik simüle edilir (gözden geçirilebilir).
Üretim ayrı bir iş parçacığında (QThread) çalışır; arayüz donmaz.
"""

from __future__ import annotations

from pathlib import Path

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QCheckBox,
    QFileDialog, QMessageBox, QProgressBar, QPlainTextEdit,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal

from core.models import ProjeVerisi
from cikti import docx_motoru as motor
from cikti import pdf_donustur as pdfm
from cikti.pdf_donustur import pdf_mevcut_mu
from moduller.widget_yardimcilari import (
    MODUL_STIL, baslik_etiketi, bolum_etiketi, ipucu_etiketi, ayirici,
)


def _dosya_adi_guvenli(metin: str) -> str:
    """
    Ürün adını dosya adı olarak güvenli hale getirir:
    - Türkçe karakterleri sadeleştirir (ı→i, ş→s, ç→c, ...)
    - Yol ayraçları (/ \\) ve geçersiz karakterleri (: * ? " < > |) '_' yapar
    - Boşlukları '_' yapar, baştaki/sondaki '_' temizlenir
    """
    tr = str.maketrans({
        "ı": "i", "İ": "I", "ş": "s", "Ş": "S", "ç": "c", "Ç": "C",
        "ğ": "g", "Ğ": "G", "ü": "u", "Ü": "U", "ö": "o", "Ö": "O",
    })
    ad = (metin or "urun").translate(tr)
    gecersiz = '/\\:*?"<>|'
    for ch in gecersiz:
        ad = ad.replace(ch, "_")
    ad = "_".join(ad.split())  # boşlukları tek '_' yap
    while "__" in ad:
        ad = ad.replace("__", "_")
    return ad.strip("_") or "urun"


class UretimIscisi(QThread):
    """Word/PDF üretimini arka planda yapar (UI donmasın)."""
    bitti = pyqtSignal(list, str)   # (uretilen_yollar, hata_mesaji)
    ilerleme = pyqtSignal(str)

    def __init__(self, proje: ProjeVerisi, dizin: Path,
                 pvp: bool, pvr: bool, pdf: bool, veri_uret: bool):
        super().__init__()
        self.proje = proje
        self.dizin = dizin
        self.pvp = pvp
        self.pvr = pvr
        self.pdf = pdf
        self.veri_uret = veri_uret

    def run(self) -> None:
        uretilen: list[str] = []
        try:
            urun = _dosya_adi_guvenli(self.proje.dokuman.urun_adi or "urun")

            if self.pvp:
                self.ilerleme.emit("PVP (Word) üretiliyor…")
                yp = motor.pvp_uret(self.proje, self.dizin / f"PVP_{urun}.docx")
                uretilen.append(str(yp))
                if self.pdf:
                    self.ilerleme.emit("PVP (PDF) üretiliyor…")
                    pp = pdfm.docx_to_pdf(yp)
                    if pp:
                        uretilen.append(str(pp))

            if self.pvr:
                if self.veri_uret:
                    self.ilerleme.emit("Sonuç verisi üretiliyor…")
                self.ilerleme.emit("PVR (Word) üretiliyor…")
                # Veri üretimi, kural motoruyla TÜRETİLMİŞ test listesi üzerinde
                # çalışmalı; bu yüzden pvr_uret'in kendi türetme context'i içinde
                # üretilir (veri_uret=True). Aksi halde sonuç hücreleri boş kalır.
                yr = motor.pvr_uret(self.proje, self.dizin / f"PVR_{urun}.docx",
                                    veri_uret=self.veri_uret)
                uretilen.append(str(yr))
                if self.pdf:
                    self.ilerleme.emit("PVR (PDF) üretiliyor…")
                    pr = pdfm.docx_to_pdf(yr)
                    if pr:
                        uretilen.append(str(pr))

            self.bitti.emit(uretilen, "")
        except Exception as e:  # üretim hatası UI'ya taşınır
            self.bitti.emit(uretilen, str(e))


class CiktiModulu(QWidget):
    def __init__(self, proje: ProjeVerisi):
        super().__init__()
        self.proje = proje
        self.setObjectName("panel")
        self.setStyleSheet(MODUL_STIL)
        self._isci: UretimIscisi | None = None
        self._arayuzu_kur()

    def _arayuzu_kur(self) -> None:
        kok = QVBoxLayout(self)
        kok.setContentsMargins(22, 22, 22, 22)
        kok.setSpacing(12)

        kok.addWidget(baslik_etiketi("Çıktı — PVP / PVR"))
        kok.addWidget(ipucu_etiketi(
            "Girdiğiniz verilerden PVP (protokol) ve PVR (rapor) dosyalarını üretir. "
            "PVR sonuç tabloları, spesifikasyon sınırlarınıza göre otomatik "
            "(spesifikasyona uygun) veriyle doldurulur; sonra Word üzerinde gözden geçirebilirsiniz."
        ))

        kok.addWidget(ayirici())
        kok.addWidget(bolum_etiketi("Ne üretilsin?"))
        self.chk_pvp = QCheckBox("PVP — Proses Validasyon Protokolü"); self.chk_pvp.setChecked(True)
        self.chk_pvr = QCheckBox("PVR — Proses Validasyon Raporu"); self.chk_pvr.setChecked(True)
        self.chk_veri = QCheckBox("PVR sonuç verisini yeniden üret"); self.chk_veri.setChecked(True)
        kok.addWidget(self.chk_pvp)
        kok.addWidget(self.chk_pvr)
        kok.addWidget(self.chk_veri)

        kok.addWidget(ayirici())
        b = QPushButton("Dosyaları Üret…")
        b.setObjectName("birincil")
        b.clicked.connect(self._uret)
        kok.addWidget(b)

        self.bar = QProgressBar()
        self.bar.setRange(0, 0)  # belirsiz mod
        self.bar.setVisible(False)
        kok.addWidget(self.bar)

        self.gunluk = QPlainTextEdit()
        self.gunluk.setReadOnly(True)
        self.gunluk.setMaximumHeight(160)
        kok.addWidget(self.gunluk)
        kok.addStretch(1)

    def _uret(self) -> None:
        if not (self.chk_pvp.isChecked() or self.chk_pvr.isChecked()):
            QMessageBox.information(self, "Çıktı", "En az bir belge türü seçin (PVP / PVR).")
            return
        if not self.proje.spek_karti.testler and self.chk_pvr.isChecked():
            c = QMessageBox.question(
                self, "Uyarı",
                "Spesifikasyon (test) tanımlı değil; PVR sonuç tabloları boş olur. Devam edilsin mi?")
            if c != QMessageBox.StandardButton.Yes:
                return

        dizin = QFileDialog.getExistingDirectory(
            self, "Çıktı klasörünü seçin", "",
            QFileDialog.Option.DontUseNativeDialog,
        )
        if not dizin:
            return

        self.gunluk.clear()
        self.bar.setVisible(True)
        self._isci = UretimIscisi(
            self.proje, Path(dizin),
            pvp=self.chk_pvp.isChecked(), pvr=self.chk_pvr.isChecked(),
            pdf=False,
            veri_uret=self.chk_veri.isChecked(),
        )
        self._isci.ilerleme.connect(lambda m: self.gunluk.appendPlainText(m))
        self._isci.bitti.connect(self._tamamlandi)
        self._isci.start()

    def _tamamlandi(self, yollar: list, hata: str) -> None:
        self.bar.setVisible(False)
        if hata:
            self.gunluk.appendPlainText(f"HATA: {hata}")
            QMessageBox.critical(self, "Üretim Hatası", hata)
            return
        self.gunluk.appendPlainText("")
        self.gunluk.appendPlainText("Tamamlandı. Üretilen dosyalar:")
        for y in yollar:
            self.gunluk.appendPlainText(f"  • {y}")
        QMessageBox.information(self, "Tamamlandı",
                                f"{len(yollar)} dosya üretildi.")
