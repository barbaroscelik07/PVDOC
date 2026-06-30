"""
Üretim Yöntemi modülü (Bölüm 5 / Proses Tanımı) — SADELEŞTİRİLMİŞ.

Üretim yöntemi artık Word dosyasından yüklenir; program 'Operasyon X: Aşama Y'
desenini otomatik çözer ve hem üretim yöntemi bölümünü hem de akış diyagramını
bundan üretir. Bu sekme yalnızca: (1) Word yükleme, (2) yüklenen adımların
salt-okunur ÖNİZLEMESİ sağlar. Elle aşama düzenleme kaldırıldı — çıktı zaten
'uretim_adimlari' (Word'den gelen) veriyi kullanıyordu.
"""

from __future__ import annotations

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QPlainTextEdit, QFileDialog, QMessageBox,
)
from PyQt6.QtCore import Qt

from core.models import ProjeVerisi
from moduller.widget_yardimcilari import (
    MODUL_STIL, baslik_etiketi, ipucu_etiketi,
)


class ProsesModulu(QWidget):
    def __init__(self, proje: ProjeVerisi):
        super().__init__()
        self.proje = proje
        self.setObjectName("panel")
        self.setStyleSheet(MODUL_STIL)
        self._arayuzu_kur()
        self._onizleme_yenile()

    def _arayuzu_kur(self) -> None:
        kok = QVBoxLayout(self)
        kok.addWidget(baslik_etiketi("Üretim Yöntemi (Proses Tanımı)"))

        # Word yükleme çubuğu
        wbar = QHBoxLayout()
        b_word = QPushButton("📄 Word'den Üretim Yöntemi Yükle")
        b_word.setObjectName("birincil")
        b_word.clicked.connect(self._word_yukle)
        wbar.addWidget(b_word)
        wbar.addStretch(1)
        kok.addLayout(wbar)

        kok.addWidget(ipucu_etiketi(
            "Üretim yöntemini Word'den yükleyin. Program 'Operasyon X: Aşama Y' "
            "desenini otomatik çözer; üretim yöntemi bölümü ve akış diyagramı "
            "bundan üretilir. Aşağıda yüklenen adımların önizlemesi görünür."))

        # Salt-okunur önizleme
        kok.addWidget(QLabel("Önizleme (yüklenen üretim yöntemi adımları):"))
        self.onizleme = QPlainTextEdit()
        self.onizleme.setReadOnly(True)
        self.onizleme.setPlaceholderText(
            "Henüz üretim yöntemi yüklenmedi. Yukarıdaki butonla Word dosyası seçin.")
        kok.addWidget(self.onizleme, 1)

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
        self._onizleme_yenile()
        QMessageBox.information(
            self, "Üretim Yöntemi Yüklendi",
            f"{len(r['adimlar'])} adım yüklendi.\n\n"
            "Çıktı alırken üretim yöntemi bölümüne bu adımlar yazılacak ve "
            "akış diyagramı bunlardan otomatik oluşturulacak.")

    def _onizleme_yenile(self) -> None:
        """Yüklenmiş üretim adımlarını salt-okunur önizlemede gösterir."""
        adimlar = getattr(self.proje, "uretim_adimlari", None) or []
        if not adimlar:
            self.onizleme.setPlainText("")
            return
        satirlar = []
        for ad in adimlar:
            # uretim_okuyucu çıktısı: (baslik, aciklama, tablo_satirlari)
            if isinstance(ad, (tuple, list)):
                baslik = ad[0] if len(ad) > 0 else ""
                aciklama = ad[1] if len(ad) > 1 else ""
                tablo = ad[2] if len(ad) > 2 else []
            elif isinstance(ad, dict):
                baslik = ad.get("baslik", "")
                aciklama = ad.get("aciklama", "")
                tablo = ad.get("tablo", [])
            else:
                baslik = getattr(ad, "baslik", "")
                aciklama = getattr(ad, "aciklama", "")
                tablo = getattr(ad, "tablo", [])
            if baslik:
                satirlar.append(f"● {baslik}")
            if aciklama:
                satirlar.append(f"   {aciklama}")
            for sat in (tablo or []):
                if isinstance(sat, (tuple, list)) and len(sat) >= 2:
                    satirlar.append(f"      • {sat[0]} : {sat[1]}")
            satirlar.append("")
        self.onizleme.setPlainText("\n".join(satirlar).strip())
