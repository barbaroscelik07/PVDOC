"""
Veri klasörü yolları.

Uygulamanın yazılabilir verileri (spek kartı kayıtları, üretilen çıktılar,
sabit şablonlar) kullanıcının ev dizinindeki ~/.pvdoc altında tutulur.
Bu sayede:
  - EXE'nin kurulduğu klasör salt-okunur olsa bile yazma sorunu olmaz,
  - repoya boş klasör / .gitkeep koymaya gerek kalmaz; klasörler ilk
    çalıştırmada otomatik oluşturulur.

Kullanım:
    from core.yollar import SPEK_KARTLARI_DIZINI, CIKTI_DIZINI
    SPEK_KARTLARI_DIZINI.mkdir(...)  -> gerek yok, import anında hazırlanır.
"""

from __future__ import annotations

from pathlib import Path


# Tüm yazılabilir veriler burada toplanır.
VERI_KOKU = Path.home() / ".pvdoc"

SPEK_KARTLARI_DIZINI = VERI_KOKU / "spek_kartlari"
CIKTI_DIZINI = VERI_KOKU / "cikti"
SABLON_DIZINI = VERI_KOKU / "sablonlar"


def dizinleri_hazirla() -> None:
    """Gerekli veri klasörlerini oluşturur (zaten varsa dokunmaz)."""
    for d in (VERI_KOKU, SPEK_KARTLARI_DIZINI, CIKTI_DIZINI, SABLON_DIZINI):
        d.mkdir(parents=True, exist_ok=True)


# Modül import edildiği an klasörler hazır olsun.
dizinleri_hazirla()
