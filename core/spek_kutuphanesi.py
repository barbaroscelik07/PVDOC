"""
Spek kartı kütüphanesi.

Bir kez girilen spesifikasyon setini (SpekKarti) diske kaydeder ve
sonraki projelerde yeniden kullanılmasını sağlar.
"Bir kez gir → kaydet → her projede çağır" akışının kalıcılık katmanı.

Kartlar tek tek JSON dosyaları olarak ~/.pvdoc/spek_kartlari/ altında tutulur.
Dosya adı kart adından güvenli biçimde türetilir (slug).
"""

from __future__ import annotations

import json
import re
import unicodedata
from dataclasses import asdict
from pathlib import Path
from typing import Optional

from core.models import SpekKarti, _to_jsonable, _from_dict
from core.yollar import SPEK_KARTLARI_DIZINI


_UZANTI = ".spek.json"


def _slug(ad: str) -> str:
    """Kart adından güvenli dosya adı türetir (Türkçe karakterleri sadeleştirir)."""
    ad = ad.strip()
    # Türkçe karakterleri ASCII'ye yaklaştır
    eslesme = {"ı": "i", "İ": "i", "ş": "s", "Ş": "s", "ğ": "g", "Ğ": "g",
               "ç": "c", "Ç": "c", "ö": "o", "Ö": "o", "ü": "u", "Ü": "u"}
    ad = ad.translate(str.maketrans(eslesme))
    ad = unicodedata.normalize("NFKD", ad).encode("ascii", "ignore").decode("ascii")
    ad = ad.lower()
    ad = re.sub(r"[^a-z0-9]+", "_", ad).strip("_")
    return ad or "kart"


def _kart_yolu(kart_adi: str) -> Path:
    return SPEK_KARTLARI_DIZINI / f"{_slug(kart_adi)}{_UZANTI}"


# ----------------------------------------------------------------------------
# CRUD
# ----------------------------------------------------------------------------

def karti_kaydet(kart: SpekKarti) -> Path:
    """
    Spek kartını kütüphaneye kaydeder (atomik yazım).
    Aynı slug'a sahip kart varsa üzerine yazar (güncelleme).
    """
    if not kart.kart_adi.strip():
        raise ValueError("Kart adı boş olamaz.")

    SPEK_KARTLARI_DIZINI.mkdir(parents=True, exist_ok=True)
    yol = _kart_yolu(kart.kart_adi)
    tmp = yol.with_suffix(yol.suffix + ".tmp")

    veri = _to_jsonable(kart)
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(veri, f, ensure_ascii=False, indent=2)
    tmp.replace(yol)
    return yol


def karti_yukle(kart_adi: str) -> SpekKarti:
    """Ada (veya doğrudan slug'a) göre kartı yükler."""
    yol = _kart_yolu(kart_adi)
    if not yol.exists():
        raise FileNotFoundError(f"Spek kartı bulunamadı: {kart_adi}")
    with open(yol, "r", encoding="utf-8") as f:
        veri = json.load(f)
    return _from_dict(SpekKarti, veri)


def karti_sil(kart_adi: str) -> bool:
    """Kartı siler. Silindiyse True, dosya yoksa False döner."""
    yol = _kart_yolu(kart_adi)
    if yol.exists():
        yol.unlink()
        return True
    return False


def kartlari_listele() -> list[dict]:
    """
    Kütüphanedeki tüm kartların özetini döndürür (dosyayı tamamen yüklemeden).
    Her öğe: {'kart_adi', 'urun_formu', 'test_sayisi', 'etkin_madde_sayisi', 'yol'}
    Bozuk dosyalar atlanır.
    """
    SPEK_KARTLARI_DIZINI.mkdir(parents=True, exist_ok=True)
    sonuc: list[dict] = []
    for yol in sorted(SPEK_KARTLARI_DIZINI.glob(f"*{_UZANTI}")):
        try:
            with open(yol, "r", encoding="utf-8") as f:
                veri = json.load(f)
            sonuc.append({
                "kart_adi": veri.get("kart_adi", yol.stem),
                "urun_formu": veri.get("urun_formu", ""),
                "test_sayisi": len(veri.get("testler", [])),
                "etkin_madde_sayisi": len(veri.get("etkin_maddeler", [])),
                "yol": str(yol),
            })
        except (json.JSONDecodeError, OSError):
            continue
    return sonuc


def kart_var_mi(kart_adi: str) -> bool:
    return _kart_yolu(kart_adi).exists()
