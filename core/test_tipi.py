"""
Test adından PVR sonuç-tablosu tipini otomatik belirleme.

PVR şablonunda tespit edilen 6 tablo tipi (bkz. core/models.TabloTipi):

  TEK_SONUC            : Görünüş, Teşhis, Aşınma, Sızdırmazlık(tekil)
                         → 3 seri × tek değer/sonuç satırı
  IKI_NUMUNE           : Miktar Tayini, İmpurite/İlgili Bileşikler
                         → Numune-1, Numune-2 + Sonuç (ortalama)
  ON_NUMUNE            : Karışım Tekdüzeliği
                         → 1..10 numune + Ortalama
  AGIRLIK_TEKDUZELIGI  : Ağırlık Tekdüzeliği
                         → 20 numune × 3 nokta(Baş/Orta/Son) + Ortalama/RSD/SD
  BOS_NOKTA            : Sertlik, Kalınlık, Çap, Dağılma, Dissolüsyon
                         → n numune × 3 nokta, nokta ortalaması + seri Sonucu
  MATRIS               : Mikrobiyolojik Kontrol
                         → çok satırlı spesifikasyon + 3 seri (×3 nokta) matrisi

Belirleme tamamen test adına dayanır (kullanıcı kararı: "test adından otomatik").
Eşleşme bulunamazsa güvenli varsayılan TEK_SONUC döner.
"""

from __future__ import annotations

import re

from core.models import TabloTipi


_TR_HARITA = str.maketrans({
    "ı": "i", "İ": "i", "I": "i",
    "ş": "s", "Ş": "s",
    "ğ": "g", "Ğ": "g",
    "ç": "c", "Ç": "c",
    "ö": "o", "Ö": "o",
    "ü": "u", "Ü": "u",
})


def _normalize(s: str) -> str:
    """Karşılaştırma için adı sadeleştirir: Türkçe harfleri ASCII'ye indir, küçük harf, tek boşluk."""
    s = s.translate(_TR_HARITA).lower()
    s = re.sub(r"\s+", " ", s).strip()
    return s


# Anahtar kelime → tablo tipi. Sıra ÖNEMLİ: daha özgül (spesifik) kurallar üstte.
# Her kural: (tablo_tipi, [bu kelimelerden biri adda geçerse eşleşir])
_KURALLAR: list[tuple[TabloTipi, list[str]]] = [
    # En spesifik olanlar önce
    (TabloTipi.AGIRLIK_TEKDUZELIGI, ["agirlik tekduzeligi", "agirlik tekduzelik"]),
    (TabloTipi.ON_NUMUNE,           ["karisim tekduzeligi", "karisim tekduzelik", "icerik tekduzeligi"]),
    (TabloTipi.BOS_NOKTA,           ["dissolusyon", "dissolüsyon", "sertlik", "kalinlik", "cap", "dagilma"]),
    (TabloTipi.IKI_NUMUNE,          ["miktar tayini", "ilgili bilesik", "impurite", "imp.", "safsizlik"]),
    (TabloTipi.MATRIS,              ["mikrobiyolojik"]),
    # Tekil / metin sonuçlar
    (TabloTipi.TEK_SONUC,           ["gorunus", "teshis", "asinma", "sizdirmazlik",
                                     "elek testi", "bulk", "tap dansite", "dansite",
                                     "ortalama agirlik"]),
]


def tablo_tipini_belirle(test_adi: str) -> TabloTipi:
    """
    Test adından sonuç tablosu tipini döndürür.

    Örnekler:
      "Etkin madde 1 Miktar Tayini"      -> IKI_NUMUNE
      "Etkin madde 2 Karışım Tekdüzeliği"-> ON_NUMUNE
      "Ağırlık Tekdüzeliği"              -> AGIRLIK_TEKDUZELIGI
      "Sertlik"                          -> BOS_NOKTA
      "Görünüş"                          -> TEK_SONUC
      "Mikrobiyolojik Kontrol"           -> MATRIS
    """
    ad = _normalize(test_adi)
    for tip, kelimeler in _KURALLAR:
        for k in kelimeler:
            if _normalize(k) in ad:
                return tip
    return TabloTipi.TEK_SONUC  # güvenli varsayılan


# Hangi testlerin "Ortalama Ağırlık" gibi tek-değer ama nokta bazlı olduğunu
# ileride ayırmak gerekirse buraya yardımcı eklenebilir. Şimdilik şablon davranışı
# yeterince kapsanıyor.


def ipk_testi_mi(test_adi: str) -> bool:
    """
    Bir testin tipik olarak IPK (In-Process Control) testi olup olmadığını
    önerir. NİHAİ KARAR kullanıcının spek kartındaki 'ipk' bayrağıdır;
    bu yalnızca form ilk doldurulurken varsayılan öneri olarak kullanılır.

    Şablon Tablo 7 (IPK Testleri): Görünüş, Ortalama Ağırlık, Ağırlık
    Tekdüzeliği, Kalınlık, Çap, Sertlik, Aşınma, Dağılma, Sızdırmazlık.
    Teşhis/Miktar/Dissolüsyon/İmpurite/Mikrobiyolojik IPK değildir.
    """
    ad = _normalize(test_adi)
    ipk_kelimeler = [
        "gorunus", "ortalama agirlik", "agirlik tekduzeligi",
        "kalinlik", "cap", "sertlik", "asinma", "dagilma", "sizdirmazlik",
    ]
    return any(_normalize(k) in ad for k in ipk_kelimeler)
