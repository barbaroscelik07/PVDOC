"""
Sonuç verisi üretici (PVR).

Kullanıcının spek kartında belirlediği sınırlardan, spesifikasyona UYGUN
(sağlıklı) simüle sonuç verisi üretir. Kullanıcı kararı: "Program rastgele/
simüle sağlıklı veri üretsin, ben gözden geçireyim."

Üretilen veri tablo tipine göre yapılandırılır ve Test.sonuc_verisi'ne yazılır.
İstatistikler (Ortalama / RSD% / SD) buradaki saf fonksiyonlarla hesaplanır
(numpy bağımlılığı yok — donma/boyut riskini azaltır).

Üretim hedefi: değerler her zaman spesifikasyon içinde kalır; kullanıcı sonra
elle düzenleyebilir.
"""

from __future__ import annotations

import math
import random

from core.models import (
    Test, Spesifikasyon, LimitTuru, TabloTipi,
    SERI_SAYISI, NOKTA_ADLARI,
)


# ----------------------------------------------------------------------------
# İstatistik (saf fonksiyonlar)
# ----------------------------------------------------------------------------

def ortalama(degerler: list[float]) -> float:
    return sum(degerler) / len(degerler) if degerler else 0.0


def std_sapma(degerler: list[float]) -> float:
    """Örneklem standart sapması (n-1)."""
    n = len(degerler)
    if n < 2:
        return 0.0
    ort = ortalama(degerler)
    return math.sqrt(sum((x - ort) ** 2 for x in degerler) / (n - 1))


def rsd_yuzde(degerler: list[float]) -> float:
    """Bağıl standart sapma (%)."""
    ort = ortalama(degerler)
    if ort == 0:
        return 0.0
    return (std_sapma(degerler) / ort) * 100.0


# ----------------------------------------------------------------------------
# Aralık belirleme: spesifikasyondan güvenli üretim aralığı
# ----------------------------------------------------------------------------

def _uretim_araligi(spek: Spesifikasyon) -> tuple[float, float]:
    """
    Spesifikasyondan, içine güvenle değer üretilebilecek (alt, üst) aralık.
    Limitin tam kenarına yapışmamak için hafif içeri çekilir.
    """
    if spek.limit_turu is LimitTuru.ARALIK and spek.alt_limit is not None and spek.ust_limit is not None:
        genislik = spek.ust_limit - spek.alt_limit
        pay = genislik * 0.20
        return spek.alt_limit + pay, spek.ust_limit - pay
    if spek.limit_turu is LimitTuru.MINIMUM and spek.minimum_deger is not None:
        # minimumun biraz üstünden, makul bir bant
        taban = spek.minimum_deger
        return taban * 1.05 + 0.01, taban * 1.05 + max(taban * 0.5, 1.0)
    if spek.limit_turu is LimitTuru.MAKSIMUM and spek.maksimum_deger is not None:
        # maksimumun altında, 0'a yakın sağlıklı bant
        tavan = spek.maksimum_deger
        return tavan * 0.05, tavan * 0.6
    if spek.hedef_deger is not None:
        return spek.hedef_deger * 0.98, spek.hedef_deger * 1.02
    return 0.0, 1.0


def _deger(spek: Spesifikasyon, alt: float, ust: float) -> float:
    """Aralıkta rastgele bir değeri spek ondalığına yuvarlayarak üretir."""
    v = random.uniform(alt, ust)
    return round(v, spek.ondalik)


# ----------------------------------------------------------------------------
# Tablo tipine göre üretim
# ----------------------------------------------------------------------------

def _tek_sonuc(spek: Spesifikasyon) -> dict:
    """3 seri × tek değer/metin."""
    if spek.limit_turu in (LimitTuru.METIN, LimitTuru.BILGI):
        deg = spek.sabit_sonuc or "Uygun"
        return {"seriler": [deg for _ in range(SERI_SAYISI)]}
    alt, ust = _uretim_araligi(spek)
    return {"seriler": [f"{_deger(spek, alt, ust)} {spek.birim}".strip()
                        for _ in range(SERI_SAYISI)]}


def _iki_numune(spek: Spesifikasyon) -> dict:
    """Numune-1, Numune-2 + Sonuç(ortalama), her seri için."""
    alt, ust = _uretim_araligi(spek)
    seriler = []
    for _ in range(SERI_SAYISI):
        n1 = _deger(spek, alt, ust)
        n2 = _deger(spek, alt, ust)
        seriler.append({
            "numune_1": n1, "numune_2": n2,
            "sonuc": round((n1 + n2) / 2, spek.ondalik),
        })
    return {"seriler": seriler}


def _on_numune(spek: Spesifikasyon) -> dict:
    """1..10 numune + Ortalama, her seri için."""
    alt, ust = _uretim_araligi(spek)
    seriler = []
    for _ in range(SERI_SAYISI):
        olcumler = [_deger(spek, alt, ust) for _ in range(10)]
        seriler.append({
            "olcumler": olcumler,
            "ortalama": round(ortalama(olcumler), spek.ondalik),
        })
    return {"seriler": seriler}


def _bos_nokta(spek: Spesifikasyon, numune_sayisi: int = 10) -> dict:
    """
    n numune × 3 nokta (Baş/Orta/Son), her nokta ortalaması + seri Sonucu.
    Sertlik, Kalınlık, Çap, Dağılma, Dissolüsyon.
    """
    alt, ust = _uretim_araligi(spek)
    seriler = []
    for _ in range(SERI_SAYISI):
        noktalar = {}
        nokta_ort = []
        for nokta in NOKTA_ADLARI:
            olcumler = [_deger(spek, alt, ust) for _ in range(numune_sayisi)]
            o = round(ortalama(olcumler), spek.ondalik)
            noktalar[nokta] = {"olcumler": olcumler, "ortalama": o}
            nokta_ort.append(o)
        seriler.append({
            "noktalar": noktalar,
            "sonuc": round(ortalama(nokta_ort), spek.ondalik),
        })
    return {"seriler": seriler}


def _agirlik_tekduzeligi(spek: Spesifikasyon) -> dict:
    """20 numune × 3 nokta + Ortalama/RSD%/SD, her nokta için."""
    alt, ust = _uretim_araligi(spek)
    seriler = []
    for _ in range(SERI_SAYISI):
        noktalar = {}
        for nokta in NOKTA_ADLARI:
            olcumler = [_deger(spek, alt, ust) for _ in range(20)]
            noktalar[nokta] = {
                "olcumler": olcumler,
                "ortalama": round(ortalama(olcumler), spek.ondalik),
                "rsd": round(rsd_yuzde(olcumler), 2),
                "sd": round(std_sapma(olcumler), 2),
            }
        seriler.append({"noktalar": noktalar})
    return {"seriler": seriler}


def _matris(spek: Spesifikasyon) -> dict:
    """Mikrobiyolojik: çok satırlı, 3 seri × Uygun/değer. Basit uygun matrisi."""
    return {"seriler": ["Uygun" for _ in range(SERI_SAYISI)]}


def test_verisi_uret(test: Test) -> dict:
    """Bir testin tablo tipine göre simüle sonuç verisini üretir ve döndürür."""
    t = test.tablo_tipi
    spek = test.spesifikasyon
    if t is TabloTipi.TEK_SONUC:
        return _tek_sonuc(spek)
    if t is TabloTipi.IKI_NUMUNE:
        return _iki_numune(spek)
    if t is TabloTipi.ON_NUMUNE:
        return _on_numune(spek)
    if t is TabloTipi.BOS_NOKTA:
        return _bos_nokta(spek)
    if t is TabloTipi.AGIRLIK_TEKDUZELIGI:
        return _agirlik_tekduzeligi(spek)
    if t is TabloTipi.MATRIS:
        return _matris(spek)
    return _tek_sonuc(spek)


def tum_testleri_uret(testler: list[Test], tohum: int | None = None) -> None:
    """
    Verilen testlerin her biri için sonuç verisi üretip Test.sonuc_verisi'ne yazar.
    tohum verilirse tekrarlanabilir sonuç üretilir (test/doğrulama için).
    """
    if tohum is not None:
        random.seed(tohum)
    for test in testler:
        test.sonuc_verisi = test_verisi_uret(test)
