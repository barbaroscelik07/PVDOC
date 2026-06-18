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
import re

from core.models import (
    Test, Spesifikasyon, LimitTuru, TabloTipi,
    SERI_SAYISI, NOKTA_ADLARI,
)


def _metinden_sayilar(metin: str) -> list[float]:
    """
    Spesifikasyon metninden sayıları çıkarır.
    "%85 – %115" -> [85, 115]
    "5.0 mg/f.tab ±%5 (4.75 – 5.25 mg/f.tab)" -> [5.0, 5, 4.75, 5.25]
    "Minimum %80.0" -> [80.0]
    Virgül ondalık ayırıcı olarak da kabul edilir.
    """
    if not metin:
        return []
    # 4,75 gibi virgüllü ondalıkları noktaya çevir (binlik ayırıcı varsayma)
    norm = re.sub(r"(\d),(\d)", r"\1.\2", metin)
    return [float(x) for x in re.findall(r"\d+(?:\.\d+)?", norm)]


def _aralik_metinden(metin: str) -> tuple:
    """
    Metinden (alt, üst) aralığı çıkarmaya çalışır.
    Parantez içindeki aralığı (4.75 – 5.25) önceliklendirir; yoksa ilk iki sayıyı kullanır.
    Tek sayı varsa (alt, None) veya (None, None) döner.
    """
    if not metin:
        return None, None
    # Önce parantez içi "a – b" ara
    par = re.search(r"\(([^)]*)\)", metin)
    kaynak = par.group(1) if par else metin
    sayilar = _metinden_sayilar(kaynak)
    # Parantez içinde 2+ sayı varsa son ikisini aralık kabul et
    if len(sayilar) >= 2:
        return sayilar[-2], sayilar[-1]
    # parantez yoksa tüm metindeki ilk iki sayı
    tum = _metinden_sayilar(metin)
    if len(tum) >= 2:
        return tum[-2], tum[-1]
    if len(tum) == 1:
        return tum[0], None
    return None, None


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
    Öncelik: açıkça girilen alt_limit/ust_limit; yoksa spesifikasyon metninden
    otomatik çıkarılan sayılar.
    """
    alt_l = spek.alt_limit
    ust_l = spek.ust_limit
    # Limitler boşsa metinden (spesifikasyon_metni / sabit_sonuc) çıkarmayı dene
    if alt_l is None or ust_l is None:
        kaynak = spek.spesifikasyon_metni or spek.sabit_sonuc or ""
        m_alt, m_ust = _aralik_metinden(kaynak)
        if alt_l is None:
            alt_l = m_alt
        if ust_l is None:
            ust_l = m_ust

    if alt_l is not None and ust_l is not None:
        genislik = ust_l - alt_l
        pay = genislik * 0.20
        return alt_l + pay, ust_l - pay
    if spek.limit_turu is LimitTuru.MINIMUM and spek.minimum_deger is not None:
        taban = spek.minimum_deger
        return taban * 1.05 + 0.01, taban * 1.05 + max(taban * 0.5, 1.0)
    if spek.limit_turu is LimitTuru.MAKSIMUM and spek.maksimum_deger is not None:
        tavan = spek.maksimum_deger
        return tavan * 0.05, tavan * 0.6
    if spek.hedef_deger is not None:
        return spek.hedef_deger * 0.98, spek.hedef_deger * 1.02
    # Son çare: minimum tek sayı varsa onun etrafı
    tek = _metinden_sayilar(spek.spesifikasyon_metni or spek.sabit_sonuc or "")
    if tek:
        return tek[0] * 1.02, tek[0] * 1.10
    return 0.0, 1.0


def _deger(spek: Spesifikasyon, alt: float, ust: float) -> float:
    """Aralıkta rastgele bir değeri spek ondalığına yuvarlayarak üretir."""
    v = random.uniform(alt, ust)
    return round(v, spek.ondalik)


# ----------------------------------------------------------------------------
# Tablo tipine göre üretim
# ----------------------------------------------------------------------------

def _tek_sonuc(spek: Spesifikasyon, test_adi: str = "") -> dict:
    """
    3 seri × tek değer/metin.
    - Görünüş → sonuç 'Uygun'  (spesifikasyon 'Beyaz toz' olsa bile).
    - Teşhis  → sonuç 'Pozitif'.
    - Diğer METIN/BILGI → sabit_sonuc (varsa) yoksa 'Uygun'.
    - Sayısal tip (Elek Testi, Bulk/Tap Dansite): spek bandından DEĞER üretilir.
    """
    ad = test_adi.lower()
    if spek.limit_turu in (LimitTuru.METIN, LimitTuru.BILGI):
        # Görünüş/Teşhis → metin sonuç
        if "görünüş" in ad or "gorunus" in ad:
            return {"seriler": ["Uygun" for _ in range(SERI_SAYISI)]}
        if "teşhis" in ad or "teshis" in ad:
            return {"seriler": ["Pozitif" for _ in range(SERI_SAYISI)]}
        # Elek/Dansite gibi sayısal sonuç: alt/üst limit verilmişse DEĞER üret
        if spek.alt_limit is not None and spek.ust_limit is not None:
            alt, ust = _uretim_araligi(spek)
            return {"seriler": [f"{_deger(spek, alt, ust)} {spek.birim}".strip()
                                for _ in range(SERI_SAYISI)]}
        return {"seriler": [(spek.sabit_sonuc or "Uygun") for _ in range(SERI_SAYISI)]}
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
    """
    1..10 numune + Ortalama, her seri için.
    Karışım Tekdüzeliği tipik olarak ortalama ~%100 civarında çıkar.
    Aralık spesifikasyondan (alt/üst limit veya metinden) belirlenir.
    """
    alt, ust = _uretim_araligi(spek)
    # Karışım gibi geniş bantlarda (örn 85-115) gerçekçi dar üretim (~%100)
    if ust - alt > 20:
        merkez = (alt + ust) / 2
        alt2 = max(alt, merkez - 5)
        ust2 = min(ust, merkez + 5)
        alt, ust = alt2, ust2
    seriler = []
    for _ in range(SERI_SAYISI):
        olcumler = [round(random.uniform(alt, ust), spek.ondalik) for _ in range(10)]
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
        return _tek_sonuc(spek, test.ad)
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
