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

    Kullanıcı limit mantığı:
      - Hem alt hem üst → ikisi arasında üret.
      - Sadece üst (alt yok) → maksimum gibi: üstün altında üret.
      - Sadece alt (üst yok) → minimum gibi: altın üstünde üret.
    Limitler boşsa spesifikasyon metninden otomatik çıkarılır.
    """
    alt_l = spek.alt_limit
    ust_l = spek.ust_limit
    # Limitler boşsa metinden çıkarmayı dene
    if alt_l is None or ust_l is None:
        kaynak = spek.spesifikasyon_metni or spek.sabit_sonuc or ""
        m_alt, m_ust = _aralik_metinden(kaynak)
        if alt_l is None:
            alt_l = m_alt
        if ust_l is None:
            ust_l = m_ust

    # Hem alt hem üst
    if alt_l is not None and ust_l is not None:
        genislik = ust_l - alt_l
        pay = genislik * 0.20
        return alt_l + pay, ust_l - pay
    # Sadece üst → maksimum mantığı (üstün altında sağlıklı bant)
    if ust_l is not None:
        return ust_l * 0.05, ust_l * 0.6
    # Sadece alt → minimum mantığı (altın üstünde bant)
    if alt_l is not None:
        return alt_l * 1.05 + 0.01, alt_l * 1.05 + max(alt_l * 0.5, 1.0)
    # Eski limit türü alanları (geri uyumluluk)
    if spek.limit_turu is LimitTuru.MINIMUM and spek.minimum_deger is not None:
        taban = spek.minimum_deger
        return taban * 1.05 + 0.01, taban * 1.05 + max(taban * 0.5, 1.0)
    if spek.limit_turu is LimitTuru.MAKSIMUM and spek.maksimum_deger is not None:
        tavan = spek.maksimum_deger
        return tavan * 0.05, tavan * 0.6
    if spek.hedef_deger is not None:
        return spek.hedef_deger * 0.98, spek.hedef_deger * 1.02
    tek = _metinden_sayilar(spek.spesifikasyon_metni or spek.sabit_sonuc or "")
    if tek:
        return tek[0] * 1.02, tek[0] * 1.10
    return 0.0, 1.0


def _bicimle(deger, ondalik: int = 2) -> str:
    """
    Sayıyı sabit ondalık + Türkçe virgülle string'e çevirir.
    99.7 -> '99,70', 1 -> '1,00', 0.6 -> '0,60'. String/None aynen döner.
    """
    if deger is None or deger == "":
        return ""
    if isinstance(deger, str):
        return deger  # T.E., Uygun, Pozitif vb. dokunma
    try:
        return f"{float(deger):.{ondalik}f}".replace(".", ",")
    except (ValueError, TypeError):
        return str(deger)


def _deger(spek: Spesifikasyon, alt: float, ust: float) -> float:
    """Aralıkta rastgele bir değeri spek ondalığına yuvarlayarak üretir."""
    v = random.uniform(alt, ust)
    return round(v, spek.ondalik)


# ----------------------------------------------------------------------------
# Test sınıflandırma yardımcıları (ad / operasyondan)
# ----------------------------------------------------------------------------

def _film_asamasi(test: Test) -> bool:
    """
    Test film kaplama aşamasına mı ait? Ağırlık Tekdüzeliği ve Ortalama Ağırlık'ın
    kısım (tablet vs film) ayrımı operasyon adından belirlenir.
    'Film Kaplama' / 'Film Kaplam' → film; aksi halde (Tablet Baskı vb.) → tablet.
    """
    op = (test.operasyon or "").lower()
    return "film" in op


def _testtipi(ad: str) -> str:
    """Test adından özel sonuç-üretim tipini döndürür (band seçimi için)."""
    a = ad.lower()
    if "sertlik" in a:
        return "sertlik"
    if "dağılma" in a or "dagilma" in a:
        return "dagilma"
    if "çap" in a or "cap" in a:
        return "cap"
    if "kalınlık" in a or "kalinlik" in a:
        return "kalinlik"
    if "dissol" in a:
        return "dissolusyon"
    if "aşınma" in a or "asinma" in a:
        return "asinma"
    if "miktar tayini" in a:
        return "miktar"
    return ""


# ----------------------------------------------------------------------------
# Sertlik bant tablosu (kullanıcı kuralı)
# ----------------------------------------------------------------------------

# Bilinen minimum spesifikasyonları → (alt, üst) gerçekçi band
_SERTLIK_MIN = {
    3.0: (5.1, 6.8),
    4.0: (6.1, 7.9),
    5.0: (7.2, 8.8),
    6.0: (8.1, 9.3),
}
# Bilinen maksimum spesifikasyonları → (alt, üst) gerçekçi band
_SERTLIK_MAX = {
    10.0: (6.2, 8.3),
    15.0: (9.7, 12.6),
    20.0: (12.5, 15.7),
}


def _interpolasyon(x: float, tablo: dict) -> tuple[float, float]:
    """
    tablo: {spek_değeri: (alt, üst)}. x için (alt, üst) bandını döndürür.
    x tabloda varsa birebir; yoksa en yakın iki nokta arasında doğrusal
    interpolasyon; aralık dışındaysa en yakın noktanın bandına oranlanır.
    """
    anahtarlar = sorted(tablo.keys())
    if x in tablo:
        return tablo[x]
    # Aralık altı → en küçük anahtarın bandını oranla
    if x < anahtarlar[0]:
        k = anahtarlar[0]
        oran = x / k if k else 1.0
        a, u = tablo[k]
        return a * oran, u * oran
    # Aralık üstü → en büyük anahtarın bandını oranla
    if x > anahtarlar[-1]:
        k = anahtarlar[-1]
        oran = x / k if k else 1.0
        a, u = tablo[k]
        return a * oran, u * oran
    # İki anahtar arası → doğrusal interpolasyon
    for i in range(len(anahtarlar) - 1):
        k0, k1 = anahtarlar[i], anahtarlar[i + 1]
        if k0 <= x <= k1:
            t = (x - k0) / (k1 - k0) if k1 != k0 else 0.0
            a0, u0 = tablo[k0]
            a1, u1 = tablo[k1]
            return a0 + (a1 - a0) * t, u0 + (u1 - u0) * t
    return tablo[anahtarlar[-1]]


def _sertlik_araligi(spek: Spesifikasyon) -> tuple[float, float]:
    """
    Sertlik için gerçekçi üretim bandı. Kullanıcı kuralı:
      - Minimum spek → bilinen bant / interpolasyon; sonuç min'in ALTINA inemez.
      - Maksimum spek → bilinen bant / interpolasyon; sonuç max'ın ÜSTÜNE çıkamaz.
      - Aralık spek (alt-üst) → ortalamaya yakın, dağılımlı band.
    """
    # Aralık verilmişse: ortalamaya yakın ama yapışık olmayan band
    if spek.alt_limit is not None and spek.ust_limit is not None:
        a, u = spek.alt_limit, spek.ust_limit
        orta = (a + u) / 2.0
        yari = (u - a) / 2.0
        return max(a, orta - yari * 0.5), min(u, orta + yari * 0.5)
    # Minimum
    if spek.minimum_deger is not None:
        alt, ust = _interpolasyon(spek.minimum_deger, _SERTLIK_MIN)
        return max(alt, spek.minimum_deger), ust   # min'in altına inme
    # Maksimum
    if spek.maksimum_deger is not None:
        alt, ust = _interpolasyon(spek.maksimum_deger, _SERTLIK_MAX)
        return alt, min(ust, spek.maksimum_deger)   # max'ın üstüne çıkma
    # Limit yoksa metinden dene
    sayilar = _metinden_sayilar(spek.spesifikasyon_metni or "")
    if sayilar:
        alt, ust = _interpolasyon(sayilar[0], _SERTLIK_MIN)
        return max(alt, sayilar[0]), ust
    return 7.0, 9.0


def _ozel_arabant(test: Test, spek: Spesifikasyon) -> tuple[float, float] | None:
    """
    Test tipine göre özel (alt, üst) band döndürür. Tip özel değilse None.
    Kullanıcı kuralları (PVR örneğinden netleşen):
      - sertlik   → _sertlik_araligi
      - dagilma   → sabit 4–7 dk
      - cap       → hedef+0.11 .. hedef+0.13
      - kalinlik  → spek aralığının iç %60'ı
      - dissolusyon → sabit %98–105
      - asinma    → max×0.2 .. max×0.3
      - miktar    → hedef .. hedef×1.025
    """
    tip = _testtipi(test.ad)
    if tip == "sertlik":
        return _sertlik_araligi(spek)
    if tip == "dagilma":
        return 4.0, 7.0
    if tip == "cap":
        hedef = spek.hedef_deger
        if hedef is None:
            hedef = _hedef_metinden(spek)
        if hedef is not None:
            return hedef + 0.11, hedef + 0.13
        return None
    if tip == "kalinlik":
        a, u = _uretim_sinirlari(spek)
        if a is not None and u is not None:
            pay = (u - a) * 0.20
            return a + pay, u - pay
        return None
    if tip == "dissolusyon":
        return 98.0, 105.0
    if tip == "asinma":
        maks = spek.maksimum_deger
        if maks is None:
            sayilar = _metinden_sayilar(spek.spesifikasyon_metni or "")
            maks = sayilar[0] if sayilar else None
        if maks is not None:
            return maks * 0.2, maks * 0.3
        return None
    if tip == "miktar":
        hedef = spek.hedef_deger
        if hedef is None:
            hedef = _hedef_metinden(spek)
        if hedef is not None:
            ust = hedef * 1.025
            if spek.ust_limit is not None:
                ust = min(ust, spek.ust_limit)
            return hedef, ust
        return None
    return None


def _hedef_metinden(spek: Spesifikasyon) -> float | None:
    """Hedef değeri sayısal alandan ya da metinden çıkarır."""
    if spek.hedef_deger is not None:
        return spek.hedef_deger
    sayilar = _metinden_sayilar(spek.hedef_metin or "")
    if sayilar:
        return sayilar[0]
    alt, ust = _aralik_metinden(spek.spesifikasyon_metni or "")
    if alt is not None and ust is not None:
        return (alt + ust) / 2.0
    return None


def _uretim_sinirlari(spek: Spesifikasyon) -> tuple:
    """alt_limit/ust_limit; boşsa metinden (alt, üst) çıkarır."""
    alt_l, ust_l = spek.alt_limit, spek.ust_limit
    if alt_l is None or ust_l is None:
        m_alt, m_ust = _aralik_metinden(spek.spesifikasyon_metni or spek.sabit_sonuc or "")
        if alt_l is None:
            alt_l = m_alt
        if ust_l is None:
            ust_l = m_ust
    return alt_l, ust_l


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
        # Görünüş/Teşhis/Boyar Madde → metin sonuç
        if "görünüş" in ad or "gorunus" in ad:
            return {"seriler": ["Uygun" for _ in range(SERI_SAYISI)]}
        if "teşhis" in ad or "teshis" in ad:
            return {"seriler": ["Pozitif" for _ in range(SERI_SAYISI)]}
        if "boyar" in ad:
            return {"seriler": ["Pozitif" for _ in range(SERI_SAYISI)]}
        # Elek/Dansite gibi sayısal sonuç: alt/üst limit verilmişse DEĞER üret
        if spek.alt_limit is not None and spek.ust_limit is not None:
            alt, ust = _uretim_araligi(spek)
            return {"seriler": [f"{_deger(spek, alt, ust)} {spek.birim}".strip()
                                for _ in range(SERI_SAYISI)]}
        return {"seriler": [(spek.sabit_sonuc or "Uygun") for _ in range(SERI_SAYISI)]}
    # Aşınma: max×0.2 – max×0.3, çıktı "%0.XXX" (3 ondalık), tek satır.
    if "aşınma" in ad or "asinma" in ad:
        maks = spek.maksimum_deger
        if maks is None:
            sayilar = _metinden_sayilar(spek.spesifikasyon_metni or spek.sabit_sonuc or "")
            maks = sayilar[0] if sayilar else 1.0
        return {"seriler": [f"%{round(random.uniform(maks * 0.2, maks * 0.3), 3):.3f}"
                            for _ in range(SERI_SAYISI)]}
    alt, ust = _uretim_araligi(spek)
    return {"seriler": [f"{_deger(spek, alt, ust)} {spek.birim}".strip()
                        for _ in range(SERI_SAYISI)]}


def _iki_numune(spek: Spesifikasyon) -> dict:
    """
    Numune-1, Numune-2 + Sonuç(ortalama), her seri için.
    Maksimum değeri 'T.E.' (tespit edilemedi) ise tüm sonuçlar 'T.E.' olur.
    """
    te = (spek.maksimum_metin or "").strip().upper().replace(" ", "")
    if te in ("T.E.", "T.E", "TE", "TESPİTEDİLEMEDİ", "TESPITEDILEMEDI"):
        return {"te": True, "seriler": [
            {"numune_1": "T.E.", "numune_2": "T.E.", "sonuc": "T.E."}
            for _ in range(SERI_SAYISI)]}
    alt, ust = _uretim_araligi(spek)
    seriler = []
    for _ in range(SERI_SAYISI):
        n1 = _deger(spek, alt, ust)
        n2 = _deger(spek, alt, ust)
        seriler.append({
            "numune_1": _bicimle(n1), "numune_2": _bicimle(n2),
            "sonuc": _bicimle((n1 + n2) / 2),
        })
    return {"seriler": seriler}


def _on_numune(spek: Spesifikasyon) -> dict:
    """
    1..10 numune + Ortalama, her seri için (Karışım Tekdüzeliği).
    Kullanıcı kuralı: her zaman değerler %97–%105 arasında, seri ortalaması
    %100–%102 arasında olmalı (spesifikasyon %85–115 olsa bile gerçekçi sağlıklı veri).
    """
    DEGER_ALT, DEGER_UST = 97.0, 105.0      # tek tek ölçümler bu bantta
    ORT_ALT, ORT_UST = 100.0, 102.0          # seri ortalaması bu bantta
    seriler = []
    for _ in range(SERI_SAYISI):
        hedef_ort = random.uniform(ORT_ALT, ORT_UST)
        # hedef ortalamayı tutturacak 10 değer üret
        olcumler = []
        for _ in range(10):
            v = random.uniform(DEGER_ALT, DEGER_UST)
            olcumler.append(v)
        # ölçümleri hedef ortalamaya kaydır
        mevcut = sum(olcumler) / len(olcumler)
        fark = hedef_ort - mevcut
        olcumler = [min(DEGER_UST, max(DEGER_ALT, round(v + fark, 2))) for v in olcumler]
        ort = ortalama(olcumler)
        seriler.append({
            "olcumler": [_bicimle(o) for o in olcumler],
            "ortalama": _bicimle(ort),
        })
    return {"seriler": seriler}


def _bos_nokta(spek: Spesifikasyon, test: Test = None, numune_sayisi: int = 10) -> dict:
    """
    n numune × 3 nokta (Baş/Orta/Son), her nokta ortalaması + seri Sonucu.
    Sertlik, Kalınlık, Çap, Dağılma, Dissolüsyon.
    Test tipine göre özel band ve nokta-başı numune sayısı uygulanır.
    """
    tip = _testtipi(test.ad) if test is not None else ""
    ozel = _ozel_arabant(test, spek) if test is not None else None
    if ozel is not None:
        alt, ust = ozel
    else:
        alt, ust = _uretim_araligi(spek)
    # Dissolüsyon nokta başına 6 değer (PVR örneği), diğerleri 10
    if tip == "dissolusyon":
        numune_sayisi = 6
    seriler = []
    for _ in range(SERI_SAYISI):
        noktalar = {}
        nokta_ort = []
        for nokta in NOKTA_ADLARI:
            olcumler = [round(random.uniform(alt, ust), spek.ondalik) for _ in range(numune_sayisi)]
            o = round(ortalama(olcumler), 2)
            noktalar[nokta] = {
                "olcumler": [_bicimle(x) for x in olcumler],
                "ortalama": _bicimle(o),
                "_ham_ortalama": o,   # türetme/eşleşme için ham değer
            }
            nokta_ort.append(o)
        seriler.append({
            "noktalar": noktalar,
            "sonuc": _bicimle(ortalama(nokta_ort)),
        })
    return {"seriler": seriler}


def _agirlik_band(spek: Spesifikasyon, hedef_override: float | None = None) -> tuple[float, float]:
    """
    Ağırlık tekdüzeliği üretim bandı: hedef – hedef×1.025.
    Alt uç hedef'tir; böylece bireysel ölçümler ve nokta/seri ortalamaları
    hedefin altına inmez ve Ortalama Ağırlık tablosuyla çelişmez.
    hedef_override verilirse (aynı aşamadaki Ortalama Ağırlık'tan) o kullanılır.
    """
    hedef = hedef_override if hedef_override is not None else _hedef_metinden(spek)
    if hedef is None:
        alt, ust = _uretim_sinirlari(spek)
        if alt is not None and ust is not None:
            hedef = (alt + ust) / 2.0
    if hedef is None:
        return _uretim_araligi(spek)
    return hedef, hedef * 1.025


def _agirlik_tekduzeligi(spek: Spesifikasyon, film: bool = False,
                         hedef_override: float | None = None) -> dict:
    """
    Ağırlık Tekdüzeliği. Kısım, ürün formundan (operasyon) belirlenir:
      - Tablet aşaması (film=False): Baş/Orta/Son × 10 numune + Ort/RSD%/SD.
      - Film aşaması  (film=True) : Baş/Orta/Son YOK, seri başına 20 düz değer
                                    + Ort/RSD%/SD (tek blok).
    Band: hedef – hedef×1.025. hedef_override verilirse o hedef kullanılır.
    """
    alt, ust = _agirlik_band(spek, hedef_override)
    seriler = []
    if film:
        # Seri başına düz 20 değer
        for _ in range(SERI_SAYISI):
            olcumler = [round(random.uniform(alt, ust), spek.ondalik) for _ in range(20)]
            o = round(ortalama(olcumler), 2)
            seriler.append({
                "duz": True,
                "olcumler": [_bicimle(x) for x in olcumler],
                "ortalama": _bicimle(o),
                "_ham_ortalama": o,
                "rsd": _bicimle(rsd_yuzde(olcumler)),
                "sd": _bicimle(std_sapma(olcumler)),
            })
        return {"film": True, "seriler": seriler}
    # Tablet: Baş/Orta/Son × 10 numune
    for _ in range(SERI_SAYISI):
        noktalar = {}
        for nokta in NOKTA_ADLARI:
            olcumler = [round(random.uniform(alt, ust), spek.ondalik) for _ in range(10)]
            o = round(ortalama(olcumler), 2)
            noktalar[nokta] = {
                "olcumler": [_bicimle(x) for x in olcumler],
                "ortalama": _bicimle(o),
                "_ham_ortalama": o,    # Ortalama Ağırlık türetmesi için
                "rsd": _bicimle(rsd_yuzde(olcumler)),
                "sd": _bicimle(std_sapma(olcumler)),
            }
        seriler.append({"noktalar": noktalar})
    return {"film": False, "seriler": seriler}


def _miktar_tayini(spek: Spesifikasyon) -> dict:
    """
    Miktar Tayini (PVR Tablo.45/46): her seri Baş/Orta/Son; her nokta altında
    Numune-1, Numune-2, Sonuç(=ort). Seri sonunda 3 noktanın Sonuç ortalaması.
    Band: hedef .. hedef×1.025 (üst spek aşılmaz).
    """
    ozel = _ozel_arabant_miktar(spek)
    if ozel is not None:
        alt, ust = ozel
    else:
        alt, ust = _uretim_araligi(spek)
    seriler = []
    for _ in range(SERI_SAYISI):
        noktalar = {}
        nokta_sonuc = []
        for nokta in NOKTA_ADLARI:
            n1 = round(random.uniform(alt, ust), spek.ondalik)
            n2 = round(random.uniform(alt, ust), spek.ondalik)
            s = round((n1 + n2) / 2, spek.ondalik)
            noktalar[nokta] = {
                "numune_1": _bicimle(n1),
                "numune_2": _bicimle(n2),
                "sonuc": _bicimle(s),
                "_ham_sonuc": s,
            }
            nokta_sonuc.append(s)
        seriler.append({
            "noktalar": noktalar,
            "ortalama": _bicimle(ortalama(nokta_sonuc)),
        })
    return {"seriler": seriler}


def _ozel_arabant_miktar(spek: Spesifikasyon) -> tuple[float, float] | None:
    """Miktar Tayini bandı: hedef .. hedef×1.025 (üst spek aşılmaz)."""
    hedef = _hedef_metinden(spek)
    if hedef is None:
        return None
    ust = hedef * 1.025
    if spek.ust_limit is not None:
        ust = min(ust, spek.ust_limit)
    return hedef, ust


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
        # Miktar Tayini: Baş/Orta/Son × Numune-1/2/Sonuç + seri ortalaması.
        # Diğer iki-numune testleri (impurite vb.): klasik düz Numune-1/2/Sonuç.
        if _testtipi(test.ad) == "miktar":
            return _miktar_tayini(spek)
        return _iki_numune(spek)
    if t is TabloTipi.ON_NUMUNE:
        return _on_numune(spek)
    if t is TabloTipi.BOS_NOKTA:
        return _bos_nokta(spek, test)
    if t is TabloTipi.AGIRLIK_TEKDUZELIGI:
        return _agirlik_tekduzeligi(spek, film=_film_asamasi(test),
                                    hedef_override=getattr(test, "_hedef_devralinan", None))
    if t is TabloTipi.MATRIS:
        return _matris(spek)
    return _tek_sonuc(spek)


def tum_testleri_uret(testler: list[Test], tohum: int | None = None) -> None:
    """
    Verilen testlerin her biri için sonuç verisi üretip Test.sonuc_verisi'ne yazar.
    tohum verilirse tekrarlanabilir sonuç üretilir (test/doğrulama için).

    Özel bağlantı: 'Ortalama Ağırlık' sonuçları, AYNI OPERASYONDAKİ 'Ağırlık
    Tekdüzeliği'nin ortalamalarıyla BİREBİR eşleşir (kullanıcı kuralı). Üründe
    hem Tablet Baskı hem Film Kaplama aşamalarında ayrı Ağırlık Tekdüzeliği /
    Ortalama Ağırlık bulunabilir; her biri kendi aşamasıyla eşlenir.
    """
    if tohum is not None:
        random.seed(tohum)

    # 0) Ağırlık Tekdüzeliği'nin değer bandı için HEDEF gerekir. Kendi
    #    spesifikasyonunda hedef yoksa AYNI OPERASYONDAKİ Ortalama Ağırlık
    #    spesifikasyonundan türetilir (örn. 700 mg ± %5 → hedef 700).
    ort_agirlik_map: dict[str, Test] = {}
    for test in testler:
        if "ortalama ağırlık" in test.ad.lower():
            ort_agirlik_map[(test.operasyon or "").lower()] = test
    for test in testler:
        if test.tablo_tipi is TabloTipi.AGIRLIK_TEKDUZELIGI:
            if _hedef_metinden(test.spesifikasyon) is None:
                esi = ort_agirlik_map.get((test.operasyon or "").lower())
                if esi is not None:
                    test._hedef_devralinan = _hedef_metinden(esi.spesifikasyon)

    # 1) Tüm Ağırlık Tekdüzeliği testlerini önce üret (operasyona göre indeksle)
    agirlik_map: dict[str, Test] = {}
    for test in testler:
        if test.tablo_tipi is TabloTipi.AGIRLIK_TEKDUZELIGI:
            test.sonuc_verisi = test_verisi_uret(test)
            agirlik_map[(test.operasyon or "").lower()] = test

    # 2) Diğer testleri üret; Ortalama Ağırlık'ı aynı operasyondaki Ağırlık
    #    Tekdüzeliği'nden türet. Aynı operasyonda eşi yoksa mevcut tek tabletten
    #    (varsa) ya da genel üretimden düşülür.
    for test in testler:
        if test.tablo_tipi is TabloTipi.AGIRLIK_TEKDUZELIGI:
            continue
        if "ortalama ağırlık" in test.ad.lower() and agirlik_map:
            esi = agirlik_map.get((test.operasyon or "").lower())
            if esi is None:
                # Aynı film/tablet kısmındaki ilk uygun testi seç
                film = _film_asamasi(test)
                for a in agirlik_map.values():
                    if _film_asamasi(a) == film:
                        esi = a
                        break
            if esi is None:
                esi = next(iter(agirlik_map.values()))
            test.sonuc_verisi = _ortalama_agirlik_turet(esi, test.spesifikasyon)
        else:
            test.sonuc_verisi = test_verisi_uret(test)


def _ortalama_agirlik_turet(agirlik_test: Test, spek: Spesifikasyon) -> dict:
    """
    Ortalama Ağırlık sonuç verisini, Ağırlık Tekdüzeliği'nin ortalamalarından
    türetir (iki tablo birebir eşleşir). Kısma göre iki yapı:
      - Tablet: her seri/nokta tek değer = o noktanın 10 tablet ortalaması;
        seri Sonuç'u 3 noktanın ortalaması. (BOS_NOKTA uyumlu)
      - Film  : her seri tek değer = o serinin 20 tablet ortalaması;
        Baş/Orta/Son yok. ('film' işaretli)
    Hedef tanımlıysa türetilen ortalamalar hedeften küçük olamaz (alta klipslenir).
    """
    hedef = _hedef_metinden(spek)
    if hedef is None:
        hedef = getattr(agirlik_test, "_hedef_devralinan", None)
    film = agirlik_test.sonuc_verisi.get("film", False)

    if film:
        seriler = []
        for sr in agirlik_test.sonuc_verisi.get("seriler", []):
            ham = sr.get("_ham_ortalama", 0.0)
            if hedef is not None and ham < hedef:
                ham = hedef
            seriler.append({"ortalama": _bicimle(ham), "_ham_ortalama": ham})
        return {"film": True, "seriler": seriler}

    # Tablet
    seriler = []
    for sr in agirlik_test.sonuc_verisi.get("seriler", []):
        noktalar = {}
        nokta_ort = []
        for nokta in NOKTA_ADLARI:
            ham = sr.get("noktalar", {}).get(nokta, {}).get("_ham_ortalama", 0.0)
            if hedef is not None and ham < hedef:
                ham = hedef
            noktalar[nokta] = {"olcumler": [_bicimle(ham)], "ortalama": _bicimle(ham),
                               "_ham_ortalama": ham}
            nokta_ort.append(ham)
        seriler.append({
            "noktalar": noktalar,
            "sonuc": _bicimle(ortalama(nokta_ort)),
        })
    return {"film": False, "seriler": seriler}
