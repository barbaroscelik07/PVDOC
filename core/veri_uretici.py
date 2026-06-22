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


def _agirlik_tekduzeligi(spek: Spesifikasyon) -> dict:
    """20 numune × 3 nokta + Ortalama/RSD%/SD, her nokta için."""
    alt, ust = _uretim_araligi(spek)
    seriler = []
    for _ in range(SERI_SAYISI):
        noktalar = {}
        for nokta in NOKTA_ADLARI:
            olcumler = [_deger(spek, alt, ust) for _ in range(20)]
            o = round(ortalama(olcumler), 2)
            noktalar[nokta] = {
                "olcumler": [_bicimle(x) for x in olcumler],
                "ortalama": _bicimle(o),
                "_ham_ortalama": o,    # Ortalama Ağırlık türetmesi için
                "rsd": _bicimle(rsd_yuzde(olcumler)),
                "sd": _bicimle(std_sapma(olcumler)),
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

    Özel bağlantı: 'Ortalama Ağırlık' sonuçları, 'Ağırlık Tekdüzeliği'nin
    nokta-ortalamalarıyla BİREBİR eşleşir (kullanıcı kuralı). Bu yüzden önce
    Ağırlık Tekdüzeliği üretilir, sonra Ortalama Ağırlık ondan türetilir.
    """
    if tohum is not None:
        random.seed(tohum)

    # 1) Ağırlık Tekdüzeliği testini önce üret
    agirlik_test = None
    for test in testler:
        if test.tablo_tipi is TabloTipi.AGIRLIK_TEKDUZELIGI:
            test.sonuc_verisi = test_verisi_uret(test)
            agirlik_test = test
            break

    # 2) Diğer testleri üret; Ortalama Ağırlık'ı Ağırlık Tekdüzeliği'nden türet
    for test in testler:
        if test is agirlik_test:
            continue
        if "ortalama ağırlık" in test.ad.lower() and agirlik_test is not None:
            test.sonuc_verisi = _ortalama_agirlik_turet(agirlik_test, test.spesifikasyon)
        else:
            test.sonuc_verisi = test_verisi_uret(test)


def _ortalama_agirlik_turet(agirlik_test: Test, spek: Spesifikasyon) -> dict:
    """
    Ortalama Ağırlık sonuç verisini, Ağırlık Tekdüzeliği'nin nokta
    ortalamalarından türetir. Her seri/nokta için tek değer = o noktanın
    20 tablet ortalaması. Böylece iki tablo birebir eşleşir.
    Yapı BOS_NOKTA ile uyumlu: seriler[i]['noktalar'][nokta]['ortalama'].
    """
    seriler = []
    for sr in agirlik_test.sonuc_verisi.get("seriler", []):
        noktalar = {}
        nokta_ort = []
        for nokta in NOKTA_ADLARI:
            # Ağırlık Tekdüzeliği'nin HAM nokta ortalamasını kullan (birebir eşleşme)
            ham = sr.get("noktalar", {}).get(nokta, {}).get("_ham_ortalama", 0.0)
            noktalar[nokta] = {"olcumler": [_bicimle(ham)], "ortalama": _bicimle(ham),
                               "_ham_ortalama": ham}
            nokta_ort.append(ham)
        seriler.append({
            "noktalar": noktalar,
            "sonuc": _bicimle(ortalama(nokta_ort)),
        })
    return {"seriler": seriler}
