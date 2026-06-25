"""
Akış Şeması veri motoru (v2).

Tablo 7'deki (IPK) operasyonlara göre gruplar. Her operasyon için:
  - operasyon adı (kutu etiketi)
  - IPK testleri (Tablo 7'den, o operasyon)
  - kimyasal testler (Tablo 6'da olup Tablo 7'de olmayan, o operasyon)

Tartım dahil edilmez. Operasyon sırası: Karıştırma, Tablet Baskı, Film Kaplama, Blisterleme.
"""

from __future__ import annotations

import re

from core.kural_motoru import turet

# Akış şemasında gösterilecek operasyon sırası ve görünen ad
_OP_SIRA = [
    ("Karıştırma", "Karıştırma"),
    ("Tablet Baskı", "Tablet Baskı"),
    ("Film Kaplama", "Film Kaplama"),
    ("Blisterleme", "Blisterleme"),
]


def _hammadde_kutulari(proje) -> list[str]:
    """
    Üretim yöntemi açıklamalarından hammadde kutularını çıkarır.
    Her aşamada işlem gören hammadde grubu ayrı bir kutu olur.
    Hammaddesi olmayan aşama atlanır.
    Dönüş: ["Parasetamol", "Kafein\\nMısır Nişastası\\n...", "Stearik asit", ...]
    """
    adimlar = getattr(proje, "uretim_adimlari", None) or []
    kutular = []
    for adim in adimlar:
        aciklama = adim[1] if len(adim) > 1 else ""
        maddeler = _maddeleri_ayikla(aciklama)
        if maddeler:
            kutular.append("\n".join("- " + m for m in maddeler))
    return kutular


# Hammadde olmayan, cümlede geçen ama madde sayılmayacak kelimeler
_HAM_DISI = ("aşama", "karışım", "konteynır", "konteyner", "dakika", "rpm",
             "operasyon", "tablet", "bulk", "makina", "süre", "hız", "ilave",
             "elek", "elenerek", "karıştırılır", "alınır", "hazırlanan",
             "baskı", "kaplama", "kaplanır", "blister", "solüsyon", "çözülerek")


def _maddeleri_ayikla(aciklama: str) -> list[str]:
    """
    Açıklamadan hammadde isimlerini çıkarır.
    Yaklaşım: Büyük harfle başlayan madde isimlerini ve 'X kg Madde' kalıplarını
    yakalar. 'elenerek', 'ilave' gibi işlem kelimelerini atar.
    """
    if not aciklama:
        return []
    maddeler = []

    # 1) "<sayı> kg <Madde Adı>" kalıpları
    for m in re.finditer(
            r"\d+[.,]?\d*\s*kg\s+([A-ZÇĞİÖŞÜ][\wçğıöşüÇĞİÖŞÜ\s./()-]*?)"
            r"(?=,|\.|;|\s+\d+[.,]?\d*\s*kg|\s+elenerek|\s+ilave|\s+poşete|$)",
            aciklama):
        ad = _ad_temizle(m.group(1))
        if ad and ad not in maddeler:
            maddeler.append(ad)

    if maddeler:
        return maddeler

    # 2) kg yoksa: Büyük harfle başlayan ardışık kelime gruplarını madde say
    #    "Parasetamol konteynıra alınır üzerine, Kafein, Mısır nişastası, ..."
    parcalar = re.split(r"[,]|\s+üzerine|\s+elenerek|\s+ilave\s+edilip|\s+ile\s+", aciklama)
    for parca in parcalar:
        ad = _ad_temizle(parca)
        if not ad:
            continue
        # ilk kelimesi büyük harf mi ve işlem kelimesi değil mi
        ilk = ad.split()[0] if ad.split() else ""
        if ilk[:1].isupper() and not any(k in ad.lower() for k in _HAM_DISI):
            if ad not in maddeler:
                maddeler.append(ad)
    return maddeler


def _ad_temizle(ad: str) -> str:
    """Madde adından işlem fiillerini ve gereksiz kuyrukları temizler."""
    ad = re.sub(r"\s+", " ", ad or "").strip(" .,;")
    # sonundaki işlem fiillerini at
    ad = re.sub(r"\s+(konteynıra|konteynera|elenerek|ilave|poşete|alınır|"
                r"çözülerek|karıştırılır|hazırlanan).*$", "", ad, flags=re.IGNORECASE)
    ad = ad.strip(" .,;-")
    # çok uzunsa (cümle parçası) madde değildir
    if len(ad.split()) > 6:
        return ""
    return ad


def akis_semasi_hazirla(proje) -> list[dict]:
    """
    Tablo 7 operasyonlarına göre akış kutularını hazırlar. Her kutu:
      {
        "operasyon": "Karıştırma",
        "ipk_testleri": [...],       # Tablo 7 (o operasyon)
        "kimyasal_testler": [...],   # Tablo 6  Tablo 7 (o operasyon)
      }
    """
    kart = proje.spek_karti
    if not kart.testler:
        return []
    ops = proje.urun_formu.operasyonlar
    tablo6 = turet(kart.testler, kart.etkin_maddeler, ops,
                   cift_katman=getattr(kart, "cift_katman", False),
                   tablet_ipk=getattr(kart, "tablet_ipk", {}),
                   ozel_test_kurallari=getattr(kart, "ozel_test_kurallari", {}))

    ipk_map = {}        # operasyon -> [test adı]
    kimyasal_map = {}   # operasyon -> [test adı]
    for t in tablo6:
        op = t.operasyon
        if getattr(t, "_impurite", False):
            continue
        ad = t.ad + ("*" if t.yildizli else "")
        if t.ipk:
            ipk_map.setdefault(op, []).append(ad)
        else:
            kimyasal_map.setdefault(op, []).append(ad)

    kutular = []
    for op_anahtar, op_ad in _OP_SIRA:
        # bu operasyonun hiç testi yoksa atla (ör. Blisterleme yoksa)
        if op_anahtar not in ipk_map and op_anahtar not in kimyasal_map:
            continue
        kutular.append({
            "operasyon": op_ad,
            "ipk_testleri": ipk_map.get(op_anahtar, []),
            "kimyasal_testler": kimyasal_map.get(op_anahtar, []),
        })

    return {
        "operasyonlar": kutular,
        "hammaddeler": _hammadde_kutulari(proje),
    }
