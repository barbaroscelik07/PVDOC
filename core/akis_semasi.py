"""
Akış Şeması veri motoru (Faz 1).

Üretim yöntemi adımlarından + Tablo 6/7'den 4 sütunlu akış şeması verisini hazırlar:
  - Kutular: her aşama bir kutu (isim açıklamadaki fiilden çıkarılır)
  - IPK işareti: açıklamada "(IPK N)" varsa o kutudan sağa ok çıkar
  - Sütun 2 (IPK testleri): kutunun operasyon adına göre Tablo 7'den
  - Sütun 3 (kimyasal): operasyonda olup Tablo 7'de olmayan testler (Tablo 6'dan)
  - Sütun 0 (hammaddeler): açıklamadan çıkarılan madde isimleri
"""

from __future__ import annotations

import re

from core.kural_motoru import turet, OP_NO


def _norm(s: str) -> str:
    tr = str.maketrans({"ı": "i", "İ": "i", "ş": "s", "Ş": "s", "ğ": "g",
                        "ç": "c", "ö": "o", "ü": "u", "I": "i"})
    return (s or "").translate(tr).lower().strip()


# Açıklamadaki ana fiil → kısa kutu ismi (öncelik sırası önemli: en belirleyici önce)
_FIIL_KUTU = [
    ("tartıl", "Tartım"),
    ("baskı", "Tablet Baskı"),
    ("baski", "Tablet Baskı"),
    ("kaplan", "Film Kaplama"),
    ("film kaplama", "Film Kaplama"),
    ("blister", "Ambalajlama"),
    ("kutulan", "Ambalajlama"),
    ("karıştırıl", "Karıştırma"),   # 'karıştır' eleme'den önce gelmeli
    ("karistiril", "Karıştırma"),
    ("karıştır", "Karıştırma"),
    ("elen", "Eleme"),
    ("elekten", "Eleme"),
]


def kutu_ismi_cikar(aciklama: str) -> str:
    """Açıklamadaki ana işlem fiilinden kısa kutu ismi çıkarır."""
    n = _norm(aciklama)  # ı→i, ş→s ...
    # anahtarlar normalize edilmiş halde aranır
    sirali = [
        ("tartil", "Tartım"),
        ("baski", "Tablet Baskı"),
        ("kaplan", "Film Kaplama"),
        ("film kaplama", "Film Kaplama"),
        ("blister", "Ambalajlama"),
        ("kutulan", "Ambalajlama"),
        ("karistiril", "Karıştırma"),
        ("karistir", "Karıştırma"),
        ("elen", "Eleme"),
        ("elekten", "Eleme"),
    ]
    for anahtar, isim in sirali:
        if anahtar in n:
            return isim
    return "İşlem"


def ipk_no_cikar(aciklama: str):
    """Açıklamada '(IPK N)' varsa numarayı döndürür, yoksa None."""
    m = re.search(r"IPK\s*[-–]?\s*(\d+)", aciklama or "", re.IGNORECASE)
    return int(m.group(1)) if m else None


def hammadde_cikar(aciklama: str) -> list[str]:
    """
    Açıklamadan hammadde isimlerini çıkarır. Genellikle:
    'X kg Madde Adı, Y kg Diğer Madde ... konteynıra alınır' biçiminde.
    Miktar+birim kalıplarını ('6.750 kg', '0.450 kg') ayraç olarak kullanır.
    """
    if not aciklama:
        return []
    maddeler = []
    # "<sayı> kg <Madde>" kalıplarını yakala
    for m in re.finditer(r"\d+[.,]?\d*\s*kg\s+([A-ZÇĞİÖŞÜ][^,.;()]*?)(?=,|\.|;|\d+[.,]?\d*\s*kg|$)",
                         aciklama):
        ad = m.group(1).strip()
        # sonundaki bağlaçları temizle
        ad = re.sub(r"\s+(ve|ile)\s*$", "", ad).strip()
        if ad and len(ad) > 2 and ad not in maddeler:
            maddeler.append(ad)
    return maddeler


def akis_semasi_hazirla(proje) -> list[dict]:
    """
    Akış şeması kutularını hazırlar. Her kutu:
      {
        "operasyon": "Karıştırma",       # kutu ismi (fiilden)
        "operasyon_adi": "Karıştırma",   # Tablo 6/7 eşleşmesi için operasyon
        "operasyon_no": 2,
        "asama_no": 2,
        "ipk_no": 1 veya None,            # (IPK N) işareti
        "hammaddeler": [...],             # sol sütun
        "ipk_testleri": [...],            # sütun 2 (Tablo 7'den) — sadece ipk_no varsa
        "kimyasal_testler": [...],        # sütun 3 (Tablo 6'dan) — sadece ipk_no varsa
      }
    """
    adimlar = getattr(proje, "uretim_adimlari", None) or []
    kart = proje.spek_karti

    # Tablo 6 (tüm aşama testleri) ve Tablo 7 (IPK testleri) türet
    ops = proje.urun_formu.operasyonlar
    tablo6 = turet(kart.testler, kart.etkin_maddeler, ops,
                   cift_katman=getattr(kart, "cift_katman", False),
                   tablet_ipk=getattr(kart, "tablet_ipk", {}),
                   ozel_test_kurallari=getattr(kart, "ozel_test_kurallari", {}))
    # operasyon -> IPK testleri (Tablo 7), kimyasal testleri (Tablo 6 \ Tablo 7)
    ipk_map = {}        # operasyon -> [test adı]
    kimyasal_map = {}   # operasyon -> [test adı]
    for t in tablo6:
        op = t.operasyon
        ad = t.ad + ("*" if t.yildizli else "")
        if getattr(t, "_impurite", False):
            continue  # impurite alt satırları şemada gösterme
        if t.ipk:
            ipk_map.setdefault(op, []).append(ad)
        else:
            kimyasal_map.setdefault(op, []).append(ad)

    kutular = []
    for adim in adimlar:
        baslik = adim[0]
        aciklama = adim[1] if len(adim) > 1 else ""
        # operasyon no / aşama no
        m = re.search(r"operasyon\s*(\d+)\s*[:：]\s*aşama\s*(\d+)", baslik, re.IGNORECASE)
        op_no = int(m.group(1)) if m else 0
        as_no = int(m.group(2)) if m else 0

        kutu_isim = kutu_ismi_cikar(aciklama)
        ipk_no = ipk_no_cikar(aciklama)

        # operasyon adı (Tablo 6/7 eşleşmesi): kutu ismine göre
        op_adi_map = {"Karıştırma": "Karıştırma", "Eleme": "Karıştırma",
                      "Tartım": "Karıştırma", "Tablet Baskı": "Tablet Baskı",
                      "Film Kaplama": "Film Kaplama", "Ambalajlama": "Blisterleme"}
        op_adi = op_adi_map.get(kutu_isim, kutu_isim)

        kutu = {
            "operasyon": kutu_isim,
            "operasyon_adi": op_adi,
            "operasyon_no": op_no,
            "asama_no": as_no,
            "ipk_no": ipk_no,
            "hammaddeler": hammadde_cikar(aciklama),
            "ipk_testleri": [],
            "kimyasal_testler": [],
        }
        # IPK işareti varsa testleri bağla
        if ipk_no is not None:
            kutu["ipk_testleri"] = ipk_map.get(op_adi, [])
            kutu["kimyasal_testler"] = kimyasal_map.get(op_adi, [])
        kutular.append(kutu)

    return kutular
