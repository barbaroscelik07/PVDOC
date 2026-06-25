"""
Akış Şeması veri motoru (v2).

Tablo 7'deki (IPK) operasyonlara göre gruplar. Her operasyon için:
  - operasyon adı (kutu etiketi)
  - IPK testleri (Tablo 7'den, o operasyon)
  - kimyasal testler (Tablo 6'da olup Tablo 7'de olmayan, o operasyon)

Tartım dahil edilmez. Operasyon sırası: Karıştırma, Tablet Baskı, Film Kaplama, Blisterleme.
"""

from __future__ import annotations

from core.kural_motoru import turet

# Akış şemasında gösterilecek operasyon sırası ve görünen ad
_OP_SIRA = [
    ("Karıştırma", "Karıştırma"),
    ("Tablet Baskı", "Tablet Baskı"),
    ("Film Kaplama", "Film Kaplama"),
    ("Blisterleme", "Blisterleme"),
]


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
    return kutular
