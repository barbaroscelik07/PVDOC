"""
Üretim Yöntemi Word okuyucu.

Kullanıcının yüklediği Word dosyasından "Operasyon X: Aşama Y" başlık +
açıklama çiftlerini ayıklar. "Operasyon 1: Aşama 1"den başlar, operasyon/aşama
deseni bitene (ana başlık veya desen-dışı içerik) kadar okur.
"""

from __future__ import annotations

import re

from docx import Document

# "Operasyon 2: Aşama 13", "Operasyon 1 : Aşama 1" gibi varyasyonları yakalar
_OP_DESEN = re.compile(r"^\s*operasyon\s*\d+\s*[:：]\s*aşama\s*\d+\s*$", re.IGNORECASE)

# Üretim yönteminin bittiğini gösteren başlıklar
_BITIS = ("proses akış", "proses akis", "kapsanan ürünler", "proses parametre",
          "risk analiz", "akış diyagram", "akis diyagram")


def _norm(s: str) -> str:
    tr = str.maketrans({"ı": "i", "İ": "i", "ş": "s", "Ş": "s", "ğ": "g",
                        "ç": "c", "ö": "o", "ü": "u", "I": "i"})
    return (s or "").translate(tr).lower().strip()


def uretim_yontemi_coz(yol: str) -> dict:
    """
    Word'den üretim yöntemi adımlarını ayıklar.
    Dönüş: {"bulundu": bool, "adimlar": [(baslik, aciklama), ...]}
    """
    d = Document(yol)
    paragraflar = [p.text.strip() for p in d.paragraphs]

    adimlar = []
    i = 0
    n = len(paragraflar)
    basladi = False

    while i < n:
        metin = paragraflar[i]
        if not metin:
            i += 1
            continue

        # Operasyon başlığı mı?
        if _OP_DESEN.match(metin):
            basladi = True
            baslik = re.sub(r"\s+", " ", metin).strip()
            # açıklama = sonraki dolu paragraf(lar), bir sonraki operasyon başlığına kadar
            aciklama_parcalari = []
            j = i + 1
            while j < n:
                sonraki = paragraflar[j]
                if not sonraki:
                    j += 1
                    continue
                if _OP_DESEN.match(sonraki):
                    break
                # bitiş başlığı mı?
                if any(b in _norm(sonraki) for b in _BITIS):
                    break
                aciklama_parcalari.append(sonraki)
                j += 1
            adimlar.append((baslik, " ".join(aciklama_parcalari).strip()))
            i = j
            continue

        # Başladıysak ve bitiş başlığı geldiyse dur
        if basladi and any(b in _norm(metin) for b in _BITIS):
            break

        i += 1

    return {"bulundu": bool(adimlar), "adimlar": adimlar}
