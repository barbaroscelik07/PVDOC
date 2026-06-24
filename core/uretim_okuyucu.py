"""
Üretim Yöntemi Word okuyucu.

"Operasyon X: Aşama Y" başlık + açıklama + (varsa) parametre tablosu üçlülerini
ayıklar. "Operasyon 1: Aşama 1"den başlar, "ÜRETİM AKIŞ ŞEMASI" görününce durur.

Dönüş adımları: [(baslik, aciklama, tablo_satirlari), ...]
  tablo_satirlari: [(sol, sag), ...]  (parametre tablosu; yoksa boş liste)
"""

from __future__ import annotations

import re

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

_OP_DESEN = re.compile(r"^\s*operasyon\s*\d+\s*[:：]\s*aşama\s*\d+\s*$", re.IGNORECASE)

# Üretim yönteminin bittiği yer
_BITIS = ("üretim akış", "uretim akis", "akış şeması", "akis semasi",
          "proses akış", "proses akis", "in-proses", "i̇n-proses",
          "in proses", "kapsanan ürünler")

# Açıklamadan atılacak ara başlık desenleri ("2. Karışımın Hazırlanması", "3. Tablet baskı")
_ARA_BASLIK = re.compile(r"\s*\d+\.\s*(karışımın hazırlanması|tablet baskı|"
                         r"film kaplama|blisterleme|tartım|ambalajlama|kutulama)"
                         r"[ ,]*", re.IGNORECASE)


def _norm(s: str) -> str:
    tr = str.maketrans({"ı": "i", "İ": "i", "ş": "s", "Ş": "s", "ğ": "g",
                        "ç": "c", "ö": "o", "ü": "u", "I": "i"})
    return (s or "").translate(tr).lower().strip()


def _aciklama_temizle(metin: str) -> str:
    """Açıklamadan bir sonraki bölüm ara başlıklarını atar."""
    metin = _ARA_BASLIK.sub(" ", metin)
    return re.sub(r"\s+", " ", metin).strip()


def uretim_yontemi_coz(yol: str) -> dict:
    d = Document(yol)
    body = list(d.element.body)

    # body'yi sırayla (paragraf / tablo) gez
    ogeler = []  # ("p", metin) veya ("t", [(sol,sag),...])
    for el in body:
        if el.tag.endswith("}p"):
            ogeler.append(("p", Paragraph(el, d).text.strip()))
        elif el.tag.endswith("}tbl"):
            tbl = Table(el, d)
            satirlar = []
            for r in tbl.rows:
                if len(r.cells) >= 2:
                    satirlar.append((r.cells[0].text.strip(), r.cells[1].text.strip()))
            ogeler.append(("t", satirlar))

    adimlar = []
    i = 0
    n = len(ogeler)
    basladi = False

    while i < n:
        tip, icerik = ogeler[i]
        if tip == "p" and _OP_DESEN.match(icerik or ""):
            basladi = True
            baslik = re.sub(r"\s+", " ", icerik).strip()
            aciklama_parcalari = []
            tablo_satirlari = []
            j = i + 1
            while j < n:
                t2, c2 = ogeler[j]
                if t2 == "p":
                    if not c2:
                        j += 1
                        continue
                    if _OP_DESEN.match(c2):
                        break
                    if any(b in _norm(c2) for b in _BITIS):
                        break
                    aciklama_parcalari.append(c2)
                    j += 1
                else:  # tablo → bu aşamaya ait parametre tablosu
                    tablo_satirlari.extend(c2)
                    j += 1
            aciklama = _aciklama_temizle(" ".join(aciklama_parcalari))
            adimlar.append((baslik, aciklama, tablo_satirlari))
            i = j
            continue

        if basladi and tip == "p" and any(b in _norm(icerik or "") for b in _BITIS):
            break
        i += 1

    return {"bulundu": bool(adimlar), "adimlar": adimlar}
