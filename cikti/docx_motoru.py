"""
DOCX üretim motoru — ŞABLON DOLDURMA yaklaşımı.

PVP/PVR şablonu açılır, placeholder'lar kullanıcı verisiyle değiştirilir ve
dinamik tablolar (formül, kapsanan ürünler, risk, proses param, ekipman,
spesifikasyon, numune planı, PVR sonuçları) şablonun kendi satır biçimi
korunarak doldurulur.

Bu yaklaşım şablonun fontunu, satır boşluklarını, kenarlıklarını, kapak
sayfasını, içindekiler bölümünü ve gömülü resimleri (IBC numune resmi vb.)
BİREBİR korur — sıfırdan üretimde kaybedilen "taslakla aynı görünüm" budur.

Tablolar indeks yerine ilk-hücre metnine göre bulunur (şablon değişse de sağlam).
"""

from __future__ import annotations

import sys
import os
from pathlib import Path

from docx import Document

from core.models import (
    ProjeVerisi, UrunFormu, Test, TabloTipi, LimitTuru,
    SERI_SAYISI, NOKTA_ADLARI,
)
from core import veri_uretici as vu
from cikti.sablon_doldur import (
    belgede_degistir, satir_klonla, hucre_yaz, satiri_bosalt,
)


# ----------------------------------------------------------------------------
# Şablon yolu (PyInstaller uyumlu)
# ----------------------------------------------------------------------------

def _sablon_yolu(ad: str) -> Path:
    taban = getattr(sys, "_MEIPASS", os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    return Path(taban) / "cikti" / "sablonlar" / ad


# ----------------------------------------------------------------------------
# Yardımcılar
# ----------------------------------------------------------------------------

def _urun(proje: ProjeVerisi) -> str:
    return proje.dokuman.urun_adi or "Xxx Film Kaplı Tablet"


def _tablo_bul(doc, ilk_hucre_metni: str, sutun_sayisi: int | None = None):
    """İlk hücresi verilen metinle başlayan ilk tabloyu döndürür."""
    hedef = ilk_hucre_metni.strip().lower()
    for t in doc.tables:
        if not t.rows:
            continue
        ilk = t.rows[0].cells[0].text.strip().lower()
        if ilk.startswith(hedef):
            if sutun_sayisi is None or len(t.columns) == sutun_sayisi:
                return t
    return None


def _tablo_basliga_gore(doc, tablo_no: int):
    """
    'Tablo N ...' başlık paragrafından HEMEN SONRA gelen tabloyu döndürür.
    Benzer yapılı tabloları (Tablo 6 vs 8.2) kesin ayırt eder — en sağlam yöntem.
    """
    from docx.table import Table as _T
    from docx.text.paragraph import Paragraph as _P
    from docx.oxml.ns import qn as _qn

    onceki_baslik_no = None
    for child in doc.element.body.iterchildren():
        if child.tag == _qn("w:p"):
            txt = _P(child, doc).text.strip()
            # "Tablo 6 ..." veya "Tablo.6 ..." veya "Tablo 6\t..."
            low = txt.lower()
            if low.startswith("tablo"):
                # ilk sayıyı çek
                import re
                m = re.search(r"tablo[\s\.]*(\d+)", low)
                onceki_baslik_no = int(m.group(1)) if m else None
        elif child.tag == _qn("w:tbl"):
            if onceki_baslik_no == tablo_no:
                return _T(child, doc)
            onceki_baslik_no = None  # tablo geçildi, başlık tüketildi
    return None


def _seri_nolar(proje: ProjeVerisi) -> list[str]:
    out = []
    for i in range(SERI_SAYISI):
        sno = proje.seriler[i].seri_no if i < len(proje.seriler) else ""
        out.append(sno or f"YYY-P{i+1:02d}")
    return out


def _veri_satirlarini_ayarla(tablo, baslik_satir_sayisi: int, gereken_veri_satiri: int):
    """
    Tablodaki veri satırı sayısını 'gereken'e eşitler:
    - fazla satırları siler,
    - eksikse son veri satırını klonlar.
    Başlık satırları (ilk N) korunur. Döndürür: veri satırlarının indeks listesi.
    """
    mevcut_veri = len(tablo.rows) - baslik_satir_sayisi
    # fazlaları sil (sondan)
    while mevcut_veri > gereken_veri_satiri and mevcut_veri > 0:
        tr = tablo.rows[-1]._tr
        tr.getparent().remove(tr)
        mevcut_veri -= 1
    # eksikleri klonla
    while mevcut_veri < gereken_veri_satiri:
        yeni = satir_klonla(tablo, kaynak_index=-1)
        satiri_bosalt(yeni)
        mevcut_veri += 1
    return list(range(baslik_satir_sayisi, baslik_satir_sayisi + gereken_veri_satiri))


# ============================================================================
# PVP tablolarını doldur
# ============================================================================

def _doldur_formul(doc, proje: ProjeVerisi) -> None:
    t = _tablo_basliga_gore(doc, 1)  # Tablo 1
    if t is None or not proje.hammaddeler:
        return
    # {adet} placeholder'ı seri boyutundaki adet değeriyle değiştir (başlık hücresi)
    adet = proje.seriler[0].seri_boyutu_adet if proje.seriler else ""
    if adet:
        for row in t.rows:
            for cell in row.cells:
                if "{adet}" in cell.text:
                    from cikti.sablon_doldur import paragraf_metni_degistir
                    for p in cell.paragraphs:
                        paragraf_metni_degistir(p, "{adet}", adet)

    son = len(proje.hammaddeler) - 1
    idxler = _veri_satirlarini_ayarla(t, 1, len(proje.hammaddeler))
    for sira, (ri, h) in enumerate(zip(idxler, proje.hammaddeler)):
        cells = t.rows[ri].cells
        # Son satır VEYA ara_toplam satırı kalın
        bold = h.ara_toplam or (sira == son)
        hucre_yaz(cells[0], h.ad, bold=bold)
        hucre_yaz(cells[1], h.fonksiyon, bold=bold)
        # Birim formül HER ZAMAN 3 ondalık (3 -> 3.000)
        hucre_yaz(cells[2], "" if h.birim_formul is None else f"{h.birim_formul:.3f}", bold=bold)
        hucre_yaz(cells[3], "" if h.yuzde_icerik is None else f"{h.yuzde_icerik:g}", bold=bold)
        hucre_yaz(cells[4], "" if h.seri_miktar is None else f"{h.seri_miktar:.3f}", bold=bold)


def _doldur_kapsanan(doc, proje: ProjeVerisi) -> None:
    t = _tablo_basliga_gore(doc, 2)  # Tablo 2
    if t is None:
        return
    # 2 başlık satırı (Ürün İsmi / Film Kaplı Tablet-kg alt başlığı)
    baslik = 2
    idxler = _veri_satirlarini_ayarla(t, baslik, SERI_SAYISI)
    for k, ri in enumerate(idxler):
        s = proje.seriler[k]
        cells = t.rows[ri].cells
        hucre_yaz(cells[0], s.urun_ismi or _urun(proje))
        hucre_yaz(cells[1], s.seri_no)
        hucre_yaz(cells[2], s.seri_boyutu_adet)
        hucre_yaz(cells[3], s.seri_boyutu_kg)


def _doldur_risk(doc, proje: ProjeVerisi) -> None:
    t = _tablo_basliga_gore(doc, 3)  # Tablo 3
    if t is None or not proje.risk_satirlari:
        return
    idxler = _veri_satirlarini_ayarla(t, 1, len(proje.risk_satirlari))
    for ri, rs in zip(idxler, proje.risk_satirlari):
        cells = t.rows[ri].cells
        hucre_yaz(cells[0], str(rs.operasyon_no or ""))
        hucre_yaz(cells[1], rs.operasyon)
        hucre_yaz(cells[2], "E" if rs.kritik else "H")
        hucre_yaz(cells[3], rs.testler)
        hucre_yaz(cells[4], rs.yorumlar)


def _doldur_proses_param(doc, proje: ProjeVerisi) -> None:
    t = _tablo_basliga_gore(doc, 4)  # Tablo 4
    if t is None or not proje.proses_parametreleri:
        return
    idxler = _veri_satirlarini_ayarla(t, 1, len(proje.proses_parametreleri))
    for ri, pp in zip(idxler, proje.proses_parametreleri):
        cells = t.rows[ri].cells
        hucre_yaz(cells[0], pp.aciklama)
        hucre_yaz(cells[1], pp.parametre)
        if len(cells) > 2:
            hucre_yaz(cells[2], pp.deger)


def _doldur_ekipman(doc, proje: ProjeVerisi) -> None:
    t = _tablo_basliga_gore(doc, 5)  # Tablo 5
    if t is None or not proje.ekipmanlar:
        return
    idxler = _veri_satirlarini_ayarla(t, 1, len(proje.ekipmanlar))
    for ri, e in zip(idxler, proje.ekipmanlar):
        cells = t.rows[ri].cells
        hucre_yaz(cells[0], str(e.operasyon_no or ""))
        hucre_yaz(cells[1], e.operasyon)
        hucre_yaz(cells[2], e.ekipman_adi)
        hucre_yaz(cells[3], e.kapasite)


def _doldur_numune(doc, proje: ProjeVerisi) -> None:
    t = _tablo_basliga_gore(doc, 10)  # Tablo 10
    if t is None or not proje.numune_plani:
        return
    idxler = _veri_satirlarini_ayarla(t, 1, len(proje.numune_plani))
    for ri, n in zip(idxler, proje.numune_plani):
        cells = t.rows[ri].cells
        hucre_yaz(cells[0], str(n.operasyon_no or ""))
        hucre_yaz(cells[1], n.operasyon)
        hucre_yaz(cells[2], n.numune_noktasi)
        hucre_yaz(cells[3], n.toplam_miktar)


def _doldur_spek(doc, proje: ProjeVerisi) -> None:
    """Tablo 6 — spesifikasyon. Yıldız test adının SONUNA; alt satırlar ayrı satır."""
    t = _tablo_basliga_gore(doc, 6)  # Tablo 6
    kart = proje.spek_karti
    if t is None or not kart.testler:
        return

    # Her testin kaç satır kaplayacağını hesapla (ana + alt satırlar + açıklama)
    satir_planı = []  # (test, alt_satir_listesi)  -> toplam satır
    toplam = 0
    for test in kart.testler:
        ekstra = list(test.alt_satirlar)
        if test.aciklama_etiketi:
            ekstra = ekstra + [(test.aciklama_etiketi, test.aciklama_spek)]
        satir_planı.append((test, ekstra))
        toplam += 1 + len(ekstra)

    idxler = _veri_satirlarini_ayarla(t, 1, toplam)
    it = iter(idxler)
    for test, ekstra in satir_planı:
        # ana satır
        ri = next(it)
        cells = t.rows[ri].cells
        ad = test.ad + ("*" if test.yildizli else "")
        hucre_yaz(cells[0], str(test.operasyon_no or ""))
        hucre_yaz(cells[1], test.operasyon)
        hucre_yaz(cells[2], ad)
        # mikrobiyolojik/ağırlık ana satırında spek hücresi boş (alt satırlarda dolu)
        ana_spek = "" if ekstra else test.spesifikasyon.metni_olustur()
        hucre_yaz(cells[3], ana_spek)
        # alt satırlar
        for etiket, spek_metni in ekstra:
            ri2 = next(it)
            c2 = t.rows[ri2].cells
            hucre_yaz(c2[0], "")
            hucre_yaz(c2[1], "")
            hucre_yaz(c2[2], etiket)
            hucre_yaz(c2[3], spek_metni)


def _doldur_ipk(doc, proje: ProjeVerisi) -> None:
    """Tablo 7 — sadece IPK testleri (2 sütun: TESTLER | SPESİFİKASYONLAR)."""
    ipk = [t for t in proje.spek_karti.testler if t.ipk]
    t = _tablo_basliga_gore(doc, 7)  # Tablo 7
    if t is None or not ipk:
        return
    # Tablo 7 şablonda 2 sütunlu (TESTLER | SPESİFİKASYONLAR)
    sut = len(t.columns)
    idxler = _veri_satirlarini_ayarla(t, 1, len(ipk))
    for ri, test in zip(idxler, ipk):
        cells = t.rows[ri].cells
        if sut == 2:
            hucre_yaz(cells[0], test.ad)
            hucre_yaz(cells[1], test.spesifikasyon.metni_olustur())
        else:
            # 4 sütunlu varyant: Op No | Operasyon | Test | Spesifikasyon
            hucre_yaz(cells[0], str(test.operasyon_no or ""))
            hucre_yaz(cells[1], test.operasyon)
            hucre_yaz(cells[2], test.ad)
            hucre_yaz(cells[3], test.spesifikasyon.metni_olustur())


# ============================================================================
# PVR sonuç tabloları (Bölüm 11) — şablonda örnek tablolar var; biz ekleyeceğiz
# ============================================================================
# Not: PVR sonuç tabloları çok sayıda ve teste göre değişken. Şablondaki örnek
# sonuç tablolarını silip, kullanıcı testlerine göre python-docx ile yeniden
# ekliyoruz (sonuç bölümü en sonda olduğu için bu güvenli).

from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH


def _sonuc_basligi(doc, no, ad):
    p = doc.add_paragraph()
    r = p.add_run(f"Tablo.{no} {ad} Sonuçları")
    r.bold = True
    r.font.size = Pt(10)
    return p


def _yeni_tablo(doc, satir, sutun):
    t = doc.add_table(rows=satir, cols=sutun)
    t.style = "Table Grid"
    return t


def _sr(cells, degerler, bold=False):
    for c, v in zip(cells, degerler):
        hucre_yaz(c, v, bold=bold) if c.paragraphs[0].runs else _yaz_bos(c, v, bold)


def _yaz_bos(cell, metin, bold=False):
    p = cell.paragraphs[0]
    r = p.add_run(str(metin))
    r.bold = bold
    r.font.size = Pt(9)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER


def _yaz_sol(cell, metin, bold=False):
    """Sola dayalı hücre yazımı (Test/Spesifikasyon değer hücreleri için)."""
    p = cell.paragraphs[0]
    r = p.add_run(str(metin))
    r.bold = bold
    r.font.size = Pt(9)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT


def _numuneler_basligi(t, ust_satir, alt_satir, proje):
    """
    Resimlerdeki ortak yapı:
      | Numuneler | Analiz Sonuçları (3 sütuna birleşik)        |
      |           | Seri No: YYY-P01 | YYY-P02 | YYY-P03         |
    ust_satir: 'Numuneler' + 'Analiz Sonuçları' satırının indeksi
    alt_satir: Seri No satırının indeksi
    """
    _yaz_bos(t.rows[ust_satir].cells[0], "Numuneler", True)
    t.rows[ust_satir].cells[1].merge(t.rows[ust_satir].cells[SERI_SAYISI])
    _yaz_bos(t.rows[ust_satir].cells[1], "Analiz Sonuçları", True)
    # 'Numuneler' hücresini iki satıra dikey birleştir
    t.rows[ust_satir].cells[0].merge(t.rows[alt_satir].cells[0])
    for c, sno in enumerate(_seri_nolar(proje), 1):
        _yaz_bos(t.rows[alt_satir].cells[c], f"Seri No: {sno}", True)


def _ekle_sonuc_tek(doc, proje, test, no):
    """Görünüş/Teşhis/Elek: Test | Spesifikasyon | Numuneler+Analiz | Sonuç."""
    seriler = test.sonuc_verisi.get("seriler", ["", "", ""])
    _sonuc_basligi(doc, no, test.ad)
    t = _yeni_tablo(doc, 5, SERI_SAYISI + 1)
    _yaz_bos(t.rows[0].cells[0], "Test", True)
    t.rows[0].cells[1].merge(t.rows[0].cells[SERI_SAYISI])
    _yaz_sol(t.rows[0].cells[1], test.ad)
    _yaz_bos(t.rows[1].cells[0], "Spesifikasyon", True)
    t.rows[1].cells[1].merge(t.rows[1].cells[SERI_SAYISI])
    _yaz_sol(t.rows[1].cells[1], test.spesifikasyon.metni_olustur())
    _numuneler_basligi(t, 2, 3, proje)
    _yaz_bos(t.rows[4].cells[0], "Sonuç", True)
    for c in range(SERI_SAYISI):
        _yaz_bos(t.rows[4].cells[c+1], seriler[c] if c < len(seriler) else "", True)


def _ekle_sonuc_iki(doc, proje, test, no):
    """Miktar Tayini / İmpurite: Numune-1/Numune-2/Sonuç."""
    seriler = test.sonuc_verisi.get("seriler", [])
    _sonuc_basligi(doc, no, test.ad)
    t = _yeni_tablo(doc, 7, SERI_SAYISI + 1)
    _yaz_bos(t.rows[0].cells[0], "Test", True)
    t.rows[0].cells[1].merge(t.rows[0].cells[SERI_SAYISI])
    _yaz_sol(t.rows[0].cells[1], test.ad)
    _yaz_bos(t.rows[1].cells[0], "Spesifikasyon", True)
    t.rows[1].cells[1].merge(t.rows[1].cells[SERI_SAYISI])
    _yaz_sol(t.rows[1].cells[1], test.spesifikasyon.metni_olustur())
    _numuneler_basligi(t, 2, 3, proje)
    for ri, key, et in [(4, "numune_1", "Numune-1"), (5, "numune_2", "Numune-2"), (6, "sonuc", "Sonuç")]:
        _yaz_sol(t.rows[ri].cells[0], et, ri == 6)
        for c in range(SERI_SAYISI):
            v = seriler[c].get(key, "") if c < len(seriler) else ""
            _yaz_bos(t.rows[ri].cells[c+1], v, ri == 6)
    # İmpurite ise T.E. notu
    if "impurite" in test.ad.lower() or "ilgili bileşik" in test.ad.lower() or "imp." in test.ad.lower():
        np = doc.add_paragraph()
        nr = np.add_run("T.E.: Tespit edilemedi.")
        nr.italic = True; nr.font.size = Pt(8)


def _ekle_sonuc_on(doc, proje, test, no):
    """Karışım Tekdüzeliği: 1-10 + Ortalama."""
    seriler = test.sonuc_verisi.get("seriler", [])
    _sonuc_basligi(doc, no, test.ad)
    t = _yeni_tablo(doc, 15, SERI_SAYISI + 1)
    _yaz_bos(t.rows[0].cells[0], "Test", True)
    t.rows[0].cells[1].merge(t.rows[0].cells[SERI_SAYISI])
    _yaz_sol(t.rows[0].cells[1], test.ad)
    _yaz_bos(t.rows[1].cells[0], "Spesifikasyon", True)
    t.rows[1].cells[1].merge(t.rows[1].cells[SERI_SAYISI])
    _yaz_sol(t.rows[1].cells[1], test.spesifikasyon.metni_olustur())
    _numuneler_basligi(t, 2, 3, proje)
    for n in range(10):
        _yaz_bos(t.rows[4+n].cells[0], str(n+1))
        for c in range(SERI_SAYISI):
            olc = seriler[c].get("olcumler", []) if c < len(seriler) else []
            _yaz_bos(t.rows[4+n].cells[c+1], olc[n] if n < len(olc) else "")
    _yaz_bos(t.rows[14].cells[0], "Ortalama", True)
    for c in range(SERI_SAYISI):
        _yaz_bos(t.rows[14].cells[c+1], seriler[c].get("ortalama", "") if c < len(seriler) else "", True)


def _ekle_sonuc_bos(doc, proje, test, no, ns=10):
    """Sertlik/Kalınlık/Çap/Dağılma/Dissolüsyon: Baş/Orta/Son × seri + Sonuç."""
    seriler = test.sonuc_verisi.get("seriler", [])
    _sonuc_basligi(doc, no, test.ad)
    # Test + Spesifikasyon + Seri başlık + nokta başlık + ns ölçüm + Ortalama + Sonuç
    t = _yeni_tablo(doc, 4 + ns + 2, SERI_SAYISI * 3 + 1)
    _yaz_bos(t.rows[0].cells[0], "Test", True)
    t.rows[0].cells[1].merge(t.rows[0].cells[SERI_SAYISI * 3])
    _yaz_sol(t.rows[0].cells[1], test.ad)
    _yaz_bos(t.rows[1].cells[0], "Spesifikasyon", True)
    t.rows[1].cells[1].merge(t.rows[1].cells[SERI_SAYISI * 3])
    _yaz_sol(t.rows[1].cells[1], test.spesifikasyon.metni_olustur())
    # seri başlıkları
    _yaz_bos(t.rows[2].cells[0], "Numuneler", True)
    col = 1
    for sno in _seri_nolar(proje):
        a = t.rows[2].cells[col]; a.merge(t.rows[2].cells[col+2])
        _yaz_bos(a, f"Seri No: {sno}", True); col += 3
    # 'Numuneler' dikey birleştir
    t.rows[2].cells[0].merge(t.rows[3].cells[0])
    col = 1
    for _ in range(SERI_SAYISI):
        for nokta in NOKTA_ADLARI:
            _yaz_bos(t.rows[3].cells[col], nokta, True); col += 1
    for n in range(ns):
        _yaz_bos(t.rows[4+n].cells[0], str(n+1))
        col = 1
        for c in range(SERI_SAYISI):
            for nokta in NOKTA_ADLARI:
                noktalar = seriler[c].get("noktalar", {}) if c < len(seriler) else {}
                olc = noktalar.get(nokta, {}).get("olcumler", [])
                _yaz_bos(t.rows[4+n].cells[col], olc[n] if n < len(olc) else ""); col += 1
    _yaz_bos(t.rows[4+ns].cells[0], "Ortalama", True)
    col = 1
    for c in range(SERI_SAYISI):
        for nokta in NOKTA_ADLARI:
            noktalar = seriler[c].get("noktalar", {}) if c < len(seriler) else {}
            _yaz_bos(t.rows[4+ns].cells[col], noktalar.get(nokta, {}).get("ortalama", ""), True); col += 1
    _yaz_bos(t.rows[5+ns].cells[0], "Sonuç", True)
    col = 1
    for c in range(SERI_SAYISI):
        sonuc = seriler[c].get("sonuc", "") if c < len(seriler) else ""
        a = t.rows[5+ns].cells[col]; a.merge(t.rows[5+ns].cells[col+2])
        _yaz_bos(a, sonuc, True); col += 3


def _ekle_sonuc_agirlik(doc, proje, test, no):
    """Ağırlık Tekdüzeliği: 20 numune × Baş/Orta/Son × seri + Ort/RSD/SD."""
    seriler = test.sonuc_verisi.get("seriler", [])
    _sonuc_basligi(doc, no, test.ad)
    t = _yeni_tablo(doc, 4 + 20 + 3, SERI_SAYISI * 3 + 1)
    _yaz_bos(t.rows[0].cells[0], "Test", True)
    t.rows[0].cells[1].merge(t.rows[0].cells[SERI_SAYISI * 3])
    _yaz_sol(t.rows[0].cells[1], test.ad)
    _yaz_bos(t.rows[1].cells[0], "Spesifikasyon", True)
    t.rows[1].cells[1].merge(t.rows[1].cells[SERI_SAYISI * 3])
    _yaz_sol(t.rows[1].cells[1], test.spesifikasyon.metni_olustur())
    _yaz_bos(t.rows[2].cells[0], "Numuneler", True)
    col = 1
    for sno in _seri_nolar(proje):
        a = t.rows[2].cells[col]; a.merge(t.rows[2].cells[col+2])
        _yaz_bos(a, f"Seri No: {sno}", True); col += 3
    t.rows[2].cells[0].merge(t.rows[3].cells[0])
    col = 1
    for _ in range(SERI_SAYISI):
        for nokta in NOKTA_ADLARI:
            _yaz_bos(t.rows[3].cells[col], nokta, True); col += 1
    for n in range(20):
        _yaz_bos(t.rows[4+n].cells[0], str(n+1))
        col = 1
        for c in range(SERI_SAYISI):
            for nokta in NOKTA_ADLARI:
                noktalar = seriler[c].get("noktalar", {}) if c < len(seriler) else {}
                olc = noktalar.get(nokta, {}).get("olcumler", [])
                _yaz_bos(t.rows[4+n].cells[col], olc[n] if n < len(olc) else ""); col += 1
    for k, (et, key) in enumerate([("Ortalama", "ortalama"), ("RSD%", "rsd"), ("SD", "sd")]):
        _yaz_bos(t.rows[24+k].cells[0], et, True)
        col = 1
        for c in range(SERI_SAYISI):
            for nokta in NOKTA_ADLARI:
                noktalar = seriler[c].get("noktalar", {}) if c < len(seriler) else {}
                _yaz_bos(t.rows[24+k].cells[col], noktalar.get(nokta, {}).get(key, ""), True); col += 1


def _ekle_sonuc_matris(doc, proje, test, no):
    """
    Mikrobiyolojik Kontrol (resimdeki Tablo 33/34):
      Test | Mikrobiyolojik Kontrol
      Spesifikasyon | (3 alt satır spek, birleşik)
      Numuneler/Analiz Sonuçları | Seri No
      -Toplam Aerobik... | <10 cfu/g × 3
      -Küf ve Maya...    | <10 cfu/g × 3
      -E. coli           | 0 cfu/g × 3
    Sonuç satırı YOK (kullanıcı kararı).
    """
    _sonuc_basligi(doc, no, test.ad)
    alt = test.alt_satirlar or [("-Toplam Aerobik Mikroorganizma Sayısı", "≤10³ cfu/g"),
                                ("-Küf ve Maya Sayısı", "≤10² cfu/g"), ("-E. coli", "0 cfu/g")]
    # Test(1) + Spesifikasyon(len alt, birleşik) + Numuneler/SeriNo(2) + ölçüm(len alt)
    n_spec = len(alt)
    t = _yeni_tablo(doc, 1 + n_spec + 2 + len(alt), SERI_SAYISI + 1)
    r = 0
    _yaz_bos(t.rows[r].cells[0], "Test", True)
    t.rows[r].cells[1].merge(t.rows[r].cells[SERI_SAYISI])
    _yaz_sol(t.rows[r].cells[1], test.ad); r += 1
    # Spesifikasyon: sol hücre dikey birleşik, sağda alt satır spek'leri
    spec_bas = r
    for i, (etiket, spek_metni) in enumerate(alt):
        t.rows[r].cells[1].merge(t.rows[r].cells[SERI_SAYISI])
        _yaz_sol(t.rows[r].cells[1], f"{etiket} : {spek_metni}")
        r += 1
    _yaz_bos(t.rows[spec_bas].cells[0], "Spesifikasyon", True)
    t.rows[spec_bas].cells[0].merge(t.rows[r-1].cells[0])
    # Numuneler / Analiz Sonuçları + Seri No
    _numuneler_basligi(t, r, r+1, proje); r += 2
    # ölçüm satırları
    for etiket, spek_metni in alt:
        _yaz_sol(t.rows[r].cells[0], etiket)
        for c in range(SERI_SAYISI):
            # değer: aerobik/küf <10 cfu/g, e.coli 0 cfu/g
            deger = "0 cfu/g" if "coli" in etiket.lower() else "<10 cfu/g"
            _yaz_bos(t.rows[r].cells[c+1], deger)
        r += 1


def _doldur_sonuclar(doc, proje: ProjeVerisi) -> None:
    """PVR Bölüm 11 sonuç tablolarını belgenin sonuna ekler."""
    doc.add_page_break()
    h = doc.add_paragraph()
    r = h.add_run("11. SONUÇLAR — PROSES VALİDASYONU TEST SONUÇLARI")
    r.bold = True; r.font.size = Pt(12)

    no = 11
    for test in proje.spek_karti.testler:
        tip = test.tablo_tipi
        if test.mikrobiyolojik or tip is TabloTipi.MATRIS:
            _ekle_sonuc_matris(doc, proje, test, no)
        elif tip is TabloTipi.TEK_SONUC:
            _ekle_sonuc_tek(doc, proje, test, no)
        elif tip is TabloTipi.IKI_NUMUNE:
            _ekle_sonuc_iki(doc, proje, test, no)
        elif tip is TabloTipi.ON_NUMUNE:
            _ekle_sonuc_on(doc, proje, test, no)
        elif tip is TabloTipi.BOS_NOKTA:
            _ekle_sonuc_bos(doc, proje, test, no)
        elif tip is TabloTipi.AGIRLIK_TEKDUZELIGI:
            _ekle_sonuc_agirlik(doc, proje, test, no)
        else:
            _ekle_sonuc_tek(doc, proje, test, no)
        doc.add_paragraph("")
        no += 1

    # Genel Değerlendirme HER ZAMAN en sonda
    doc.add_paragraph("")
    gh = doc.add_paragraph()
    gr = gh.add_run("12. GENEL DEĞERLENDİRME")
    gr.bold = True; gr.font.size = Pt(12)
    doc.add_paragraph(f"Sapmalar: {proje.sapmalar}")
    doc.add_paragraph(f"Sonuç: {proje.sonuc_degerlendirme}")
    if proje.yorum:
        doc.add_paragraph(f"Yorum: {proje.yorum}")


# ============================================================================
# Ana üretim
# ============================================================================

def _placeholder_eslemeleri(proje: ProjeVerisi, rapor: bool) -> dict[str, str]:
    d = proje.dokuman
    urun = _urun(proje)
    dok_no = (d.pvr_dokuman_no if rapor else d.pvp_dokuman_no) or "AG-PV-xxx"
    es = {
        "XxxFilm Kaplı Tablet": urun,
        "XxxFİLM TABLET": urun.upper(),
        "AG-PV-xxx": dok_no,
    }
    for i in range(SERI_SAYISI):
        sno = proje.seriler[i].seri_no
        if sno:
            es[f"yyy-P0{i+1}"] = sno
    if d.firma_ismi:
        es["{Firma ismi}"] = d.firma_ismi
    return es


def _tolerans_oran(tolerans: str):
    """'±%5' / '%7.5' / '±%10' gibi metinden oranı (0.05) çıkarır."""
    import re
    m = re.search(r"(\d+(?:[.,]\d+)?)", tolerans or "")
    if not m:
        return None
    return float(m.group(1).replace(",", ".")) / 100.0


def _miktar_spek_uret(hedef: float, tolerans: str, birim: str) -> str:
    """Hedef + tolerans → '10.0 mg/f.tab ±%7.5 (9.25 – 10.75 mg/f.tab)'."""
    oran = _tolerans_oran(tolerans)
    if oran is None:
        return f"{hedef:g} {birim}".strip()
    alt = round(hedef * (1 - oran), 4)
    ust = round(hedef * (1 + oran), 4)
    return f"{hedef:g} {birim} {tolerans} ({alt:g} – {ust:g} {birim})".strip()


def _doldur_tablo89(doc, proje: ProjeVerisi) -> None:
    """
    Tablo 8 (Serbest Bırakma) ve Tablo 9 (Raf Ömrü): Tablo 6'daki bitmiş ürün
    testlerini temel alır; Miktar Tayini toleransı her tabloda farklıdır.
    """
    kart = proje.spek_karti
    if not kart.tablo89_ekle:
        return
    bitmis = [t for t in kart.testler
              if t.operasyon in ("Tablet Baskı", "Film Kaplama", "Dolum")
              and not t.mikrobiyolojik]
    for tablo_no, tol in [(8, kart.serbest_birakma_tolerans), (9, kart.raf_omru_tolerans)]:
        t = _tablo_basliga_gore(doc, tablo_no)
        if t is None or not bitmis:
            continue
        sut = len(t.columns)
        idxler = _veri_satirlarini_ayarla(t, 1, len(bitmis))
        for ri, test in zip(idxler, bitmis):
            cells = t.rows[ri].cells
            spek_metni = test.spesifikasyon.metni_olustur()
            if "miktar tayini" in test.ad.lower() and test.spesifikasyon.hedef_deger:
                spek_metni = _miktar_spek_uret(
                    test.spesifikasyon.hedef_deger, tol, test.spesifikasyon.birim)
            if sut == 2:
                hucre_yaz(cells[0], test.ad)
                hucre_yaz(cells[1], spek_metni)
            else:
                hucre_yaz(cells[0], test.ad)
                hucre_yaz(cells[-1], spek_metni)


def _ortak_doldur(doc, proje: ProjeVerisi, rapor: bool) -> None:
    belgede_degistir(doc, _placeholder_eslemeleri(proje, rapor))
    _doldur_formul(doc, proje)
    _doldur_kapsanan(doc, proje)
    _doldur_risk(doc, proje)
    _doldur_proses_param(doc, proje)
    _doldur_ekipman(doc, proje)
    _doldur_spek(doc, proje)
    _doldur_ipk(doc, proje)
    _doldur_tablo89(doc, proje)
    _doldur_numune(doc, proje)


def pvp_uret(proje: ProjeVerisi, cikti_yolu: str | Path) -> Path:
    doc = Document(str(_sablon_yolu("PVP_sablon.docx")))
    _ortak_doldur(doc, proje, rapor=False)
    yol = Path(cikti_yolu)
    doc.save(str(yol))
    return yol


def pvr_uret(proje: ProjeVerisi, cikti_yolu: str | Path, veri_uret: bool = True,
             tohum: int | None = None) -> Path:
    if veri_uret:
        vu.tum_testleri_uret(proje.spek_karti.testler, tohum=tohum)
    # PVR, TEMIZ PVP şablonundan üretilir; sonuç tabloları (Bölüm 11) temiz eklenir.
    # (PVR şablonu zaten dolu örnek sonuç tabloları içerdiği için kullanılmaz —
    #  aksi halde çift içerik oluşurdu.)
    doc = Document(str(_sablon_yolu("PVP_sablon.docx")))
    _ortak_doldur(doc, proje, rapor=True)
    _doldur_sonuclar(doc, proje)
    yol = Path(cikti_yolu)
    doc.save(str(yol))
    return yol
