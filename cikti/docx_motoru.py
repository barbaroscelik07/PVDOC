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
        out.append(sno or f"P{i+1:02d}")
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
    idxler = _veri_satirlarini_ayarla(t, 1, len(proje.hammaddeler))
    for ri, h in zip(idxler, proje.hammaddeler):
        cells = t.rows[ri].cells
        bold = h.ara_toplam
        hucre_yaz(cells[0], h.ad, bold=bold)
        hucre_yaz(cells[1], h.fonksiyon, bold=bold)
        hucre_yaz(cells[2], "" if h.birim_formul is None else f"{h.birim_formul:g}", bold=bold)
        hucre_yaz(cells[3], "" if h.yuzde_icerik is None else f"{h.yuzde_icerik:g}", bold=bold)
        hucre_yaz(cells[4], "" if h.seri_miktar is None else f"{h.seri_miktar:g}", bold=bold)


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
    """Tablo 6 — spesifikasyon. Yıldız test adının SONUNA eklenir (ayrı sütun yok)."""
    t = _tablo_basliga_gore(doc, 6)  # Tablo 6
    kart = proje.spek_karti
    if t is None or not kart.testler:
        return
    idxler = _veri_satirlarini_ayarla(t, 1, len(kart.testler))
    for ri, test in zip(idxler, kart.testler):
        cells = t.rows[ri].cells
        ad = test.ad + ("*" if test.yildizli else "")
        hucre_yaz(cells[0], str(test.operasyon_no or ""))
        hucre_yaz(cells[1], test.operasyon)
        hucre_yaz(cells[2], ad)
        hucre_yaz(cells[3], test.spesifikasyon.metni_olustur())


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


def _ekle_sonuc_tek(doc, proje, test, no):
    seriler = test.sonuc_verisi.get("seriler", ["", "", ""])
    _sonuc_basligi(doc, no, test.ad)
    t = _yeni_tablo(doc, 3, SERI_SAYISI + 1)
    _yaz_bos(t.rows[0].cells[0], "Test", True)
    _yaz_bos(t.rows[0].cells[1], test.ad, False)
    _yaz_bos(t.rows[1].cells[0], "Spesifikasyon", True)
    _yaz_bos(t.rows[1].cells[1], test.spesifikasyon.metni_olustur(), False)
    _yaz_bos(t.rows[2].cells[0], "Sonuç", True)
    for c in range(SERI_SAYISI):
        _yaz_bos(t.rows[2].cells[c+1], seriler[c] if c < len(seriler) else "", True)


def _ekle_sonuc_iki(doc, proje, test, no):
    seriler = test.sonuc_verisi.get("seriler", [])
    _sonuc_basligi(doc, no, test.ad)
    t = _yeni_tablo(doc, 6, SERI_SAYISI + 1)
    _yaz_bos(t.rows[0].cells[0], "Test", True)
    t.rows[0].cells[1].merge(t.rows[0].cells[SERI_SAYISI])
    _yaz_bos(t.rows[0].cells[1], test.ad)
    _yaz_bos(t.rows[1].cells[0], "Spesifikasyon", True)
    t.rows[1].cells[1].merge(t.rows[1].cells[SERI_SAYISI])
    _yaz_bos(t.rows[1].cells[1], test.spesifikasyon.metni_olustur())
    _yaz_bos(t.rows[2].cells[0], "Numuneler", True)
    for c, sno in enumerate(_seri_nolar(proje), 1):
        _yaz_bos(t.rows[2].cells[c], f"Seri No: {sno}", True)
    for ri, key, et in [(3, "numune_1", "Numune-1"), (4, "numune_2", "Numune-2"), (5, "sonuc", "Sonuç")]:
        _yaz_bos(t.rows[ri].cells[0], et, ri == 5)
        for c in range(SERI_SAYISI):
            v = seriler[c].get(key, "") if c < len(seriler) else ""
            _yaz_bos(t.rows[ri].cells[c+1], v, ri == 5)


def _ekle_sonuc_on(doc, proje, test, no):
    seriler = test.sonuc_verisi.get("seriler", [])
    _sonuc_basligi(doc, no, test.ad)
    t = _yeni_tablo(doc, 14, SERI_SAYISI + 1)
    # Test ve Spesifikasyon satırlarında değer hücresi tüm seri sütunlarına yayılır
    _yaz_bos(t.rows[0].cells[0], "Test", True)
    t.rows[0].cells[1].merge(t.rows[0].cells[SERI_SAYISI])
    _yaz_bos(t.rows[0].cells[1], test.ad)
    _yaz_bos(t.rows[1].cells[0], "Spesifikasyon", True)
    t.rows[1].cells[1].merge(t.rows[1].cells[SERI_SAYISI])
    _yaz_bos(t.rows[1].cells[1], test.spesifikasyon.metni_olustur())
    _yaz_bos(t.rows[2].cells[0], "Numuneler", True)
    for c, sno in enumerate(_seri_nolar(proje), 1):
        _yaz_bos(t.rows[2].cells[c], f"Seri No: {sno}", True)
    for n in range(10):
        _yaz_bos(t.rows[3+n].cells[0], str(n+1))
        for c in range(SERI_SAYISI):
            olc = seriler[c].get("olcumler", []) if c < len(seriler) else []
            _yaz_bos(t.rows[3+n].cells[c+1], olc[n] if n < len(olc) else "")
    _yaz_bos(t.rows[13].cells[0], "Ortalama", True)
    for c in range(SERI_SAYISI):
        _yaz_bos(t.rows[13].cells[c+1], seriler[c].get("ortalama", "") if c < len(seriler) else "", True)


def _ekle_sonuc_bos(doc, proje, test, no, ns=10):
    seriler = test.sonuc_verisi.get("seriler", [])
    _sonuc_basligi(doc, no, test.ad)
    t = _yeni_tablo(doc, 2 + ns + 2, SERI_SAYISI * 3 + 1)
    _yaz_bos(t.rows[0].cells[0], "Numuneler", True)
    col = 1
    for sno in _seri_nolar(proje):
        _yaz_bos(t.rows[0].cells[col], f"Seri: {sno}", True); col += 3
    col = 1
    for _ in range(SERI_SAYISI):
        for nokta in NOKTA_ADLARI:
            _yaz_bos(t.rows[1].cells[col], nokta, True); col += 1
    for n in range(ns):
        _yaz_bos(t.rows[2+n].cells[0], str(n+1))
        col = 1
        for c in range(SERI_SAYISI):
            for nokta in NOKTA_ADLARI:
                noktalar = seriler[c].get("noktalar", {}) if c < len(seriler) else {}
                olc = noktalar.get(nokta, {}).get("olcumler", [])
                _yaz_bos(t.rows[2+n].cells[col], olc[n] if n < len(olc) else ""); col += 1
    # ortalama
    _yaz_bos(t.rows[2+ns].cells[0], "Ortalama", True)
    col = 1
    for c in range(SERI_SAYISI):
        for nokta in NOKTA_ADLARI:
            noktalar = seriler[c].get("noktalar", {}) if c < len(seriler) else {}
            _yaz_bos(t.rows[2+ns].cells[col], noktalar.get(nokta, {}).get("ortalama", ""), True); col += 1
    # sonuç (her seri tek)
    _yaz_bos(t.rows[3+ns].cells[0], "Sonuç", True)
    col = 1
    for c in range(SERI_SAYISI):
        sonuc = seriler[c].get("sonuc", "") if c < len(seriler) else ""
        _yaz_bos(t.rows[3+ns].cells[col], sonuc, True)
        col += 3


def _ekle_sonuc_agirlik(doc, proje, test, no):
    seriler = test.sonuc_verisi.get("seriler", [])
    _sonuc_basligi(doc, no, test.ad)
    t = _yeni_tablo(doc, 2 + 20 + 3, SERI_SAYISI * 3 + 1)
    _yaz_bos(t.rows[0].cells[0], "Numuneler", True)
    col = 1
    for sno in _seri_nolar(proje):
        _yaz_bos(t.rows[0].cells[col], f"Seri: {sno}", True); col += 3
    col = 1
    for _ in range(SERI_SAYISI):
        for nokta in NOKTA_ADLARI:
            _yaz_bos(t.rows[1].cells[col], nokta, True); col += 1
    for n in range(20):
        _yaz_bos(t.rows[2+n].cells[0], str(n+1))
        col = 1
        for c in range(SERI_SAYISI):
            for nokta in NOKTA_ADLARI:
                noktalar = seriler[c].get("noktalar", {}) if c < len(seriler) else {}
                olc = noktalar.get(nokta, {}).get("olcumler", [])
                _yaz_bos(t.rows[2+n].cells[col], olc[n] if n < len(olc) else ""); col += 1
    for k, (et, key) in enumerate([("Ortalama", "ortalama"), ("RSD%", "rsd"), ("SD", "sd")]):
        _yaz_bos(t.rows[22+k].cells[0], et, True)
        col = 1
        for c in range(SERI_SAYISI):
            for nokta in NOKTA_ADLARI:
                noktalar = seriler[c].get("noktalar", {}) if c < len(seriler) else {}
                _yaz_bos(t.rows[22+k].cells[col], noktalar.get(nokta, {}).get(key, ""), True); col += 1


def _doldur_sonuclar(doc, proje: ProjeVerisi) -> None:
    """PVR Bölüm 11 sonuç tablolarını belgenin sonuna ekler."""
    doc.add_page_break()
    h = doc.add_paragraph()
    r = h.add_run("11. SONUÇLAR — PROSES VALİDASYONU TEST SONUÇLARI")
    r.bold = True; r.font.size = Pt(12)

    no = 11
    for test in proje.spek_karti.testler:
        tip = test.tablo_tipi
        if tip is TabloTipi.TEK_SONUC:
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


def _ortak_doldur(doc, proje: ProjeVerisi, rapor: bool) -> None:
    belgede_degistir(doc, _placeholder_eslemeleri(proje, rapor))
    _doldur_formul(doc, proje)
    _doldur_kapsanan(doc, proje)
    _doldur_risk(doc, proje)
    _doldur_proses_param(doc, proje)
    _doldur_ekipman(doc, proje)
    _doldur_spek(doc, proje)
    _doldur_ipk(doc, proje)
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
