"""
DOCX üretim motoru (python-docx).

ProjeVerisi'nden PVP (protokol) ve PVR (rapor) Word dosyalarını üretir.
- PVP: spesifikasyonlar + boş/öngörülen yapı.
- PVR: PVP içeriği + dolu sonuç tabloları (Bölüm 11), simüle veriden.

Tasarım: tek dil (Python), tek motor. Sabit metinler cikti/sabit_metinler.py'den;
dinamik tablolar bu modülde. Çıktı UTF-8 Türkçe karakter güvenli.
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION

from core.models import (
    ProjeVerisi, UrunFormu, Test, TabloTipi, LimitTuru,
    SERI_SAYISI, NOKTA_ADLARI,
)
from core import veri_uretici as vu
from cikti import sabit_metinler as sm


# Tema renkleri (lacivert başlıklar)
_LACIVERT = RGBColor(0x00, 0x29, 0x5C)
_KENAR = "BFBFBF"


# ============================================================================
# Yardımcılar
# ============================================================================

def _urun_adi(proje: ProjeVerisi) -> str:
    return proje.dokuman.urun_adi or "Ürün"


def _doldur(metin: str, proje: ProjeVerisi) -> str:
    return metin.replace("{urun}", _urun_adi(proje)).replace(
        "{firma}", proje.dokuman.firma_ismi or "{Firma ismi}")


def _baslik(doc, metin: str, seviye: int = 1) -> None:
    h = doc.add_heading(metin, level=seviye)
    for run in h.runs:
        run.font.color.rgb = _LACIVERT


def _p(doc, metin: str = "", bold: bool = False):
    p = doc.add_paragraph()
    r = p.add_run(metin)
    r.bold = bold
    return p


def _madde(doc, metin: str) -> None:
    doc.add_paragraph(metin, style="List Bullet")


def _seri_basliklari(proje: ProjeVerisi) -> list[str]:
    """Seri numaralarını başlık olarak döndürür (boşsa P01/P02/P03)."""
    out = []
    for i in range(SERI_SAYISI):
        sno = proje.seriler[i].seri_no if i < len(proje.seriler) else ""
        out.append(sno or f"P{i+1:02d}")
    return out


def _hucre(cell, metin: str, bold: bool = False, ortala: bool = True) -> None:
    cell.text = ""
    p = cell.paragraphs[0]
    if ortala:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(str(metin))
    r.bold = bold
    r.font.size = Pt(9)


def _tablo(doc, satir: int, sutun: int):
    t = doc.add_table(rows=satir, cols=sutun)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    return t


# ============================================================================
# Üst bilgi (her sayfada doküman no tablosu)
# ============================================================================

def _ustbilgi_kur(doc, proje: ProjeVerisi, rapor: bool) -> None:
    section = doc.sections[0]
    header = section.header
    header.is_linked_to_previous = False
    # mevcut paragrafı temizle
    for p in list(header.paragraphs):
        p.clear()

    d = proje.dokuman
    dok_no = (d.pvr_dokuman_no if rapor else d.pvp_dokuman_no) or "AG-PV-xxx"
    tip = "RAPOR" if rapor else "PROTOKOL"

    p = header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f"{d.firma_ismi or '{Firma ismi}'}  |  PİLOT ÜRETİM PROSES VALİDASYON {tip}")
    r.bold = True
    r.font.size = Pt(9)

    p2 = header.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run(f"Doküman No: {dok_no}   |   Revizyon No: {d.revizyon_no}   "
                    f"|   Revizyon Tarihi: {d.revizyon_tarihi}   |   {_urun_adi(proje)}")
    r2.font.size = Pt(8)


# ============================================================================
# Ortak (sabit) bölümler
# ============================================================================

def _dokumantasyon(doc, proje: ProjeVerisi) -> None:
    _baslik(doc, "1. DÖKÜMANTASYON", 1)

    _baslik(doc, "1.1 Kısaltma ve Tanımlar", 2)
    t = _tablo(doc, len(sm.KISALTMALAR), 2)
    for i, (k, v) in enumerate(sm.KISALTMALAR):
        _hucre(t.rows[i].cells[0], k, bold=True, ortala=False)
        _hucre(t.rows[i].cells[1], v, ortala=False)

    _baslik(doc, "1.2 Revizyon Detayları", 2)
    t2 = _tablo(doc, 2, 5)
    for c, h in enumerate(["Revizyon No", "Sıra No", "Revizyon Tarihi", "Revizyon Nedeni", "Revize Eden"]):
        _hucre(t2.rows[0].cells[c], h, bold=True)
    for c, v in enumerate(["U.Y.", "U.Y.", "U.Y.", "İlk Yayın", "U.Y."]):
        _hucre(t2.rows[1].cells[c], v)

    _baslik(doc, "1.3 Referanslar", 2)
    for ref in sm.REFERANSLAR:
        _madde(doc, ref)


def _amac_kapsam_sorumluluk(doc, proje: ProjeVerisi) -> None:
    _baslik(doc, "2. AMAÇ", 1)
    _p(doc, _doldur(sm.AMAC_GIRIS, proje))
    _madde(doc, _doldur(sm.AMAC_URUN, proje))
    _p(doc, "Bu protokol ışığında pilot üretim proses validasyon çalışmasının amacı;")
    for m in sm.AMAC_MADDELER:
        _madde(doc, m)

    _baslik(doc, "3. KAPSAM", 1)
    _p(doc, _doldur(sm.KAPSAM, proje))

    _baslik(doc, "4. SORUMLULUKLAR", 1)
    for dept, maddeler in sm.SORUMLULUKLAR:
        _p(doc, dept, bold=True)
        for m in maddeler:
            _madde(doc, m)


# ============================================================================
# Bölüm 5 — Formül + Proses Tanımı + Kapsanan Ürünler
# ============================================================================

def _formul_proses(doc, proje: ProjeVerisi) -> None:
    _baslik(doc, "5. BİRİM FORMÜL ve PROSES TANIMI", 1)

    # Tablo 1 — Birim/Seri Formül
    _p(doc, f"Tablo 1 {_urun_adi(proje)} Birim ve Seri Formül", bold=True)
    if proje.hammaddeler:
        t = _tablo(doc, len(proje.hammaddeler) + 1, 5)
        for c, h in enumerate(["Hammadde / Yardımcı Madde", "Fonksiyon", "Birim Formül (mg/tb)", "% İçerik", "kg / seri"]):
            _hucre(t.rows[0].cells[c], h, bold=True)
        for i, h in enumerate(proje.hammaddeler, start=1):
            _hucre(t.rows[i].cells[0], h.ad, ortala=False)
            _hucre(t.rows[i].cells[1], h.fonksiyon, ortala=False)
            _hucre(t.rows[i].cells[2], "" if h.birim_formul is None else f"{h.birim_formul:g}")
            _hucre(t.rows[i].cells[3], "" if h.yuzde_icerik is None else f"{h.yuzde_icerik:g}")
            _hucre(t.rows[i].cells[4], "" if h.seri_miktar is None else f"{h.seri_miktar:g}")

    # Üretim yöntemi (aşamalar)
    _p(doc, "")
    _p(doc, "Üretim Yöntemi:", bold=True)
    for a in proje.asamalar:
        baslik = f"Operasyon {a.operasyon_no} — Aşama {a.asama_no}"
        if a.ipk_etiketi:
            baslik += f"  ({a.ipk_etiketi})"
        _p(doc, baslik, bold=True)
        if a.metin:
            _p(doc, a.metin)
        if a.parametreler:
            tp = _tablo(doc, len(a.parametreler), 2)
            for i, par in enumerate(a.parametreler):
                _hucre(tp.rows[i].cells[0], par.etiket, ortala=False)
                _hucre(tp.rows[i].cells[1], par.deger)

    # 5.2 Kapsanan Ürünler — Tablo 2
    _baslik(doc, "5.2 Kapsanan Ürünler", 2)
    _p(doc, "Bu validasyon planı çerçevesinde kapsanan ürünler Tablo 2'de gösterilmiştir.")
    _p(doc, "Tablo 2 Kapsanan ürünler", bold=True)
    t2 = _tablo(doc, SERI_SAYISI + 1, 4)
    for c, h in enumerate(["Ürün İsmi", "Seri No", "Seri Boyutu (adet)", "Seri Boyutu (kg)"]):
        _hucre(t2.rows[0].cells[c], h, bold=True)
    for i in range(SERI_SAYISI):
        s = proje.seriler[i]
        _hucre(t2.rows[i+1].cells[0], s.urun_ismi or _urun_adi(proje), ortala=False)
        _hucre(t2.rows[i+1].cells[1], s.seri_no)
        _hucre(t2.rows[i+1].cells[2], s.seri_boyutu_adet)
        _hucre(t2.rows[i+1].cells[3], s.seri_boyutu_kg)


# ============================================================================
# Bölüm 6/7 — Risk + Ekipman
# ============================================================================

def _risk_ekipman(doc, proje: ProjeVerisi) -> None:
    _baslik(doc, "6. PROSES PARAMETRELERİNİN DEĞERLENDİRMESİ (RİSK ANALİZİ)", 1)
    _p(doc, "Tablo 3 Proses için Kritik/Kritik olmayan parametrelerin değerlendirilmesi", bold=True)
    if proje.risk_satirlari:
        t = _tablo(doc, len(proje.risk_satirlari) + 1, 5)
        for c, h in enumerate(["Op. No", "Operasyon", "Kritik (E/H)", "Testler", "Yorumlar"]):
            _hucre(t.rows[0].cells[c], h, bold=True)
        for i, rs in enumerate(proje.risk_satirlari, start=1):
            _hucre(t.rows[i].cells[0], rs.operasyon_no or "")
            _hucre(t.rows[i].cells[1], rs.operasyon, ortala=False)
            _hucre(t.rows[i].cells[2], "E" if rs.kritik else "H")
            _hucre(t.rows[i].cells[3], rs.testler, ortala=False)
            _hucre(t.rows[i].cells[4], rs.yorumlar, ortala=False)

    _baslik(doc, "6.1 Proses Parametreleri", 2)
    _p(doc, "Tablo 4 Öngörülen Proses Parametreleri", bold=True)
    if proje.proses_parametreleri:
        t = _tablo(doc, len(proje.proses_parametreleri) + 1, 3)
        for c, h in enumerate(["Açıklama", "Parametre", "Değer"]):
            _hucre(t.rows[0].cells[c], h, bold=True)
        for i, pp in enumerate(proje.proses_parametreleri, start=1):
            _hucre(t.rows[i].cells[0], pp.aciklama, ortala=False)
            _hucre(t.rows[i].cells[1], pp.parametre, ortala=False)
            _hucre(t.rows[i].cells[2], pp.deger)

    _baslik(doc, "7. EKİPMANLAR", 1)
    _p(doc, "Tablo 5 Üretimde kullanılacak ekipman listesi", bold=True)
    if proje.ekipmanlar:
        t = _tablo(doc, len(proje.ekipmanlar) + 1, 4)
        for c, h in enumerate(["Op. No", "Operasyon", "Ekipman Adı", "Ekipman Kapasitesi"]):
            _hucre(t.rows[0].cells[c], h, bold=True)
        for i, e in enumerate(proje.ekipmanlar, start=1):
            _hucre(t.rows[i].cells[0], e.operasyon_no or "")
            _hucre(t.rows[i].cells[1], e.operasyon, ortala=False)
            _hucre(t.rows[i].cells[2], e.ekipman_adi, ortala=False)
            _hucre(t.rows[i].cells[3], e.kapasite, ortala=False)


# ============================================================================
# Bölüm 8 — Spesifikasyonlar (Tablo 6 / 7)
# ============================================================================

def _spesifikasyonlar(doc, proje: ProjeVerisi) -> None:
    kart = proje.spek_karti
    _baslik(doc, "8. KABUL KRİTERLERİ", 1)
    _baslik(doc, "8.1 IPK, Rutin ve Validasyon Testleri Spesifikasyonları", 2)
    _p(doc, f"Tablo 6 {_urun_adi(proje)}'e ait IPK, Rutin ve Validasyon Testleri Spesifikasyonları", bold=True)

    if kart.testler:
        t = _tablo(doc, len(kart.testler) + 1, 4)
        for c, h in enumerate(["Operasyon", "Test", "Spesifikasyon", "*"]):
            _hucre(t.rows[0].cells[c], h, bold=True)
        for i, test in enumerate(kart.testler, start=1):
            _hucre(t.rows[i].cells[0], test.operasyon, ortala=False)
            _hucre(t.rows[i].cells[1], test.ad, ortala=False)
            _hucre(t.rows[i].cells[2], test.spesifikasyon.metni_olustur(), ortala=False)
            _hucre(t.rows[i].cells[3], "*" if test.yildizli else "")
        _p(doc, "* Proses validasyonu serilerinde uygulanmaktadır.")

    # Tablo 7 — sadece IPK testleri
    ipk = [t for t in kart.testler if t.ipk]
    _p(doc, f"Tablo 7 {_urun_adi(proje)}'e ait IPK Testleri Spesifikasyonları", bold=True)
    if ipk:
        t = _tablo(doc, len(ipk) + 1, 2)
        _hucre(t.rows[0].cells[0], "TESTLER", bold=True)
        _hucre(t.rows[0].cells[1], "SPESİFİKASYONLAR", bold=True)
        for i, test in enumerate(ipk, start=1):
            _hucre(t.rows[i].cells[0], test.ad, ortala=False)
            _hucre(t.rows[i].cells[1], test.spesifikasyon.metni_olustur(), ortala=False)
    else:
        _p(doc, "(IPK olarak işaretlenmiş test yok.)")


# ============================================================================
# Bölüm 9/10 — Numune Alma + Stabilite
# ============================================================================

def _numune_stabilite(doc, proje: ProjeVerisi) -> None:
    _baslik(doc, "9. NUMUNE ALMA PLANI", 1)
    _p(doc, f"Tablo 10 {_urun_adi(proje)}'e ait numune alma planı", bold=True)
    if proje.numune_plani:
        t = _tablo(doc, len(proje.numune_plani) + 1, 4)
        for c, h in enumerate(["Op. No", "Operasyon", "Numune Alma Noktası", "Toplam Numune Miktarı"]):
            _hucre(t.rows[0].cells[c], h, bold=True)
        for i, n in enumerate(proje.numune_plani, start=1):
            _hucre(t.rows[i].cells[0], n.operasyon_no or "")
            _hucre(t.rows[i].cells[1], n.operasyon, ortala=False)
            _hucre(t.rows[i].cells[2], n.numune_noktasi, ortala=False)
            _hucre(t.rows[i].cells[3], n.toplam_miktar)

    _baslik(doc, "10. STABİLİTE", 1)
    _p(doc, sm.STABILITE)


# ============================================================================
# Bölüm 11 — SONUÇLAR (sadece PVR) — 6 tablo tipi
# ============================================================================

def _sonuc_tablo_tek(doc, proje, test, no) -> None:
    seriler = test.sonuc_verisi.get("seriler", ["", "", ""])
    _p(doc, f"Tablo.{no} {test.ad} Sonuçları", bold=True)
    t = _tablo(doc, 2, SERI_SAYISI + 1)
    _hucre(t.rows[0].cells[0], "Numuneler", bold=True)
    for c, sno in enumerate(_seri_basliklari(proje), start=1):
        _hucre(t.rows[0].cells[c], f"Seri No: {sno}", bold=True)
    _hucre(t.rows[1].cells[0], "Sonuç", bold=True)
    for c in range(SERI_SAYISI):
        _hucre(t.rows[1].cells[c+1], seriler[c] if c < len(seriler) else "", bold=True)


def _sonuc_tablo_iki_numune(doc, proje, test, no) -> None:
    seriler = test.sonuc_verisi.get("seriler", [])
    _p(doc, f"Tablo.{no} {test.ad} Sonuçları", bold=True)
    t = _tablo(doc, 5, SERI_SAYISI + 1)
    _hucre(t.rows[0].cells[0], "Test", bold=True)
    _hucre(t.rows[0].cells[1], test.ad, ortala=False)
    _hucre(t.rows[1].cells[0], "Spesifikasyon", bold=True)
    _hucre(t.rows[1].cells[1], test.spesifikasyon.metni_olustur(), ortala=False)
    _hucre(t.rows[2].cells[0], "Numuneler", bold=True)
    for c, sno in enumerate(_seri_basliklari(proje), start=1):
        _hucre(t.rows[2].cells[c], f"Seri No: {sno}", bold=True)
    for ri, anahtar, etiket in [(3, "numune_1", "Numune-1"), (4, "numune_2", "Numune-2")]:
        _hucre(t.rows[ri].cells[0], etiket, ortala=False)
        for c in range(SERI_SAYISI):
            v = seriler[c].get(anahtar, "") if c < len(seriler) else ""
            _hucre(t.rows[ri].cells[c+1], v)
    # Sonuç satırı ekle
    sonuc_row = t.add_row()
    _hucre(sonuc_row.cells[0], "Sonuç", bold=True)
    for c in range(SERI_SAYISI):
        v = seriler[c].get("sonuc", "") if c < len(seriler) else ""
        _hucre(sonuc_row.cells[c+1], v, bold=True)


def _sonuc_tablo_on_numune(doc, proje, test, no) -> None:
    seriler = test.sonuc_verisi.get("seriler", [])
    _p(doc, f"Tablo.{no} {test.ad} Sonuçları", bold=True)
    t = _tablo(doc, 13, SERI_SAYISI + 1)
    _hucre(t.rows[0].cells[0], "Test", bold=True)
    _hucre(t.rows[0].cells[1], test.ad, ortala=False)
    _hucre(t.rows[1].cells[0], "Spesifikasyon", bold=True)
    _hucre(t.rows[1].cells[1], test.spesifikasyon.metni_olustur(), ortala=False)
    for c, sno in enumerate(_seri_basliklari(proje), start=1):
        _hucre(t.rows[2].cells[c], f"Seri No: {sno}", bold=True)
    _hucre(t.rows[2].cells[0], "Numuneler", bold=True)
    for n in range(10):
        _hucre(t.rows[3+n].cells[0], str(n+1))
        for c in range(SERI_SAYISI):
            olc = seriler[c].get("olcumler", []) if c < len(seriler) else []
            _hucre(t.rows[3+n].cells[c+1], olc[n] if n < len(olc) else "")
    for c in range(SERI_SAYISI):
        ort = seriler[c].get("ortalama", "") if c < len(seriler) else ""
        _hucre(t.rows[12].cells[c+1], ort, bold=True)
    _hucre(t.rows[12].cells[0], "Ortalama", bold=True)


def _sonuc_tablo_bos_nokta(doc, proje, test, no, numune_sayisi=10) -> None:
    seriler = test.sonuc_verisi.get("seriler", [])
    _p(doc, f"Tablo.{no} {test.ad} Sonuçları", bold=True)
    # üst başlık (2 satır) + numune satırları + ortalama + sonuç
    t = _tablo(doc, 2, SERI_SAYISI * 3 + 1)
    _hucre(t.rows[0].cells[0], "Numuneler", bold=True)
    # Seri başlıkları (her seri 3 nokta kapsar) — basitlik için tek satırda nokta adları
    col = 1
    for sno in _seri_basliklari(proje):
        for nokta in NOKTA_ADLARI:
            _hucre(t.rows[1].cells[col], nokta, bold=True)
            col += 1
    # seri adlarını üst satıra (3'er hücreye) yaz
    col = 1
    for sno in _seri_basliklari(proje):
        _hucre(t.rows[0].cells[col], f"Seri: {sno}", bold=True)
        col += 3
    # ölçüm satırları
    for n in range(numune_sayisi):
        row = t.add_row()
        _hucre(row.cells[0], str(n+1))
        col = 1
        for c in range(SERI_SAYISI):
            for nokta in NOKTA_ADLARI:
                noktalar = seriler[c].get("noktalar", {}) if c < len(seriler) else {}
                olc = noktalar.get(nokta, {}).get("olcumler", [])
                _hucre(row.cells[col], olc[n] if n < len(olc) else "")
                col += 1
    # ortalama satırı
    row = t.add_row()
    _hucre(row.cells[0], "Ortalama", bold=True)
    col = 1
    for c in range(SERI_SAYISI):
        for nokta in NOKTA_ADLARI:
            noktalar = seriler[c].get("noktalar", {}) if c < len(seriler) else {}
            _hucre(row.cells[col], noktalar.get(nokta, {}).get("ortalama", ""), bold=True)
            col += 1


def _sonuc_tablo_agirlik(doc, proje, test, no) -> None:
    seriler = test.sonuc_verisi.get("seriler", [])
    _p(doc, f"Tablo.{no} {test.ad} Sonuçları", bold=True)
    t = _tablo(doc, 2, SERI_SAYISI * 3 + 1)
    _hucre(t.rows[0].cells[0], "Numuneler", bold=True)
    col = 1
    for sno in _seri_basliklari(proje):
        _hucre(t.rows[0].cells[col], f"Seri: {sno}", bold=True)
        col += 3
    col = 1
    for sno in _seri_basliklari(proje):
        for nokta in NOKTA_ADLARI:
            _hucre(t.rows[1].cells[col], nokta, bold=True)
            col += 1
    for n in range(20):
        row = t.add_row()
        _hucre(row.cells[0], str(n+1))
        col = 1
        for c in range(SERI_SAYISI):
            for nokta in NOKTA_ADLARI:
                noktalar = seriler[c].get("noktalar", {}) if c < len(seriler) else {}
                olc = noktalar.get(nokta, {}).get("olcumler", [])
                _hucre(row.cells[col], olc[n] if n < len(olc) else "")
                col += 1
    for etiket, anahtar in [("Ortalama", "ortalama"), ("RSD%", "rsd"), ("SD", "sd")]:
        row = t.add_row()
        _hucre(row.cells[0], etiket, bold=True)
        col = 1
        for c in range(SERI_SAYISI):
            for nokta in NOKTA_ADLARI:
                noktalar = seriler[c].get("noktalar", {}) if c < len(seriler) else {}
                _hucre(row.cells[col], noktalar.get(nokta, {}).get(anahtar, ""), bold=True)
                col += 1


def _sonuclar(doc, proje: ProjeVerisi) -> None:
    _baslik(doc, "11. SONUÇLAR", 1)
    _baslik(doc, "11.1 PROSES VALİDASYONU TEST SONUÇLARI", 2)

    no = 11
    for test in proje.spek_karti.testler:
        tip = test.tablo_tipi
        if tip is TabloTipi.TEK_SONUC:
            _sonuc_tablo_tek(doc, proje, test, no)
        elif tip is TabloTipi.IKI_NUMUNE:
            _sonuc_tablo_iki_numune(doc, proje, test, no)
        elif tip is TabloTipi.ON_NUMUNE:
            _sonuc_tablo_on_numune(doc, proje, test, no)
        elif tip is TabloTipi.BOS_NOKTA:
            _sonuc_tablo_bos_nokta(doc, proje, test, no)
        elif tip is TabloTipi.AGIRLIK_TEKDUZELIGI:
            _sonuc_tablo_agirlik(doc, proje, test, no)
        else:
            _sonuc_tablo_tek(doc, proje, test, no)
        _p(doc, "")
        no += 1

    _baslik(doc, "12. GENEL DEĞERLENDİRME", 1)
    _p(doc, f"Sapmalar: {proje.sapmalar}")
    _p(doc, f"Sonuç: {proje.sonuc_degerlendirme}")
    if proje.yorum:
        _p(doc, f"Yorum: {proje.yorum}")


# ============================================================================
# Ana üretim fonksiyonları
# ============================================================================

def _stil_kur(doc) -> None:
    st = doc.styles["Normal"]
    st.font.name = "Arial"
    st.font.size = Pt(10)


def belge_uret(proje: ProjeVerisi, rapor: bool) -> Document:
    doc = Document()
    _stil_kur(doc)
    _ustbilgi_kur(doc, proje, rapor)

    _dokumantasyon(doc, proje)
    _amac_kapsam_sorumluluk(doc, proje)
    _formul_proses(doc, proje)
    _risk_ekipman(doc, proje)
    _spesifikasyonlar(doc, proje)
    _numune_stabilite(doc, proje)

    if rapor:
        _sonuclar(doc, proje)

    return doc


def pvp_uret(proje: ProjeVerisi, cikti_yolu: str | Path) -> Path:
    """PVP (protokol) Word dosyası üretir."""
    doc = belge_uret(proje, rapor=False)
    yol = Path(cikti_yolu)
    doc.save(str(yol))
    return yol


def pvr_uret(proje: ProjeVerisi, cikti_yolu: str | Path, veri_uret: bool = True,
             tohum: int | None = None) -> Path:
    """
    PVR (rapor) Word dosyası üretir.
    veri_uret=True ise testlerin sonuç verisi yoksa simüle veri üretir.
    """
    if veri_uret:
        for test in proje.spek_karti.testler:
            if not test.sonuc_verisi:
                test.sonuc_verisi = vu.test_verisi_uret(test)
    doc = belge_uret(proje, rapor=True)
    yol = Path(cikti_yolu)
    doc.save(str(yol))
    return yol
