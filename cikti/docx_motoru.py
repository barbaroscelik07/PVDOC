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


def _tablo_genislik_duzelt(tablo):
    """
    Şablonda tblW=0/auto olan tablolarda LibreOffice sütunları çökertir.
    tblW'yi tblGrid toplamına (dxa) sabitler — fixed layout ile birlikte düzgün render.
    """
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    tbl = tablo._tbl
    tblPr = tbl.tblPr
    grid = tbl.find(qn('w:tblGrid'))
    if grid is None:
        return
    toplam = sum(int(gc.get(qn('w:w'))) for gc in grid.findall(qn('w:gridCol'))
                 if gc.get(qn('w:w')))
    if not toplam:
        return
    tblW = tblPr.find(qn('w:tblW'))
    if tblW is None:
        tblW = OxmlElement('w:tblW'); tblPr.append(tblW)
    tblW.set(qn('w:w'), str(toplam)); tblW.set(qn('w:type'), 'dxa')


def _tablo_render_duzelt(tablo):
    """
    LibreOffice PDF render'ında çok satırlı tabloların sütun çökmesini önler:
    fixed layout + her hücreye tblGrid'den gelen sabit genişlik uygular.
    """
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    tbl = tablo._tbl
    tblPr = tbl.tblPr
    layout = tblPr.find(qn('w:tblLayout'))
    if layout is None:
        layout = OxmlElement('w:tblLayout'); tblPr.append(layout)
    layout.set(qn('w:type'), 'fixed')
    grid = tbl.find(qn('w:tblGrid'))
    if grid is None:
        return
    genislikler = [int(gc.get(qn('w:w'))) for gc in grid.findall(qn('w:gridCol'))
                   if gc.get(qn('w:w'))]
    if not genislikler:
        return
    toplam = sum(genislikler)
    tblW = tblPr.find(qn('w:tblW'))
    if tblW is None:
        tblW = OxmlElement('w:tblW'); tblPr.append(tblW)
    tblW.set(qn('w:w'), str(toplam)); tblW.set(qn('w:type'), 'dxa')
    # her hücreye sabit genişlik
    for row in tablo.rows:
        for ci, cell in enumerate(row.cells):
            if ci >= len(genislikler):
                break
            tcPr = cell._tc.get_or_add_tcPr()
            tcW = tcPr.find(qn('w:tcW'))
            if tcW is None:
                tcW = OxmlElement('w:tcW'); tcPr.append(tcW)
            tcW.set(qn('w:w'), str(genislikler[ci])); tcW.set(qn('w:type'), 'dxa')


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


def _tablo_sabit_layout(tablo):
    """
    Tabloya sabit (fixed) layout ve tablo genişliği uygular; LibreOffice'in
    çok satırlı tablolarda sütunları çökertmesini önler.
    """
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    tbl = tablo._tbl
    tblPr = tbl.tblPr
    # tblLayout = fixed
    mevcut = tblPr.find(qn('w:tblLayout'))
    if mevcut is None:
        layout = OxmlElement('w:tblLayout')
        layout.set(qn('w:type'), 'fixed')
        tblPr.append(layout)
    else:
        mevcut.set(qn('w:type'), 'fixed')
    # tblW = grid toplamı (dxa)
    grid = tbl.find(qn('w:tblGrid'))
    if grid is not None:
        toplam = sum(int(gc.get(qn('w:w'))) for gc in grid.findall(qn('w:gridCol'))
                     if gc.get(qn('w:w')))
        tblW = tblPr.find(qn('w:tblW'))
        if tblW is None:
            tblW = OxmlElement('w:tblW'); tblPr.append(tblW)
        tblW.set(qn('w:w'), str(toplam))
        tblW.set(qn('w:type'), 'dxa')


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
    """
    Tablo 6 — spesifikasyon. Türetilmiş test listesi tek geçişte yazılır.
    Her satırda Op No + Operasyon dolu (alt satırlar dahil). İlgili Bileşikler
    grup başlığı + impurite satırları test listesinde sırayla gelir.
    """
    t = _tablo_basliga_gore(doc, 6)
    kart = proje.spek_karti
    if t is None or not kart.testler:
        return

    # Satır planı: (op_no, op, ad, spek, alt_satirlar)
    plan = []  # her eleman: (opno, op, ad, spek)
    for test in kart.testler:
        opno = str(test.operasyon_no or "")
        op = test.operasyon
        yildiz = "*" if test.yildizli else ""
        ad = test.ad + yildiz
        # ekstra alt satırlar (mikrobiyolojik 3 alt, ağırlık tekdüzeliği 2 alt)
        ekstra = list(test.alt_satirlar)
        if test.aciklama_etiketi:
            ekstra.append((test.aciklama_etiketi, test.aciklama_spek))
        if test.aciklama2_etiketi:
            ekstra.append((test.aciklama2_etiketi, test.aciklama2_spek))
        if ekstra:
            # başlık satırı (spek boş), sonra alt satırlar
            plan.append((opno, op, ad, ""))
            for et, sp in ekstra:
                plan.append((opno, op, et, sp))
        else:
            # grup başlığı testinin spek'i boş olmalı (İlgili Bileşikler başlığı)
            if getattr(test, "_grup_baslik", False):
                plan.append((opno, op, ad, ""))
            elif getattr(test, "_impurite", False):
                plan.append((opno, op, ad, test.spesifikasyon.spesifikasyon_metni or ""))
            else:
                plan.append((opno, op, ad, test.spesifikasyon.metni_olustur()))

    idxler = _veri_satirlarini_ayarla(t, 1, len(plan))
    for ri, (opno, op, ad, spek) in zip(idxler, plan):
        cells = t.rows[ri].cells
        hucre_yaz(cells[0], opno)
        hucre_yaz(cells[1], op)
        hucre_yaz(cells[2], ad)
        hucre_yaz(cells[3], spek)
    return


def _doldur_spek_ESKI(doc, proje: ProjeVerisi) -> None:
    """Tablo 6 — spesifikasyon. Yıldız test adının SONUNA; alt satırlar ayrı satır.
    İlgili Bileşikler etkin madde başına GRUPLU eklenir (operasyon hücresi dikey birleşik)."""
    t = _tablo_basliga_gore(doc, 6)  # Tablo 6
    kart = proje.spek_karti
    if t is None:
        return

    # 1) Normal testlerin satır planı
    satir_planı = []  # (tip, veri)
    for test in kart.testler:
        ekstra = list(test.alt_satirlar)
        if test.aciklama_etiketi:
            ekstra = ekstra + [(test.aciklama_etiketi, test.aciklama_spek)]
        if test.aciklama2_etiketi:
            ekstra = ekstra + [(test.aciklama2_etiketi, test.aciklama2_spek)]
        satir_planı.append(("test", (test, ekstra)))

    # 2) İlgili Bileşikler grupları (etkin madde başına)
    #    Her grup: başlık satırı + impurite satırları; operasyon dikey birleşik.
    ilgili_gruplar = []
    for em in kart.etkin_maddeler:
        if not em.impuriteler:
            continue
        ilk = em.impuriteler[0]
        yildiz = "*" if any(i.yildizli for i in em.impuriteler) else ""
        ilgili_gruplar.append({
            "baslik": f"{em.ad} İlgili Bileşikler{yildiz}",
            "operasyon": ilk.operasyon,
            "operasyon_no": ilk.operasyon_no,
            "impuriteler": em.impuriteler,
        })

    # Toplam satır sayısı
    toplam = 0
    for tip, veri in satir_planı:
        _, ekstra = veri
        toplam += 1 + len(ekstra)
    for g in ilgili_gruplar:
        toplam += 1 + len(g["impuriteler"])  # başlık + impurite satırları

    if toplam == 0:
        return

    idxler = _veri_satirlarini_ayarla(t, 1, toplam)
    it = iter(idxler)

    # Normal testler
    for tip, veri in satir_planı:
        test, ekstra = veri
        ri = next(it)
        cells = t.rows[ri].cells
        ad = test.ad + ("*" if test.yildizli else "")
        hucre_yaz(cells[0], str(test.operasyon_no or ""))
        hucre_yaz(cells[1], test.operasyon)
        hucre_yaz(cells[2], ad)
        ana_spek = "" if ekstra else test.spesifikasyon.metni_olustur()
        hucre_yaz(cells[3], ana_spek)
        for etiket, spek_metni in ekstra:
            ri2 = next(it)
            c2 = t.rows[ri2].cells
            hucre_yaz(c2[0], ""); hucre_yaz(c2[1], "")
            hucre_yaz(c2[2], etiket); hucre_yaz(c2[3], spek_metni)

    # İlgili Bileşikler grupları
    for g in ilgili_gruplar:
        grup_satir_idx = []
        # başlık satırı
        rb = next(it)
        grup_satir_idx.append(rb)
        cb = t.rows[rb].cells
        hucre_yaz(cb[0], str(g["operasyon_no"] or ""))
        hucre_yaz(cb[1], g["operasyon"])
        hucre_yaz(cb[2], g["baslik"])
        hucre_yaz(cb[3], "")
        # impurite satırları
        for imp in g["impuriteler"]:
            ri = next(it)
            grup_satir_idx.append(ri)
            c = t.rows[ri].cells
            hucre_yaz(c[0], ""); hucre_yaz(c[1], "")
            ad = imp.ad if imp.ad.startswith("—") or imp.ad.startswith("-") else f"—{imp.ad}"
            hucre_yaz(c[2], ad)
            hucre_yaz(c[3], imp.limit_metni or "")
        # Operasyon No + Operasyon hücrelerini grup boyunca DİKEY birleştir
        if len(grup_satir_idx) > 1:
            ust, alt = grup_satir_idx[0], grup_satir_idx[-1]
            t.rows[ust].cells[0].merge(t.rows[alt].cells[0])
            t.rows[ust].cells[1].merge(t.rows[alt].cells[1])
            hucre_yaz(t.rows[ust].cells[0], str(g["operasyon_no"] or ""))
            hucre_yaz(t.rows[ust].cells[1], g["operasyon"])


def _doldur_ipk(doc, proje: ProjeVerisi) -> None:
    """Tablo 7 — sadece IPK testleri. Ağırlık Tekdüzeliği başlık + 2 alt satır."""
    ipk = [t for t in proje.spek_karti.testler if t.ipk]
    t = _tablo_basliga_gore(doc, 7)  # Tablo 7
    if t is None or not ipk:
        return
    sut = len(t.columns)

    # Satır planı: her test 1 satır; Ağırlık Tekdüzeliği için + alt satırlar
    plan = []  # (op_no, op, ad, spek)
    for test in ipk:
        n = test.ad.lower()
        if "ağırlık tekdüzeliği" in n or "agirlik tekduzeligi" in n:
            plan.append((test.operasyon_no, test.operasyon, test.ad, ""))  # başlık boş
            if test.aciklama_etiketi:
                plan.append((test.operasyon_no, test.operasyon, test.aciklama_etiketi, test.aciklama_spek))
            if test.aciklama2_etiketi:
                plan.append((test.operasyon_no, test.operasyon, test.aciklama2_etiketi, test.aciklama2_spek))
        else:
            plan.append((test.operasyon_no, test.operasyon, test.ad, test.spesifikasyon.metni_olustur()))

    idxler = _veri_satirlarini_ayarla(t, 1, len(plan))
    for ri, (opno, op, ad, spek) in zip(idxler, plan):
        cells = t.rows[ri].cells
        if sut == 2:
            hucre_yaz(cells[0], ad)
            hucre_yaz(cells[1], spek)
        else:
            hucre_yaz(cells[0], str(opno or ""))
            hucre_yaz(cells[1], op)
            hucre_yaz(cells[2], ad)
            hucre_yaz(cells[3], spek)


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


def _seri_dict(seriler, c):
    """seriler[c] güvenli dict erişimi (string/eksikse boş dict)."""
    if c < len(seriler) and isinstance(seriler[c], dict):
        return seriler[c]
    return {}


def _sr(cells, degerler, bold=False):
    for c, v in zip(cells, degerler):
        hucre_yaz(c, v, bold=bold) if c.paragraphs[0].runs else _yaz_bos(c, v, bold)


def _bicimle(metin):
    """Sayısal değerleri her zaman 2 ondalıkla biçimler (1 -> 1,00, 0.6 -> 0,60).
    Metin/None ise olduğu gibi döner. Türkçe ondalık ayracı (,) kullanılır."""
    if metin is None or metin == "":
        return ""
    if isinstance(metin, (int, float)):
        return f"{metin:.2f}".replace(".", ",")
    return str(metin)


def _yaz_bos(cell, metin, bold=False):
    p = cell.paragraphs[0]
    r = p.add_run(_bicimle(metin))
    r.bold = bold
    r.font.size = Pt(12)
    r.font.name = "Times New Roman"
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER


def _yaz_sol(cell, metin, bold=False):
    """Sola dayalı hücre yazımı (Test/Spesifikasyon değer hücreleri için)."""
    p = cell.paragraphs[0]
    r = p.add_run(str(metin))
    r.bold = bold
    r.font.size = Pt(12)
    r.font.name = "Times New Roman"
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
            v = _seri_dict(seriler, c).get(key, "")
            _yaz_bos(t.rows[ri].cells[c+1], v, ri == 6)
    # T.E. (tespit edilemedi) verisi varsa tablonun ALTINA not ekle
    if test.sonuc_verisi.get("te"):
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
            olc = _seri_dict(seriler, c).get("olcumler", [])
            _yaz_bos(t.rows[4+n].cells[c+1], olc[n] if n < len(olc) else "")
    _yaz_bos(t.rows[14].cells[0], "Ortalama", True)
    for c in range(SERI_SAYISI):
        _yaz_bos(t.rows[14].cells[c+1], _seri_dict(seriler, c).get("ortalama", ""), True)


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
                noktalar = _seri_dict(seriler, c).get("noktalar", {})
                olc = noktalar.get(nokta, {}).get("olcumler", [])
                _yaz_bos(t.rows[4+n].cells[col], olc[n] if n < len(olc) else ""); col += 1
    _yaz_bos(t.rows[4+ns].cells[0], "Ortalama", True)
    col = 1
    for c in range(SERI_SAYISI):
        for nokta in NOKTA_ADLARI:
            noktalar = _seri_dict(seriler, c).get("noktalar", {})
            _yaz_bos(t.rows[4+ns].cells[col], noktalar.get(nokta, {}).get("ortalama", ""), True); col += 1
    _yaz_bos(t.rows[5+ns].cells[0], "Sonuç", True)
    col = 1
    for c in range(SERI_SAYISI):
        sonuc = _seri_dict(seriler, c).get("sonuc", "")
        a = t.rows[5+ns].cells[col]; a.merge(t.rows[5+ns].cells[col+2])
        _yaz_bos(a, sonuc, True); col += 3


def _ekle_sonuc_ortalama_agirlik(doc, proje, test, no):
    """
    Ortalama Ağırlık (resim 36): Test / Spesifikasyon / Numuneler+Analiz /
    Seri No + Baş/Orta/Son / Sonuç (her nokta tek değer) / Ortalama (seri ort).
    Değerler Ağırlık Tekdüzeliği nokta-ortalamalarından gelir (birebir eşleşir).
    """
    seriler = test.sonuc_verisi.get("seriler", [])
    # Güvenlik: veri beklenen dict yapısında değilse (örn. test yanlış tiple
    # eklenmiş ve string listesi gelmiş) boş nokta yapısına çevir.
    def _nokta(c):
        if c < len(seriler) and isinstance(seriler[c], dict):
            return seriler[c].get("noktalar", {})
        return {}
    def _seri_sonuc(c):
        if c < len(seriler) and isinstance(seriler[c], dict):
            return seriler[c].get("sonuc", "")
        return ""
    _sonuc_basligi(doc, no, test.ad)
    # Test + Spesifikasyon + Numuneler + nokta başlık + Sonuç + Ortalama = 6 satır
    t = _yeni_tablo(doc, 6, SERI_SAYISI * 3 + 1)
    _yaz_bos(t.rows[0].cells[0], "Test", True)
    t.rows[0].cells[1].merge(t.rows[0].cells[SERI_SAYISI * 3])
    _yaz_sol(t.rows[0].cells[1], test.ad)
    _yaz_bos(t.rows[1].cells[0], "Spesifikasyon", True)
    t.rows[1].cells[1].merge(t.rows[1].cells[SERI_SAYISI * 3])
    _yaz_sol(t.rows[1].cells[1], test.spesifikasyon.metni_olustur())
    # Numuneler / Seri No (3'er birleşik) + nokta başlıkları
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
    # Sonuç satırı: her nokta tek değer (= Ağırlık Tekdüzeliği nokta ortalaması)
    _yaz_bos(t.rows[4].cells[0], "Sonuç", True)
    col = 1
    for c in range(SERI_SAYISI):
        noktalar = _nokta(c)
        for nokta in NOKTA_ADLARI:
            _yaz_bos(t.rows[4].cells[col], noktalar.get(nokta, {}).get("ortalama", ""), False); col += 1
    # Ortalama satırı: seri başına tek (3 nokta birleşik), seri ortalaması
    _yaz_bos(t.rows[5].cells[0], "Ortalama", True)
    col = 1
    for c in range(SERI_SAYISI):
        sonuc = _seri_sonuc(c)
        a = t.rows[5].cells[col]; a.merge(t.rows[5].cells[col+2])
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
                noktalar = _seri_dict(seriler, c).get("noktalar", {})
                olc = noktalar.get(nokta, {}).get("olcumler", [])
                _yaz_bos(t.rows[4+n].cells[col], olc[n] if n < len(olc) else ""); col += 1
    for k, (et, key) in enumerate([("Ortalama", "ortalama"), ("RSD%", "rsd"), ("SD", "sd")]):
        _yaz_bos(t.rows[24+k].cells[0], et, True)
        col = 1
        for c in range(SERI_SAYISI):
            for nokta in NOKTA_ADLARI:
                noktalar = _seri_dict(seriler, c).get("noktalar", {})
                _yaz_bos(t.rows[24+k].cells[col], noktalar.get(nokta, {}).get(key, ""), True); col += 1


def _ekle_sonuc_impurite(doc, proje, imp, baslik, no, tohum_ek=0):
    """
    Bir impurite için sonuç tablosu (resim — Tablo 25-28):
      Test | <başlık>  (örn. 'Etkin madde 1 imp. a')
      Spesifikasyon | <limit metni>
      Numuneler/Analiz | Seri No
      Numune 1 / Numune 2 / Sonuç
    Maksimum değere uyar; T.E. ise hepsi 'T.E.' + alt not.
    """
    import random as _r
    _sonuc_basligi(doc, no, baslik)
    t = _yeni_tablo(doc, 7, SERI_SAYISI + 1)
    _yaz_bos(t.rows[0].cells[0], "Test", True)
    t.rows[0].cells[1].merge(t.rows[0].cells[SERI_SAYISI])
    _yaz_sol(t.rows[0].cells[1], baslik)
    _yaz_bos(t.rows[1].cells[0], "Spesifikasyon", True)
    t.rows[1].cells[1].merge(t.rows[1].cells[SERI_SAYISI])
    _yaz_sol(t.rows[1].cells[1], imp.limit_metni or "")
    # Numuneler / Analiz Sonuçları + Seri No
    _yaz_bos(t.rows[2].cells[0], "Numuneler", True)
    t.rows[2].cells[1].merge(t.rows[2].cells[SERI_SAYISI])
    _yaz_bos(t.rows[2].cells[1], "Analiz Sonuçları", True)
    t.rows[2].cells[0].merge(t.rows[3].cells[0])
    for c, sno in enumerate(_seri_nolar(proje), 1):
        _yaz_bos(t.rows[3].cells[c], f"Seri No: {sno}", True)

    te = imp.te or (str(imp.limit_metni).upper().replace(" ", "").endswith("T.E.")
                    if imp.limit_metni else False)
    maks = imp.maksimum_deger
    for ri, et in [(4, "Numune 1"), (5, "Numune 2"), (6, "Sonuç")]:
        _yaz_sol(t.rows[ri].cells[0], et, ri == 6)
        for c in range(SERI_SAYISI):
            if te:
                deg = "T.E."
            elif maks:
                # maksimumun altında sağlıklı değer (örn maks*0.05–0.6)
                deg = _bicimle_sayi(_r.uniform(maks * 0.05, maks * 0.6))
            else:
                deg = "T.E."
            _yaz_bos(t.rows[ri].cells[c+1], deg, ri == 6)
    if te:
        np = doc.add_paragraph()
        nr = np.add_run("T.E.: Tespit edilemedi.")
        nr.italic = True; nr.font.size = Pt(8)


def _bicimle_sayi(deger, ondalik=2) -> str:
    """Çıktı motoru içi: sayıyı 2 ondalık + Türkçe virgülle string yapar."""
    try:
        return f"{float(deger):.{ondalik}f}".replace(".", ",")
    except (ValueError, TypeError):
        return str(deger)


def _ekle_limit_tablosu(doc, proje, baslik, limit_metni, no):
    """
    Ağırlık Tekdüzeliği limit tablosu (resim 2 — Tablo 35/36):
      Test | <başlık>
      Spesifikasyon | <limit metni>
      Numuneler / Analiz Sonuçları | Seri No
      Sonuç | <limit metni> × 3 seri
    """
    doc.add_paragraph("")
    _sonuc_basligi(doc, no, baslik)
    t = _yeni_tablo(doc, 5, SERI_SAYISI + 1)
    _yaz_bos(t.rows[0].cells[0], "Test", True)
    t.rows[0].cells[1].merge(t.rows[0].cells[SERI_SAYISI])
    _yaz_sol(t.rows[0].cells[1], baslik)
    _yaz_bos(t.rows[1].cells[0], "Spesifikasyon", True)
    t.rows[1].cells[1].merge(t.rows[1].cells[SERI_SAYISI])
    _yaz_sol(t.rows[1].cells[1], limit_metni)
    _yaz_bos(t.rows[2].cells[0], "Numuneler", True)
    t.rows[2].cells[1].merge(t.rows[2].cells[SERI_SAYISI])
    _yaz_bos(t.rows[2].cells[1], "Analiz Sonuçları", True)
    t.rows[2].cells[0].merge(t.rows[3].cells[0])
    for c, sno in enumerate(_seri_nolar(proje), 1):
        _yaz_bos(t.rows[3].cells[c], f"Seri No: {sno}", True)
    _yaz_bos(t.rows[4].cells[0], "Sonuç", True)
    for c in range(SERI_SAYISI):
        _yaz_bos(t.rows[4].cells[c+1], limit_metni, False)


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


def _genel_degerlendirme_paragrafi(doc):
    """Şablondaki 'GENEL DEĞERLENDİRME' başlık paragrafını bulur (varsa)."""
    for p in doc.paragraphs:
        if "GENEL DEĞERLENDİRME" in p.text.strip().upper():
            return p
    return None


def _doldur_sonuclar(doc, proje: ProjeVerisi) -> None:
    """
    PVR sonuç tablolarını şablondaki 'GENEL DEĞERLENDİRME' başlığının ÖNÜNE ekler.
    Böylece Genel Değerlendirme HER ZAMAN en sonda kalır.
    """
    hedef = _genel_degerlendirme_paragrafi(doc)

    # Sonuç bloğunu belge SONUNA üret (geçici), sonra hedefin önüne taşı.
    # Üretilen yeni elemanları belirlemek için mevcut body çocuklarını işaretle.
    from docx.oxml.ns import qn
    body = doc.element.body
    onceki = list(body)

    # başlık
    h = doc.add_paragraph()
    r = h.add_run("PROSES VALİDASYONU TEST SONUÇLARI")
    r.bold = True; r.font.size = Pt(12)

    no = 11
    for test in proje.spek_karti.testler:
        tip = test.tablo_tipi
        if test.mikrobiyolojik or tip is TabloTipi.MATRIS:
            _ekle_sonuc_matris(doc, proje, test, no)
        elif "ortalama ağırlık" in test.ad.lower():
            _ekle_sonuc_ortalama_agirlik(doc, proje, test, no)
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
            # Resim 2: ayrıca 2 limit tablosu (sapabilir + sapmamalıdır)
            if test.aciklama_etiketi:
                no += 1
                baslik1 = test.aciklama_etiketi.lstrip("—- ").strip()
                _ekle_limit_tablosu(doc, proje, baslik1, test.aciklama_spek, no)
            if test.aciklama2_etiketi:
                no += 1
                baslik2 = test.aciklama2_etiketi.lstrip("—- ").strip()
                _ekle_limit_tablosu(doc, proje, baslik2, test.aciklama2_spek, no)
        else:
            _ekle_sonuc_tek(doc, proje, test, no)
        doc.add_paragraph("")
        no += 1

    # İlgili Bileşikler (impurite) sonuç tabloları — her etkin maddenin her impuritesi
    for em in proje.spek_karti.etkin_maddeler:
        for imp in em.impuriteler:
            ad = imp.ad if imp.ad.startswith("—") or imp.ad.startswith("-") else imp.ad
            baslik = f"{em.ad} {ad}".replace("—", "").replace("- ", "").strip()
            _ekle_sonuc_impurite(doc, proje, imp, baslik, no)
            doc.add_paragraph("")
            no += 1

    # Yeni eklenen elemanlar (başlık + tablolar + boş paragraflar)
    yeni_elemanlar = [el for el in body if el not in onceki]

    if hedef is not None:
        hedef_el = hedef._p
        for el in yeni_elemanlar:
            body.remove(el)
            hedef_el.addprevious(el)
    # hedef yoksa zaten en sonda kalır (sorun değil)


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


def _test_bul(testler, *anahtarlar):
    """Adında verilen anahtar kelimeleri içeren ilk testi bulur."""
    for t in testler:
        ad = t.ad.lower()
        if all(a.lower() in ad for a in anahtarlar):
            return t
    return None


def _doldur_tablo89(doc, proje: ProjeVerisi) -> None:
    """
    Tablo 8 (Serbest Bırakma) ve Tablo 9 (Raf Ömrü).
    SABİT şablon sırası: Görünüş → Ortalama Ağırlık → Ağırlık Tekdüzeliği (2 alt) →
    Dağılma → her etkin madde (Teşhis / Miktar Tayini) → Dissolüsyon →
    İlgili Bileşikler (etken başına gruplu) → Mikrobiyolojik Kontrol (3 alt).
    İçerik Tablo 6'daki testlerden ve etkin maddelerden gelir.
    Tablo 9 farkları: Ağırlık Tekdüzeliği* ve Teşhis* yıldızlı, Miktar Tayini
    toleransı farklı, en altta '* Stabilite analizlerinde bakılmayacaktır' notu.
    """
    kart = proje.spek_karti
    if not kart.tablo89_ekle:
        return

    testler = kart.testler
    etkenler = kart.etkin_maddeler

    def _satirlari_uret(raf_omru: bool, tol: str):
        """(etiket, spek, girinti) üçlülerinden satır listesi üretir."""
        yildiz = "*" if raf_omru else ""
        satirlar = []  # (sol_metin, sag_metin)

        # Görünüş
        g = _test_bul(testler, "görünüş")
        satirlar.append(("Görünüş", g.spesifikasyon.metni_olustur() if g else ""))
        # Ortalama Ağırlık
        oa = _test_bul(testler, "ortalama ağırlık")
        satirlar.append(("Ortalama Ağırlık", oa.spesifikasyon.metni_olustur() if oa else ""))
        # Ağırlık Tekdüzeliği (başlık + 2 alt)
        at = _test_bul(testler, "ağırlık tekdüzeliği")
        satirlar.append((f"Ağırlık Tekdüzeliği{yildiz}", ""))
        if at:
            satirlar.append((at.aciklama_etiketi or "—20 tablette tek tek tabletlerden maksimum 2 tanesi bu limitten sapabilir.", at.aciklama_spek))
            if at.aciklama2_etiketi:
                satirlar.append((at.aciklama2_etiketi, at.aciklama2_spek))
        # Dağılma
        dg = _test_bul(testler, "dağılma")
        satirlar.append(("Dağılma", dg.spesifikasyon.metni_olustur() if dg else "Maksimum 30 dakika"))
        # Her etkin madde: Teşhis + Miktar Tayini
        for em in etkenler:
            satirlar.append((em.ad, ""))  # ara başlık
            tes = _test_bul(testler, em.ad, "teşhis")
            satirlar.append((f"— Teşhis{yildiz}",
                             tes.spesifikasyon.metni_olustur() if tes else "Standart ve numune alıkonma zamanı aynı olmalıdır."))
            mt = _test_bul(testler, em.ad, "miktar tayini")
            if mt:
                mt_spek = mt.spesifikasyon.metni_olustur()
                if raf_omru and mt.spesifikasyon.hedef_deger:
                    mt_spek = _miktar_spek_uret(mt.spesifikasyon.hedef_deger, tol, mt.spesifikasyon.birim)
                satirlar.append(("—Miktar Tayini", mt_spek))
        # Dissolüsyon (her etken)
        for em in etkenler:
            ds = _test_bul(testler, em.ad, "dissolüsyon")
            if ds:
                satirlar.append((f"{em.ad} Dissolüsyon (Q)", ds.spesifikasyon.metni_olustur()))
        # İlgili Bileşikler (etken başına gruplu)
        if any(em.impuriteler for em in etkenler):
            satirlar.append(("İlgili Bileşikler", ""))
            for em in etkenler:
                if not em.impuriteler:
                    continue
                satirlar.append((f"{em.ad}'e Ait", ""))  # italik ara başlık
                for imp in em.impuriteler:
                    ad = imp.ad if imp.ad.startswith("—") or imp.ad.startswith("-") else f"—{imp.ad}"
                    satirlar.append((ad, imp.limit_metni or ""))
        # Mikrobiyolojik Kontrol (3 alt)
        mik = next((t for t in testler if t.mikrobiyolojik), None)
        if mik:
            satirlar.append(("Mikrobiyolojik Kontrol", ""))
            for etiket, spek_metni in (mik.alt_satirlar or []):
                ad = etiket if etiket.startswith("—") or etiket.startswith("-") else f"—{etiket}"
                satirlar.append((ad, spek_metni))
        return satirlar

    for tablo_no, raf, tol in [(8, False, kart.serbest_birakma_tolerans),
                               (9, True, kart.raf_omru_tolerans)]:
        t = _tablo_basliga_gore(doc, tablo_no)
        if t is None:
            continue
        satirlar = _satirlari_uret(raf, tol)
        if not satirlar:
            continue
        idxler = _veri_satirlarini_ayarla(t, 1, len(satirlar))
        for ri, (sol, sag) in zip(idxler, satirlar):
            cells = t.rows[ri].cells
            hucre_yaz(cells[0], sol)
            hucre_yaz(cells[-1], sag)
        # Not: Tablo 9 '* Stabilite analizlerinde bakılmayacaktır.' notu şablonda
        # zaten mevcut; tekrar eklenmez.


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


def _turetilmis_testlerle(proje: ProjeVerisi):
    """
    Context manager: otomatik_turet açıksa, spek_karti.testler (bitmiş ürün
    listesi) kural motoruyla tüm aşamalara dağıtılır. Çıkışta orijinal liste
    geri yüklenir (kullanıcının girdiği liste bozulmaz).
    """
    import contextlib

    @contextlib.contextmanager
    def _ctx():
        kart = proje.spek_karti
        if not getattr(kart, "otomatik_turet", False):
            yield
            return
        from core.kural_motoru import turet
        orijinal = kart.testler
        ops = proje.urun_formu.operasyonlar
        kart.testler = turet(orijinal, kart.etkin_maddeler, ops,
                             cift_katman=getattr(kart, "cift_katman", False),
                             tablet_ipk=getattr(kart, "tablet_ipk", {}),
                             ozel_test_kurallari=getattr(kart, "ozel_test_kurallari", {}))
        try:
            yield
        finally:
            kart.testler = orijinal
    return _ctx()


def pvp_uret(proje: ProjeVerisi, cikti_yolu: str | Path) -> Path:
    with _turetilmis_testlerle(proje):
        doc = Document(str(_sablon_yolu("PVP_sablon.docx")))
        _ortak_doldur(doc, proje, rapor=False)
        yol = Path(cikti_yolu)
        doc.save(str(yol))
    return yol


def pvr_uret(proje: ProjeVerisi, cikti_yolu: str | Path, veri_uret: bool = True,
             tohum: int | None = None) -> Path:
    with _turetilmis_testlerle(proje):
        if veri_uret:
            vu.tum_testleri_uret(proje.spek_karti.testler, tohum=tohum)
        # PVR, TEMIZ PVP şablonundan üretilir; sonuç tabloları (Bölüm 11) temiz eklenir.
        doc = Document(str(_sablon_yolu("PVP_sablon.docx")))
        _ortak_doldur(doc, proje, rapor=True)
        _doldur_sonuclar(doc, proje)
        yol = Path(cikti_yolu)
        doc.save(str(yol))
    return yol
