"""
Tablo 8 (Bitmiş Ürün Serbest Bırakma Spesifikasyonları) Word okuyucu.

Kullanıcı tüm PVR/PVP Word dosyasını yükler; bu modül "Tablo 8" başlıklı
2 sütunlu (TESTLER | SPESİFİKASYONLAR) tabloyu bulur ve içeriği yapısal
olarak çözer:
  - Görünüş, Ortalama Ağırlık, Dağılma → tek test
  - Ağırlık Tekdüzeliği → 2 alt satır (sapabilir / sapmamalıdır)
  - Etkin madde başlıkları → o etkene ait Teşhis/Miktar Tayini
  - Dissolüsyon
  - İlgili Bileşikler → etkin madde 'e Ait' grupları + impuriteler
  - Mikrobiyolojik Kontrol → 3 alt satır
"""

from __future__ import annotations

from docx import Document
from docx.table import Table

from core.models import (
    Test, Spesifikasyon, LimitTuru, TabloTipi, EtkinMadde, Impurite,
)


def _norm(s: str) -> str:
    tr = str.maketrans({"ı": "i", "İ": "i", "ş": "s", "Ş": "s", "ğ": "g",
                        "ç": "c", "ö": "o", "ü": "u", "I": "i"})
    return (s or "").translate(tr).lower().strip()


def tablo8_bul(yol: str) -> Table | None:
    """Word dosyasında 'Tablo 8 ... Bitmiş Ürün' tablosunu bulur."""
    d = Document(yol)
    body = list(d.element.body)
    bulundu = False
    for el in body:
        if el.tag.endswith("}p"):
            txt = _norm("".join(el.itertext()))
            if "tablo" in txt and "8" in txt and "bitmis urun" in txt:
                bulundu = True
        elif el.tag.endswith("}tbl") and bulundu:
            return Table(el, d)
    # Yedek: 2 sütunlu, ilk satırı TESTLER/SPESİFİKASYON olan ilk tablo
    for tbl in d.tables:
        if len(tbl.columns) == 2 and tbl.rows:
            h0 = _norm(tbl.rows[0].cells[0].text)
            h1 = _norm(tbl.rows[0].cells[1].text)
            if "test" in h0 and "spesifikasyon" in h1:
                return tbl
    return None


def _sayi(metin: str):
    import re
    m = re.search(r"(\d+(?:[.,]\d+)?)", metin or "")
    return float(m.group(1).replace(",", ".")) if m else None


def tablo8_coz(yol: str) -> dict:
    """
    Tablo 8'i çözüp şu yapıyı döndürür:
      {
        "testler": [Test, ...],          # bitmiş ürün testleri (etken adı gömülü)
        "etkin_maddeler": [EtkinMadde],  # impurite grupları ile
        "bulundu": bool,
      }
    """
    t = tablo8_bul(yol)
    if t is None:
        return {"bulundu": False, "testler": [], "etkin_maddeler": []}

    testler: list[Test] = []
    etkenler: dict[str, EtkinMadde] = {}
    aktif_etken_imp = None     # İlgili Bileşikler altında aktif etken
    ilgili_bolumde = False
    mikro_test = None
    agirlik_test = None
    aktif_etken_adi = None     # "Etkin madde 1" başlığı altındayız

    def _etken(ad):
        if ad not in etkenler:
            etkenler[ad] = EtkinMadde(ad=ad)
        return etkenler[ad]

    rows = t.rows[1:]  # başlık satırını atla
    for row in rows:
        sol = row.cells[0].text.strip()
        sag = row.cells[1].text.strip()
        n = _norm(sol)
        if not sol and not sag:
            continue

        # --- Bölüm başlıkları ---
        if n == "ilgili bilesikler" or n.startswith("ilgili bilesik"):
            ilgili_bolumde = True
            aktif_etken_imp = None
            continue

        if ilgili_bolumde:
            # "Etkin madde 1'e Ait" → grup başlığı
            if "e ait" in n or "a ait" in n or "'e ait" in sol.lower() or "'a ait" in sol.lower():
                ad = sol.split("'")[0].split("’")[0].strip()
                if not ad:
                    ad = sol.replace("e Ait", "").replace("a Ait", "").strip()
                aktif_etken_imp = _etken(ad)
                continue
            # impurite satırı
            if aktif_etken_imp is not None and sol.startswith(("—", "-", "–")):
                ad = sol.lstrip("—-– ").strip()
                te = "t.e" in _norm(sag)
                aktif_etken_imp.impuriteler.append(Impurite(
                    ad=ad, limit_metni=sag, maksimum_deger=_sayi(sag), te=te))
                continue
            # İlgili Bileşikler bölümü bitmiş olabilir (mikrobiyolojik başladı)
            if "mikrobiyolojik" in n:
                ilgili_bolumde = False
                # aşağıda mikro işlenecek

        # --- Mikrobiyolojik ---
        if "mikrobiyolojik" in n:
            ilgili_bolumde = False
            mikro_test = Test(
                ad="Mikrobiyolojik Kontrol", tablo_tipi=TabloTipi.MATRIS,
                mikrobiyolojik=True,
                spesifikasyon=Spesifikasyon(limit_turu=LimitTuru.METIN, sabit_sonuc="Uygun"),
                alt_satirlar=[])
            testler.append(mikro_test)
            continue
        if mikro_test is not None and sol.startswith(("—", "-", "–")) and not ilgili_bolumde:
            mikro_test.alt_satirlar.append((sol, sag))
            continue

        # --- Ağırlık Tekdüzeliği (başlık + 2 alt) ---
        if "agirlik tekduzeligi" in n and not sag:
            agirlik_test = Test(
                ad="Ağırlık Tekdüzeliği", tablo_tipi=TabloTipi.AGIRLIK_TEKDUZELIGI,
                spesifikasyon=Spesifikasyon(limit_turu=LimitTuru.ARALIK, birim="mg"))
            testler.append(agirlik_test)
            continue
        if agirlik_test is not None and sol.startswith(("—", "-", "–")):
            if not agirlik_test.aciklama_etiketi:
                agirlik_test.aciklama_etiketi = sol
                agirlik_test.aciklama_spek = sag
            else:
                agirlik_test.aciklama2_etiketi = sol
                agirlik_test.aciklama2_spek = sag
                # alt/üst limiti açıklamadan türet
                sayilar = _alt_ust_metinden(sag) or _alt_ust_metinden(agirlik_test.aciklama_spek)
            continue
        else:
            if agirlik_test is not None and not sol.startswith(("—", "-", "–")):
                agirlik_test = None  # ağırlık bölümü bitti

        # --- Etkin madde başlığı ("Etkin madde 1") ---
        if n.startswith("etkin madde") and not sag and "ilgili" not in n and "dissol" not in n:
            aktif_etken_adi = sol.strip()
            continue

        # --- Alt satır: Teşhis / Miktar Tayini (aktif etken altında) ---
        if sol.startswith(("—", "-", "–")) and aktif_etken_adi:
            alt_ad = sol.lstrip("—-– ").strip()
            tam_ad = f"{aktif_etken_adi} {alt_ad}"
            tip = TabloTipi.IKI_NUMUNE if "miktar" in _norm(alt_ad) else TabloTipi.TEK_SONUC
            testler.append(_test_yap(tam_ad, sag, tip))
            continue

        # --- Düz testler (Görünüş, Ortalama Ağırlık, Dağılma, Dissolüsyon) ---
        if sol and not sol.startswith(("—", "-", "–")):
            aktif_etken_adi = None
            tip = _tip_tahmin(n)
            testler.append(_test_yap(sol, sag, tip))

    return {"bulundu": True, "testler": testler,
            "etkin_maddeler": list(etkenler.values())}


def _alt_ust_metinden(metin):
    import re
    s = re.sub(r"(\d),(\d)", r"\1.\2", metin or "")
    return [float(x) for x in re.findall(r"\d+(?:\.\d+)?", s)]


def _tip_tahmin(n: str) -> TabloTipi:
    if "ortalama agirlik" in n:
        return TabloTipi.BOS_NOKTA
    if "dagilma" in n or "dissol" in n:
        return TabloTipi.BOS_NOKTA
    if "miktar" in n:
        return TabloTipi.IKI_NUMUNE
    return TabloTipi.TEK_SONUC


def _test_yap(ad: str, spek_metni: str, tip: TabloTipi) -> Test:
    sp = Spesifikasyon(limit_turu=LimitTuru.ARALIK, spesifikasyon_metni=spek_metni)
    # alt/üst limiti metinden türet
    sayilar = _alt_ust_metinden(spek_metni)
    if len(sayilar) >= 2:
        sp.alt_limit = sayilar[-2]
        sp.ust_limit = sayilar[-1]
        sp.alt_metin = str(sayilar[-2])
        sp.ust_metin = str(sayilar[-1])
        sp.hedef_deger = sayilar[0]
    elif len(sayilar) == 1:
        if "maksimum" in _norm(spek_metni) or "max" in _norm(spek_metni):
            sp.ust_limit = sayilar[0]; sp.maksimum_deger = sayilar[0]
        elif "minimum" in _norm(spek_metni) or "min" in _norm(spek_metni):
            sp.alt_limit = sayilar[0]; sp.minimum_deger = sayilar[0]
    # metinsel testler
    if _norm(ad).endswith("gorunus") or "teshis" in _norm(ad):
        sp.limit_turu = LimitTuru.METIN
        sp.sabit_sonuc = spek_metni
    return Test(ad=ad, tablo_tipi=tip, spesifikasyon=sp)
