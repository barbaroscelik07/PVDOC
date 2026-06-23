"""
Kural motoru — Bitmiş ürün spesifikasyonlarından (Tablo 8) Tablo 6/7/9 türetir.

Kullanıcı kararı: Sadece Tablo 8 (bitmiş ürün serbest bırakma) girilir; program
ürün formuna gömülü kurallarla testleri aşamalara dağıtır, yıldız/IPK atar.

FİLM TABLET test haritası (kullanıcı onaylı):
  Görünüş             : Karışım, Tablet Baskı, Film Kaplama (ayrı) — IPK
  Elek Testi          : Karışım — *
  Karışım Tekdüzeliği : Karışım (spek sabit %85–115) — *
  Ortalama Ağırlık    : Tablet Baskı, Film Kaplama (ayrı) — IPK
  Ağırlık Tekdüzeliği : Tablet Baskı, Film Kaplama (ayrı) — IPK
  Sertlik/Kalınlık/Çap/Aşınma/Dağılma : Tablet Baskı — IPK
  Teşhis              : tüm aşamalar
  Miktar Tayini       : tüm aşamalar
  Dissolüsyon         : Tablet Baskı, Film Kaplama
  İlgili Bileşikler   : tüm aşamalar (Film Kaplama'da * yok, diğerlerinde *)
  Mikrobiyolojik      : tüm aşamalar + Blisterleme (Blisterleme'de * yok)
"""

from __future__ import annotations

from core.models import (
    Test, Spesifikasyon, LimitTuru, TabloTipi, EtkinMadde, Impurite,
)

# Operasyon adı → numara
OP_NO = {"Karıştırma": 2, "Tablet Baskı": 3, "Film Kaplama": 4,
         "Dolum": 3, "Blisterleme": 5}


def _norm(s: str) -> str:
    tr = str.maketrans({"ı": "i", "İ": "i", "ş": "s", "Ş": "s", "ğ": "g",
                        "ç": "c", "ö": "o", "ü": "u", "I": "i"})
    return (s or "").translate(tr).lower().strip()


# Test anahtar kelimesi → kural
#   asamalar: bu testin görüneceği operasyonlar (None = forma göre tümü)
#   ipk: Tablo 7'ye de girer mi
#   yildiz_haric: bu operasyon(lar)da * KONMAZ; diğerlerinde konur
#   yildiz_hep: her zaman * (yildiz_haric ile birlikte değerlendirilir)
#   tip: sonuç tablosu tipi
def _kurallar(operasyonlar: list[str]) -> dict:
    KARISIM = "Karıştırma"
    TABLET = "Tablet Baskı"
    FILM = "Film Kaplama"
    BLISTER = "Blisterleme"
    var = lambda *adlar: [a for a in adlar if a in operasyonlar]

    return {
        "gorunus": dict(asamalar=var(KARISIM, TABLET, FILM), ipk=True, yildiz="yok",
                        tip=TabloTipi.TEK_SONUC),
        "elek": dict(asamalar=var(KARISIM), ipk=False, yildiz="hep",
                     tip=TabloTipi.TEK_SONUC),
        "karisim tekduzeligi": dict(asamalar=var(KARISIM), ipk=False, yildiz="hep",
                                    tip=TabloTipi.ON_NUMUNE, sabit_spek="%85 – %115"),
        "ortalama agirlik": dict(asamalar=var(TABLET, FILM), ipk=True, yildiz="yok",
                                 tip=TabloTipi.BOS_NOKTA),
        "agirlik tekduzeligi": dict(asamalar=var(TABLET, FILM), ipk=True, yildiz="yok",
                                    tip=TabloTipi.AGIRLIK_TEKDUZELIGI),
        "sertlik": dict(asamalar=var(TABLET), ipk=True, yildiz="yok", tip=TabloTipi.BOS_NOKTA),
        "kalinlik": dict(asamalar=var(TABLET), ipk=True, yildiz="yok", tip=TabloTipi.BOS_NOKTA),
        "cap": dict(asamalar=var(TABLET), ipk=True, yildiz="yok", tip=TabloTipi.BOS_NOKTA),
        "asinma": dict(asamalar=var(TABLET), ipk=True, yildiz="yok", tip=TabloTipi.BOS_NOKTA),
        "dagilma": dict(asamalar=var(TABLET), ipk=True, yildiz="yok", tip=TabloTipi.BOS_NOKTA),
        "teshis": dict(asamalar=var(KARISIM, TABLET, FILM), ipk=False, yildiz="yok",
                       tip=TabloTipi.TEK_SONUC),
        "miktar tayini": dict(asamalar=var(KARISIM, TABLET, FILM), ipk=False, yildiz="yok",
                              tip=TabloTipi.IKI_NUMUNE),
        "dissolusyon": dict(asamalar=var(TABLET, FILM), ipk=False, yildiz="yok",
                            tip=TabloTipi.BOS_NOKTA),
        "ilgili bilesik": dict(asamalar=var(KARISIM, TABLET, FILM), ipk=False,
                               yildiz="haric", yildiz_haric=[FILM], tip=TabloTipi.IKI_NUMUNE),
        "mikrobiyolojik": dict(asamalar=var(KARISIM, TABLET, FILM, BLISTER), ipk=False,
                               yildiz="haric", yildiz_haric=[BLISTER], tip=TabloTipi.MATRIS),
    }


def _kural_bul(ad: str, kurallar: dict):
    n = _norm(ad)
    # en spesifik eşleşmeler önce
    sirali = ["agirlik tekduzeligi", "karisim tekduzeligi", "ortalama agirlik",
              "ilgili bilesik", "miktar tayini", "mikrobiyolojik", "dissolusyon",
              "gorunus", "teshis", "elek", "sertlik", "kalinlik", "cap", "asinma", "dagilma"]
    for anahtar in sirali:
        if anahtar in n:
            return kurallar.get(anahtar)
    return None


def _yildiz_belirle(kural: dict, operasyon: str) -> bool:
    mod = kural.get("yildiz", "yok")
    if mod == "hep":
        return True
    if mod == "yok":
        return False
    if mod == "haric":
        return operasyon not in kural.get("yildiz_haric", [])
    return False


def testleri_turet(bitmis_testler: list[Test], operasyonlar: list[str],
                   etkin_maddeler: list[EtkinMadde]) -> list[Test]:
    """
    Bitmiş ürün test listesinden (Tablo 8) tüm aşamalara dağıtılmış Tablo 6
    test listesini üretir. Her test, kuralına göre ilgili operasyonlara kopyalanır;
    yıldız/IPK/tip otomatik atanır.
    """
    kurallar = _kurallar(operasyonlar)
    uretilen: list[Test] = []

    for bt in bitmis_testler:
        kural = _kural_bul(bt.ad, kurallar)
        if kural is None:
            # kuralsız test: olduğu gibi bırak (kullanıcı elle ayarlamış)
            uretilen.append(bt)
            continue
        asamalar = kural["asamalar"] or operasyonlar
        # test adından etkin madde önekini ayıkla (örn "Etkin madde 1 Görünüş")
        for op in asamalar:
            t = Test(
                ad=bt.ad,
                operasyon=op,
                operasyon_no=OP_NO.get(op, 0),
                tablo_tipi=kural["tip"],
                ipk=kural["ipk"],
                yildizli=_yildiz_belirle(kural, op),
                spesifikasyon=_spek_kopya(bt.spesifikasyon, kural),
                mikrobiyolojik=(kural["tip"] is TabloTipi.MATRIS),
                alt_satirlar=list(bt.alt_satirlar),
                aciklama_etiketi=bt.aciklama_etiketi,
                aciklama_spek=bt.aciklama_spek,
                aciklama2_etiketi=bt.aciklama2_etiketi,
                aciklama2_spek=bt.aciklama2_spek,
            )
            uretilen.append(t)

    # operasyon sırasına göre sırala (sabit sonuç düzeni için)
    uretilen.sort(key=lambda t: (t.operasyon_no or 99))
    return uretilen


def _spek_kopya(spek: Spesifikasyon, kural: dict) -> Spesifikasyon:
    """Spesifikasyonu kopyalar; sabit spek kuralı varsa onu uygular."""
    import copy
    yeni = copy.deepcopy(spek)
    if kural.get("sabit_spek"):
        yeni.spesifikasyon_metni = kural["sabit_spek"]
    return yeni
