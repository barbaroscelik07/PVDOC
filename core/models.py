"""
PV-DOC merkezi veri modeli.

Tüm modüllerin okuyup yazdığı tek kaynak: ProjeVerisi.
JSON olarak diske kaydedilir / yüklenir (bkz. core/proje_io.py).

Tasarım ilkeleri:
- Her dataclass JSON-serileştirilebilir (sadece basit tipler + iç içe dataclass/list).
- to_dict() / from_dict() ile sözlüğe çevrim. (json modülü doğrudan dataclass bilmez.)
- Sabitler: seri sayısı = 3, nokta (Baş/Orta/Son) = 3.
"""

from __future__ import annotations

from dataclasses import dataclass, field, asdict, fields, is_dataclass
from enum import Enum
from typing import Any, Optional


# ----------------------------------------------------------------------------
# Sabitler
# ----------------------------------------------------------------------------

SERI_SAYISI = 3
NOKTA_SAYISI = 3  # Baş / Orta / Son
NOKTA_ADLARI = ("Baş", "Orta", "Son")

UYGULAMA_YOK = "U.Y."
TESPIT_EDILMEDI = "T.E."


# ----------------------------------------------------------------------------
# Enum'lar
# ----------------------------------------------------------------------------

class UrunFormu(str, Enum):
    """Ürün dozaj formu. Tablo 6/7/10 ve sonuç operasyonlarını forma göre süzer."""
    TABLET = "Tablet"
    FILM_TABLET = "Film Kaplı Tablet"
    KAPSUL = "Kapsül"

    @property
    def operasyonlar(self) -> list[str]:
        """Bu ürün formunda yer alan üretim operasyonları (sonuç bölümü için)."""
        if self is UrunFormu.TABLET:
            return ["Karıştırma", "Tablet Baskı", "Blisterleme"]
        if self is UrunFormu.FILM_TABLET:
            return ["Karıştırma", "Tablet Baskı", "Film Kaplama", "Blisterleme"]
        if self is UrunFormu.KAPSUL:
            return ["Karıştırma", "Dolum", "Blisterleme"]
        return []


class TabloTipi(str, Enum):
    """
    PVR sonuç tablosu tipleri. Test adından otomatik belirlenir
    (bkz. core/test_tipi.py). Hangi UI/render şablonunun kullanılacağını söyler.
    """
    TEK_SONUC = "tek_sonuc"            # Görünüş, Teşhis, Aşınma → 3 seri × tek değer
    IKI_NUMUNE = "iki_numune"          # Miktar Tayini, İmpurite → Numune-1/2 + Sonuç(ort)
    ON_NUMUNE = "on_numune"            # Karışım Tekdüzeliği → 1..10 + Ortalama
    AGIRLIK_TEKDUZELIGI = "agirlik_tekduzeligi"  # 20 numune × 3 nokta + Ort/RSD/SD
    BOS_NOKTA = "bos_nokta"            # Sertlik/Kalınlık/Çap/Dağılma/Dissolüsyon → n×3 nokta + Sonuç
    MATRIS = "matris"                  # Mikrobiyolojik / Sızdırmazlık → çok satırlı spek + matris


class LimitTuru(str, Enum):
    """Spesifikasyon limitinin yapısı — hangi alanların anlamlı olduğunu belirler."""
    ARALIK = "aralik"          # alt-üst aralık: "4.75 – 5.25"
    MINIMUM = "minimum"        # "Minimum %80.0"
    MAKSIMUM = "maksimum"      # "Maksimum %1.0"
    METIN = "metin"            # serbest metin: "Beyaz renkli toz", "Pozitif"
    BILGI = "bilgi"            # "Bilgi amaçlıdır." (limitsiz, sadece kayıt)


# ----------------------------------------------------------------------------
# Spesifikasyon / Test modelleri
# ----------------------------------------------------------------------------

@dataclass
class Spesifikasyon:
    """
    Bir testin kabul kriteri.

    Kullanıcının spek formunda girdiği değerler buraya yazılır. Program PVR'de
    simüle/gerçek sonuç üretirken bu sınırları kullanır:
      - ARALIK: alt_limit / ust_limit arasında değer üretilir.
      - MINIMUM: minimum_deger üstünde değer üretilir.
      - MAKSIMUM: maksimum_deger altında değer üretilir.
      - METIN/BILGI: sabit_sonuc kullanılır.
    """
    limit_turu: LimitTuru = LimitTuru.METIN

    # Sayısal limitler (veri üretimi için; LimitTuru'na göre anlamlı olanlar dolar)
    hedef_deger: Optional[float] = None
    alt_limit: Optional[float] = None
    ust_limit: Optional[float] = None
    minimum_deger: Optional[float] = None
    maksimum_deger: Optional[float] = None

    # Kullanıcının GİRDİĞİ ham metin (ondalık birebir korunur: "5,0" -> "5,0").
    # Görüntüde bunlar kullanılır; sayısal alanlar yalnızca veri üretimi içindir.
    hedef_metin: str = ""
    alt_metin: str = ""
    ust_metin: str = ""
    minimum_metin: str = ""
    maksimum_metin: str = ""

    tolerans: str = ""            # örn. "±%5" — Miktar Tayini vb. için
    birim: str = ""               # "mg/f.tab", "%", "kP", "dakika", "mm" ...
    ondalik: int = 2              # üretilen SONUÇ verisinin ondalık basamağı

    # Spesifikasyon hücresinde gösterilecek hazır metin. Boşsa otomatik biçimlenir.
    spesifikasyon_metni: str = ""

    # Metin tipli testler için sabit sonuç (örn. "Beyaz renkli toz", "Uygun", "Pozitif")
    sabit_sonuc: str = ""

    def _g(self, metin: str, sayi: Optional[float]) -> str:
        """Ham metin doluysa onu, değilse sayıyı döndürür (ondalık koruma)."""
        if metin.strip():
            return metin.strip()
        return "" if sayi is None else f"{sayi:g}"

    def metni_olustur(self) -> str:
        """
        Spesifikasyon hücre metnini üretir. Ham metin alanları varsa onları
        kullanır (kullanıcının girdiği ondalık birebir korunur).
        Tolerans varsa hedef ile birlikte gösterilir:
          "5,0 mg/f.tab ±%5 (4,75 – 5,25 mg/f.tab)"
        """
        if self.spesifikasyon_metni:
            return self.spesifikasyon_metni
        b = f" {self.birim}".rstrip()
        tol = f" {self.tolerans}".rstrip() if self.tolerans.strip() else ""

        if self.limit_turu is LimitTuru.ARALIK:
            hedef = self._g(self.hedef_metin, self.hedef_deger)
            alt = self._g(self.alt_metin, self.alt_limit)
            ust = self._g(self.ust_metin, self.ust_limit)
            if hedef:
                aralik = f" ({alt} – {ust}{b})" if (alt or ust) else ""
                return f"{hedef}{b}{tol}{aralik}"
            return f"{alt} – {ust}{b}"
        if self.limit_turu is LimitTuru.MINIMUM:
            return f"Minimum {self._g(self.minimum_metin, self.minimum_deger)}{b}"
        if self.limit_turu is LimitTuru.MAKSIMUM:
            return f"Maksimum {self._g(self.maksimum_metin, self.maksimum_deger)}{b}"
        if self.limit_turu is LimitTuru.BILGI:
            return "Bilgi amaçlıdır."
        return self.sabit_sonuc


@dataclass
class Test:
    """
    Tek bir analiz testi (spek kartının ve sonuç tablolarının yapı taşı).

    operasyon: "Karıştırma" / "Tablet Baskı" / "Film Kaplama" / "Blisterleme" / "Dolum"
    operasyon_no: şablondaki sayısal operasyon numarası (2,3,4,5).
    """
    ad: str                                       # "Etkin madde 1 Miktar Tayini"
    operasyon: str = ""                           # üretim operasyonu adı
    operasyon_no: int = 0
    spesifikasyon: Spesifikasyon = field(default_factory=Spesifikasyon)
    tablo_tipi: TabloTipi = TabloTipi.TEK_SONUC   # otomatik atanır, override edilebilir

    ipk: bool = False                 # Tablo 7 (IPK Testleri) bu bayraktan süzülür
    yildizli: bool = False            # "*": proses validasyonu serilerinde uygulanır
    etkin_madde_index: int = -1       # hangi etkin maddeye ait (-1 = ürüne ait genel)

    # Alt satırlar: ana test başlık satırının altına gelen, kendi spesifikasyonu
    # olan satırlar. Örn. Mikrobiyolojik Kontrol:
    #   ("-Toplam Aerobik Mikroorganizma Sayısı", "≤10³ cfu/g")
    # Her öğe (etiket, spek_metni).
    alt_satirlar: list[tuple] = field(default_factory=list)

    # Açıklama satırı: ana başlığın altına gelen, sadece açıklama içeren satır.
    # Örn. Ağırlık Tekdüzeliği:
    #   ("—20 tablette tek tek tabletlerden maksimum 2 tanesi bu limitten sapabilir.",
    #    "≤ 270.75 veya ≥ 299.25 mg")
    aciklama_etiketi: str = ""
    aciklama_spek: str = ""
    aciklama2_etiketi: str = ""
    aciklama2_spek: str = ""

    # Bu test özel bir kalıp mı (mikrobiyolojik)? Sonuç hep "Uygun".
    mikrobiyolojik: bool = False

    # Üretilen ham sonuç verisi (PVR). Yapı tablo_tipi'ne göre değişir.
    sonuc_verisi: dict[str, Any] = field(default_factory=dict)


# ----------------------------------------------------------------------------
# Formül / etkin madde modelleri
# ----------------------------------------------------------------------------

@dataclass
class Impurite:
    """Bir etkin maddeye ait safsızlık (İlgili Bileşikler tablosu satırı)."""
    ad: str                           # "imp. a", "Sülfoksit imp.", "Toplam imp."
    limit_metni: str = ""             # "Maksimum %1.0"
    maksimum_deger: Optional[float] = None
    operasyon: str = ""               # impuritenin ölçüldüğü operasyon (grup için)
    operasyon_no: int = 0
    yildizli: bool = False            # bu grup * (validasyon) mı
    te: bool = False                  # sonuçta T.E. (tespit edilemedi) mi


@dataclass
class Hammadde:
    """Tablo 1 (Birim ve Seri Formül) satırı."""
    ad: str
    fonksiyon: str = ""               # "Etkin madde", "Dolgu Ajanı", "Lubrikant" ...
    birim_formul: Optional[float] = None   # mg/tb
    yuzde_icerik: Optional[float] = None
    seri_miktar: Optional[float] = None    # kg / {adet}
    katman: int = 0                   # çift katman tablet için: 0=yok, 1=Katman I, 2=Katman II
    ara_toplam: bool = False          # "Katman I Ağırlık" gibi toplam satırı mı


@dataclass
class EtkinMadde:
    """Üründeki bir etkin madde. Kullanıcı dinamik olarak ekler/çıkarır."""
    ad: str                           # "Etkin madde 1" veya gerçek INN
    impuriteler: list[Impurite] = field(default_factory=list)
    # Enantiomerik impurite alt satırları (İlgili Bileşikler ile aynı yapı).
    # Örn. [Impurite(ad='Linezolid R-İzomer', maksimum_deger=0.3, limit_metni='Maksimum % 0.3')]
    enantiomerik: list[Impurite] = field(default_factory=list)


# ----------------------------------------------------------------------------
# Üretim yöntemi (proses tanımı) modelleri
# ----------------------------------------------------------------------------

@dataclass
class ParametreSatiri:
    """Üretim aşaması altındaki mini tablo satırı (örn. 'Elek açıklığı | 0,8 mm')."""
    etiket: str
    deger: str


@dataclass
class Asama:
    """Bir operasyon aşaması: serbest metin + opsiyonel mini parametre tablosu."""
    operasyon_no: int
    asama_no: int
    metin: str = ""
    parametreler: list[ParametreSatiri] = field(default_factory=list)
    ipk_etiketi: str = ""             # kullanıcı girer: "IPK-1", "IPK-2" ...


# ----------------------------------------------------------------------------
# Ekipman / risk analizi / numune planı
# ----------------------------------------------------------------------------

@dataclass
class Ekipman:
    """Tablo 5 satırı."""
    operasyon_no: int
    operasyon: str
    ekipman_adi: str
    kapasite: str = ""


@dataclass
class RiskSatiri:
    """Tablo 3 (Kritik/Kritik olmayan parametre değerlendirmesi) satırı."""
    operasyon_no: int
    operasyon: str
    kritik: bool = False              # E (Evet) / H (Hayır)
    testler: str = ""
    yorumlar: str = ""


@dataclass
class ProsesParametresi:
    """Tablo 4 (Öngörülen Proses Parametreleri) satırı."""
    aciklama: str                     # "Operasyon 2: Aşama 8"
    parametre: str = ""               # "Karıştırma Süresi"
    deger: str = ""                   # "10 dk"


@dataclass
class NumuneAlmaSatiri:
    """Tablo 10 satırı (forma göre süzülür)."""
    operasyon_no: int
    operasyon: str
    numune_noktasi: str = ""          # "Baş, Orta, Son" veya "1,2,...,10"
    toplam_miktar: str = ""


# ----------------------------------------------------------------------------
# Spek kartı (kütüphaneye kaydedilebilir, yeniden kullanılabilir birim)
# ----------------------------------------------------------------------------

@dataclass
class SpekKarti:
    """
    Bir ürünün tüm spesifikasyon setini taşıyan, kütüphaneye kaydedilip
    yeniden kullanılabilen birim. Bir kez gir → kaydet → her projede çağır.

    Hem PVP (Tablo 6/7/8/9) hem PVR (sonuç tabloları) bu karttan beslenir.
    """
    kart_adi: str                     # kütüphanede görünen ad
    urun_formu: UrunFormu = UrunFormu.FILM_TABLET
    etkin_maddeler: list[EtkinMadde] = field(default_factory=list)
    testler: list[Test] = field(default_factory=list)

    # Tablo 8 (Serbest Bırakma) / Tablo 9 (Raf Ömrü) otomatik üretilsin mi?
    # Miktar Tayini toleransı bu tablolarda farklıdır; aşağıdaki toleranslar kullanılır.
    tablo89_ekle: bool = False
    serbest_birakma_tolerans: str = "±%5"     # Tablo 8 miktar tayini toleransı
    raf_omru_tolerans: str = "±%7.5"          # Tablo 9 miktar tayini toleransı

    # Açıksa: testler[] BİTMİŞ ÜRÜN listesi kabul edilir; çıktıda kural motoruyla
    # tüm aşamalara dağıtılır (Tablo 6/7/9 otomatik türetilir).
    otomatik_turet: bool = True

    # Çift katmanlı tablet mi? Çift ise Karışım'da Görünüş/Elek/Bulk-Tap her etken
    # için ayrı; tek katmanda bunlar birleşik (tek satır).
    cift_katman: bool = False

    # Tablet-only IPK spesifikasyonları (Tablo 8'de YOK, Tablo 6 Tablet'te VAR).
    # Kullanıcı girer; örn. {"Kalınlık": "4.75 mm (4.45 – 5.05 mm)", ...}
    tablet_ipk: dict = field(default_factory=dict)

    # Tanımsız testler için kullanıcının verdiği kurallar (Word yüklemede sorulur).
    # {test_adı: {"asamalar": [...], "yildiz": [...], "ipk": bool, "spek": str}}
    ozel_test_kurallari: dict = field(default_factory=dict)

    # Eski alanlar (manuel/kopyala-yapıştır) — ileride kullanılabilir
    serbest_birakma: list[dict[str, str]] = field(default_factory=list)
    raf_omru: list[dict[str, str]] = field(default_factory=list)


# ----------------------------------------------------------------------------
# Diyagram (ayrı pencere serbest tuval) — model JSON olarak proje içinde saklanır
# ----------------------------------------------------------------------------

@dataclass
class DiyagramDugumu:
    """Akış diyagramı düğümü (dikdörtgen, karar elması vb.)."""
    id: int
    tip: str = "dikdortgen"           # "dikdortgen" | "karar" | "baslangic"
    metin: str = ""
    x: float = 0.0
    y: float = 0.0
    en: float = 160.0
    boy: float = 60.0


@dataclass
class DiyagramOku:
    """İki düğüm arasındaki yönlü bağlantı."""
    kaynak_id: int
    hedef_id: int
    etiket: str = ""


@dataclass
class Diyagram:
    """Proses akış diyagramı (Bölüm 5.1). render → PNG, model → JSON."""
    dugumler: list[DiyagramDugumu] = field(default_factory=list)
    oklar: list[DiyagramOku] = field(default_factory=list)


# ----------------------------------------------------------------------------
# Doküman üst bilgileri
# ----------------------------------------------------------------------------

@dataclass
class DokumanBilgisi:
    """Şablon başlığı / alt bilgi alanları. Doküman no formatı firmaya özgü, kullanıcı girer."""
    firma_ismi: str = ""
    urun_adi: str = ""                # "Xxx Film Kaplı Tablet"
    pvp_dokuman_no: str = ""          # "AG-PV-xxx"
    pvr_dokuman_no: str = ""          # "AG-PV-xxx-R"
    revizyon_no: str = "03"           # otomatik 03; kullanıcı değiştirebilir
    revizyon_tarihi: str = UYGULAMA_YOK
    # PVP ve PVR için AYRI form numaraları. Otomatik varsayılanlar; kullanıcı
    # genel bilgi sayfasından değiştirebilir.
    pvp_form_no: str = "N-15-506"
    pvr_form_no: str = "N-15-507"
    form_no: str = ""                 # (eski tek alan; geriye dönük uyumluluk)


@dataclass
class Seri:
    """Kapsanan ürün serisi (Tablo 2). Sabit 3 adet."""
    urun_ismi: str = ""
    seri_no: str = ""                 # "yyy-P01"
    seri_boyutu_adet: str = ""        # "150.000"
    seri_boyutu_kg: str = ""          # "43.500"


# ----------------------------------------------------------------------------
# KÖK: ProjeVerisi
# ----------------------------------------------------------------------------

@dataclass
class ProjeVerisi:
    """
    Tüm modüllerin okuyup yazdığı merkezi veri sınıfı.
    JSON olarak kaydedilip yüklenir; aynı anda tek proje açık olur.
    """
    sema_versiyonu: int = 1

    dokuman: DokumanBilgisi = field(default_factory=DokumanBilgisi)
    urun_formu: UrunFormu = UrunFormu.FILM_TABLET

    seriler: list[Seri] = field(default_factory=lambda: [Seri() for _ in range(SERI_SAYISI)])

    # Bölüm 5
    hammaddeler: list[Hammadde] = field(default_factory=list)
    asamalar: list[Asama] = field(default_factory=list)
    diyagram: Diyagram = field(default_factory=Diyagram)

    # Bölüm 6/7
    risk_satirlari: list[RiskSatiri] = field(default_factory=list)
    proses_parametreleri: list[ProsesParametresi] = field(default_factory=list)

    # Üretim yöntemi adımları: [(operasyon_baslik, aciklama), ...]
    # Word'den yüklenir; örn. ("Operasyon 2: Aşama 1", "6.750 kg ... tartılır.")
    uretim_adimlari: list = field(default_factory=list)
    ekipmanlar: list[Ekipman] = field(default_factory=list)

    # Bölüm 8 — spek kartından gelir
    spek_karti: SpekKarti = field(default_factory=lambda: SpekKarti(kart_adi=""))

    # Bölüm 9
    numune_plani: list[NumuneAlmaSatiri] = field(default_factory=list)

    # Bölüm 12 — Genel değerlendirme (PVR)
    sapmalar: str = "Sapma gözlenmemiştir."
    sapma_gerekce: str = UYGULAMA_YOK
    sonuc_degerlendirme: str = "Belirtilen spesifikasyonlara uygun şekilde sonuçlar elde edilmiştir."
    yorum: str = ""

    # ---- serileştirme ----
    def to_dict(self) -> dict[str, Any]:
        return _to_jsonable(self)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "ProjeVerisi":
        return _from_dict(cls, data)


# ----------------------------------------------------------------------------
# Serileştirme yardımcıları (Enum + dataclass'ı güvenli JSON'a çevirir)
# ----------------------------------------------------------------------------

def _to_jsonable(obj: Any) -> Any:
    """Dataclass/Enum/list/dict ağacını JSON uyumlu yapıya çevirir."""
    if is_dataclass(obj) and not isinstance(obj, type):
        return {f.name: _to_jsonable(getattr(obj, f.name)) for f in fields(obj)}
    if isinstance(obj, Enum):
        return obj.value
    if isinstance(obj, list):
        return [_to_jsonable(x) for x in obj]
    if isinstance(obj, dict):
        return {k: _to_jsonable(v) for k, v in obj.items()}
    return obj


# Enum alanlarını from_dict sırasında doğru tipe çevirmek için kayıt
_ENUM_ALANLARI = {
    "urun_formu": UrunFormu,
    "limit_turu": LimitTuru,
    "tablo_tipi": TabloTipi,
}


def _from_dict(cls: type, data: Any) -> Any:
    """JSON sözlüğünü dataclass ağacına geri çevirir (iç içe yapıları çözer)."""
    if not is_dataclass(cls):
        return data
    if data is None:
        return cls()

    kwargs: dict[str, Any] = {}
    type_hints = {f.name: f.type for f in fields(cls)}

    for f in fields(cls):
        if f.name not in data:
            continue
        ham = data[f.name]
        kwargs[f.name] = _coerce_field(f.name, type_hints[f.name], ham)

    return cls(**kwargs)


def _coerce_field(ad: str, tip: Any, ham: Any) -> Any:
    """Tek bir alanı tip ipucuna göre uygun nesneye dönüştürür."""
    # Enum alanı
    if ad in _ENUM_ALANLARI and isinstance(ham, str):
        return _ENUM_ALANLARI[ad](ham)

    # list[...] alanı
    if isinstance(ham, list):
        ic_tip = _liste_eleman_tipi(tip)
        if ic_tip is not None and is_dataclass(ic_tip):
            return [_from_dict(ic_tip, x) for x in ham]
        return ham

    # iç içe dataclass
    hedef = _tip_coz(tip)
    if hedef is not None and is_dataclass(hedef) and isinstance(ham, dict):
        return _from_dict(hedef, ham)

    return ham


def _tip_coz(tip: Any) -> Optional[type]:
    """'TipAdı' (string forward-ref) veya gerçek tipi sınıfa çevirir."""
    if isinstance(tip, type):
        return tip
    if isinstance(tip, str):
        return _GLOBAL_TIPLER.get(tip.split("[")[0].strip())
    return None


def _liste_eleman_tipi(tip: Any) -> Optional[type]:
    """'list[Hammadde]' gibi bir ipucundan eleman tipini çıkarır."""
    s = tip if isinstance(tip, str) else getattr(tip, "__name__", str(tip))
    if "[" in s and "]" in s:
        ic = s[s.index("[") + 1:s.rindex("]")].strip()
        return _GLOBAL_TIPLER.get(ic)
    return None


# from_dict'in forward-ref string'lerini çözebilmesi için tip kayıt defteri
_GLOBAL_TIPLER: dict[str, type] = {
    "DokumanBilgisi": DokumanBilgisi,
    "Seri": Seri,
    "Hammadde": Hammadde,
    "Asama": Asama,
    "ParametreSatiri": ParametreSatiri,
    "Diyagram": Diyagram,
    "DiyagramDugumu": DiyagramDugumu,
    "DiyagramOku": DiyagramOku,
    "RiskSatiri": RiskSatiri,
    "ProsesParametresi": ProsesParametresi,
    "Ekipman": Ekipman,
    "SpekKarti": SpekKarti,
    "EtkinMadde": EtkinMadde,
    "Impurite": Impurite,
    "Test": Test,
    "Spesifikasyon": Spesifikasyon,
    "NumuneAlmaSatiri": NumuneAlmaSatiri,
}
