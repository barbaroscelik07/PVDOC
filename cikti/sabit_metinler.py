"""
PVP/PVR şablonundaki SABİT metinler.

Bu bölümler her protokolde aynıdır (kullanıcı kararı: Stabilite ve giriş
bölümleri sabit). Ürün adı vb. yer tutucular {urun} ile işaretlenir ve
docx_motoru tarafından doldurulur.
"""

from __future__ import annotations


KISALTMALAR = [
    ("Ar-Ge", "Araştırma Geliştirme"),
    ("mg", "miligram"),
    ("g", "gram"),
    ("kg", "kilogram"),
    ("dk", "dakika"),
    ("min", "minimum"),
    ("maks.", "maksimum"),
    ("HPLC", "Yüksek Performans Sıvı Kromatografisi"),
    ("EP", "Avrupa Farmakopesi"),
    ("E", "evet"),
    ("H", "hayır"),
    ("No", "Numara"),
    ("U.Y.", "Uygulama Yoktur"),
    ("Tb", "Tablet"),
    ("Ftb", "Film Kaplı Tablet"),
    ("k.m.", "kafi miktarda"),
    ("T.E", "Tespit edilmedi"),
]

REFERANSLAR = [
    "EEC Guide to Good Manufacturing Practice for Medical Products – III/2244/87-EN, Rev. 3. January 1989 incl. Supplementary Guidelines",
    "U.S. FDA. SUPAC-IR/MR: Immediate release and modified release solid oral dosage forms manufacturing equipment addendum. Guide Indus Jan. (1999).",
    "U.S. FDA. Process Validation: General Principles and Practices. Guide Indus Jan. (2011).",
    "U.S. FDA. Guide to inspections of oral solid dosage forms pre/post approval issues for development and validation. Guide Indus Jan. (1994).",
    "U.S. FDA. Immediate release solid oral dosage forms scale-up and postapproval changes: Chemistry, manufacturing, and controls. In vitro dissolution testing, and in vivo bioequivalence documentation. Guide Indus Nov. (1995).",
    "ICH Harmonised tripartite guideline - good manufacturing practice guide for active pharmaceutical ingredients - Q7. (2000)",
    "EMEA - Note For Guidance On Process Validation - CPMP/QWP/848/96, EMEA/CVMP/598/99 (2001)",
]

AMAC_GIRIS = (
    "Bu validasyon protokolü, öngörülen üretim protokolü, tanımlanan sistem ve "
    "ekipmanlar ile üretilecek ürünün sürekli olarak öngörülen spesifikasyon "
    "aralığında üretiminin sağlanabileceğini kanıtlamak / dokümante etmek amacıyla "
    "oluşturulmuştur. Pilot Üretim Proses validasyonu {firma} İlaç üretim "
    "fabrikasında gerçekleştirilecektir."
)

AMAC_URUN = "{urun}'e ait üretim prosedürü Üretim Protokolü'nde detaylandırılmıştır."

AMAC_MADDELER = [
    "Risk analizi gerçekleştirerek ürünün kalite parametrelerini etkileyecek kritik proses basamaklarını belirlemek",
    "İzlenmesi ve kontrol edilmesi gereken kilit proses değişkenlerini belirlemek",
    "Önceden belirlenmiş proses parametrelerini/değişkenlerin datalarının sağlanabilmesi için uygun basamaklardan yeterli miktarda değerlendirecek bir numune alma planı geliştirmek",
    "Uygun istatistiksel araçlarla kabul kriterlerini belirlemek",
]

KAPSAM = "{urun} ürününe ait pilot üretim proses validasyon çalışması ardışık 3 seriye uygulanacaktır."

SORUMLULUKLAR = [
    ("Ar-Ge Departmanı", [
        "Pilot Üretim Proses Validasyon Protokolünün yazılması, kontrol edilmesi ve onaylanması",
        "Çalışmaların protokole uygun yürütülmesi",
        "Çalışma sırasında gereken numunelerin alınması, analizlerin yapılması ve raporlanması",
        "Pilot Üretim Proses Validasyon Raporunun hazırlanması, kontrol edilmesi ve onaylanması",
    ]),
    ("Kalite Güvence Departmanı", [
        "Pilot Üretim proses validasyon protokolünün, yürürlülük onayının verilmesi",
        "Pilot Üretim proses validasyon raporunun, yürürlülük onayının verilmesi",
    ]),
    ("Kalite Kontrol Departmanı", [
        "Mikrobiyolojik analizlerin ve Ar-Ge alt yapısında olmayan testlerin yapılması",
    ]),
]

STABILITE = (
    "Validasyon serilerinden alınan numuneler, ürünün raf ömrü boyunca stabilitesinin "
    "izlenmesi amacıyla stabilite çalışmasına alınacaktır. Stabilite çalışması, ürünün "
    "onaylı stabilite protokolüne uygun olarak yürütülür."
)
