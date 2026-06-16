"""
Proje kaydetme / yükleme ve oturum sürdürme.

- ProjeVerisi <-> .pvdoc dosyası (JSON) arası dönüşüm.
- "Son oturum" yolu bir ayar dosyasında tutulur; uygulama açılışta
  kaldığı projeyi otomatik geri yükleyebilir.
- Aynı anda tek proje açık olur (tasarım kararı).

Dosya uzantısı: .pvdoc  (içerik düz JSON)
"""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Optional

from core.models import ProjeVerisi


PROJE_UZANTISI = ".pvdoc"

# Ayar/oturum dosyası kullanıcının ev dizininde tutulur (terminal erişimi gerekmez).
_AYAR_DIZINI = Path.home() / ".pvdoc"
_OTURUM_DOSYASI = _AYAR_DIZINI / "oturum.json"


# ----------------------------------------------------------------------------
# Kaydet / Yükle
# ----------------------------------------------------------------------------

def projeyi_kaydet(proje: ProjeVerisi, yol: str | os.PathLike) -> Path:
    """
    Projeyi verilen yola JSON olarak yazar. Uzantı eksikse .pvdoc eklenir.
    Atomik yazım: önce .tmp dosyaya yazıp sonra taşır (yazım sırasında
    çökme olursa eski dosya bozulmaz).
    """
    p = Path(yol)
    if p.suffix.lower() != PROJE_UZANTISI:
        p = p.with_suffix(PROJE_UZANTISI)

    p.parent.mkdir(parents=True, exist_ok=True)
    tmp = p.with_suffix(p.suffix + ".tmp")

    veri = proje.to_dict()
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(veri, f, ensure_ascii=False, indent=2)
    os.replace(tmp, p)  # atomik

    _son_oturumu_yaz(p)
    return p


def projeyi_yukle(yol: str | os.PathLike) -> ProjeVerisi:
    """Verilen .pvdoc dosyasını okuyup ProjeVerisi döndürür."""
    p = Path(yol)
    if not p.exists():
        raise FileNotFoundError(f"Proje dosyası bulunamadı: {p}")
    with open(p, "r", encoding="utf-8") as f:
        veri = json.load(f)

    _sema_gocu(veri)  # ileride versiyon farkı olursa burada düzeltilir
    proje = ProjeVerisi.from_dict(veri)
    _son_oturumu_yaz(p)
    return proje


# ----------------------------------------------------------------------------
# Son oturum (kaldığın yerden devam)
# ----------------------------------------------------------------------------

def son_oturum_yolu() -> Optional[Path]:
    """En son açılan/kaydedilen projenin yolunu döndürür (yoksa None)."""
    if not _OTURUM_DOSYASI.exists():
        return None
    try:
        with open(_OTURUM_DOSYASI, "r", encoding="utf-8") as f:
            d = json.load(f)
        yol = d.get("son_proje")
        if yol and Path(yol).exists():
            return Path(yol)
    except (json.JSONDecodeError, OSError):
        return None
    return None


def son_oturumu_geri_yukle() -> Optional[ProjeVerisi]:
    """Son projeyi otomatik yükler; yoksa veya bozuksa None döndürür."""
    yol = son_oturum_yolu()
    if yol is None:
        return None
    try:
        return projeyi_yukle(yol)
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return None


def _son_oturumu_yaz(yol: Path) -> None:
    """Son açılan proje yolunu oturum ayarına kaydeder."""
    try:
        _AYAR_DIZINI.mkdir(parents=True, exist_ok=True)
        with open(_OTURUM_DOSYASI, "w", encoding="utf-8") as f:
            json.dump({"son_proje": str(yol.resolve())}, f, ensure_ascii=False)
    except OSError:
        # Oturum hatırlama kritik değil; başarısızsa sessizce geç.
        pass


def oturumu_temizle() -> None:
    """Son oturum kaydını siler (örn. yeni boş projeye geçişte)."""
    try:
        if _OTURUM_DOSYASI.exists():
            _OTURUM_DOSYASI.unlink()
    except OSError:
        pass


# ----------------------------------------------------------------------------
# Şema göçü (ileri uyumluluk)
# ----------------------------------------------------------------------------

def _sema_gocu(veri: dict) -> None:
    """
    Eski sürümde kaydedilmiş projeleri güncel şemaya taşır.
    Şimdilik tek sürüm (1) var; sonraki sürümlerde buraya dönüşüm eklenir.
    Yerinde (in-place) değiştirir.
    """
    versiyon = veri.get("sema_versiyonu", 1)
    if versiyon < 1:
        veri["sema_versiyonu"] = 1
    # gelecekte: if versiyon < 2: ... ; veri["sema_versiyonu"] = 2
