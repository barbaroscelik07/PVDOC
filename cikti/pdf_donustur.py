"""
Word (.docx) → PDF dönüştürücü.

LibreOffice (soffice) komut satırı ile dönüştürür. LibreOffice kurulu değilse
(çoğu Windows iş bilgisayarında olmayabilir) hata fırlatmaz; None döner ve
çağıran taraf kullanıcıya "PDF için LibreOffice gerekli" bilgisini verir.

Word her durumda üretilir; PDF opsiyoneldir.
"""

from __future__ import annotations

import os
import shutil
import subprocess
from pathlib import Path
from typing import Optional


def _soffice_bul() -> Optional[str]:
    """soffice/libreoffice çalıştırılabilirini bulur."""
    for ad in ("soffice", "libreoffice"):
        yol = shutil.which(ad)
        if yol:
            return yol
    # Windows tipik kurulum yolları
    adaylar = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for a in adaylar:
        if os.path.exists(a):
            return a
    return None


def pdf_mevcut_mu() -> bool:
    """PDF dönüşümü yapılabilir mi (LibreOffice var mı)?"""
    return _soffice_bul() is not None


def docx_to_pdf(docx_yolu: str | Path, cikti_dizini: str | Path | None = None,
                zaman_asimi: int = 120) -> Optional[Path]:
    """
    .docx dosyasını PDF'e çevirir. Başarılıysa PDF yolunu, LibreOffice yoksa
    veya dönüşüm başarısızsa None döner.
    """
    soffice = _soffice_bul()
    if soffice is None:
        return None

    docx_yolu = Path(docx_yolu)
    if cikti_dizini is None:
        cikti_dizini = docx_yolu.parent
    cikti_dizini = Path(cikti_dizini)
    cikti_dizini.mkdir(parents=True, exist_ok=True)

    try:
        subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf",
             "--outdir", str(cikti_dizini), str(docx_yolu)],
            check=True, capture_output=True, timeout=zaman_asimi,
        )
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired, OSError):
        return None

    pdf = cikti_dizini / (docx_yolu.stem + ".pdf")
    return pdf if pdf.exists() else None
