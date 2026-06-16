"""
PV-DOC — Pilot Üretim Proses Validasyon doküman üreticisi.

Uygulama girişi. Tek görev: QApplication kur, ana pencereyi aç,
varsa son oturumu (kaldığı proje) geri yükle.

Çalıştırma:
    python main.py
"""

from __future__ import annotations

import sys


def main() -> int:
    # PyQt6 importu fonksiyon içinde: paket kurulu değilse anlaşılır mesaj ver.
    try:
        from PyQt6.QtWidgets import QApplication
    except ImportError:
        sys.stderr.write(
            "PyQt6 bulunamadı. Kurmak için:\n"
            "    pip install PyQt6\n"
        )
        return 1

    from ana_pencere import AnaPencere

    app = QApplication(sys.argv)
    app.setApplicationName("PV-DOC")
    app.setOrganizationName("PV-DOC")

    pencere = AnaPencere()
    pencere.show()

    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
