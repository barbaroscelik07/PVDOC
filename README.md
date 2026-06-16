# PV-DOC

Pilot Üretim Proses Validasyon Protokolü (PVP) ve Raporu (PVR) doküman üreticisi.
PyQt6 masaüstü uygulaması. Kullanıcı ürün/proses bilgilerini girer; program
TİTCK formatına uygun Word (ve PDF) çıktısı üretir.

## Yerel çalıştırma

```bash
pip install -r requirements.txt
python main.py
```

## Windows EXE üretimi (GitHub Actions ile — terminal gerekmez)

1. Bu dosyaları repoya yükle/commit'le (örn. `barbaroscelik07/pvdoc`).
2. `main` (veya `master`) dalına push et.
3. GitHub'da **Actions** sekmesine git → **Windows EXE Build** çalışması otomatik başlar.
4. Çalışma bitince (yeşil tik) → çalışmaya tıkla → en altta **Artifacts** altında
   **PV-DOC-windows** dosyasını indir. İçinde `PV-DOC.exe` var.

### Sürüm yayını (otomatik Release)
`v` ile başlayan bir etiket (tag) push edersen EXE otomatik olarak GitHub Release'e eklenir:
örneğin `v0.1` etiketi → Releases sayfasında indirilebilir `PV-DOC.exe`.

## Klasör yapısı

```
main.py              Uygulama girişi
ana_pencere.py       Hub + sekme yapısı, proje aç/kaydet
core/                Veri modeli (models.py) ve proje IO (proje_io.py)
moduller/            Sekme modülleri (Faz 1+ ile dolacak)
diyagram/            Akış diyagramı editörü (ayrı pencere)
cikti/               Word/PDF üretim motoru
kaynaklar/           Spek kartı kayıtları, sabit kaynaklar
PVDOC.spec           PyInstaller yapılandırması
.github/workflows/   Otomatik build
```

## Durum

- [x] Faz 0: Veri modeli, proje kaydet/yükle, hub iskeleti, build altyapısı
- [ ] Faz 1: Spek kartı sistemi
- [ ] Faz 2: Diğer modüller (formül, proses, risk, ekipman, numune)
- [ ] Faz 3: Akış diyagramı editörü
- [ ] Faz 4: Word/PDF çıktı motoru
