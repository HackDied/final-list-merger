# Final List Merger Pro

Birden fazla Excel fiyat teklifi dosyasini tek bir formatli Excel ciktisinda birlestiren masaustu uygulamasi.

## Ekran Goruntusu

Uygulama modern bir CustomTkinter arayuzu ile calisir. Dosyalarinizi surekleyip birakin veya dosya secici ile ekleyin, ardindan tek tikla birlestirin.

## Ozellikler

- **Coklu Excel Birlestirme** — Birden fazla fiyat teklifi dosyasini tek bir profesyonel Excel'de birlestirin
- **Grand Summary** — Tum siparislerin TOTAL, DISCOUNT ve GRAND TOTAL toplamlari otomatik hesaplanir
- **Surukle & Birak** — Dosyalari dogrudan uygulamaya surukleyip birakin
- **Dosya Siralama** — Yukari/asagi butonlari ile dosya sirasini ayarlayin
- **Coklu Secim & Silme** — Ctrl+Click ile birden fazla dosya secip tek seferde kaldirin
- **Onizleme** — Birlestirmeden once dosya icerigini kontrol edin
- **Header Hucreleri** — Tarih, referans numarasi gibi bilgileri ciktiya dahil edin (opsiyonel)
- **Cikti Klasoru Secimi** — Birlestirilmis dosyanin kaydedilecegi klasoru belirleyin
- **Otomatik Acma** — Birlestirme tamamlaninca Excel otomatik acilir
- **Klasor Hafizasi** — Son kullanilan klasoru hatirlar
- **Dosya Kilit Kontrolu** — Acik dosyalara yazma hatalarini onler
- **Dogrulama Uyarisi** — Birlestirme sonrasi toplam tutarlarin elle kontrol edilmesi icin uyari

## Kurulum

### Hazir EXE (Onerilen)

1. [Releases](https://github.com/HackDied/final-list-merger/releases) sayfasindan en son surumu indirin
2. ZIP dosyasini bir klasore cikartin
3. `Final_List_Template.xlsx` dosyasinin `Final List Merger Pro.exe` ile ayni klasorde oldugundan emin olun
4. `Final List Merger Pro.exe` dosyasini calistirin

### Kaynaktan Calistirma

```bash
# Gerekli kutuphaneler
pip install customtkinter openpyxl pandas

# Opsiyonel (surukle & birak destegi)
pip install tkinterdnd2

# Calistir
python final_list_merger.py
```

## Kullanim

1. **Dosya Ekle** — "Dosya Secmek Icin Tikla" butonuna basin veya dosyalari surukleyip birakin
2. **Siralama** — Dosyalari yukari/asagi butonlari ile istediginiz siraya getirin
3. **Ayarlar** — Header hucreleri, cikti klasoru ve otomatik acma tercihlerini yapin
4. **Birlestir** — "BIRLESTIR" butonuna basin
5. **Kontrol** — Uyari ciktiktan sonra toplam tutarlari mutlaka elle dogrulayin

## Dosya Formati

Uygulama asagidaki yapida Excel dosyalari bekler:

- **A3:B5** arasi header bilgileri (tarih, referans vb.)
- **A sutunu** — Sira numarasi (1, 2, 3 veya 1A, 1B, 2A gibi)
- **B sutunu** — Urun aciklamasi
- **C sutunu** — Birim
- **D sutunu** — Miktar
- **G sutunu** — Birim fiyat
- TOTAL, DISCOUNT ve GRAND TOTAL satirlari otomatik algilanir

## EXE Olusturma

```bash
pip install pyinstaller
pyinstaller final_list_merger.spec
```

Olusturulan exe `dist/Final List Merger Pro/` klasorunde bulunur. `Final_List_Template.xlsx` dosyasini bu klasore kopylamayi unutmayin.

## Lisans

Bu proje serbestce kullanilabilir.
