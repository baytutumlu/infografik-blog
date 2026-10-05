# Bilanço Okuma Eğitimi, Kaynaklar ve Ek Materyal

Eğitimle birlikte verilecek indirilebilir dosyalar. Hepsi katılımcının derste gördüğü hesabı kendi şirketinde tekrarlayabilmesi için tasarlandı. Excel dosyalarında sadece sarı zeminli mavi hücreler doldurulur, geri kalan her şey formüldür.

| No | Dosya | Açıklama |
|----|-------|----------|
| 1 | `01-rasyo-hesap-sablonu.xlsx` | Üç döneme kadar tablo rakamlarını girin. Likidite, borçluluk, faaliyet, karlılık ve nakit akışı oranlarından 37 gösterge otomatik hesaplanır. Her girdinin yanında finansal tablodaki kalem adı, her oranın yanında yorumlama notu var. Kontrol sayfası bilanço denkliğini ve işaret hatalarını yakalar. |
| 2 | `02-bilanco-okuma-kontrol-listesi.pdf` | 1 sayfa. Tabloya oturunca sırayla bakılacak 35 madde, 8 başlıkta. Her maddenin altında nereye bakılacağı yazıyor. |
| 3 | `03-kirmizi-bayrak-kontrol-listesi.pdf` | 1 sayfa. Ders 21 uyarı işaretlerinin işaretlenebilir hali, 5 başlıkta 26 madde. |
| 4 | `04-8-kategorili-puanlama-sablonu.xlsx` | Kategori, ölçüt, puan ve gerekçe sütunları. Her kategoride 4 ölçüt, kategori puanı ölçütlerin ortalaması. Özet sayfası toplam puanı, yüzdelik genel puanı ve yorum bandını verir. Puan hücreleri 1 ile 10 arası tam sayıya kilitli. Karlılık satırları örnek olarak CCOLA ile dolu. |
| 5 | `05-altman-z-beneish-m-sablonu.xlsx` | Girdi hücreleri finansal tablo kalemleriyle eşleştirilmiş. Altman Z orijinal model ve gelişmekte olan piyasa şirketleri için Z'' modeli, Beneish M 8 değişkenli model. TMS 29 için piyasa değeri endeksleme katsayısı satırı var. |
| 6 | `06-terim-sozlugu.pdf` | 101 terim, Türkçe alfabetik sırayla, İngilizce karşılığı ve kısa tanımıyla. |
| 7 | `07-vaka-sirketi-CCOLA/` | Ders 23 vaka şirketi Coca-Cola İçecek. Aşağıya bakın. |

## Vaka klasörü (Ders 23, CCOLA)

| Dosya | İçerik |
|-------|--------|
| `CCOLA-finansal-tablolar-2023-2025.xlsx` | 2023, 2024 ve 2025 yıllık konsolide bilanço, gelir tablosu ve nakit akış tablosu, milyon TL. Bilgi sayfasında KAP ve faaliyet raporu bağlantıları var. |
| `CCOLA-faaliyet-raporu-ozeti-ve-vaka-sorulari.pdf` | 2025 entegre faaliyet raporunun özeti (şirket kimliği, segment kırılımı, yatırımlar, 2026 beklentileri, öncelikli riskler) ve 10 vaka çalışma sorusu. |
| `CCOLA-rasyo-hesabi-2023-2025.xlsx` | Rasyo şablonunun CCOLA ile doldurulmuş hali. Katılımcı kendi hesabını bununla kontrol eder. |
| `CCOLA-altman-beneish-2023-2025.xlsx` | Altman ve Beneish şablonunun CCOLA ile doldurulmuş hali. |

Veriler KAP'ta yayımlanan tablolardan Fintables üzerinden 05.10.2026 tarihinde alındı. Fintables tutarları TMS 29 kapsamında son açıklanan dönemin (2026/06) satın alma gücüyle verdiği için faaliyet raporundaki TL tutarlardan yaklaşık 1,1776 kat yüksektir. Oranlar değişmez. Bu fark vaka sorularından biri olarak bilinçli şekilde bırakıldı.

Orijinal faaliyet raporu PDF'i telif ve dosya boyutu nedeniyle repoya konmadı, bağlantısı vaka dosyalarında var. Derste dağıtmak isterseniz KAP'tan indirip klasöre ekleyebilirsiniz.

## Yeniden üretme

Dosyalar `_uretim/` klasöründeki Python betikleriyle üretildi. Bir maddeyi, terimi veya rakamı değiştirmek için ilgili betiği düzenleyip şunu çalıştırın.

```
cd egitim-materyalleri/_uretim
python3 hepsini_uret.py
```

Gerekenler openpyxl, reportlab ve formül sonuçlarını dosyaya yazmak için LibreOffice Calc.

| Betik | Ürettiği |
|-------|----------|
| `rasyo_sablonu.py`, `uret_rasyo.py` | 1 numaralı şablon ve CCOLA rasyo dosyası |
| `kontrol_listeleri.py` | 2 ve 3 numaralı PDF'ler, maddeler dosyanın içinde liste olarak duruyor |
| `puanlama_sablonu.py` | 4 numaralı şablon |
| `altman_beneish.py`, `uret_altman.py` | 5 numaralı şablon ve CCOLA Altman Beneish dosyası |
| `sozluk.py` | 6 numaralı sözlük |
| `ccola_veri.py`, `vaka.py`, `vaka_ozet.py` | Vaka klasörü |

Tüm materyaller eğitim amaçlıdır, yatırım tavsiyesi değildir.
