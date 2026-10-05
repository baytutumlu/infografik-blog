import os
from reportlab.platypus import FrameBreak, BaseDocTemplate, Frame, PageTemplate, Paragraph, Table, TableStyle, Spacer, KeepTogether
from pdf_ortak import *
OUT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

h = stil("h", 9.5, "DVB", color=LACI, spaceBefore=5, spaceAfter=2)
it = stil("it", 9.3, leading=11.2)
ipucu = stil("ip", 7.7, "DVI", leading=9.4, color=GRI)
KUTU = "☐"

def madde_tablosu(maddeler, genislik, ikinci=True):
    rows = []
    for m in maddeler:
        metin, nereye = m if isinstance(m, tuple) else (m, None)
        hucre = [Paragraph(metin, it)]
        if nereye and ikinci:
            hucre.append(Paragraph(nereye, ipucu))
        rows.append([Paragraph(KUTU, stil("k", 11, leading=11.5)), hucre])
    t = Table(rows, colWidths=[5 * mm, genislik - 5 * mm])
    t.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 1), ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 2.2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2.2),
        ("LINEBELOW", (0, 0), (-1, -1), 0.25, colors.HexColor("#D0D7E2")),
    ]))
    return t

def iki_sutun(path, baslik, alt, bolumler, dip, renk=LACI, sonnot=None, bol=4):
    gutter = 6 * mm
    colw = (W - 28 * mm - gutter) / 2
    top = H - 26 * mm
    bottom = 14 * mm
    f1 = Frame(14 * mm, bottom, colw, top - bottom, leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
    f2 = Frame(14 * mm + colw + gutter, bottom, colw, top - bottom, leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)

    def sayfa(c, d):
        ust_bilgi(c, baslik, alt, renk)
        alt_bilgi(c, dip)
    doc = BaseDocTemplate(path, pagesize=A4, title=baslik, author="Bilanço Okuma Eğitimi")
    doc.addPageTemplates([PageTemplate(frames=[f1, f2], onPage=sayfa)])
    story = []
    hs = stil("h2", 11, "DVB", color=renk, spaceBefore=7, spaceAfter=3)
    for i, (bas, maddeler) in enumerate(bolumler):
        if i == bol:
            story.append(FrameBreak())
        story.append(KeepTogether([Paragraph(bas, hs), madde_tablosu(maddeler, colw)]))
    if sonnot:
        story.append(Spacer(1, 4))
        box = Table([[Paragraph(sonnot, stil("sn", 8.8, leading=11))]], colWidths=[colw])
        box.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), ACIK), ("BOX", (0, 0), (-1, -1), 0.5, renk),
                                 ("LEFTPADDING", (0, 0), (-1, -1), 5), ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                                 ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]))
        story.append(box)
    doc.build(story)

BILANCO_OKUMA = [
    ("1. Masaya oturmadan önce", [
        ("Raporun türü belli mi", "Konsolide mi solo mu, yıllık denetimli mi, ara dönem sınırlı denetimli mi"),
        ("Bağımsız denetçi görüşü olumlu mu", "Şartlı görüş, olumsuz görüş veya dikkat çekilen hususlar paragrafı var mı"),
        ("Raporlama para birimi ve TMS 29 durumu", "Tutarlar hangi tarihin satın alma gücüyle ifade ediliyor"),
        ("Karşılaştırmalı dönemler aynı esasla mı", "Önceki yılı eski rapordan değil, son raporun karşılaştırmalı sütunundan alın"),
        ("Şirketin iş modelini bir cümleyle anlatabiliyor musunuz", "Faaliyet raporu ve sektör sunumu"),
    ]),
    ("2. Bilançonun büyük resmi", [
        ("Aktif büyüklüğü ve yıllık değişimi", "Büyüme hasılatla orantılı mı"),
        ("Dönen ve duran varlık dengesi", "Sektörün sermaye yoğunluğuyla uyumlu mu"),
        ("Özkaynak / Aktif oranı", "Varlıkların ne kadarı ortak parasıyla finanse ediliyor"),
        ("Net borç mu net nakit mi", "Finansal borçlar eksi nakit ve finansal yatırımlar"),
        ("Cari oran ve kısa vadeli borç payı", "Bir yıl içinde ödenecekler, bir yıl içinde nakde dönecek varlıklarla karşılanıyor mu"),
    ]),
    ("3. Varlıkların kalitesi", [
        ("Nakit gerçekten serbest mi", "Bloke mevduat, yurt dışı nakit, kur riski dipnotu"),
        ("Ticari alacaklar ve tahsil süresi", "İlişkili taraf alacağı, vadesi geçmiş alacak, şüpheli alacak karşılığı"),
        ("Stoklar ve devir süresi", "Stok değer düşüklüğü karşılığı, stok artışı satışla uyumlu mu"),
        ("Maddi duran varlıklar ve yeniden değerleme", "Değer artışı özkaynağa mı yansımış, yatırım takvimi"),
        ("Şerefiye ve maddi olmayan varlıklar", "Değer düşüklüğü testi varsayımları, aktifleştirilen giderler"),
    ]),
    ("4. Kaynak tarafı", [
        ("Finansal borçların vade ve döviz dağılımı", "Finansal borçlar dipnotu, yabancı para pozisyonu tablosu"),
        ("Ticari borç ödeme süresi", "Tedarikçi finansmanı kalıcı mı geçici mi"),
        ("Karşılıklar, davalar ve koşullu yükümlülükler", "Karşılıklar dipnotu, TRİ tablosu"),
        ("Azınlık payları ve özkaynak hareketleri", "Özkaynak değişim tablosu, sermaye artırımı ve temettü"),
    ]),
    ("5. Gelir tablosuyla bağlantı", [
        ("Hasılat büyümesi hacim mi fiyat mı", "Segment dipnotu, faaliyet raporundaki hacim verisi"),
        ("Brüt marj ve FAVÖK marjı trendi", "En az 3 yıl yan yana"),
        ("Finansman gideri, kur farkı ve parasal kazanç", "Esas faaliyet karı ile net kar arasındaki makas nereden geliyor"),
        ("Tek seferlik kalemler ve efektif vergi oranı", "Diğer faaliyet gelirleri, yatırım faaliyeti gelirleri, vergi dipnotu"),
    ]),
    ("6. Nakit akışıyla doğrulama", [
        ("İşletme nakit akışı net karla uyumlu mu", "Kar kalitesi 1 civarında mı"),
        ("İşletme sermayesindeki değişim", "Alacak ve stok artışı nakdi ne kadar yiyor"),
        ("CAPEX ve amortisman ilişkisi", "Bakım yatırımı mı büyüme yatırımı mı"),
        ("Serbest nakit akışı ve temettünün kaynağı", "Temettü SNA ile mi borçla mı ödeniyor"),
    ]),
    ("7. Dipnotlarda mutlaka bakılacaklar", [
        ("İlişkili taraf işlemleri", "Hacim, yön ve koşullar"),
        ("Muhasebe politikası ve tahmin değişiklikleri", "Faydalı ömür, gelir tanıma, sınıflandırma"),
        ("Olay sonrası hususlar", "Bilanço tarihinden sonra açıklanan önemli gelişmeler"),
    ]),
    ("8. Sonuca bağla", [
        "Rasyo hesap şablonunu doldur ve trendleri işaretle",
        "Kırmızı bayrak kontrol listesinden geç",
        "Altman Z ve Beneish M skorlarını hesapla",
        "8 kategorili puanlama şablonuyla genel puanı çıkar",
        "Üç cümlelik özet yaz. Güçlü yan, zayıf yan, takip edilecek kalem",
    ]),
]

KIRMIZI_BAYRAK = [
    ("Gelir ve kar kalitesi", [
        ("Alacaklar satışlardan belirgin hızlı büyüyor", "DSRI 1 üzerinde, alacak tahsil süresi uzuyor"),
        ("Net kar artıyor, işletme nakit akışı düşüyor", "Kar kalitesi birkaç yıldır 1 altında"),
        ("Karın büyük kısmı esas faaliyet dışından geliyor", "Kur farkı, değerleme kazancı, varlık satışı, parasal kazanç"),
        ("Dönem sonunda olağandışı satış artışı", "Son çeyrek hasılat payı, sonraki dönem iadeleri"),
        ("Brüt marj sektörden ayrışarak yükseliyor", "Maliyet aktifleştirme veya stok değerleme şüphesi"),
        ("Hasılat tanıma politikası değişti", "Muhasebe politikaları dipnotu, müşteri sözleşmeleri"),
    ]),
    ("Varlık tarafı", [
        ("Stoklar satışlardan hızlı artıyor", "Stokta kalma süresi uzuyor, değer düşüklüğü karşılığı azalıyor"),
        ("Aktifleştirilen giderler ve maddi olmayan varlıklar şişiyor", "AQI 1 üzerinde, geliştirme maliyetleri dipnotu"),
        ("Amortisman oranı düşüyor", "Faydalı ömür uzatıldı mı, DEPI 1 üzerinde"),
        ("Yeniden değerlemeyle şişen özkaynak", "Değer artış fonu, nakit üretmeyen varlık"),
        ("İlişkili taraflardan alacak büyüyor", "İlişkili taraf dipnotu, ortaklara verilen borçlar"),
        ("Şerefiye testinde iyimser varsayımlar", "Büyüme ve iskonto oranı varsayımları, hiç değer düşüklüğü yok"),
    ]),
    ("Borç ve likidite", [
        ("Kısa vadeli borç payı hızla artıyor", "Yeniden finansman riski, cari oran 1 altında"),
        ("Açık döviz pozisyonu büyük", "Yabancı para pozisyonu tablosu, hedge oranı"),
        ("Faiz karşılama oranı 1,5 altına iniyor", "FVÖK / Finansman gideri"),
        ("Bilanço dışı yükümlülük ve verilen teminatlar yüksek", "TRİ tablosu, özkaynağa oranı"),
        ("Faktoring ve tedarikçi finansmanı ile borç gizleme", "Ticari borç süresi aniden uzuyor"),
    ]),
    ("Nakit akışı", [
        ("Serbest nakit akışı yıllardır negatif", "Büyüme sürekli borç veya sermaye artırımıyla finanse ediliyor"),
        ("Temettü SNA yerine borçla ödeniyor", "Temettü / SNA yüzde 100 üzerinde"),
        ("Faiz ve vergi ödemeleri sınıflandırması değişti", "İşletme nakit akışını yapay olarak iyileştirebilir"),
        ("Yıl sonunda işletme sermayesi makyajı", "Ticari borç yıl sonunda şişiyor, ara dönemlerde eriyor"),
    ]),
    ("Yönetişim ve denetim", [
        ("Şartlı denetim görüşü veya denetçi değişikliği", "Denetçi raporu, KAP açıklamaları"),
        ("Muhasebe politikası veya tahmin değişikliği sık", "Politika değişikliği dipnotu"),
        ("Finansal tabloların gecikmeli açıklanması", "KAP yayın tarihleri, ek süre talepleri"),
        ("Yönetimde ve finans direktörlüğünde sık değişim", "KAP özel durum açıklamaları"),
        ("Hedeflerle gerçekleşmeler sürekli ayrışıyor", "Önceki yıl beklentileri ile sonuçlar"),
    ]),
]


def build():
    iki_sutun(f"{OUT}/02-bilanco-okuma-kontrol-listesi.pdf", "Bilanço Okuma Kontrol Listesi",
              "Bir finansal tabloya oturduğunuzda sırayla bakılacak maddeler. Gri satırlar nereye bakacağınızı gösterir.",
              BILANCO_OKUMA, "Bilanço Okuma Eğitimi, ek materyal 2. Eğitim amaçlıdır, yatırım tavsiyesi değildir.")
    iki_sutun(f"{OUT}/03-kirmizi-bayrak-kontrol-listesi.pdf", "Kırmızı Bayrak Kontrol Listesi",
              "Ders 21'de işlenen uyarı işaretlerinin taranabilir hali. İşaretlenen her madde için ilgili dipnotu okuyun.",
              KIRMIZI_BAYRAK, "Bilanço Okuma Eğitimi, ek materyal 3. Eğitim amaçlıdır, yatırım tavsiyesi değildir.", renk=KIRMIZI,
              sonnot="<b>Nasıl kullanılır</b><br/>Tek bir bayrak tek başına sorun anlamına gelmez, sektör ve dönem koşullarıyla birlikte okuyun. "
                     "Aynı başlık altında üç veya daha fazla işaret birikiyorsa, ya da işaretler birbirini doğruluyorsa (örneğin DSRI yüksek ve kar kalitesi düşük) "
                     "o şirketi derin incelemeye alın. Beneish M ve Altman Z sonuçlarını bu listeyle birlikte değerlendirin.<br/><br/><b>İşaretlenen madde sayısı</b> ______ / " + str(sum(len(m) for _, m in KIRMIZI_BAYRAK)), bol=3)

if __name__ == "__main__":
    build()
