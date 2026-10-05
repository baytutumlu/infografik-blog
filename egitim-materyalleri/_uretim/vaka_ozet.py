import os
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, KeepTogether
from pdf_ortak import *
from vaka import LINKLER
OUT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "07-vaka-sirketi-CCOLA"))

h1 = stil("h1", 12, "DVB", color=LACI, spaceBefore=8, spaceAfter=4)
p = stil("p", 9, leading=12.5, spaceAfter=4)
kucuk = stil("kc", 7.6, leading=9.6, color=GRI)
th = stil("th", 8.3, "DVB", color=colors.white, alignment=1)
td = stil("td", 8.3, alignment=2)
tdl = stil("tdl", 8.3)

def tablo(basliklar, satirlar, genis):
    rows = [[Paragraph(b, th) for b in basliklar]]
    for s in satirlar:
        rows.append([Paragraph(str(s[0]), tdl)] + [Paragraph(str(x), td) for x in s[1:]])
    t = Table(rows, colWidths=genis)
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), LACI), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                           ("LINEBELOW", (0, 1), (-1, -1), 0.25, colors.HexColor("#D0D7E2")),
                           ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F4F7FB")]),
                           ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3)]))
    return t

def madde(liste):
    return [Paragraph(f"• {x}", stil("m", 9, leading=12.3, leftIndent=8, firstLineIndent=-8, spaceAfter=2)) for x in liste]

def build():
    path = f"{OUT}/CCOLA-faaliyet-raporu-ozeti-ve-vaka-sorulari.pdf"
    def sayfa(c, d):
        ust_bilgi(c, "Vaka Şirketi  Coca-Cola İçecek (CCOLA)".replace("  ", ", "), "Ders 23 vaka dosyası. 2025 entegre faaliyet raporu özeti ve çalışma soruları")
        alt_bilgi(c, "Bilanço Okuma Eğitimi, ek materyal 7. Eğitim amaçlıdır, yatırım tavsiyesi değildir.")
    doc = SimpleDocTemplate(path, pagesize=A4, leftMargin=16 * mm, rightMargin=16 * mm, topMargin=28 * mm, bottomMargin=16 * mm,
                            title="CCOLA Vaka Dosyası", author="Bilanço Okuma Eğitimi")
    tw = W - 32 * mm
    s = []
    s.append(Paragraph("Bu dosyanın içeriği", h1))
    s.append(Paragraph("Vaka klasöründe dört dosya var. Bu PDF faaliyet raporunun özetini ve derste birlikte çözeceğimiz soruları içerir. "
                       "CCOLA-finansal-tablolar dosyası 2023, 2024 ve 2025 yıllık konsolide bilanço, gelir tablosu ve nakit akış tablosunu milyon TL olarak verir. "
                       "Diğer iki Excel dosyası ise aynı rakamlarla doldurulmuş rasyo şablonu ve Altman Beneish şablonudur, kendi hesabınızı kontrol etmek için kullanın.", p))
    s.append(Paragraph("Şirket kimliği", h1))
    s.append(tablo(["Gösterge", "2025"], [
        ("Faaliyet gösterilen ülke", "12"), ("Çalışan", "10 bin+"), ("Şişeleme fabrikası", "36"), ("Meyve işleme fabrikası", "3"),
        ("Üretim hattı", "172"), ("Satış noktası", "yaklaşık 1,4 milyon"), ("Marka sayısı", "25+"), ("Hitap edilen nüfus", "600 milyon+"),
    ], [tw * 0.6, tw * 0.4]))
    s.append(Spacer(1, 4))
    s.append(Paragraph("Şirket Türkiye ile birlikte Orta Asya, Orta Doğu ve Güney Asya'da toplam 12 ülkede faaliyet gösteren bir şişeleyici. "
                       "Raporda yatırım haberleriyle adı geçen ülkeler arasında Azerbaycan, Özbekistan, Irak, Tacikistan, Kazakistan, Pakistan, Ürdün ve Bangladeş var. "
                       "Şişeleyici iş modelinde konsantre marka sahibinden alınır, üretim, dağıtım ve satış şirkette kalır. "
                       "Bu yapı brüt marjın konsantre fiyatına ve ambalaj maliyetine duyarlı olduğu anlamına gelir.", p))
    s.append(Paragraph("2025'in öne çıkanları (faaliyet raporundan, 2025/12 alım gücüyle)", h1))
    s.append(tablo(["milyon TL", "2024", "2025", "Değişim"], [
        ("Satış hacmi (milyon ünite kasa)", "1.501", "1.622", "%8,0"),
        ("Net satış gelirleri", "180.216", "187.185", "%3,9"),
        ("FVÖK", "24.719", "25.156", "%1,8"),
        ("FVÖK marjı", "%13,7", "%13,4", "-0,3 puan"),
        ("FAVÖK", "33.177", "33.197", "%0,1"),
        ("FAVÖK marjı", "%18,4", "%17,7", "-0,7 puan"),
        ("Net kar", "19.390", "14.072", "-%27,4"),
        ("Net kar marjı", "%10,8", "%7,5", "-3,3 puan"),
        ("Toplam özkaynaklar", "70.294", "75.840", "%7,9"),
        ("Toplam varlıklar", "194.148", "190.872", "-%1,7"),
        ("Net borç", "33.764", "25.593", "-%24,2"),
        ("Net borç / FAVÖK", "1,02", "0,77", ""),
    ], [tw * 0.46, tw * 0.18, tw * 0.18, tw * 0.18]))
    s.append(Paragraph("Not. Faaliyet raporundaki özkaynak ve varlık rakamları azınlık payları ve sunum farkları nedeniyle Fintables tablolarıyla birebir aynı olmayabilir. "
                       "Hesaplarınızda finansal tablo dosyasını esas alın.", kucuk))
    s.append(Paragraph("Segment kırılımı", h1))
    s.append(tablo(["", "Türkiye 2024", "Türkiye 2025", "Uluslararası 2024", "Uluslararası 2025"], [
        ("Hacim (milyon UC)", "568", "562", "934", "1.060"),
        ("Net satış (milyon TL)", "81.665", "81.582", "98.821", "106.263"),
        ("FVÖK (milyon TL)", "6.166", "4.993", "16.338", "18.663"),
        ("FAVÖK (milyon TL)", "9.889", "8.472", "20.603", "22.822"),
    ], [tw * 0.28] + [tw * 0.18] * 4))
    s.append(Spacer(1, 4))
    s.append(Paragraph("Büyümenin tamamı uluslararası operasyonlardan geliyor. Türkiye'de hacim yüzde 1 geriledi, FVÖK yüzde 19 düştü. "
                       "Uluslararası tarafta hacim yüzde 13,5, FVÖK yüzde 14,2 arttı. Konsolide FVÖK'ün yaklaşık dörtte üçü artık yurt dışından geliyor.", p))
    s.append(Paragraph("Yönetimin söyledikleri", h1))
    s += madde([
        "Yatırım tutarı 13,9 milyar TL. Azerbaycan İsmayıllı, Özbekistan Namangan ve Irak Bağdat fabrikaları açıldı, Tacikistan'da ikinci hat devreye alındı, Kazakistan Aktöbe fabrikası için çalışma başladı.",
        "Talepten önce yatırım yaklaşımı sürüyor. Azerbaycan'da kapasite yüzde 50 arttı, Özbekistan'daki toplam yatırım 550 milyon doları aştı.",
        "Fitch notu BBB, görünüm durağan. S&P notu BB+, görünüm durağana yükseltildi. Şirketin notu ülke notunun iki basamak üzerinde.",
        "Türkiye'de su kategorisinde bilinçli daralma ve katma değerli kategorilere odaklanma stratejisi uygulanıyor.",
    ])
    s.append(Paragraph("2026 beklentileri (guidance)", h1))
    s.append(tablo(["Gösterge", "Beklenti"], [
        ("Konsolide satış hacmi", "Orta tek haneli büyüme"),
        ("Ünite kasa başına net satış geliri", "Yatay ile orta tek haneli artış arası"),
        ("FVÖK marjı", "Yatay"),
    ], [tw * 0.5, tw * 0.5]))
    s.append(Paragraph("Yönetimin öncelikli risk listesi", h1))
    s += madde([
        "Makroekonomik yavaşlama ve istikrarsızlık. Yüksek faiz, kalıcı enflasyon, talep daralması, bayi ve tedarikçi likiditesi.",
        "Politik istikrarsızlık ve güvenlik. Rusya Ukrayna savaşı, Orta Doğu, yaptırımlar ve ticaret savaşları.",
        "İnsan ve yetenek yönetimi. Yetenek kazanımı, iş sağlığı ve güvenliği.",
        "Etik, uyum, düzenleme ve vergiler. Ülkeler arası değişken vergi ve ambalaj, etiketleme düzenlemeleri.",
    ])
    sorular = [
        "Net kar 2023'ten 2025'e yaklaşık üçte birine inerken FAVÖK neredeyse yatay kaldı. Gelir tablosunda bu makası hangi iki kalem açıklıyor? Net parasal pozisyon kazancı ve finansman giderlerini yıllar boyunca karşılaştırın.",
        "Efektif vergi oranı yüzde 18,5'ten yüzde 33,1'e çıktı. Hangi kalemler bu artışa yol açmış olabilir, vergi dipnotunda neye bakarsınız?",
        "Cari oran 1,15 civarında. Bir içecek şişeleyicisi için bu düşük mü? Nakit dönüşüm süresinin negatife dönmesiyle birlikte yorumlayın.",
        "Ticari borç ödeme süresi 100 günün üzerinde. Bu bir güç göstergesi mi yoksa risk mi? Kırmızı bayrak listesindeki işletme sermayesi makyajı maddesini kontrol edin.",
        "Yeniden sınıflandırılacak birikmiş diğer kapsamlı gider 3 yılda 17,7 milyardan 46,7 milyar TL'ye büyüdü. Bu kalem nedir, neden çok uluslu bir şirkette büyür, nakit akışını etkiler mi?",
        "CAPEX / Amortisman 1,7 ile 2,2 arasında. Bu ne kadar büyüme yatırımı, ne kadar bakım yatırımı? Segment büyümesiyle ilişkilendirin.",
        "2023 ve 2024'te bağlı ortaklıklarda pay alımına toplam yaklaşık 13 milyar TL nakit çıkışı var. Azınlık paylarının seyriyle birlikte yorumlayın.",
        "Hedge dahil net yabancı para pozisyonu eksi 8,2 milyar TL. Özkaynağa oranı nedir, kur riskini nasıl değerlendirirsiniz?",
        "Fintables tablolarındaki 2025 hasılatı 220,4 milyar TL, faaliyet raporundaki 187,2 milyar TL. Bu fark neden oluşuyor ve hangi oranları etkiliyor?",
        "Rasyo şablonu, Altman Z, Beneish M ve kırmızı bayrak listesinden sonra 8 kategorili puanlama şablonunu doldurun. Genel puanınızı ve en zayıf iki kategoriyi gerekçeleriyle yazın.",
    ]
    s.append(KeepTogether([Paragraph("Vaka çalışma soruları", h1)] + [Paragraph(f"<b>{i+1}.</b> {q}", stil("q", 9, leading=12.5, spaceAfter=5)) for i, q in enumerate(sorular[:2])]))
    for i, q in enumerate(sorular[2:], start=3):
        s.append(Paragraph(f"<b>{i}.</b> {q}", stil("q", 9, leading=12.5, spaceAfter=5)))
    kay = [Paragraph("Kaynaklar", h1)]
    for a, u in LINKLER:
        kay.append(Paragraph(f'{a}<br/><a href="{u}" color="#0563C1">{u}</a>', stil("l", 8, leading=10.5, spaceAfter=3)))
    s.append(KeepTogether(kay))
    s.append(Paragraph("Faaliyet raporu verileri CCI 2025 Entegre Faaliyet Raporu sayfa 4, 13, 14, 50, 55 ve 56'dan alınmıştır. "
                       "Finansal tablo verileri Fintables üzerinden 05.10.2026 tarihinde çekilmiştir.", kucuk))
    doc.build(s, onFirstPage=sayfa, onLaterPages=sayfa)

if __name__ == "__main__":
    build()
