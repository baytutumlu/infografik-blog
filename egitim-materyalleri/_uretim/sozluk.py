import os
from reportlab.platypus import KeepTogether, SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer
from pdf_ortak import *
OUT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

T = [
    ("Aktif devir hızı", "Asset turnover", "Hasılatın ortalama toplam varlıklara oranı. Varlıkların satış üretme verimliliğini gösterir."),
    ("Aktif karlılığı", "Return on assets (ROA)", "Net karın ortalama toplam varlıklara oranı."),
    ("Aktifleştirme", "Capitalisation", "Bir harcamanın gider yerine varlık olarak bilançoya alınması."),
    ("Alacak tahsil süresi", "Days sales outstanding (DSO)", "Ticari alacakların ortalama kaç günde tahsil edildiği."),
    ("Altman Z skoru", "Altman Z-score", "Beş finansal oranı ağırlıklandırarak iflas riskini tahmin eden model."),
    ("Amortisman", "Depreciation", "Maddi duran varlık maliyetinin faydalı ömrü boyunca gidere dağıtılması."),
    ("Ana ortaklık payı", "Attributable to owners of the parent", "Konsolide kar veya özkaynağın ana şirket ortaklarına düşen kısmı."),
    ("Asit test oranı", "Quick ratio, acid-test ratio", "Stoklar hariç dönen varlıkların kısa vadeli yükümlülüklere oranı."),
    ("Azınlık payları", "Non-controlling interests", "Bağlı ortaklıklarda ana şirket dışındaki ortaklara ait pay."),
    ("Bağımsız denetim raporu", "Independent auditor's report", "Denetçinin finansal tabloların gerçeğe uygunluğuna dair görüşü."),
    ("Bağlı ortaklık", "Subsidiary", "Ana şirketin kontrol ettiği ve konsolide ettiği şirket."),
    ("Beneish M skoru", "Beneish M-score", "Sekiz endeksle kar manipülasyonu olasılığını tahmin eden model."),
    ("Bilanço", "Balance sheet, statement of financial position", "Belirli bir tarihte varlıkları, yükümlülükleri ve özkaynakları gösteren tablo."),
    ("Bilanço dışı yükümlülük", "Off-balance sheet liability", "Bilançoda görünmeyen ama dipnotlarda açıklanan taahhüt ve garantiler."),
    ("Borç / Özkaynak oranı", "Debt to equity ratio", "Finansal borçların özkaynaklara oranı."),
    ("Brüt kar", "Gross profit", "Hasılattan satışların maliyetinin düşülmesiyle bulunan kar."),
    ("Brüt kar marjı", "Gross margin", "Brüt karın hasılata oranı."),
    ("Cari oran", "Current ratio", "Dönen varlıkların kısa vadeli yükümlülüklere oranı."),
    ("Dağıtılmamış karlar", "Retained earnings", "Geçmiş dönemlerde elde edilip ortaklara dağıtılmayan karlar."),
    ("Değer düşüklüğü", "Impairment", "Varlığın geri kazanılabilir tutarının defter değerinin altına düşmesiyle ayrılan zarar."),
    ("Dipnotlar", "Notes to the financial statements", "Tablolardaki rakamların arkasındaki ayrıntıyı ve politikaları açıklayan bölüm."),
    ("Dönen varlıklar", "Current assets", "Bir yıl içinde nakde dönmesi beklenen varlıklar."),
    ("Duran varlıklar", "Non-current assets", "Bir yıldan uzun süre kullanılacak veya elde tutulacak varlıklar."),
    ("Efektif vergi oranı", "Effective tax rate", "Vergi giderinin vergi öncesi kara oranı."),
    ("Ertelenmiş vergi", "Deferred tax", "Muhasebe ve vergi matrahı arasındaki geçici farklardan doğan vergi varlığı veya yükümlülüğü."),
    ("Esas faaliyet karı", "Operating profit", "Brüt kardan faaliyet giderleri ve diğer faaliyet gelir giderlerinin netlenmesiyle bulunan kar."),
    ("FAVÖK", "EBITDA", "Faiz, amortisman ve vergi öncesi kar."),
    ("FAVÖK marjı", "EBITDA margin", "FAVÖK'ün hasılata oranı."),
    ("Faaliyet raporu", "Annual report, integrated report", "Yönetimin dönem faaliyetlerini, stratejisini ve risklerini anlattığı rapor."),
    ("Faiz karşılama oranı", "Interest coverage ratio", "Faiz ve vergi öncesi karın finansman giderlerine oranı."),
    ("Faktoring", "Factoring", "Alacakların vadesinden önce bir finans kuruluşuna devredilerek nakde çevrilmesi."),
    ("Finansal borç", "Financial debt, borrowings", "Banka kredileri, tahviller ve kiralama yükümlülükleri gibi faizli borçlar."),
    ("Finansal kiralama", "Lease", "Bir varlığın kullanım hakkının kira karşılığı edinilmesi. TFRS 16 ile bilançoya alınır."),
    ("Finansman gideri", "Finance cost", "Faiz, kur farkı ve diğer finansman maliyetleri."),
    ("FVÖK", "EBIT", "Faiz ve vergi öncesi kar."),
    ("FD/FAVÖK", "EV/EBITDA", "Firma değerinin FAVÖK'e oranı. Borç yapısından bağımsız değerleme çarpanı."),
    ("F/K oranı", "Price to earnings ratio (P/E)", "Piyasa değerinin net kara oranı."),
    ("Firma değeri", "Enterprise value (EV)", "Piyasa değeri artı net borç artı azınlık payları."),
    ("Gelir tablosu", "Income statement, profit or loss statement", "Bir dönemin gelir, gider ve karını gösteren tablo."),
    ("Gerçeğe uygun değer", "Fair value", "Piyasa katılımcıları arasında olağan işlemde oluşacak fiyat."),
    ("Geçmiş yıllar karları", "Prior years' profits", "Önceki dönemlerden devreden ve dağıtılmamış birikmiş karlar."),
    ("Hasılat", "Revenue", "Esas faaliyetlerden elde edilen satış gelirleri."),
    ("İlişkili taraf", "Related party", "Şirketi kontrol eden veya şirketle önemli etkisi olan kişi ve kuruluşlar."),
    ("İşletme faaliyetlerinden nakit akışı", "Cash flow from operating activities (CFO)", "Esas faaliyetlerden elde edilen net nakit."),
    ("İşletme sermayesi", "Working capital", "Dönen varlıklar eksi kısa vadeli yükümlülükler. Dar tanımı alacak artı stok eksi ticari borç."),
    ("Kaldıraç", "Leverage", "Varlıkların borçla finanse edilme derecesi."),
    ("Kar kalitesi", "Earnings quality", "Raporlanan karın nakitle desteklenme derecesi."),
    ("Karşılık", "Provision", "Zamanı veya tutarı belirsiz yükümlülükler için ayrılan tutar."),
    ("Kısa vadeli yükümlülükler", "Current liabilities", "Bir yıl içinde ödenecek borçlar."),
    ("Konsolide finansal tablo", "Consolidated financial statements", "Ana şirket ve bağlı ortaklıklarının tek bir ekonomik birim gibi sunulduğu tablolar."),
    ("Koşullu yükümlülük", "Contingent liability", "Gerçekleşmesi gelecekteki belirsiz olaylara bağlı olası yükümlülük."),
    ("Kullanım hakkı varlığı", "Right-of-use asset", "Kiralanan varlığın kullanım hakkının bilançodaki karşılığı."),
    ("Kur farkı", "Foreign exchange gain or loss", "Döviz cinsi varlık ve borçların değerlemesinden doğan kar veya zarar."),
    ("Maddi duran varlıklar", "Property, plant and equipment (PP&E)", "Arazi, bina, makine gibi üretimde kullanılan varlıklar."),
    ("Maddi olmayan duran varlıklar", "Intangible assets", "Marka, lisans, yazılım gibi fiziksel olmayan varlıklar."),
    ("Nakit dönüşüm süresi", "Cash conversion cycle", "Stok süresi artı alacak süresi eksi borç süresi."),
    ("Nakit akış tablosu", "Statement of cash flows", "Dönem içindeki nakit giriş ve çıkışlarını üç faaliyet başlığında gösteren tablo."),
    ("Nakit oranı", "Cash ratio", "Nakit ve kısa vadeli yatırımların kısa vadeli yükümlülüklere oranı."),
    ("Net borç", "Net debt", "Finansal borçlar eksi nakit ve nakit benzerleri."),
    ("Net kar marjı", "Net profit margin", "Net karın hasılata oranı."),
    ("Net parasal pozisyon kazancı", "Monetary gain or loss", "TMS 29 kapsamında parasal kalemlerin enflasyon karşısındaki kazanç veya kaybı."),
    ("NOPAT", "Net operating profit after tax", "Vergi sonrası faaliyet karı."),
    ("Olay sonrası hususlar", "Events after the reporting period", "Bilanço tarihinden sonra, tablolar yayımlanmadan önce gerçekleşen önemli olaylar."),
    ("Ödenmiş sermaye", "Share capital", "Ortakların taahhüt edip ödediği nominal sermaye."),
    ("Özkaynak karlılığı", "Return on equity (ROE)", "Net karın ortalama özkaynaklara oranı."),
    ("Özkaynaklar", "Equity", "Varlıklardan yükümlülükler düşüldükten sonra ortaklara kalan pay."),
    ("Özkaynak değişim tablosu", "Statement of changes in equity", "Özkaynak kalemlerinin dönem içindeki hareketini gösteren tablo."),
    ("Özkaynak yöntemi", "Equity method", "İştirak ve iş ortaklıklarının kar payının tek satırla muhasebeleştirilmesi."),
    ("PD/DD oranı", "Price to book ratio (P/B)", "Piyasa değerinin özkaynak defter değerine oranı."),
    ("Pazarlama giderleri", "Selling and marketing expenses", "Satış, dağıtım ve pazarlama faaliyetlerine ilişkin giderler."),
    ("Peşin ödenmiş giderler", "Prepaid expenses", "Gelecek dönemlere ait olup önceden ödenen giderler."),
    ("Piyasa değeri", "Market capitalisation", "Hisse fiyatı çarpı toplam pay adedi."),
    ("Rasyo", "Financial ratio", "İki finansal tablo kalemi arasındaki oran."),
    ("Sahip kazancı", "Owner earnings", "İşletme nakit akışı artı amortisman eksi bakım yatırımları. Buffett'ın tanımıdır."),
    ("Satışların maliyeti", "Cost of sales, cost of goods sold (COGS)", "Satılan mal veya hizmetin üretim ya da alış maliyeti."),
    ("Segment raporlaması", "Segment reporting", "Faaliyetlerin coğrafya veya iş koluna göre ayrı ayrı sunulması."),
    ("Serbest nakit akışı", "Free cash flow (FCF)", "İşletme nakit akışı eksi yatırım harcamaları."),
    ("Stok devir hızı", "Inventory turnover", "Satışların maliyetinin ortalama stoklara oranı."),
    ("Stokta kalma süresi", "Days inventory outstanding (DIO)", "Stokların ortalama kaç gün depoda kaldığı."),
    ("Şartlı görüş", "Qualified opinion", "Denetçinin belirli konular dışında tabloları uygun bulduğu görüş."),
    ("Şerefiye", "Goodwill", "Satın alma bedelinin edinilen net varlıkların gerçeğe uygun değerini aşan kısmı."),
    ("Taahhüt", "Commitment", "Gelecekte yerine getirilmesi gereken sözleşmesel yükümlülük."),
    ("Tahakkuk", "Accrual", "Gelir ve giderin nakit hareketinden bağımsız olarak ait olduğu dönemde kaydedilmesi."),
    ("Temettü", "Dividend", "Karın ortaklara dağıtılan kısmı."),
    ("Temettü dağıtım oranı", "Payout ratio", "Dağıtılan temettünün net kara oranı."),
    ("Temettü verimi", "Dividend yield", "Hisse başı temettünün hisse fiyatına oranı."),
    ("TFRS", "TFRS, Turkish IFRS", "Türkiye Finansal Raporlama Standartları. UFRS ile uyumlu standart seti."),
    ("Ticari alacaklar", "Trade receivables", "Satışlardan doğan ve henüz tahsil edilmemiş alacaklar."),
    ("Ticari borç ödeme süresi", "Days payables outstanding (DPO)", "Tedarikçilere ortalama kaç günde ödeme yapıldığı."),
    ("Ticari borçlar", "Trade payables", "Mal ve hizmet alımlarından doğan, henüz ödenmemiş borçlar."),
    ("TMS 29", "IAS 29, hyperinflation accounting", "Yüksek enflasyonlu ekonomilerde finansal tabloların satın alma gücüne göre yeniden ifade edilmesi."),
    ("TRİ", "Collaterals, pledges and mortgages (CPM)", "Teminat, rehin ve ipotekler. Şirketin verdiği güvencelerin dipnot tablosu."),
    ("Türev araçlar", "Derivatives", "Değeri bir dayanak varlığa bağlı forward, swap ve opsiyon gibi sözleşmeler."),
    ("Uzun vadeli yükümlülükler", "Non-current liabilities", "Vadesi bir yıldan uzun borçlar."),
    ("Varlık kalitesi", "Asset quality", "Varlıkların gerçekten nakit üretme veya nakde dönme kabiliyeti."),
    ("Yabancı para pozisyonu", "Foreign currency position", "Döviz cinsi varlıklar ile yükümlülükler arasındaki fark."),
    ("Yatırılan sermaye", "Invested capital", "Özkaynaklar artı net borç. Faaliyete bağlanmış toplam sermaye."),
    ("Yatırılan sermaye getirisi", "Return on invested capital (ROIC)", "NOPAT'ın yatırılan sermayeye oranı."),
    ("Yatırım harcaması", "Capital expenditure (CAPEX)", "Maddi ve maddi olmayan duran varlık alımları için yapılan nakit çıkışı."),
    ("Yeniden değerleme", "Revaluation", "Duran varlıkların gerçeğe uygun değerle yeniden ölçülmesi."),
    ("Yükümlülük", "Liability", "Geçmiş olaylardan doğan ve kaynak çıkışı gerektiren mevcut borç."),
]

ALFABE = "abcçdefgğhıijklmnoöprsştuüvyz"
def anahtar(s):
    s = s.replace("I", "ı").replace("İ", "i").lower()
    return [ALFABE.index(ch) if ch in ALFABE else -1 for ch in s]

def build():
    path = f"{OUT}/06-terim-sozlugu.pdf"
    def sayfa(c, d):
        ust_bilgi(c, "Terim Sözlüğü", "Bilanço okumada sık geçen Türkçe terimler ve İngilizce karşılıkları")
        alt_bilgi(c, "Bilanço Okuma Eğitimi, ek materyal 6")
    doc = SimpleDocTemplate(path, pagesize=A4, leftMargin=14 * mm, rightMargin=14 * mm, topMargin=28 * mm, bottomMargin=15 * mm,
                            title="Terim Sözlüğü", author="Bilanço Okuma Eğitimi")
    tr = stil("tr", 8.8, "DVB", leading=10.8)
    en = stil("en", 8.3, leading=10.3, color=LACI)
    ac = stil("ac", 8.3, leading=10.3)
    harf = stil("harf", 10, "DVB", color=colors.white, leading=12)
    kafa = stil("kafa", 8.5, "DVB", color=colors.white)
    cw = [46 * mm, 50 * mm, W - 28 * mm - 96 * mm]
    temel = [("VALIGN", (0, 0), (-1, -1), "TOP"),
             ("LINEBELOW", (0, 0), (-1, -1), 0.25, colors.HexColor("#D0D7E2")),
             ("TOPPADDING", (0, 0), (-1, -1), 2.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5)]
    story = []
    k = Table([[Paragraph("Türkçe terim", kafa), Paragraph("İngilizce karşılığı", kafa), Paragraph("Kısa tanım", kafa)]], colWidths=cw)
    k.setStyle(TableStyle(temel + [("BACKGROUND", (0, 0), (-1, -1), LACI)]))
    story.append(k)
    gruplar = {}
    for t, e, a in sorted(T, key=lambda x: anahtar(x[0])):
        gruplar.setdefault(t[0].upper(), []).append([Paragraph(t, tr), Paragraph(e, en), Paragraph(a, ac)])
    for h, satirlar in gruplar.items():
        ilk = Table([[Paragraph(h, harf), "", ""]] + satirlar[:2], colWidths=cw)
        ilk.setStyle(TableStyle(temel + [("SPAN", (0, 0), (-1, 0)), ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#5B7BA6"))]))
        parca = [KeepTogether([ilk])]
        if satirlar[2:]:
            kalan = Table(satirlar[2:], colWidths=cw)
            kalan.setStyle(TableStyle(temel))
            parca.append(kalan)
        story += parca
    doc.build(story, onFirstPage=sayfa, onLaterPages=sayfa)
    print(len(T), "terim")

if __name__ == "__main__":
    build()
