from ccola_veri import BILANCO, GELIR, NAKIT, PIYASA, bul

def mn(v):
    return [round(x / 1e6, 1) for x in v]

def neg(v):
    return [-x for x in v]

def girdiler():
    B = lambda k: mn(bul(BILANCO, k))
    G = lambda k: mn(bul(GELIR, k))
    N = lambda k: mn(bul(NAKIT, k))
    d = {
        "nakit": B("Nakit ve Nakit Benzerleri"),
        "fin_yat": B("Finansal Yatırımlar (KV)"),
        "tic_alacak": B("Ticari Alacaklar"),
        "stok": B("Stoklar"),
        "donen": B("Toplam Dönen Varlıklar"),
        "mdv": B("Maddi Duran Varlıklar"),
        "tv": B("TOPLAM VARLIKLAR"),
        "kv_fin": B("Finansal Borçlar (KV)"),
        "tic_borc": B("Ticari Borçlar (KV)"),
        "kvyk": B("Toplam Kısa Vadeli Yükümlülükler"),
        "uv_fin": B("Finansal Borçlar (UV)"),
        "uvyk": B("Toplam Uzun Vadeli Yükümlülükler"),
        "ana_oz": B("Ana Ortaklığa Ait Özkaynaklar"),
        "top_oz": B("Toplam Özkaynaklar"),
        "gecmis": B("Geçmiş Yıllar Kar/Zararları"),
        "donem_net_bil": B("Dönem Net Kar/Zararı"),
        "hasilat": G("Satış Gelirleri"),
        "smm": neg(G("Satışların Maliyeti (-)")),
        "gyg": neg(G("Genel Yönetim Giderleri (-)")),
        "pazarlama": neg(G("Pazarlama, Satış ve Dağıtım Giderleri (-)")),
        "arge": [0, 0, 0],
        "faal_kar": G("Faaliyet Karı"),
        "fvok": G("Finansman Geliri (Gideri) Öncesi Faaliyet Karı"),
        "amort": N("Amortisman ve İtfa Gideri"),
        "fin_gider": neg(G("Finansal Giderler (-)")),
        "vok": G("Vergi Öncesi Kar"),
        "vergi": neg(G("Vergi Gideri (-)")),
        "net_kar": G("Dönem Karı"),
        "ana_net": G("Ana Ortaklık Payları"),
        "ocf": N("A. İşletme Faaliyetlerinden Nakit Akışları"),
        "capex": neg(N("MDV ve MODV Alımından Çıkışlar (CAPEX)")),
        "temettu": neg(N("Ödenen Temettüler")),
        "_kaynak": "Kaynak KAP finansal tabloları, Fintables üzerinden 05.10.2026 tarihinde alındı. Tutarlar TMS 29 kapsamında 2026/06 satın alma gücüne endekslidir.",
    }
    return d

def ornek_2025():
    return {k: v[2] for k, v in girdiler().items() if not k.startswith("_")}
