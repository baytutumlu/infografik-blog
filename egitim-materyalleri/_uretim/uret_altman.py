import os
from altman_beneish import build
from ccola_veri import BILANCO, GELIR, NAKIT, PIYASA, bul
OUT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
mn = lambda v: [round(x / 1e6, 1) for x in v]
neg = lambda v: [-x for x in v]
B = lambda k: mn(bul(BILANCO, k)); G = lambda k: mn(bul(GELIR, k)); N = lambda k: mn(bul(NAKIT, k))
PAY = PIYASA["pay_adedi"]
FIYAT = [46.32, 58.71, 58.80]  # 2023, 2024, 2025 son işlem günü kapanışları
# TÜFE (2003=100) Aralık değerleri, kaynak TÜİK, Fintables üzerinden
TUFE = {2023: 1859.38, 2024: 2684.55, 2025: 3513.87}
K_FINTABLES = round(220428307256 / 187185e6, 4)  # 2025/12 alım gücünden 2026/06 alım gücüne, hasılat oranından
kat = [round(TUFE[2025] / TUFE[y] * K_FINTABLES, 4) for y in (2023, 2024, 2025)]
veri = {
    "donen": B("Toplam Dönen Varlıklar"), "tv": B("TOPLAM VARLIKLAR"),
    "kvyk": B("Toplam Kısa Vadeli Yükümlülükler"), "uvyk": B("Toplam Uzun Vadeli Yükümlülükler"),
    "uv_fin": B("Finansal Borçlar (UV)"), "top_oz": B("Toplam Özkaynaklar"),
    "gecmis": B("Geçmiş Yıllar Kar/Zararları"), "donem_bil": B("Dönem Net Kar/Zararı"),
    "tic_alacak": B("Ticari Alacaklar"), "mdv": B("Maddi Duran Varlıklar"),
    "fin_yat": [a + b for a, b in zip(B("Finansal Yatırımlar (KV)"), [0, 0, 0])],
    "hasilat": G("Satış Gelirleri"), "smm": neg(G("Satışların Maliyeti (-)")),
    "gyg": neg(G("Genel Yönetim Giderleri (-)")), "pazarlama": neg(G("Pazarlama, Satış ve Dağıtım Giderleri (-)")),
    "fvok": G("Finansman Geliri (Gideri) Öncesi Faaliyet Karı"), "surd_kar": G("Dönem Karı"),
    "amort": N("Amortisman ve İtfa Gideri"), "ocf": N("A. İşletme Faaliyetlerinden Nakit Akışları"),
    "pd_nom": [round(f * PAY / 1e6, 1) for f in FIYAT], "pd_kat": kat,
    "_kaynak": "Kaynak KAP finansal tabloları ve BIST kapanış fiyatları, Fintables üzerinden 05.10.2026 tarihinde alındı. Tablolar 2026/06 alım gücüyle endekslidir.",
}
notlar = {
    "pd_nom": "Kapanış fiyatları 2023 için 29.12.2023, 2024 için 31.12.2024, 2025 için 31.12.2025 (bedelsiz düzeltilmiş). Pay adedi 2.798.078.602.",
    "pd_kat": f"Katsayı = TÜFE(Ara 2025) / TÜFE(Ara yıl) x {K_FINTABLES}. Son çarpan Fintables tutarlarının faaliyet raporundaki 2025/12 alım gücü tutarlarına oranıdır (hasılat 220.428 / 187.185).",
}
build(f"{OUT}/05-altman-z-beneish-m-sablonu.xlsx")
build(f"{OUT}/07-vaka-sirketi-CCOLA/CCOLA-altman-beneish-2023-2025.xlsx", donemler=["2023/12", "2024/12", "2025/12"],
      veri=veri, notlar=notlar, baslik_ek=", CCOLA")
print(kat)
