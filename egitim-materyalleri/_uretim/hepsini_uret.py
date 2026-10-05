"""Tüm ek materyalleri yeniden üretir.

Gereksinimler: openpyxl, reportlab, LibreOffice Calc (formül önbelleği için).
Çalıştırma: python3 hepsini_uret.py
"""
import glob, os, subprocess, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import uret_rasyo, uret_altman  # noqa: F401  (içe aktarıldığında dosyaları üretir)
import puanlama_sablonu, kontrol_listeleri, sozluk, vaka, vaka_ozet

KOK = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
puanlama_sablonu.build(f"{KOK}/04-8-kategorili-puanlama-sablonu.xlsx", ornek=True)
kontrol_listeleri.build()
sozluk.build()
vaka.tablolar()
vaka_ozet.build()

# openpyxl formüllerin sonucunu yazmaz. LibreOffice ile hesaplatıp kaydediyoruz ki
# dosya önizleyicilerde ve Excel'in korumalı görünümünde de değerler görünsün.
for f in sorted(glob.glob(f"{KOK}/**/*.xlsx", recursive=True)):
    subprocess.run(["soffice", "--headless", "--norestore", "--convert-to", "xlsx", "--outdir", os.path.dirname(f), f],
                   check=False, capture_output=True)
print("Tamam")
