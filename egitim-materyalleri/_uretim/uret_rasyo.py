import os
from rasyo_sablonu import build
from ccola_girdi import girdiler, ornek_2025
OUT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
build(f"{OUT}/01-rasyo-hesap-sablonu.xlsx", ornek=ornek_2025())
build(f"{OUT}/07-vaka-sirketi-CCOLA/CCOLA-rasyo-hesabi-2023-2025.xlsx",
      donemler=["2023/12", "2024/12", "2025/12"], veri=girdiler(), ornek=ornek_2025(), baslik_ek=", CCOLA")
