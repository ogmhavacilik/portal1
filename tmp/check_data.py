import json

with open('data/can_yelegi_stok.json', 'r', encoding='utf-8') as f:
    cy = json.load(f)
with open('data/spare_air_stok.json', 'r', encoding='utf-8') as f:
    sa = json.load(f)
with open('data/helmet_kit_stok.json', 'r', encoding='utf-8') as f:
    hk = json.load(f)
with open('data/personnel_list.json', 'r', encoding='utf-8') as f:
    p = json.load(f)

print(f"Loaded: CY={len(cy)}, SA={len(sa)}, HK={len(hk)}, P={len(p)}")
