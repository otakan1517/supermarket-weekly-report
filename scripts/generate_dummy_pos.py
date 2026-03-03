import argparse
from datetime import datetime, timedelta
import numpy as np
import pandas as pd

def make_master():
    stores = ["品川店", "大井町店", "五反田店"]
    dept_cat = {
        "青果": ["野菜", "果物"],
        "精肉": ["牛肉", "豚肉", "鶏肉"],
        "鮮魚": ["刺身", "切り身"],
        "惣菜": ["弁当", "揚げ物", "サラダ"],
        "日配": ["牛乳・乳製品", "豆腐", "納豆"],
        "ベーカリー": ["パン", "菓子パン"],
        "加工食品": ["麺類", "調味料", "缶詰"],
        "飲料": ["水", "お茶", "炭酸"],
    }

    products = []
    rng = np.random.default_rng(42)
    p = 1
    for dept, cats in dept_cat.items():
        for cat in cats:
            n = rng.integers(8, 18)
            for i in range(n):
                products.append({
                    "department": dept,
                    "category": cat,
                    "product_code": f"P{p:06d}",
                    "product_name": f"{cat}商品{i+1}"
                })
                p += 1
    return stores, pd.DataFrame(products)

def generate(rows: int, start_date: str, days: int, seed: int):
    rng = np.random.default_rng(seed)
    stores, prod = make_master()

    start = datetime.fromisoformat(start_date)
    dt_days = rng.integers(0, days, size=rows)
    dt_hours = rng.integers(9, 22, size=rows)
    dt_mins = rng.integers(0, 60, size=rows)
    dt = [start + timedelta(days=int(d), hours=int(h), minutes=int(m))
          for d, h, m in zip(dt_days, dt_hours, dt_mins)]

    prod_idx = rng.integers(0, len(prod), size=rows)
    picked = prod.iloc[prod_idx].reset_index(drop=True)

    qty = rng.integers(1, 4, size=rows)
    base_price = rng.integers(80, 1200, size=rows)
    unit_price = (base_price // 10) * 10

    disc_rate = rng.choice([0.0, 0.0, 0.0, 0.1, 0.2, 0.3], size=rows)
    discount_yen = ((unit_price * qty) * disc_rate).astype(int)

    sales_yen = unit_price * qty - discount_yen

    margin = rng.uniform(0.15, 0.35, size=rows)
    cost_yen = (sales_yen * (1 - margin)).astype(int)

    store = rng.choice(stores, size=rows)
    tx = [f"T{d.strftime('%Y%m%d%H%M')}_{s}_{i//3:06d}"
          for i, (d, s) in enumerate(zip(dt, store))]

    df = pd.DataFrame({
        "datetime": dt,
        "store": store,
        "transaction_id": tx,
        "department": picked["department"],
        "category": picked["category"],
        "product_code": picked["product_code"],
        "product_name": picked["product_name"],
        "qty": qty,
        "sales_yen": sales_yen,
        "cost_yen": cost_yen,
        "discount_yen": discount_yen,
    })
    return df.sort_values(["datetime", "store"]).reset_index(drop=True)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--rows", type=int, default=20000)
    ap.add_argument("--start_date", type=str, default="2026-01-01")
    ap.add_argument("--days", type=int, default=56)
    ap.add_argument("--seed", type=int, default=42)
    ap.add_argument("--out", type=str, default="data/raw/sample_pos.xlsx")
    args = ap.parse_args()

    df = generate(args.rows, args.start_date, args.days, args.seed)
    df.to_excel(args.out, index=False)
    print(f"saved: {args.out} / rows={len(df)}")

if __name__ == "__main__":
    main()
