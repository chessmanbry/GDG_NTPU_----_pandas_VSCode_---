#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""使用 pandas 重建四份教學用假資料（已隨 ZIP 附上，不必先執行）。

執行方式：python generate_fake_data.py
若要覆寫原有假資料：python generate_fake_data.py --force
"""
import argparse
from pathlib import Path
import pandas as pd

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
# 固定、虛構的發票資料，與教材隨附四份 Excel 的資料內容相同。
RECORDS = {
    "活動部": [
        (1, "INV-2026-0881", "慶功宴飲料與點心", 1200, "陳大文"),
        (2, "INV-2026-0102", "團康破冰道具一批", 850, "黃小安"),
        (3, "INV-2026-0304", "活動主持無線麥克風租借", 1500, "陳大文"),
    ],
    "器材部": [
        (1, "INV-2026-0881", "高功率延長線與轉接頭", 1200, "林小豪"),
        (2, "INV-2026-0505", "網路交換器與跳線組", 3200, "張阿強"),
        (3, "INV-2026-0606", "大會專用簡報筆", 680, "林小豪"),
    ],
    "公關部": [
        (1, "INV-2026-0943", "大會宣傳手冊印刷", 4500, "王小美"),
        (2, "INV-2026-0707", "講師感謝狀與相框", 950, "王小美"),
        (3, "INV-2026-0808", "贊助商專屬迎賓禮盒", 2200, "蔡小彤"),
    ],
    "總務部": [
        (1, "INV-2026-0943", "全校宣傳大型海報輸出", 4500, "張志明"),
        (2, "INV-2026-1111", "活動工作人員便當 (80份)", 7200, "張志明"),
        (3, "INV-2026-1212", "慶祝拉炮與文具消耗品", 540, "周小玲"),
    ],
}
COLS = ["項次", "發票號碼", "品名", "金額", "申請人"]


def main():
    parser = argparse.ArgumentParser(description="重新產生四份練習用虛構報銷 Excel")
    parser.add_argument("--force", action="store_true", help="覆寫已存在的 data/報銷_*.xlsx")
    args = parser.parse_args()
    DATA.mkdir(exist_ok=True)
    targets = [DATA / f"報銷_{dept}.xlsx" for dept in RECORDS]
    exists = [f.name for f in targets if f.exists()]
    if exists and not args.force:
        print("資料檔已隨專案提供，沒有覆寫：" + "、".join(exists))
        print("若要重新產生並覆寫原始假資料，請加 --force。")
        return
    for department, rows in RECORDS.items():
        out = DATA / f"報銷_{department}.xlsx"
        pd.DataFrame(rows, columns=COLS).to_excel(out, sheet_name="報銷明細", index=False, engine="openpyxl")
        print(f"產生：data/{out.name}（{len(rows)} 筆）")
    print("完成：共 4 份、12 筆假資料；其中兩組發票重複。")


if __name__ == "__main__":
    main()
