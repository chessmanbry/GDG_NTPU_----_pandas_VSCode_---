#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""供講師／學員驗收報銷抓鬼成果。執行：python verify_results.py"""
from pathlib import Path
import pandas as pd
from pandas.testing import assert_frame_equal

ROOT = Path(__file__).resolve().parent
OUTPUTS = ROOT / "outputs"
REPORT = OUTPUTS / "跨部門報銷稽核報告.xlsx"
EXPECTED_INVOICES = {"INV-2026-0881", "INV-2026-0943"}
EXPECTED_TOTALS = {"活動部": 3550, "器材部": 5080, "公關部": 7650, "總務部": 12240}


def main():
    if not REPORT.is_file():
        raise FileNotFoundError("找不到稽核報告，請先執行 run_audit.py 或雙擊執行腳本。")
    detail = pd.read_excel(REPORT, sheet_name="稽核明細", dtype={"發票號碼":str})
    suspicious = pd.read_excel(REPORT, sheet_name="異常清單", dtype={"發票號碼":str})
    summary = pd.read_excel(REPORT, sheet_name="部門摘要")
    assert len(detail)==12, f"稽核明細應為 12 筆，實際 {len(detail)}"
    assert len(suspicious)==4, f"異常清單應為 4 筆，實際 {len(suspicious)}"
    assert len(summary)==4, f"部門摘要應為 4 筆，實際 {len(summary)}"
    assert set(suspicious["發票號碼"]) == EXPECTED_INVOICES, "疑似重複發票號碼不一致"
    assert detail["稽核狀態"].eq("疑似重複／需人工覆核").sum() == 4, "待覆核標記有誤"
    actual_totals = dict(zip(summary["部門"],summary["申請金額合計"]))
    assert actual_totals == EXPECTED_TOTALS, f"部門金額不一致：{actual_totals}"
    assert int(summary["申請金額合計"].sum()) == 28520, "申請金額合計不一致"
    expected_files = [
        ("01_合併總表.csv", 12), ("02_疑似重複發票.csv", 4),
        ("03_稽核明細.csv", 12), ("04_部門摘要.csv", 4),
    ]
    for name, expected_n in expected_files:
        p = OUTPUTS / name
        assert p.exists(), f"缺少階段資料：{name}"
        csv_df = pd.read_csv(p, dtype={"發票號碼":str})
        assert len(csv_df) == expected_n, f"{name} 應為 {expected_n} 筆"
        if name == "03_稽核明細.csv":
            assert_frame_equal(csv_df, detail, check_dtype=False)
    print("[PASS] 四份假資料合併成 12 筆。")
    print("[PASS] 發現兩組發票號碼：INV-2026-0881、INV-2026-0943。")
    print("[PASS] 保留全部原始紀錄，其中 4 筆標記待覆核。")
    print("[PASS] 四個部門申請金額合計：28,520 元。")
    print("[PASS] 四份中途 CSV 及三張 Excel 工作表符合預期。")
    print("注意：重複發票只是疑點，不足以單獨認定違規。")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"[FAIL] {type(exc).__name__}：{exc}")
        raise SystemExit(1) from exc
