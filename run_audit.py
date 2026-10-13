#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""GDG NTPU W4：跨部門報銷抓鬼（pandas 版）。

直接執行：python run_audit.py
原始資料在 data/；所有產出只寫入 outputs/，不覆寫假資料。
"""

from pathlib import Path
import pandas as pd

HERE = Path(__file__).resolve().parent
INPUT_DIR = HERE / "data"
OUTPUT_DIR = HERE / "outputs"
DEPARTMENTS = ["活動部", "器材部", "公關部", "總務部"]
EXPECTED_COLUMNS = ["項次", "發票號碼", "品名", "金額", "申請人"]
AUDIT_LABEL = "疑似重複／需人工覆核"
NORMAL_LABEL = "待一般審核"


def step_1_merge() -> pd.DataFrame:
    """讀取四個部門的 Excel，新增部門欄後合併。"""
    parts = []
    for department in DEPARTMENTS:
        path = INPUT_DIR / f"報銷_{department}.xlsx"
        if not path.is_file():
            raise FileNotFoundError(f"缺少假資料：{path.name}。請確認 data 資料夾是否完整。")
        part = pd.read_excel(path, engine="openpyxl", dtype={"發票號碼": str})
        if part.columns.tolist() != EXPECTED_COLUMNS:
            raise ValueError(f"{path.name} 欄位不符合規格，實際為：{part.columns.tolist()}")
        if part[EXPECTED_COLUMNS].isna().any().any():
            raise ValueError(f"{path.name} 有空白欄位，本練習未設計處理缺值。")
        part["發票號碼"] = part["發票號碼"].str.strip()
        if part["發票號碼"].eq("").any():
            raise ValueError(f"{path.name} 有空白發票號碼。")
        part["金額"] = pd.to_numeric(part["金額"], errors="raise")
        if part["金額"].le(0).any():
            raise ValueError(f"{path.name} 有非正數金額。")
        part["部門"] = department
        parts.append(part)
        print(f"  已讀取：{path.name}，{len(part)} 筆")

    merged = pd.concat(parts, ignore_index=True)
    # 保留原始工作表中的「項次」，便於追溯各部門的原始列。
    merged = merged[["部門", "項次", "發票號碼", "品名", "金額", "申請人"]]
    print(f"[G1] 合併完成：{len(parts)} 份 Excel，共 {len(merged)} 筆、{len(merged.columns)} 欄。")
    return merged


def step_2_and_3_audit(merged: pd.DataFrame):
    """找出同發票的所有紀錄，標記而非刪除，並製作部門統計。"""
    duplicate_mask = merged.duplicated(subset=["發票號碼"], keep=False)
    # 保留來源記錄，避免直接刪掉可能需要人工覆核的申請。
    audit = merged.copy()
    audit["稽核狀態"] = NORMAL_LABEL
    audit.loc[duplicate_mask, "稽核狀態"] = AUDIT_LABEL
    suspicious = (audit.loc[duplicate_mask]
                  .sort_values(["發票號碼", "部門"])
                  .reset_index(drop=True))
    summary = (audit.groupby("部門", as_index=False)
               .agg(申請筆數=("發票號碼", "size"), 申請金額合計=("金額", "sum")))
    summary["待覆核筆數"] = summary["部門"].map(
        audit.loc[duplicate_mask].groupby("部門").size()).fillna(0).astype(int)
    print(f"[G2] 重複發票：{suspicious['發票號碼'].nunique()} 組，涉及 {len(suspicious)} 筆。");
    for invoice, group in suspicious.groupby("發票號碼", sort=True):
        print(f"  發票 {invoice}：")
        for r in group.itertuples(index=False):
            print(f"    {r.部門}／{r.申請人}／{r.品名}／{r.金額} 元")
    print("[G3] 稽核標記完成，原始 12 筆都保留；以下金額是申請總額，不是已核准支出。")
    for r in summary.itertuples(index=False):
        print(f"  {r.部門}：{r.申請筆數} 筆，{r.申請金額合計} 元，待覆核 {r.待覆核筆數} 筆")
    return audit, suspicious, summary


def step_4_export(merged, audit, suspicious, summary):
    OUTPUT_DIR.mkdir(exist_ok=True)
    for filename, frame in [
        ("01_合併總表.csv", merged),
        ("02_疑似重複發票.csv", suspicious),
        ("03_稽核明細.csv", audit),
        ("04_部門摘要.csv", summary),
    ]:
        frame.to_csv(OUTPUT_DIR / filename, index=False, encoding="utf-8", lineterminator="\n")
        print(f"[G4] 已輸出 outputs/{filename}：{len(frame)} 筆")

    report = OUTPUT_DIR / "跨部門報銷稽核報告.xlsx"
    with pd.ExcelWriter(report, engine="openpyxl") as writer:
        audit.to_excel(writer, sheet_name="稽核明細", index=False)
        suspicious.to_excel(writer, sheet_name="異常清單", index=False)
        summary.to_excel(writer, sheet_name="部門摘要", index=False)

    note = (
        "GDG NTPU W4｜pandas 跨部門報銷稽核\n"
        f"合併筆數：{len(merged)}\n"
        f"疑似重複發票組數：{suspicious['發票號碼'].nunique()}\n"
        f"涉及紀錄：{len(suspicious)}\n"
        f"申請金額合計：{audit['金額'].sum()} 元（非實際核准支出）\n"
        "警語：同號發票不代表已證實重複請款，必須進行人工覆核。\n"
    )
    (OUTPUT_DIR / "稽核結果摘要.txt").write_text(note, encoding="utf-8")
    print(f"[G4] Excel 報告：outputs/{report.name}")
    return report


def step_5_verify(report: Path, audit: pd.DataFrame, suspicious: pd.DataFrame, summary: pd.DataFrame):
    """重新讀取 Excel，驗證所有儲存格資料，而非只檢查檔案存在。"""
    from pandas.testing import assert_frame_equal
    for sheet_name, expected in [("稽核明細", audit), ("異常清單", suspicious), ("部門摘要", summary)]:
        actual = pd.read_excel(report, sheet_name=sheet_name, engine="openpyxl", dtype={"發票號碼":str})
        assert_frame_equal(actual, expected, check_dtype=False)
        print(f"[G5] PASS：{sheet_name}，{len(actual)} 筆，內容一致。")
    print("[完成] 請開啟 outputs/跨部門報銷稽核報告.xlsx；原始假資料未修改。")


def main():
    print("=" * 55)
    print("GDG NTPU｜報銷抓鬼遊戲：跨部門發票稽核（pandas 版）")
    print("=" * 55)
    merged = step_1_merge()
    audit, suspicious, summary = step_2_and_3_audit(merged)
    report = step_4_export(merged, audit, suspicious, summary)
    step_5_verify(report, audit, suspicious, summary)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"[失敗] {type(exc).__name__}: {exc}")
        raise SystemExit(1) from exc
