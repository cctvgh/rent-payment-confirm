# -*- coding: utf-8 -*-
"""
构建/修正 10040102 房屋/土地/场地租赁 租金付款确认表（GDZDY-011 模板）。

核心职责（含跨工作簿外部链接的 COM 修正流程）：
  1. 在报账总表(.xls)的「202503报账总表湛江电信基站租金费用确认表」中，
     按合同编号（D 列）定位该合同对应的记录行号。
  2. 打开目标付款确认表(.xlsx)，把 A6/C6/D6/G6/H6/I6 的外部链接公式
     行号统一改为该记录行（用 Excel COM 写公式，保留 externalLinks 部件）。
  3. 从合同台账(.xlsx)补填 B6 详细地点（房屋地址列）与 F6 银行帐号（开户账号列）。
  4. 可选：按 --rename 重命名文件为「10040102房屋土地场地租赁<地点><金额>.xlsx」。

注意：
  - 禁止用 openpyxl 保存含外部链接的目标文件（会丢 externalLinks）。
  - 目标文件若被 WPS/Excel 占用（wb.ReadOnly 为 True），需先让用户关闭。
  - 报账总表默认 sheet 名为「202503报账总表湛江电信基站租金费用确认表」。
"""

import argparse
import os
import re
import sys

try:
    import pandas as pd
    import win32com.client
except ImportError as e:  # pragma: no cover
    sys.exit("缺少依赖: %s （需要 pandas 与 pywin32）" % e)

# 目标单元格 -> 报账总表源列字母（表头：A序号 B报账期间 C基站地址 D合同编号 E合同名称
# F租赁期间起 G租赁期间止 H电信物理站址名称 I业主名称 J付款金额）
CELL_COL_MAP = {"A6": "C", "C6": "D", "D6": "I", "G6": "F", "H6": "G", "I6": "J"}
# 报账总表 pandas 列索引（0 基）
COL_CONTRACT_NO = 3      # D 合同编号
COL_PERIOD = 1           # B 报账期间
COL_AMOUNT = 9           # J 付款金额
COL_BASE_STATION = 2     # C 基站地址

REPORT_SHEET = "202503报账总表湛江电信基站租金费用确认表"
LEDGER_SHEET = "合同台账查询"


def find_report_files(folder):
    """在文件夹内定位报账总表(.xls)与合同台账(.xlsx)。"""
    report = ledger = None
    for f in os.listdir(folder):
        low = f.lower()
        if low.endswith(".xls") and ("费用确认表" in f or "报帐总表" in f or "报账总表" in f):
            report = os.path.join(folder, f)
        elif low.endswith(".xlsx") and ("合同" in f and ("台账" in f or "台帐" in f)):
            ledger = os.path.join(folder, f)
    return report, ledger


def locate_record_row(report_path, contract_no, period=None):
    """在报账总表中按合同编号（可加报账期间）定位 Excel 行号。
    返回 (excel_row, row_info_dict)；找不到返回 (None, None)。
    """
    df = pd.read_excel(report_path, sheet_name=REPORT_SHEET, header=None)
    hits = []
    for idx, row in df.iterrows():
        if str(row.iloc[COL_CONTRACT_NO]).strip() == contract_no.strip():
            if period is None or str(row.iloc[COL_PERIOD]).strip() == str(period).strip():
                hits.append((idx, row))
    if not hits:
        return None, None
    # 优先取报账期间最大（最新）的一条
    idx, row = max(hits, key=lambda t: str(t[1].iloc[COL_PERIOD]))
    excel_row = idx + 1  # pandas 0 基 -> Excel 1 基
    info = {
        "period": row.iloc[COL_PERIOD],
        "station": row.iloc[COL_BASE_STATION],
        "contract": row.iloc[COL_CONTRACT_NO],
        "amount": row.iloc[COL_AMOUNT],
    }
    return excel_row, info


def load_contract(ledger_path, contract_no):
    """从合同台账提取合同信息：房屋地址、开户名称、开户账号等。
    台账 sheet 名自动匹配（含"合同"且含"台账/台帐"）。
    """
    xl = pd.ExcelFile(ledger_path)
    sheet = next((s for s in xl.sheet_names if "合同" in s and ("台账" in s or "台帐" in s)), None)
    if sheet is None:
        sys.exit("台账文件中未找到合同台账 sheet，当前 sheet: %s" % xl.sheet_names)
    df = xl.parse(sheet, header=0)
    col = {c: df.columns.get_loc(c) for c in ["合同编码", "合同名称", "对方名称",
                                               "开户名称", "开户账号", "房屋地址", "租赁面积"]}
    for _, row in df.iterrows():
        if str(row.iloc[col["合同编码"]]).strip() == contract_no.strip():
            return {c: ("" if pd.isna(row.iloc[col[c]]) else str(row.iloc[col[c]]).strip())
                    for c in col}
    return None


def fix_formulas_and_fill(target_path, report_path, excel_row, detail_addr, bank_acct):
    """用 Excel COM 修正公式行号并补填手工字段，保留外部链接。"""
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    try:
        wb = excel.Workbooks.Open(target_path, UpdateLinks=0, ReadOnly=False)
        if wb.ReadOnly:
            sys.exit("目标文件正被 WPS/Excel 占用（ReadOnly=True），请先关闭再运行。")
        ws = wb.Worksheets(1)
        for cell, col in CELL_COL_MAP.items():
            old = ws.Range(cell).Formula
            # 匹配 "[文件名]Sheet名'!C12" 末尾行号
            new = re.sub(r"!([A-Z]+)\d+$", r"!\g<1>%d" % excel_row, old)
            ws.Range(cell).Formula = new
            print("%s: %s -> %s" % (cell, old, new))
        if detail_addr:
            ws.Range("B6").Value = detail_addr
        if bank_acct:
            ws.Range("F6").Value = bank_acct
        wb.Save()
        print("已保存: %s" % target_path)
    finally:
        if wb is not None:
            try:
                wb.Close(False)
            except Exception:
                pass
        excel.Quit()


def verify_values(target_path):
    """重新打开（更新外部链接）并打印第 6 行各单元格取值。"""
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    try:
        wb = excel.Workbooks.Open(target_path, UpdateLinks=3, ReadOnly=True)
        ws = wb.Worksheets(1)
        for cell in ["A6", "B6", "C6", "D6", "E6", "F6", "G6", "H6", "I6"]:
            print("%s = %r" % (cell, ws.Range(cell).Value))
    finally:
        if wb is not None:
            try:
                wb.Close(False)
            except Exception:
                pass
        excel.Quit()


def main():
    ap = argparse.ArgumentParser(description="构建 10040102 租金付款确认表（GDZDY-011）")
    ap.add_argument("--folder", required=True, help="报账工作文件夹")
    ap.add_argument("--contract", required=True, help="合同编号，如 GDZJG2600487AGN00")
    ap.add_argument("--period", default=None, help="报账期间（可选），如 202605")
    ap.add_argument("--report", default=None, help="报账总表路径（默认自动查找）")
    ap.add_argument("--ledger", default=None, help="合同台账路径（默认自动查找）")
    ap.add_argument("--file", default=None, help="目标确认表 xlsx（默认按文件名含 10040102 自动匹配）")
    ap.add_argument("--rename", default=None, help="保存后重命名为该文件名（含扩展名）")
    args = ap.parse_args()

    report = args.report or find_report_files(args.folder)[0]
    ledger = args.ledger or find_report_files(args.folder)[1]
    if not report or not os.path.exists(report):
        sys.exit("未找到报账总表 .xls，请用 --report 指定")
    if not ledger or not os.path.exists(ledger):
        sys.exit("未找到合同台账 .xlsx，请用 --ledger 指定")

    excel_row, info = locate_record_row(report, args.contract, args.period)
    if excel_row is None:
        sys.exit("报账总表未找到合同编号 %s 的记录" % args.contract)
    print("定位记录行号(Excel): %d, 报账期间: %s, 金额: %s" %
          (excel_row, info["period"], info["amount"]))

    contract = load_contract(ledger, args.contract)
    if contract is None:
        print("警告: 台账未找到合同 %s，详细地点/银行账号将不填充" % args.contract)

    # 目标文件
    target = args.file
    if not target:
        candidates = [os.path.join(args.folder, f) for f in os.listdir(args.folder)
                      if f.lower().endswith(".xlsx") and "10040102" in f]
        if len(candidates) != 1:
            sys.exit("文件夹内需恰好一个 10040102*.xlsx，或用 --file 指定，当前候选: %s" % candidates)
        target = candidates[0]

    fix_formulas_and_fill(target, report, excel_row,
                          contract.get("房屋地址") if contract else None,
                          contract.get("开户账号") if contract else None)

    if args.rename:
        new_path = os.path.join(os.path.dirname(target), args.rename)
        if new_path.lower() != target.lower():
            os.rename(target, new_path)
            target = new_path
            print("已重命名 -> %s" % target)

    print("--- 验证结果 ---")
    verify_values(target)


if __name__ == "__main__":
    main()