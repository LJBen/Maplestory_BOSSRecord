import os
from typing import Dict, List, Optional

import pandas as pd
import xlwings as xw

TARGET_SHEET = "挑戰者S3"

# 將 GUI 內的核心欄位對應到 Excel 裡的六轉位置
CORE_COLUMN_MAP = {
    "hex_skill": "六轉起源-左上",
    "hex_master": "六轉精通-右上",
    "hex_enhance": "六轉強化-左下",
    "hex_shared": "六轉強化-右下",
}

# 預設 GUI 字段與 Excel 標題的對應
DEFAULT_COLUMN_MAP = {
    "character_class": "職業",
    "character_level": "等級",
    "stat_value_display": "戰力值",
    "boss": "BOSS",
    "clear_time": "通關時間",
    "note": "備註",
    "規範": "規範",
    "永續": "永續",
    "祕笈(紅)": "祕笈(紅/綠/橘)",
    "祕笈(綠)": "祕笈(紅/綠/橘)",
    "祕笈(橘)": "祕笈(紅/綠/橘)",
    "創世": "創世",
    "武公": "武公",
    "天上": "天上",
    "挑戰者": "挑戰者",
    "一般服": "一般服",
}

SECRET_SCROLL_KEYS = ["祕笈(紅)", "祕笈(綠)", "祕笈(橘)"]


def load_challenger_sheet(excel_path: str, sheet_name: str = TARGET_SHEET) -> pd.DataFrame:
    """讀取挑戰者S3 工作表並回傳 DataFrame 。"""
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel 檔案不存在: {excel_path}")
    app = xw.App(visible=False)
    wb = None
    try:
        wb = app.books.open(excel_path)
        if sheet_name not in [s.name for s in wb.sheets]:
            raise ValueError(f"工作表 {sheet_name} 不存在於 {excel_path}")
        sht = wb.sheets[sheet_name]
        # 以 A1 區域展開為表格讀入 DataFrame
        df = sht.range("A1").options(pd.DataFrame, header=1, index=False, expand="table").value
        return df
    finally:
        if wb is not None:
            wb.close()
        app.quit()


def get_sheet_headers(sht) -> List[str]:
    """取得指定 xlwings 工作表的欄位標題列，並填補合併儲存格空白值。"""
    last_col = sht.api.UsedRange.Columns.Count
    header_range = sht.range((1, 1), (1, last_col))
    values = header_range.value
    if values is None:
        return []

    headers = []
    if isinstance(values, list):
        for raw in values:
            headers.append(str(raw).strip() if raw is not None else "")
    else:
        headers = [str(values).strip()]

    # 若欄位標題列出現合併儲存格的 None/空值，向前補值
    last_seen = ""
    for idx, value in enumerate(headers):
        if value:
            last_seen = value
        else:
            headers[idx] = last_seen

    # 若展開後出現連續相同標題，視為同一個合併欄位，僅保留第一個欄位標題
    deduped_headers: List[str] = []
    for header in headers:
        if not deduped_headers or header != deduped_headers[-1]:
            deduped_headers.append(header)
    return deduped_headers


def build_excel_row(gui_data: Dict[str, object], column_map: Optional[Dict[str, str]] = None) -> Dict[str, object]:
    """將 GUI 資料轉換成 Excel 欄位對應的字典。"""
    if column_map is None:
        column_map = DEFAULT_COLUMN_MAP

    row: Dict[str, object] = {}

    for gui_key, excel_key in column_map.items():
        if gui_key in gui_data:
            row[excel_key] = gui_data.get(gui_key, "")

    # 六轉核心對應
    for gui_key, excel_key in CORE_COLUMN_MAP.items():
        if gui_key in gui_data:
            row[excel_key] = gui_data.get(gui_key, "")

    # 組合 難度 + 王 -> BOSS（中間一個空格）
    diff = gui_data.get("difficulty")
    boss_name = gui_data.get("boss")
    if diff or boss_name:
        combined = ""
        if diff:
            combined += str(diff)
        if boss_name:
            if combined:
                combined += " " + str(boss_name)
            else:
                combined = str(boss_name)
        row["BOSS"] = combined

    # 若 Excel 使用合併標題 祕笈(紅/綠/橘)，只要任一祕笈選項勾選就視為該欄為 True
    secret = list()
    for key in SECRET_SCROLL_KEYS:
        if key in gui_data:
            secret.append(gui_data.get(key, False))
    row["祕笈(紅/綠/橘)"] = secret 

    # 確保其餘勾選欄位以 True/False 寫入 Excel（使用 Excel 的核取方塊連結儲存格）
    checkbox_keys = [
        "創世", "武公", "天上", "挑戰者", "一般服",
    ]
    for gui_key in checkbox_keys:
        excel_key = column_map.get(gui_key) if column_map else DEFAULT_COLUMN_MAP.get(gui_key)
        if excel_key:
            row[excel_key] = bool(gui_data.get(gui_key, False))

    return row


def append_to_challenger_sheet(
    excel_path: str,
    gui_data: Dict[str, object],
    sheet_name: str = TARGET_SHEET,
) -> None:
    """將 GUI 資料寫入挑戰者S3 工作表的下一個空白列。"""
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel 檔案不存在: {excel_path}")

    app = xw.App(visible=False)
    wb = None
    try:
        wb = app.books.open(excel_path)
        if sheet_name not in [s.name for s in wb.sheets]:
            raise ValueError(f"工作表 {sheet_name} 不存在於 {excel_path}")
        sht = wb.sheets[sheet_name]

        headers = get_sheet_headers(sht)
        row_data = build_excel_row(gui_data)
        
        data=[]
        for header in headers:
            if header == "祕笈(紅/綠/橘)":
                for key in row_data.get(header, ""):
                    data.append(key)
            else:
                data.append(row_data.get(header, ""))

        new_row = data

        # 從 A 欄找出第一個空白列作為新資料列，只看已有資料範圍
        used = sht.api.UsedRange
        last_row = int(used.Row + used.Rows.Count - 1)
        next_row = 2
        if last_row >= 2:
            a_values = sht.range((2, 1), (last_row, 1)).value
            if not isinstance(a_values, list):
                a_values = [a_values]
            for offset, cell_value in enumerate(a_values, start=2):
                if cell_value is None or str(cell_value).strip() == "":
                    next_row = offset
                    break
            else:
                next_row = last_row + 1

        # 寫入從第一欄開始的一整列
        sht.range((next_row, 1)).value = new_row
        wb.save()
    finally:
        if wb is not None:
            wb.close()
        app.quit()


