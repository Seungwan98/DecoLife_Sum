# -*- coding: utf-8 -*-
"""
엑셀 정산 집계기 (단일 파일 버전)
- GUI(Tkinter) + 엑셀 집계 로직(pandas) 하나로 통합
- 기능:
  1) 파일 선택(.xlsx) → 파일명 표시
  2) 시트 이름(옵션) 입력
  3) "집계 실행 → 저장" 버튼으로 결과 엑셀 저장
"""

import os
import re
import traceback
from typing import Optional, Tuple, List

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# ====== 집계 로직 ======

REQUIRED = ["등록상품명", "정산대상액", "옵션명"]

def _norm(s: str) -> str:
    s = str(s)
    s = s.replace("\u200b", "").replace("\ufeff", "")
    s = re.sub(r"\s+", "", s)
    return s

def resolve_sheet_name(excel_path: str, sheet_name: Optional[str]) -> str:
    if sheet_name:
        return sheet_name
    xls = pd.ExcelFile(excel_path)
    if not xls.sheet_names:
        raise ValueError("엑셀 파일에 시트가 없습니다.")
    return xls.sheet_names[0]

def detect_header_row(
    excel_path: str,
    sheet_name: Optional[str] = None,
    search_rows: int = 50
) -> Tuple[int, pd.DataFrame]:
    sheet = resolve_sheet_name(excel_path, sheet_name)
    raw = pd.read_excel(excel_path, sheet_name=sheet, header=None, dtype=str)
    targets = [_norm(t) for t in REQUIRED]
    header_idx = None
    for i in range(min(search_rows, len(raw))):
        row_norm = [_norm(v) for v in raw.iloc[i].tolist()]
        if all(any(t in c for c in row_norm) for t in targets):
            header_idx = i
            break
    if header_idx is None:
        header_idx = 0
    df = pd.read_excel(excel_path, sheet_name=sheet, header=header_idx)
    return header_idx, df

def map_columns_loose(df: pd.DataFrame) -> List[str]:
    clean_to_orig = {_norm(c): c for c in df.columns}
    found = {}
    for key in REQUIRED:
        target = _norm(key)
        for clean, orig in clean_to_orig.items():
            if target in clean:  # 부분 일치 허용
                found[key] = orig
                break
        if key not in found:
            raise KeyError(f"'{key}' 컬럼을 찾지 못했습니다. 현재 컬럼: {list(df.columns)}")
    return [found[k] for k in REQUIRED]  # [등록상품명, 정산대상액, 옵션명]

def count_combo(excel_path: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
    _, df = detect_header_row(excel_path, sheet_name=sheet_name)
    name_col, amount_col, option_col = map_columns_loose(df)

    # 정산대상액 숫자 변환
    df[amount_col] = (
        df[amount_col]
        .astype(str)
        .replace({r"[^0-9.\-]": ""}, regex=True)  # 콤마/원 문자 제거
        .replace("", "0")
        .astype(float)
    )

    # 조합별 갯수
    result = (
        df.groupby([name_col, amount_col, option_col])
          .size()
          .reset_index(name="갯수")
    )

    # 등록상품명별 총갯수(정렬용)
    result["__total_per_name"] = result.groupby(name_col)["갯수"].transform("sum")

    # 합계금액 = 정산대상액 * 갯수
    result["합계금액"] = result[amount_col] * result["갯수"]

    # 정렬 (상품별 총갯수 ↓, 등록상품명/옵션명/금액 ↑)
    result = result.sort_values(
        ["__total_per_name", name_col, option_col, amount_col],
        ascending=[False, True, True, True]
    ).drop(columns="__total_per_name")

    # 컬럼 순서
    result = result[[name_col, option_col, amount_col, "갯수", "합계금액"]]

    # 총합계 행 추가
    total_sum = result["합계금액"].sum()
    total_row = pd.DataFrame([[None, None, "총 합계금액", None, total_sum]],
                             columns=result.columns)
    result = pd.concat([result, total_row], ignore_index=True)

    return result

# ====== GUI(App) ======

APP_TITLE = "엑셀 정산 집계기 (단일 파일)"
DEFAULT_OUT_NAME = "result_output.xlsx"

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("580x260")
        self.resizable(False, False)

        self.selected_path: Optional[str] = None

        self._build_widgets()

    def _build_widgets(self):
        pad = {"padx": 12, "pady": 8}

        # 파일 선택
        frm_file = tk.Frame(self)
        frm_file.pack(fill="x", **pad)
        tk.Label(frm_file, text="엑셀 파일(.xlsx):", anchor="w", width=16).pack(side="left")
        self.lbl_filename = tk.Label(frm_file, text="(선택 안 됨)", anchor="w")
        self.lbl_filename.pack(side="left", fill="x", expand=True, padx=(6, 6))
        tk.Button(frm_file, text="파일 선택…", command=self.on_pick_file).pack(side="right")

        # 시트 이름
        frm_sheet = tk.Frame(self)
        frm_sheet.pack(fill="x", **pad)
        tk.Label(frm_sheet, text="시트 이름(옵션):", anchor="w", width=16).pack(side="left")
        self.ent_sheet = tk.Entry(frm_sheet)
        self.ent_sheet.pack(side="left", fill="x", expand=True, padx=(6, 6))
        tk.Label(frm_sheet, text="(비우면 첫 번째 시트)", fg="#666").pack(side="left")

        # 버튼들
        frm_actions = tk.Frame(self)
        frm_actions.pack(fill="x", **pad)
        self.btn_run = tk.Button(frm_actions, text="집계 실행 → 저장", command=self.on_run)
        self.btn_run.pack(side="left")
        tk.Button(frm_actions, text="종료", command=self.destroy).pack(side="right")

        # 상태바
        self.status = tk.StringVar(value="준비됨")
        tk.Label(self, textvariable=self.status, anchor="w", fg="#444").pack(fill="x", padx=12, pady=(12, 10))

    def on_pick_file(self):
        path = filedialog.askopenfilename(
            title="엑셀 파일 선택",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if path:
            self.selected_path = path
            self.lbl_filename.config(text=os.path.basename(path))
            self.status.set("파일 선택 완료")

    def on_run(self):
        if not self.selected_path:
            messagebox.showwarning(APP_TITLE, "먼저 엑셀 파일을 선택하세요.")
            return

        initial_dir = os.path.dirname(self.selected_path) if self.selected_path else os.getcwd()
        out_path = filedialog.asksaveasfilename(
            title="결과 저장 위치 선택",
            initialdir=initial_dir,
            initialfile=DEFAULT_OUT_NAME,
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not out_path:
            self.status.set("저장 취소됨")
            return

        sheet_name = self.ent_sheet.get().strip() or None

        try:
            self._toggle_ui(False)
            self.status.set("집계 중...")

            df_result = count_combo(self.selected_path, sheet_name=sheet_name)
            df_result.to_excel(out_path, index=False)

            self.status.set(f"완료: {os.path.basename(out_path)} 저장됨")
            messagebox.showinfo(APP_TITLE, f"저장 완료:\n{out_path}")

        except Exception as e:
            traceback.print_exc()
            self.status.set("실패")
            messagebox.showerror(APP_TITLE, f"에러 발생:\n{e}")

        finally:
            self._toggle_ui(True)

    def _toggle_ui(self, enable: bool):
        state = tk.NORMAL if enable else tk.DISABLED
        self.btn_run.config(state=state)

def main():
    app = App()
    app.mainloop()

if __name__ == "__main__":
    main()
