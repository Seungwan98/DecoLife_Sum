# -*- coding: utf-8 -*-
"""
엑셀 변환기 (헤더 자동 탐지 + 옵션ID 매핑 + 단가 계산 + 묶기)
- 첫번째 엑셀: 옵션ID, 매출인식일, 판매 수량(B), 정산대상액, 등록상품명
- 두번째 엑셀: 옵션ID, 코드, 윈윈상품명

처리 로직:
 1) 옵션ID 기준으로 코드/상품명 매핑
 2) 상품명(1): 기본은 2번 엑셀 윈윈상품명/상품명, 없으면 1번 엑셀 등록상품명(빨간색 표시)
 3) 상품코드(1): 기본은 2번 엑셀 코드, 없으면 옵션ID
 4) 정산대상액(상품금액)이 - 인 경우 절댓값 사용
 5) 수량이 1보다 크면 단가 = (절댓값 정산대상액 / 수량), 수량이 1이면 단가 = 절댓값 정산대상액
 6) 단가는 소수점 첫째 자리에서 반올림해서 int 값으로
 7) 같은 (거래일자, 거래처명, 상품코드(1), 상품명(1), 단가(1)) 조합은 수량(1) 합쳐서 한 줄로 묶기
"""

import os
import re
import traceback
from typing import Optional, List

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl.styles import Font


# ====== 공통 유틸 ======

def _norm(s: str) -> str:
    s = str(s)
    s = s.replace("\u200b", "").replace("\ufeff", "")
    s = re.sub(r"\s+", "", s)
    return s.lower()


def _find_col(df: pd.DataFrame, *keywords) -> str:
    for c in df.columns:
        nc = _norm(c)
        for kw in keywords:
            if _norm(kw) in nc:
                return c
    raise KeyError(f"컬럼을 찾지 못했습니다. 필요 키워드: {keywords}, 현재 컬럼: {list(df.columns)}")


def _to_number(series: pd.Series) -> pd.Series:
    return (
        series.astype(str)
              .replace({r"[^0-9.\-]": ""}, regex=True)  # 숫자/마이너스만 남기기
              .replace("", "0")
              .astype(float)
    )


def _read_with_header_detection(path: str,
                                sheet_name: Optional[str],
                                keyword_candidates: List[str],
                                search_rows: int = 50) -> pd.DataFrame:
    sheet_arg = sheet_name if sheet_name else 0
    raw = pd.read_excel(path, sheet_name=sheet_arg, header=None, dtype=str)
    targets = [_norm(k) for k in keyword_candidates]
    header_idx = 0
    for i in range(min(search_rows, len(raw))):
        row_norm = [_norm(v) for v in raw.iloc[i].tolist()]
        if any(t in cell for t in targets for cell in row_norm):
            header_idx = i
            break
    df = pd.read_excel(path, sheet_name=sheet_arg, header=header_idx)
    return df


# ====== 핵심 로직 ======

def build_result(main_path: str, map_path: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
    # 1) 원본 엑셀 읽기
    df_main = _read_with_header_detection(
        main_path, sheet_name,
        ["옵션id", "optionid", "매출인식일", "판매수량", "수량", "정산대상액", "정산 대상액", "등록상품명"]
    )
    df_map = _read_with_header_detection(
        map_path, None,
        ["옵션id", "optionid", "코드", "상품코드", "윈윈상품명", "윈윈 상품명", "상품명"]
    )

    # 2) 필요한 컬럼 찾기
    col_optid  = _find_col(df_main, "옵션id", "optionid")
    col_date   = _find_col(df_main, "매출인식일")
    col_qty    = _find_col(df_main, "판매수량", "수량")
    col_price  = _find_col(df_main, "정산대상액", "정산 대상액")
    col_regnm  = _find_col(df_main, "등록상품명")

    col_optid2 = _find_col(df_map, "옵션id", "optionid")
    col_code   = _find_col(df_map, "코드", "상품코드")
    col_name   = _find_col(df_map, "윈윈상품명", "윈윈 상품명", "상품명")

    # 3) 옵션ID로 조인
    df_main["_optkey"] = df_main[col_optid].astype(str).map(_norm)
    df_map["_optkey"]  = df_map[col_optid2].astype(str).map(_norm)

    df = pd.merge(df_main, df_map[["_optkey", col_code, col_name]], on="_optkey", how="left")

    # 4) 숫자 변환
    qty = _to_number(df[col_qty])
    amount_raw = _to_number(df[col_price])  # 정산대상액 (음수 가능)

    # 5) 상품명 / 코드 확정
    mapped_name = df[col_name]
    reg_name    = df[col_regnm]
    # 매핑 안 된 행(True) → 등록상품명 사용 + 상품명 빨간색 표시 대상
    fallback_mask = mapped_name.isna() | (mapped_name.astype(str).str.strip() == "")

    final_name = mapped_name.copy()
    final_name[fallback_mask] = reg_name[fallback_mask]

    final_code = df[col_code].copy()
    final_code[fallback_mask] = df[col_optid].astype(str)[fallback_mask]

    # 6) 단가 계산
    #   - 금액 절댓값 사용
    #   - 수량 > 1 이면 금액 / 수량
    #   - 최종 단가는 반올림 후 int
    amount_abs = amount_raw.abs()

    unit_price = amount_abs.copy()
    multi_mask = qty > 1
    unit_price[multi_mask] = amount_abs[multi_mask] / qty[multi_mask]

    # 소수점 첫째 자리에서 반올림해서 int로 변환
    unit_price = unit_price.round().astype(int)

    # 7) 중간 테이블 만들기
    mid = pd.DataFrame({
        "거래일자": df[col_date],
        "거래처명": "쿠팡-제트배송",
        "상품코드(1)": final_code.astype(str),
        "상품명(1)":   final_name,
        "수량(1)":     qty,
        "단가(1)":     unit_price,
        "상품비고(1)": "",
        "__fallback":  fallback_mask.values,   # 상품명이 등록상품명으로 대체된 경우
    })

    # 8) 같은 (거래일자, 거래처명, 상품코드(1), 상품명(1), 단가(1)) 끼리 수량 합치기
    grouped = (
        mid
        .groupby(["거래일자", "거래처명", "상품코드(1)", "상품명(1)", "단가(1)"], dropna=False)
        .agg(
            수량_합=("수량(1)", "sum"),
            상품비고_첫=("상품비고(1)", "first"),
            fb_any=("__fallback", "max"),
        )
        .reset_index()
    )

    n = len(grouped)
    blank = [""] * n

    # 수량도 int로 정리
    qty_sum_int = grouped["수량_합"].round().astype(int)
    price_int   = grouped["단가(1)"].round().astype(int)

    # 9) 최종 ERP 포맷 테이블
    result = pd.DataFrame({
        "거래일자":    grouped["거래일자"],
        "거래처명":    grouped["거래처명"],
        # 상품코드 .0 제거
        "상품코드(1)": grouped["상품코드(1)"].astype(str).str.replace(r"\.0$", "", regex=True),
        "상품명(1)":   grouped["상품명(1)"],
        "수량(1)":     qty_sum_int,
        "단가(1)":     price_int,
        "상품비고(1)": grouped["상품비고_첫"],

        "상품코드(2)": blank,
        "상품명(2)":   blank,
        "수량(2)":     blank,
        "단가(2)":     blank,
        "상품비고(2)": blank,

        "상품코드(3)": blank,
        "상품명(3)":   blank,
        "수량(3)":     blank,
        "단가(3)":     blank,
        "상품비고(3)": blank,

        "상품코드(4)": blank,
        "상품명(4)":   blank,
        "수량(4)":     blank,
        "단가(4)":     blank,
        "상품비고(4)": blank,

        "상품코드(5)": blank,
        "상품명(5)":   blank,
        "수량(5)":     blank,
        "단가(5)":     blank,
        "상품비고(5)": blank,

        "전표비고(1)": blank,
        "전표비고(2)": blank,
        "전표비고(3)": blank,
        "전표비고(4)": blank,
        "전표비고(5)": blank,
    })

    # 색칠용 플래그
    result["__fallback"] = grouped["fb_any"].astype(bool).values

    return result


def save_result_with_style(df: pd.DataFrame, out_path: str):
    fb_mask = df["__fallback"].fillna(False).tolist() if "__fallback" in df.columns else [False]*len(df)
    df = df.drop(columns="__fallback") if "__fallback" in df.columns else df

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
        ws = writer.book.active
        name_col = next((i for i, c in enumerate(ws[1], start=1) if c.value == "상품명(1)"), None)
        if name_col:
            for idx, fb in enumerate(fb_mask, start=2):
                if fb:
                    ws.cell(row=idx, column=name_col).font = Font(color="FFFF0000")


# ====== GUI ======

APP_TITLE = "엑셀 변환기 (옵션ID 매핑 + 단가계산+묶기)"
DEFAULT_OUT_NAME = "result_output.xlsx"

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("700x340")
        self.resizable(False, False)
        self.main_path = None
        self.map_path = None
        self._build()

    def _build(self):
        pad = {"padx": 12, "pady": 8}
        f1 = tk.Frame(self); f1.pack(fill="x", **pad)
        tk.Label(f1, text="1) 주문/정산 엑셀:", width=20, anchor="w").pack(side="left")
        self.lbl1 = tk.Label(f1, text="(선택 안 됨)", anchor="w"); self.lbl1.pack(side="left", fill="x", expand=True)
        tk.Button(f1, text="파일 선택…", command=self.pick_main).pack(side="right")

        f2 = tk.Frame(self); f2.pack(fill="x", **pad)
        tk.Label(f2, text="2) 옵션ID↔코드/상품명 엑셀:", width=20, anchor="w").pack(side="left")
        self.lbl2 = tk.Label(f2, text="(선택 안 됨)", anchor="w"); self.lbl2.pack(side="left", fill="x", expand=True)
        tk.Button(f2, text="파일 선택…", command=self.pick_map).pack(side="right")

        f3 = tk.Frame(self); f3.pack(fill="x", **pad)
        tk.Label(f3, text="1) 시트 이름(옵션):", width=20, anchor="w").pack(side="left")
        self.ent = tk.Entry(f3); self.ent.pack(side="left", fill="x", expand=True, padx=(6,6))
        tk.Label(f3, text="(비우면 첫 번째 시트)", fg="#666").pack(side="left")

        f4 = tk.Frame(self); f4.pack(fill="x", **pad)
        self.btn = tk.Button(f4, text="변환 실행 → 저장", command=self.run); self.btn.pack(side="left")
        tk.Button(f4, text="종료", command=self.destroy).pack(side="right")

        self.status = tk.StringVar(value="준비됨")
        tk.Label(self, textvariable=self.status, anchor="w", fg="#444").pack(fill="x", padx=12, pady=(12,10))

    def pick_main(self):
        path = filedialog.askopenfilename(title="1) 주문/정산 엑셀 선택", filetypes=[("Excel files","*.xlsx")])
        if path:
            self.main_path = path
            self.lbl1.config(text=os.path.basename(path))
            self.status.set("1) 엑셀 선택 완료")

    def pick_map(self):
        path = filedialog.askopenfilename(title="2) 매핑 엑셀 선택", filetypes=[("Excel files","*.xlsx")])
        if path:
            self.map_path = path
            self.lbl2.config(text=os.path.basename(path))
            self.status.set("2) 엑셀 선택 완료")

    def run(self):
        if not self.main_path or not self.map_path:
            messagebox.showwarning(APP_TITLE, "엑셀 파일 두 개 모두 선택하세요.")
            return
        out = filedialog.asksaveasfilename(
            title="결과 저장 위치 선택",
            initialfile=DEFAULT_OUT_NAME,
            defaultextension=".xlsx",
            filetypes=[("Excel","*.xlsx")]
        )
        if not out:
            return
        try:
            self._toggle(False)
            self.status.set("변환 중…")
            df = build_result(self.main_path, self.map_path, self.ent.get().strip() or None)
            save_result_with_style(df, out)
            self.status.set(f"완료: {os.path.basename(out)}")
            messagebox.showinfo(APP_TITLE, f"저장 완료:\n{out}")
        except Exception as e:
            traceback.print_exc()
            self.status.set("실패")
            messagebox.showerror(APP_TITLE, f"에러 발생:\n{e}")
        finally:
            self._toggle(True)

    def _toggle(self, en):
        self.btn.config(state=tk.NORMAL if en else tk.DISABLED)


def main():
    App().mainloop()


if __name__ == "__main__":
    main()
