# -*- coding: utf-8 -*-
"""
엑셀 변환기 (헤더 자동 탐지 + 옵션ID 매핑 + 상품명 + ERP 포맷 + 옵션ID 대체)
- 첫번째 엑셀: 옵션ID, 매출인식일, 판매 수량(B), 정산대상액, 등록상품명
- 두번째 엑셀: 옵션ID, 코드, 윈윈상품명
→ 옵션ID로 매칭해서 결과 엑셀에
   거래일자(매출인식일), 거래처명(쿠팡-제트배송),
   상품코드(1)(코드 or 옵션ID), 상품명(1)(윈윈상품명 or 등록상품명),
   수량(1)(판매수량), 단가(1)(정산대상액)
   나머지 컬럼은 비워두고 매핑 실패 상품명은 빨간색 표시
"""

import os, re, traceback
from typing import Optional, List
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl.styles import Font


# ================= 유틸 =================

def _norm(s: str) -> str:
    s = str(s)
    s = s.replace("\u200b", "").replace("\ufeff", "")
    return re.sub(r"\s+", "", s).lower()

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
              .replace({r"[^0-9.\-]": ""}, regex=True)
              .replace("", "0")
              .astype(float)
    )

def _read_with_header_detection(path: str, sheet_name: Optional[str], keyword_candidates: List[str], search_rows: int = 50) -> pd.DataFrame:
    sheet_arg = sheet_name if sheet_name else 0
    raw = pd.read_excel(path, sheet_name=sheet_arg, header=None, dtype=str)
    targets = [_norm(k) for k in keyword_candidates]
    for i in range(min(search_rows, len(raw))):
        row_norm = [_norm(v) for v in raw.iloc[i].tolist()]
        if any(t in cell for t in targets for cell in row_norm):
            return pd.read_excel(path, sheet_name=sheet_arg, header=i)
    return pd.read_excel(path, sheet_name=sheet_arg, header=0)


# ================= 핵심 로직 =================

def build_result(main_path: str, map_path: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
    df_main = _read_with_header_detection(
        main_path, sheet_name,
        ["옵션id", "optionid", "매출인식일", "판매수량", "정산대상액", "등록상품명"]
    )

    col_optid  = _find_col(df_main, "옵션id", "optionid")
    col_date   = _find_col(df_main, "매출인식일")
    col_qty    = _find_col(df_main, "판매수량", "수량")
    col_price  = _find_col(df_main, "정산대상액", "정산 대상액")
    col_regnm  = _find_col(df_main, "등록상품명")

    df_map = _read_with_header_detection(map_path, None, ["옵션id", "optionid", "코드", "상품코드", "윈윈상품명", "상품명"])
    col_optid2 = _find_col(df_map, "옵션id", "optionid")
    col_code   = _find_col(df_map, "코드", "상품코드")
    col_name   = _find_col(df_map, "윈윈상품명", "윈윈 상품명")

    # 옵션ID로 조인
    df_main["_optkey"] = df_main[col_optid].astype(str).map(_norm)
    df_map["_optkey"]  = df_map[col_optid2].astype(str).map(_norm)

    df = pd.merge(df_main, df_map[["_optkey", col_code, col_name]], on="_optkey", how="left")
    df[col_qty]   = _to_number(df[col_qty])
    df[col_price] = _to_number(df[col_price])

    mapped_name = df[col_name]
    reg_name    = df[col_regnm]
    fallback_mask = mapped_name.isna() | (mapped_name.astype(str).str.strip() == "")

    final_name = mapped_name.copy()
    final_name[fallback_mask] = reg_name[fallback_mask]

    # 상품코드: 매핑 없으면 옵션ID로 대체
    final_code = df[col_code].copy()
    final_code[fallback_mask] = df[col_optid].astype(str)[fallback_mask]

    n = len(df)
    blank = [""] * n

    result = pd.DataFrame({
        "거래일자": df[col_date],
        "거래처명": ["쿠팡-제트배송"] * n,
        "상품코드(1)": final_code,
        "상품명(1)": final_name,
        "수량(1)": df[col_qty],
        "단가(1)": df[col_price],
        "상품비고(1)": blank,
        **{f"상품코드({i})": blank for i in range(2,6)},
        **{f"상품명({i})": blank for i in range(2,6)},
        **{f"수량({i})": blank for i in range(2,6)},
        **{f"단가({i})": blank for i in range(2,6)},
        **{f"상품비고({i})": blank for i in range(2,6)},
        **{f"전표비고({i})": blank for i in range(1,6)},
    })

    result["__fallback"] = fallback_mask.values
    return result


def save_result_with_style(df: pd.DataFrame, out_path: str):
    if "__fallback" in df.columns:
        fallback_mask = df["__fallback"].fillna(False).tolist()
        df = df.drop(columns="__fallback")
    else:
        fallback_mask = [False]*len(df)

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
        ws = writer.book.active
        # 상품명(1) 열 찾기
        name_col_idx = next((i for i, c in enumerate(ws[1], start=1) if c.value == "상품명(1)"), None)
        if name_col_idx:
            for row_idx, fb in enumerate(fallback_mask, start=2):
                if fb:
                    ws.cell(row=row_idx, column=name_col_idx).font = Font(color="FFFF0000")


# ================= GUI =================

APP_TITLE = "엑셀 변환기 (옵션ID 매핑 + ERP포맷 + 대체옵션ID)"
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
        pad = {"padx":12,"pady":8}
        f1=tk.Frame(self);f1.pack(fill="x",**pad)
        tk.Label(f1,text="1) 주문/정산 엑셀:",width=20,anchor="w").pack(side="left")
        self.lbl1=tk.Label(f1,text="(선택 안 됨)",anchor="w");self.lbl1.pack(side="left",fill="x",expand=True,padx=(6,6))
        tk.Button(f1,text="파일 선택…",command=self.pick_main).pack(side="right")
        f2=tk.Frame(self);f2.pack(fill="x",**pad)
        tk.Label(f2,text="2) 옵션ID↔코드/상품명 엑셀:",width=20,anchor="w").pack(side="left")
        self.lbl2=tk.Label(f2,text="(선택 안 됨)",anchor="w");self.lbl2.pack(side="left",fill="x",expand=True,padx=(6,6))
        tk.Button(f2,text="파일 선택…",command=self.pick_map).pack(side="right")
        f3=tk.Frame(self);f3.pack(fill="x",**pad)
        tk.Label(f3,text="1) 시트 이름(옵션):",width=20,anchor="w").pack(side="left")
        self.ent=tk.Entry(f3);self.ent.pack(side="left",fill="x",expand=True,padx=(6,6))
        tk.Label(f3,text="(비우면 첫 번째 시트)",fg="#666").pack(side="left")
        f4=tk.Frame(self);f4.pack(fill="x",**pad)
        self.btn=tk.Button(f4,text="변환 실행 → 저장",command=self.run);self.btn.pack(side="left")
        tk.Button(f4,text="종료",command=self.destroy).pack(side="right")
        self.status=tk.StringVar(value="준비됨")
        tk.Label(self,textvariable=self.status,anchor="w",fg="#444").pack(fill="x",padx=12,pady=(12,10))

    def pick_main(self):
        path=filedialog.askopenfilename(title="1) 주문/정산 엑셀 선택",filetypes=[("Excel files","*.xlsx")])
        if path:self.main_path=path;self.lbl1.config(text=os.path.basename(path));self.status.set("1) 주문/정산 엑셀 선택 완료")

    def pick_map(self):
        path=filedialog.askopenfilename(title="2) 옵션ID↔코드/상품명 엑셀 선택",filetypes=[("Excel files","*.xlsx")])
        if path:self.map_path=path;self.lbl2.config(text=os.path.basename(path));self.status.set("2) 매핑 엑셀 선택 완료")

    def run(self):
        if not self.main_path:return messagebox.showwarning(APP_TITLE,"먼저 1) 주문/정산 엑셀 선택")
        if not self.map_path:return messagebox.showwarning(APP_TITLE,"먼저 2) 매핑 엑셀 선택")
        out=filedialog.asksaveasfilename(title="결과 저장 위치",initialfile=DEFAULT_OUT_NAME,defaultextension=".xlsx",filetypes=[("Excel","*.xlsx")])
        if not out:return
        try:
            self._toggle(False);self.status.set("변환 중…")
            df=build_result(self.main_path,self.map_path,self.ent.get().strip() or None)
            save_result_with_style(df,out)
            self.status.set("완료: "+os.path.basename(out))
            messagebox.showinfo(APP_TITLE,f"저장 완료:\n{out}")
        except Exception as e:
            traceback.print_exc();self.status.set("실패");messagebox.showerror(APP_TITLE,f"에러 발생:\n{e}")
        finally:self._toggle(True)

    def _toggle(self,en):self.btn.config(state=tk.NORMAL if en else tk.DISABLED)


if __name__ == "__main__":
    App().mainloop()
