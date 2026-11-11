#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
adr_search.py - Tìm báo cáo ADR theo quy trình 2 lượt (thay cho thao tác filter tay trong Access).

NGUỒN DỮ LIỆU (chọn 1):
  --xlsx file.xlsx [--sheet File_test]   : file Excel xuất từ màn hình "search theo năm" (có --sheet: đọc 1 sheet; không: gộp mọi sheet)
  --access file.accdb --table TÊN        : đọc thẳng từ Access qua pyodbc (chỉ chạy được trên Windows có
                                           "Microsoft Access Driver"), lọc năm theo 2 số đầu của STT báo cáo
  --access file.accdb --list-tables      : liệt kê bảng/query để biết điền --table gì

QUY TRÌNH
  Bước 0: Thuốc NN chỉ giữ 1 hoặc rỗng (Blank)
  Lượt 1: Tên thương mại chứa BẤT KỲ mảnh nhập vào (vd "phle","bodia")  VÀ  Hoạt chất hiệu chỉnh chứa hoạt chất
  Lượt 2: Hoạt chất hiệu chỉnh chứa hoạt chất, loại các dòng đã có ở lượt 1, rồi giữ dòng nếu
            (a) Tên thương mại GẦN GIỐNG tên đầy đủ (sai tối đa N ký tự), HOẶC
            (b) Nhà sản xuất chứa BẤT KỲ mảnh nhập vào (vd "inno","hera")
  Xuất: Excel gồm sheet Tổng hợp / Lượt 1 / Lượt 2 / Tham số, mặc định ra D:\\Python_Hải\\Lọc báo cáo ADR.

VÍ DỤ
  python adr_search.py --xlsx data.xlsx --tm phle bodia --hoat-chat diosmin \
      --ten-day-du Phlebodia --nsx inno hera --nam 2024 2026
"""
from __future__ import annotations

import argparse
import sys
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd


# ----------------------------------------------------------------- chuẩn hoá
def norm(x) -> str:
    """Chữ thường, bỏ dấu tiếng Việt, bỏ tab/khoảng trắng thừa. So khớp không phân biệt dấu/hoa-thường."""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    s = str(x).replace("đ", "d").replace("Đ", "D")
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return " ".join(s.lower().split())


def contains_any(text: str, terms: list[str]) -> bool:
    return any(t in text for t in terms)


def contains_all(text: str, terms: list[str]) -> bool:
    return all(t in text for t in terms)


def best_substring_distance(pattern: str, text: str) -> int:
    """Số ký tự sai (Levenshtein) nhỏ nhất giữa `pattern` và MỘT ĐOẠN BẤT KỲ của `text` (thuật toán Sellers).
    Nhờ vậy 'Plebodia' vẫn khớp 'PHLEBODIA 600mg' hay 'Diosmin (Phlebodia) 600mg'."""
    if not pattern:
        return 0
    prev = list(range(len(pattern) + 1))  # cột 0 = chi phí xoá pattern
    best = prev[-1]
    for ch in text:
        cur = [0]  # bắt đầu khớp ở vị trí bất kỳ trong text => chi phí 0
        for i, pc in enumerate(pattern, 1):
            cur.append(min(prev[i] + 1, cur[i - 1] + 1, prev[i - 1] + (pc != ch)))
        best = min(best, cur[-1])
        prev = cur
    return best


# ----------------------------------------------------------------- lọc năm theo STT báo cáo
COL_STT, COL_NAM = "STT báo cáo", "Năm"


def nam_tu_stt(v):
    """2 số đầu của 'STT báo cáo' = 2 số cuối của năm. Vd 2412540 -> 24 -> 2024. Không đọc được -> None."""
    digits = "".join(ch for ch in str(v).split(".")[0] if ch.isdigit()) if v is not None else ""
    return 2000 + int(digits[:2]) if len(digits) >= 2 else None


def loc_nam(df, y_from, y_to):
    """Giữ dòng có năm (suy ra từ STT báo cáo) nằm trong [y_from, y_to].
    STT không đọc được -> dùng cột 'Năm'; cả hai đều trống -> giữ lại (để bạn tự kiểm)."""
    if y_from < 100:
        y_from += 2000
    if y_to < 100:
        y_to += 2000
    if COL_STT in df.columns:
        nam = df[COL_STT].map(nam_tu_stt)
    else:
        print(f"[Lưu ý] Không có cột '{COL_STT}' -> lọc năm theo cột '{COL_NAM}'.")
        nam = pd.Series([None] * len(df), index=df.index)
    if COL_NAM in df.columns:
        nam = nam.where(nam.notna(), pd.to_numeric(df[COL_NAM], errors="coerce"))
    ok = nam.isna() | ((nam >= y_from) & (nam <= y_to))
    print(f"Lọc năm {y_from}-{y_to} (theo 2 số đầu STT báo cáo): giữ {int(ok.sum())}/{len(df)} dòng.")
    return df[ok]


# ----------------------------------------------------------------- đọc dữ liệu
def load_xlsx(path: str, sheet: str | None = None) -> pd.DataFrame:
    """Có --sheet: chỉ đọc đúng sheet đó (vd File_test). Không có: đọc mọi sheet rồi gộp."""
    if sheet:
        names = pd.ExcelFile(path).sheet_names
        if sheet not in names:
            sys.exit(f"Không thấy sheet '{sheet}'. Các sheet trong file: {names}")
        return pd.read_excel(path, sheet_name=sheet, dtype=object)
    sheets = pd.read_excel(path, sheet_name=None, dtype=object)
    return pd.concat(sheets.values(), ignore_index=True)


def load_access(path: str, table: str) -> pd.DataFrame:
    try:
        import pyodbc
    except ImportError:
        sys.exit("Thiếu pyodbc: pip install pyodbc (chỉ chạy trên Windows có Access ODBC driver)")
    conn = pyodbc.connect(
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" + str(Path(path).resolve()) + ";"
    )
    df = pd.read_sql(f"SELECT * FROM [{table}]", conn)
    conn.close()
    return df


def list_access_tables(path: str):
    import pyodbc
    conn = pyodbc.connect(
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=" + str(Path(path).resolve()) + ";"
    )
    cur = conn.cursor()
    for t in cur.tables():
        if t.table_type in ("TABLE", "VIEW"):
            print(t.table_type, "-", t.table_name)
    conn.close()


# ----------------------------------------------------------------- lõi tìm kiếm
# Thư mục mặc định để xuất kết quả (sửa ở đây nếu đổi chỗ)
OUT_DIR = Path(r"D:\Python_Hải\Lọc báo cáo ADR")

COL_NN, COL_TM, COL_HC, COL_NSX = "Thuốc NN", "Tên thương mại", "Hoạt chất hiệu chỉnh", "Nhà sản xuất"


def search(df, tm, hoat_chat, ten_day_du, nsx, tm_mode="any", max_err=None, hc_mode="all"):
    # Cột bắt buộc: Tên thương mại, Hoạt chất hiệu chỉnh. Thiếu Thuốc NN / Nhà sản xuất thì bỏ qua tiêu chí đó.
    for c in (COL_TM, COL_HC):
        if c not in df.columns:
            sys.exit(f"Không thấy cột bắt buộc '{c}'. Các cột hiện có: {list(df.columns)}")
    if COL_NN in df.columns:
        # Bước 0: Thuốc NN = 1 hoặc rỗng
        nn = df[COL_NN].map(lambda v: norm(v))
        df = df[nn.isin(["1", "1.0", ""])].copy()
    else:
        print(f"[Lưu ý] Không có cột '{COL_NN}' -> bỏ qua bước lọc Thuốc NN.")
    if COL_NSX not in df.columns:
        print(f"[Lưu ý] Không có cột '{COL_NSX}' -> lượt 2 chỉ dò theo tên thương mại gần giống.")
        df[COL_NSX] = ""
        added_nsx = True
    else:
        added_nsx = False

    # Bỏ trùng 1 dòng thuốc (nếu gộp nhiều sheet) theo 'ID thuốc'
    # (chỉ xét các dòng CÓ ID; dòng để trống ID thì giữ nguyên, không coi là trùng nhau)
    if "ID thuốc" in df.columns:
        trung = df["ID thuốc"].notna() & df.duplicated(subset="ID thuốc", keep="first")
        df = df[~trung]
    df = df.reset_index(drop=True)

    n_tm = df[COL_TM].map(norm)
    n_hc = df[COL_HC].map(norm)
    n_nsx = df[COL_NSX].map(norm)

    tm_t = [norm(t) for t in tm if norm(t)]
    hc_t = [norm(t) for t in hoat_chat if norm(t)]
    nsx_t = [norm(t) for t in nsx if norm(t)]
    full_t = [norm(t) for t in ten_day_du if norm(t)] or tm_t  # không nhập tên đầy đủ -> dùng chính các mảnh
    match_tm = contains_any if tm_mode == "any" else contains_all

    match_hc = contains_all if hc_mode == "all" else contains_any  # all: chứa đủ mọi hoạt chất; any: chứa 1 trong các hoạt chất
    hc_ok = n_hc.map(lambda s: match_hc(s, hc_t) if hc_t else True)

    # ---- Lượt 1
    tm_ok = n_tm.map(lambda s: match_tm(s, tm_t))
    m1 = tm_ok & hc_ok
    r1 = df[m1].copy()
    r1.insert(0, "Lý do khớp", "Lượt 1: tên TM chứa " + "/".join(tm) + " + hoạt chất")

    # ---- Lượt 2 (đảo chiều: hoạt chất trước, loại lượt 1)
    pool = hc_ok & ~m1
    reasons, keep = {}, []
    for i in df.index[pool]:
        why = []
        # (a) tên TM gần giống (cho sai tối đa k ký tự)
        best = None
        for f in full_t:
            k = max_err if max_err is not None else 5
            d = best_substring_distance(f, n_tm[i])
            if d <= k and (best is None or d < best[0]):
                best = (d, f)
        if best is not None and n_tm[i]:
            why.append(f"tên TM gần giống '{best[1]}' (sai {best[0]} ký tự)")
        # (b) nhà sản xuất chứa mảnh (không đụng tới cột tên TM)
        if nsx_t and contains_any(n_nsx[i], nsx_t):
            hit = [t for t in nsx_t if t in n_nsx[i]]
            why.append("NSX chứa " + "/".join(hit))
        if why:
            keep.append(i)
            reasons[i] = "Lượt 2: " + "; ".join(why)
    r2 = df.loc[keep].copy()
    r2.insert(0, "Lý do khớp", [reasons[i] for i in keep])
    if added_nsx:  # cột NSX giả thêm vào để tính toán -> bỏ đi khi xuất
        r1, r2 = r1.drop(columns=COL_NSX), r2.drop(columns=COL_NSX)
    return r1, r2


# ----------------------------------------------------------------- xuất Excel
def export(r1, r2, original_cols, params: dict, out: Path):
    def restore(d):  # trả tên cột về như gốc (pandas đổi cột trùng thành 'xxx.1')
        d = d.copy()
        d.columns = ["Lý do khớp"] + original_cols[: d.shape[1] - 1]
        return d

    r1, r2 = restore(r1), restore(r2)
    allr = pd.concat([r1.assign(**{"Lượt": 1}), r2.assign(**{"Lượt": 2})], ignore_index=True)
    allr.insert(0, "Lượt", allr.pop("Lượt"))
    ps = pd.DataFrame({"Tham số": list(params), "Giá trị": [str(v) for v in params.values()]})
    with pd.ExcelWriter(out, engine="openpyxl", datetime_format="dd/mm/yyyy") as w:
        allr.to_excel(w, sheet_name="Tổng hợp", index=False)
        r1.to_excel(w, sheet_name="Lượt 1", index=False)
        r2.to_excel(w, sheet_name="Lượt 2", index=False)
        ps.to_excel(w, sheet_name="Tham số", index=False)
        for ws in w.book.worksheets:
            ws.freeze_panes = "C2" if ws.title != "Tham số" else "A2"
            ws.auto_filter.ref = ws.dimensions
            for col in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in list(col)[:50])
                ws.column_dimensions[col[0].column_letter].width = min(max(10, width + 2), 45)
    return allr


# ----------------------------------------------------------------- CLI
def ask(prompt):  # nhập nhiều mảnh cách nhau bằng dấu phẩy hoặc khoảng trắng
    return [x for x in input(prompt).replace(",", " ").split() if x]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--xlsx")
    ap.add_argument("--sheet", help="tên sheet cần đọc trong file Excel, vd File_test (bỏ trống = đọc mọi sheet)")
    ap.add_argument("--access")
    ap.add_argument("--table")
    ap.add_argument("--list-tables", action="store_true")
    ap.add_argument("--nam", nargs=2, type=int, metavar=("TỪ", "ĐẾN"), help="khoảng năm, vd 2024 2026 (đọc từ 2 số đầu của STT báo cáo)")
    ap.add_argument("--tm", nargs="*", help="các mảnh tên thương mại, vd phle bodia")
    ap.add_argument("--tm-mode", choices=["any", "all"], default="any", help="any: chứa 1 trong các mảnh (mặc định); all: chứa đủ")
    ap.add_argument("--hoat-chat", nargs="*", help='các hoạt chất; cụm nhiều chữ đặt trong ngoặc kép, vd "peppermint oil"')
    ap.add_argument("--hc-mode", choices=["all", "any"], default="all", help="all: ô hoạt chất phải chứa ĐỦ mọi hoạt chất (mặc định); any: chứa 1 trong các hoạt chất")
    ap.add_argument("--ten-day-du", nargs="*", help="tên đầy đủ để dò gần giống ở lượt 2, vd Phlebodia")
    ap.add_argument("--nsx", nargs="*", help="các mảnh nhà sản xuất, vd inno hera")
    ap.add_argument("--max-err", type=int, help="số ký tự sai tối đa khi dò gần giống (mặc định: 5)")
    ap.add_argument("--out", help="đường dẫn file Excel xuất ra (mặc định: thư mục OUT_DIR)")
    a = ap.parse_args()

    if a.list_tables:
        return list_access_tables(a.access)

    if a.tm is None:
        a.tm = ask("Tên thương mại (các mảnh, cách nhau dấu phẩy/khoảng trắng): ")
    if a.hoat_chat is None:
        a.hoat_chat = ask("Hoạt chất hiệu chỉnh: ")
    if a.ten_day_du is None:
        a.ten_day_du = ask("Tên thương mại đầy đủ để dò gần giống (Enter = bỏ qua): ")
    if a.nsx is None:
        a.nsx = ask("Nhà sản xuất (các mảnh, Enter = bỏ qua): ")

    if a.access:
        if not a.table:
            sys.exit("Thiếu --table (dùng --list-tables để xem).")
        df = load_access(a.access, a.table)
    elif a.xlsx:
        df = load_xlsx(a.xlsx, a.sheet)
    else:
        sys.exit("Cần --xlsx hoặc --access.")
    if a.nam:
        df = loc_nam(df, a.nam[0], a.nam[1])

    original_cols = list(df.columns)
    r1, r2 = search(df, a.tm, a.hoat_chat, a.ten_day_du, a.nsx, a.tm_mode, a.max_err, a.hc_mode)

    out = Path(a.out) if a.out else OUT_DIR / f"ADR_{'_'.join(a.tm) or 'ketqua'}_{datetime.now():%Y%m%d_%H%M}.xlsx"
    out.parent.mkdir(parents=True, exist_ok=True)
    params = {"Năm": a.nam, "Tên TM (mảnh)": a.tm, "Chế độ mảnh": a.tm_mode, "Hoạt chất": a.hoat_chat,
              "Tên đầy đủ (dò gần giống)": a.ten_day_du, "NSX (mảnh)": a.nsx, "Sai tối đa": a.max_err if a.max_err is not None else 5,
              "Lượt 1": len(r1), "Lượt 2": len(r2), "Chạy lúc": datetime.now().strftime("%d/%m/%Y %H:%M")}
    export(r1, r2, original_cols, params, out)
    print(f"Lượt 1: {len(r1)} dòng | Lượt 2: {len(r2)} dòng | Tổng: {len(r1) + len(r2)}")
    print("Đã xuất:", out)


if __name__ == "__main__":
    main()