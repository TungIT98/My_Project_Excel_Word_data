import os
from pathlib import Path
from typing import Dict, Any, List

import pandas as pd
from docxtpl import DocxTemplate, Subdoc
from datetime import datetime
from unicodedata import normalize, combining


# =========================
# Cấu hình đường dẫn (sửa lại cho phù hợp)
# =========================
EXCEL_PATH = "/absolute/path/to/input.xlsx"
TEMPLATE_DIR = "/absolute/path/to/templates"  # Thư mục chứa 15 file .docx mẫu
OUTPUT_DIR = "/absolute/path/to/output"       # Thư mục xuất kết quả


# =========================
# Tiện ích xử lý chuỗi/định dạng
# =========================
def strip_accents(text: str) -> str:
    if not isinstance(text, str):
        return text
    return "".join(c for c in normalize("NFD", text) if not combining(c))


def normalize_key(text: str) -> str:
    if not isinstance(text, str):
        return text
    return strip_accents(text).lower().replace(" ", "")


def safe_filename(text: str) -> str:
    if not isinstance(text, str):
        text = str(text)
    keep = "-_.() []"
    return "".join(ch if ch.isalnum() or ch in keep else "_" for ch in text)


def format_date(value: Any, fmt: str = "%d/%m/%Y") -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, (datetime, )):
        return value.strftime(fmt)
    # pandas có thể đọc ngày thành Timestamp hoặc số Excel
    try:
        dt = pd.to_datetime(value, errors="coerce")
        if pd.isna(dt):
            return str(value)
        return dt.strftime(fmt)
    except Exception:
        return str(value)


def format_int(value: Any) -> str:
    if pd.isna(value) or value == "":
        return ""
    try:
        return f"{int(round(float(value))):,}".replace(",", ".")
    except Exception:
        return str(value)


def format_currency(value: Any) -> str:
    if pd.isna(value) or value == "":
        return ""
    try:
        return f"{float(value):,.0f}".replace(",", ".")
    except Exception:
        return str(value)


# =========================
# Đọc Excel
# =========================
def find_sheet_names(xlsx_path: str) -> Dict[str, str]:
    """
    Tìm sheet 'Hồ sơ' và 'Hàng hoá' bất kể dấu/viết hoa.
    Trả về dict {'hoso': real_name, 'hanghoa': real_name}
    """
    xls = pd.ExcelFile(xlsx_path)
    wanted = {"hoso": None, "hanghoa": None}
    for s in xls.sheet_names:
        key = normalize_key(s)
        if "hoso" in key and wanted["hoso"] is None:
            wanted["hoso"] = s
        if ("hanghoa" in key or "hanghoa" in key or "hanghoá" in key or "hanghoà" in key) and wanted["hanghoa"] is None:
            wanted["hanghoa"] = s
        # Bao phủ thêm trường hợp viết "hàng hoá"
        if "hang" in key and "hoa" in key and wanted["hanghoa"] is None:
            wanted["hanghoa"] = s
    if not wanted["hoso"] or not wanted["hanghoa"]:
        raise ValueError(f"Không tìm thấy đủ 2 sheet trong file Excel: {xls.sheet_names}")
    return wanted


def read_excel_data(xlsx_path: str) -> Dict[str, pd.DataFrame]:
    sheets = find_sheet_names(xlsx_path)
    df_hoso = pd.read_excel(xlsx_path, sheet_name=sheets["hoso"], dtype={"Mã KH": str})
    df_hanghoa = pd.read_excel(xlsx_path, sheet_name=sheets["hanghoa"], dtype={"Mã KH": str})

    # Chuẩn hóa tên cột tối thiểu (giữ nguyên tiếng Việt, chỉ strip spaces)
    df_hoso.columns = [c.strip() if isinstance(c, str) else c for c in df_hoso.columns]
    df_hanghoa.columns = [c.strip() if isinstance(c, str) else c for c in df_hanghoa.columns]

    return {"hoso": df_hoso, "hanghoa": df_hanghoa}


# =========================
# Tạo Subdoc bảng hàng hoá
# =========================
def build_goods_table_subdoc(doc: DocxTemplate, items_df: pd.DataFrame) -> Subdoc:
    """
    Sinh Subdoc bảng hàng hoá để chèn vào {{BảngHàngHoá}}.
    Yêu cầu template đặt placeholder {{BảngHàngHoá}} trên một dòng riêng.
    """
    sd = doc.new_subdoc()

    # Không có dòng hàng hoá -> để trống hoặc ghi chú
    if items_df is None or items_df.empty:
        p = sd.add_paragraph("Không có hàng hoá.")
        return sd

    # Xác định các cột mong muốn (linh hoạt tên cột)
    # Mặc định: 'Tên hàng', 'Số lượng', 'Đơn giá', 'Thành tiền'
    # Chấp nhận vài biến thể phổ biến
    cols_map = {
        "ten": None,
        "soluong": None,
        "dongia": None,
        "thanhtien": None,
    }

    def pick_col(df: pd.DataFrame, candidates: List[str]) -> str:
        available = {normalize_key(c): c for c in df.columns if isinstance(c, str)}
        for cand in candidates:
            k = normalize_key(cand)
            if k in available:
                return available[k]
        return None

    cols_map["ten"] = pick_col(items_df, ["Tên hàng", "Ten hang", "Ten hàng", "Tên Hàng", "TênHH", "TenHH"])
    cols_map["soluong"] = pick_col(items_df, ["Số lượng", "So luong", "Số Lượng", "SL", "SoLuong"])
    cols_map["dongia"] = pick_col(items_df, ["Đơn giá", "Don gia", "Đơn Giá", "DonGia"])
    cols_map["thanhtien"] = pick_col(items_df, ["Thành tiền", "Thanh tien", "Thành Tiền", "ThanhTien"])

    # Tính 'Thành tiền' nếu thiếu
    working = items_df.copy()
    if cols_map["thanhtien"] is None and cols_map["soluong"] and cols_map["dongia"]:
        try:
            working["_SoluongNum"] = pd.to_numeric(working[cols_map["soluong"]], errors="coerce").fillna(0)
            working["_DongiaNum"] = pd.to_numeric(working[cols_map["dongia"]], errors="coerce").fillna(0)
            working["_ThanhTien"] = working["_SoluongNum"] * working["_DongiaNum"]
            cols_map["thanhtien"] = "_ThanhTien"
        except Exception:
            pass

    # Tạo bảng
    headers = ["Tên hàng", "Số lượng", "Đơn giá", "Thành tiền"]
    table = sd.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    for i, h in enumerate(headers):
        table.rows[0].cells[i].text = h

    for _, r in working.iterrows():
        row = table.add_row().cells

        ten_val = r[cols_map["ten"]] if cols_map["ten"] in r else ""
        sl_val = r[cols_map["soluong"]] if cols_map["soluong"] in r else ""
        dg_val = r[cols_map["dongia"]] if cols_map["dongia"] in r else ""
        tt_val = r[cols_map["thanhtien"]] if cols_map["thanhtien"] in r else ""

        row[0].text = "" if pd.isna(ten_val) else str(ten_val)
        row[1].text = format_int(sl_val)
        row[2].text = format_currency(dg_val)
        row[3].text = format_currency(tt_val)

    return sd


# =========================
# Render 15 template cho mỗi khách hàng
# =========================
def build_context_for_customer(doc: DocxTemplate, customer_row: pd.Series, items_df: pd.DataFrame) -> Dict[str, Any]:
    """
    Tạo context truyền vào docxtpl.render().
    Các key trùng với placeholder trong Word.
    """
    # Đọc các cột chuẩn từ Hồ sơ
    val = lambda col: ("" if col not in customer_row or pd.isna(customer_row[col]) else customer_row[col])

    # Mapping cột -> placeholder (điền thêm nếu bạn có nhiều trường hơn trong template)
    ho_ten = val("Họ tên")
    ngay_sinh = format_date(val("Ngày sinh"))
    dia_chi = val("Địa chỉ")
    so_dt = val("Số điện thoại")
    ma_kh = val("Mã KH")

    context = {
        "TênKH": ho_ten,
        "NgàySinh": ngay_sinh,
        "ĐịaChỉ": dia_chi,
        "SốĐiệnThoại": so_dt,
        "MãKH": ma_kh,
        # Bảng hàng hoá sinh động
        "BảngHàngHoá": build_goods_table_subdoc(doc, items_df),
    }

    return context


def render_templates_for_customer(
    templates: List[Path],
    output_root: Path,
    customer_row: pd.Series,
    items_all: pd.DataFrame,
):
    """
    Với 1 khách hàng, render toàn bộ 15 template và lưu ra thư mục con riêng.
    """
    customer_id = customer_row.get("Mã KH", "")
    customer_name = customer_row.get("Họ tên", "")
    folder_name = f"{safe_filename(str(customer_id))}_{safe_filename(str(customer_name))}".strip("_")
    customer_out_dir = output_root / folder_name
    customer_out_dir.mkdir(parents=True, exist_ok=True)

    # Lọc hàng hoá theo Mã KH
    items_df = pd.DataFrame()
    if items_all is not None and "Mã KH" in items_all.columns:
        items_df = items_all[items_all["Mã KH"].astype(str) == str(customer_id)].copy()

    for tpl_path in templates:
        doc = DocxTemplate(str(tpl_path))
        context = build_context_for_customer(doc, customer_row, items_df)

        # Render
        doc.render(context)

        # Xuất file với tên gốc + mã KH
        out_name = f"{tpl_path.stem}__{safe_filename(str(customer_id))}.docx"
        out_path = customer_out_dir / out_name
        doc.save(str(out_path))
        print(f"Đã tạo: {out_path}")


def main():
    # Kiểm tra thư mục/đường dẫn
    template_dir = Path(TEMPLATE_DIR)
    output_dir = Path(OUTPUT_DIR)
    xlsx_path = Path(EXCEL_PATH)

    if not xlsx_path.exists():
        raise FileNotFoundError(f"Không thấy file Excel: {xlsx_path}")
    if not template_dir.exists():
        raise FileNotFoundError(f"Không thấy thư mục template: {template_dir}")
    output_dir.mkdir(parents=True, exist_ok=True)

    # Lấy danh sách template .docx
    templates = sorted([p for p in template_dir.glob("*.docx") if p.is_file()])
    if len(templates) == 0:
        raise FileNotFoundError(f"Không tìm thấy template .docx trong {template_dir}")
    print(f"Tìm thấy {len(templates)} template(s).")

    # Đọc Excel
    data = read_excel_data(str(xlsx_path))
    df_hoso = data["hoso"]
    df_hanghoa = data["hanghoa"]

    # Kiểm tra cột tối thiểu
    required_cols = ["Mã KH", "Họ tên"]
    for col in required_cols:
        if col not in df_hoso.columns:
            raise ValueError(f"Sheet 'Hồ sơ' thiếu cột bắt buộc: {col}")

    # Duyệt từng khách hàng
    for idx, row in df_hoso.iterrows():
        render_templates_for_customer(
            templates=templates,
            output_root=output_dir,
            customer_row=row,
            items_all=df_hanghoa,
        )

    print("Hoàn thành tất cả.")


if __name__ == "__main__":
    main()
