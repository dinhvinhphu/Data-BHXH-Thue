import os
import re
import json
import time
import urllib3
import requests
import pandas as pd
import pdfplumber
from io import BytesIO
from datetime import datetime
from bs4 import BeautifulSoup
from urllib.parse import urljoin

# Tắt cảnh báo chứng chỉ SSL cũ thường gặp trên cổng thông tin hành chính công VN
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
}

OUTPUT_JSON = "KHO_DU_LIEU_BHXH_THUE.json"
DOWNLOAD_DIR = "raw_files"
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# Danh sách các nguồn công bố trọng điểm (Có thể mở rộng thêm các tỉnh khác theo cùng cấu trúc)
SOURCES = [
    {
        "tinh_thanh": "TP.HCM",
        "loai": "BHXH",
        "url": "https://tphcm.baohiemxahoi.gov.vn/Pages/don-vi-no-bhxh-bhyt-bhtn.aspx?CateID=269",
        "keywords": ["nợ", "chậm đóng", "bhxh", "danh sách"]
    },
    {
        "tinh_thanh": "Bình Dương",
        "loai": "BHXH",
        "url": "https://binhduong.baohiemxahoi.gov.vn/Pages/don-vi-no-bhxh-bhyt-bhtn.aspx?CateID=269",
        "keywords": ["nợ", "chậm đóng", "bhxh", "danh sách"]
    },
    {
        "tinh_thanh": "Đồng Nai",
        "loai": "BHXH",
        "url": "https://dongnai.baohiemxahoi.gov.vn/Pages/don-vi-no-bhxh-bhyt-bhtn.aspx?CateID=269",
        "keywords": ["nợ", "chậm đóng", "bhxh", "danh sách"]
    },
    {
        "tinh_thanh": "TP.HCM",
        "loai": "THUE",
        "url": "https://tphcm.gdt.gov.vn/wps/portal",
        "keywords": ["nợ thuế", "công khai", "người nộp thuế"]
    }
]


def clean_mst(raw_val):
    """Chuẩn hóa Mã số thuế về chuỗi 10 ký tự chuẩn, giữ nguyên số 0 ở đầu."""
    if pd.isna(raw_val):
        return ""
    s = str(raw_val).strip()
    # Loại bỏ phần thập phân .0 nếu Excel đọc nhầm thành số thực
    if s.endswith(".0"):
        s = s[:-2]
    # Lấy 10 chữ số đầu tiên (quy MST chi nhánh 13 số về công ty mẹ 10 số)
    digits = re.sub(r"\D", "", s)
    if len(digits) >= 10:
        return digits[:10]
    if len(digits) == 9:
        return "0" + digits  # Bù số 0 đầu bị Excel cắt mất
    return ""


def clean_number(val):
    """Chuyển đổi chuỗi tiền tệ / số tháng nợ thành số thực chuẩn."""
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = re.sub(r"[^\d.,-]", "", str(val)).replace(",", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


def find_header_row(df):
    """Tự động tìm dòng tiêu đề thực sự trong bảng Excel bị chèn quốc hiệu/tiêu ngữ phía trên."""
    header_keywords = ["mã số thuế", "mst", "tên đơn vị", "tên doanh nghiệp", "người nộp thuế", "số tiền", "tháng nợ"]
    for idx, row in df.head(15).iterrows():
        row_text = " ".join([str(x).lower() for x in row.values if pd.notna(x)])
        matches = sum(1 for kw in header_keywords if kw in row_text)
        if matches >= 2:
            return idx
    return 0


def map_columns(columns):
    """Ánh xạ tên cột động sang tên trường chuẩn của Kho dữ liệu."""
    col_map = {}
    for col in columns:
        c = str(col).lower().strip()
        if any(k in c for k in ["mã số thuế", "mst"]):
            col_map["mst"] = col
        elif any(k in c for k in ["mã đơn vị", "mã đv"]):
            col_map["ma_don_vi"] = col
        elif any(k in c for k in ["tên đơn vị", "tên doanh nghiệp", "người nộp thuế", "tên nnt", "tên công ty"]):
            col_map["ten_dn"] = col
        elif any(k in c for k in ["địa chỉ"]):
            col_map["dia_chi"] = col
        elif any(k in c for k in ["số tháng", "tháng nợ", "thời gian nợ"]):
            col_map["so_thang_no"] = col
        elif any(k in c for k in ["tổng tiền", "số tiền", "tiền nợ", "cộng nợ", "tổng cộng"]):
            col_map["tong_tien_no"] = col
    return col_map


def parse_excel_bytes(file_bytes, meta):
    """Bóc tách dữ liệu từ file Excel (.xlsx / .xls)."""
    records = []
    try:
        excel_file = pd.ExcelFile(BytesIO(file_bytes))
        for sheet in excel_file.sheet_names:
            raw_df = pd.read_excel(excel_file, sheet_name=sheet, header=None)
            if raw_df.empty or len(raw_df) < 3:
                continue

            header_idx = find_header_row(raw_df)
            df = pd.read_excel(excel_file, sheet_name=sheet, header=header_idx)
            col_map = map_columns(df.columns)

            if "ten_dn" not in col_map and "mst" not in col_map:
                continue

            for _, row in df.iterrows():
                ten_dn = str(row.get(col_map.get("ten_dn", ""), "")).strip()
                mst = clean_mst(row.get(col_map.get("mst", ""), ""))

                if not ten_dn or ten_dn.lower() in ["nan", "none", "tổng cộng", "cộng"]:
                    continue

                records.append({
                    "mst": mst,
                    "ma_don_vi": str(row.get(col_map.get("ma_don_vi", ""), "")).strip().replace("nan", ""),
                    "ten_dn": ten_dn.upper(),
                    "dia_chi": str(row.get(col_map.get("dia_chi", ""), "")).strip().replace("nan", ""),
                    "tinh_thanh": meta["tinh_thanh"],
                    "nguon_du_lieu": meta["loai"],
                    "ky_cong_bo": meta["ky_cong_bo"],
                    "so_thang_no": int(clean_number(row.get(col_map.get("so_thang_no", ""), 0))),
                    "tong_tien_no": clean_number(row.get(col_map.get("tong_tien_no", ""), 0)),
                    "tieu_de_nguon": meta["tieu_de"],
                    "ngay_cap_nhat": datetime.now().strftime("%Y-%m-%d")
                })
    except Exception as e:
        print(f"  [!] Lỗi đọc Excel ({meta['tieu_de']}): {e}")
    return records


def parse_pdf_bytes(file_bytes, meta):
    """Bóc tách bảng biểu từ file PDF dạng text (không phải ảnh scan)."""
    records = []
    try:
        with pdfplumber.open(BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if not table or len(table) < 2:
                        continue
                    df = pd.DataFrame(table[1:], columns=table[0])
                    col_map = map_columns(df.columns)
                    if "ten_dn" not in col_map and "mst" not in col_map:
                        continue

                    for _, row in df.iterrows():
                        ten_dn = str(row.get(col_map.get("ten_dn", ""), "")).replace("\n", " ").strip()
                        mst = clean_mst(row.get(col_map.get("mst", ""), ""))
                        if not ten_dn or ten_dn.lower() in ["none", "tổng cộng"]:
                            continue

                        records.append({
                            "mst": mst,
                            "ma_don_vi": str(row.get(col_map.get("ma_don_vi", ""), "")).strip(),
                            "ten_dn": ten_dn.upper(),
                            "dia_chi": str(row.get(col_map.get("dia_chi", ""), "")).replace("\n", " ").strip(),
                            "tinh_thanh": meta["tinh_thanh"],
                            "nguon_du_lieu": meta["loai"],
                            "ky_cong_bo": meta["ky_cong_bo"],
                            "so_thang_no": int(clean_number(row.get(col_map.get("so_thang_no", ""), 0))),
                            "tong_tien_no": clean_number(row.get(col_map.get("tong_tien_no", ""), 0)),
                            "tieu_de_nguon": meta["tieu_de"],
                            "ngay_cap_nhat": datetime.now().strftime("%Y-%m-%d")
                        })
    except Exception as e:
        print(f"  [!] Lỗi đọc PDF ({meta['tieu_de']}): {e}")
    return records


def extract_period(title_text):
    """Trích xuất kỳ công bố (YYYY-MM) từ tiêu đề bài viết."""
    match = re.search(r"tháng\s*(\d{1,2})[/.-](\d{4})", title_text.lower())
    if match:
        return f"{match.group(2)}-{int(match.group(1)):02d}"
    return datetime.now().strftime("%Y-%m")


def crawl_source(source):
    """Quét trang danh mục, tìm 3 bài đăng mới nhất và tải file đính kèm."""
    print(f"\n>>> Đang quét: {source['loai']} - {source['tinh_thanh']} ({source['url']})")
    extracted_records = []
    try:
        resp = requests.get(source["url"], headers=HEADERS, verify=False, timeout=25)
        resp.encoding = "utf-8"
        soup = BeautifulSoup(resp.text, "html.parser")

        article_links = []
        for a in soup.find_all("a", href=True):
            text = a.get_text(strip=True)
            href = urljoin(source["url"], a["href"])
            if len(text) > 20 and any(kw in text.lower() for kw in source["keywords"]):
                if href not in [x[1] for x in article_links]:
                    article_links.append((text, href))

        # Lấy tối đa 3 bài công bố mới nhất mỗi lần chạy
        for title, link in article_links[:3]:
            print(f"  + Kiểm tra bài: {title[:70]}...")
            time.sleep(2)  # Giãn cách request tránh bị tường lửa chặn
            art_resp = requests.get(link, headers=HEADERS, verify=False, timeout=25)
            art_soup = BeautifulSoup(art_resp.text, "html.parser")

            ky_cong_bo = extract_period(title)
            meta = {
                "tinh_thanh": source["tinh_thanh"],
                "loai": source["loai"],
                "ky_cong_bo": ky_cong_bo,
                "tieu_de": title
            }

            # Tìm link tải file Excel hoặc PDF trong bài viết
            for file_a in art_soup.find_all("a", href=True):
                file_href = urljoin(link, file_a["href"])
                lower_href = file_href.lower()

                if any(lower_href.endswith(ext) for ext in [".xlsx", ".xls", ".pdf"]):
                    print(f"    -> Đang tải file: {file_href.split('/')[-1]}")
                    time.sleep(1.5)
                    f_resp = requests.get(file_href, headers=HEADERS, verify=False, timeout=40)
                    if f_resp.status_code == 200:
                        if lower_href.endswith((".xlsx", ".xls")):
                            extracted_records.extend(parse_excel_bytes(f_resp.content, meta))
                        elif lower_href.endswith(".pdf"):
                            extracted_records.extend(parse_pdf_bytes(f_resp.content, meta))

    except Exception as e:
        print(f"  [X] Lỗi kết nối nguồn {source['tinh_thanh']}: {e}")

    return extracted_records


def update_master_warehouse(new_records):
    """Gộp dữ liệu mới vào Kho dữ liệu lịch sử, tự động khử trùng lặp theo (MST + Kỳ + Nguồn)."""
    existing_data = []
    if os.path.exists(OUTPUT_JSON):
        try:
            with open(OUTPUT_JSON, "r", encoding="utf-8") as f:
                payload = json.load(f)
                existing_data = payload.get("records", [])
        except Exception:
            existing_data = []

    # Tạo từ điển khử trùng lặp: Key = (MST hoặc Tên DN) + Tỉnh + Nguồn + Kỳ công bố
    warehouse_map = {}
    for item in existing_data:
        key_id = item["mst"] if item["mst"] else item["ten_dn"]
        unique_key = f"{key_id}|{item['tinh_thanh']}|{item['nguon_du_lieu']}|{item['ky_cong_bo']}"
        warehouse_map[unique_key] = item

    added_count = 0
    for item in new_records:
        key_id = item["mst"] if item["mst"] else item["ten_dn"]
        unique_key = f"{key_id}|{item['tinh_thanh']}|{item['nguon_du_lieu']}|{item['ky_cong_bo']}"
        if unique_key not in warehouse_map:
            added_count += 1
        warehouse_map[unique_key] = item  # Cập nhật bản mới nhất nếu trùng kỳ

    final_records = list(warehouse_map.values())
    output_payload = {
        "metadata": {
            "last_updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "total_records": len(final_records),
            "new_records_added": added_count
        },
        "records": final_records
    }

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(output_payload, f, ensure_ascii=False, indent=2)

    print(f"\n==========================================")
    print(f"HOÀN TẤT CẬP NHẬT KHO DỮ LIỆU!")
    print(f"- Tổng bản ghi trong kho: {len(final_records):,}")
    print(f"- Bản ghi mới thêm kỳ này: {added_count:,}")
    print(f"- File xuất ra: {OUTPUT_JSON}")
    print(f"==========================================")


if __name__ == "__main__":
    all_new = []
    for src in SOURCES:
        all_new.extend(crawl_source(src))
    update_master_warehouse(all_new)
