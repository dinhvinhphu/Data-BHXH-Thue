import os
import re
import json
import time
import pandas as pd
import pdfplumber
from io import BytesIO
from datetime import datetime
from urllib.parse import urljoin
from playwright.sync_api import sync_playwright

OUTPUT_JSON = "KHO_DU_LIEU_BHXH_THUE.json"

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
    if pd.isna(raw_val):
        return ""
    s = str(raw_val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    digits = re.sub(r"\D", "", s)
    if len(digits) >= 10:
        return digits[:10]
    if len(digits) == 9:
        return "0" + digits
    return ""


def clean_number(val):
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
    header_keywords = ["mã số thuế", "mst", "tên đơn vị", "tên doanh nghiệp", "người nộp thuế", "số tiền", "tháng nợ", "mã đơn vị"]
    for idx, row in df.head(15).iterrows():
        row_text = " ".join([str(x).lower() for x in row.values if pd.notna(x)])
        if sum(1 for kw in header_keywords if kw in row_text) >= 2:
            return idx
    return 0


def map_columns(columns):
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


def normalize_dataframe(df, meta):
    records = []
    col_map = map_columns(df.columns)
    if "ten_dn" not in col_map and "mst" not in col_map:
        return records

    for _, row in df.iterrows():
        ten_dn = str(row.get(col_map.get("ten_dn", ""), "")).replace("\n", " ").strip()
        mst = clean_mst(row.get(col_map.get("mst", ""), ""))
        if not ten_dn or ten_dn.lower() in ["nan", "none", "tổng cộng", "cộng", "stt"]:
            continue

        records.append({
            "mst": mst,
            "ma_don_vi": str(row.get(col_map.get("ma_don_vi", ""), "")).strip().replace("nan", ""),
            "ten_dn": ten_dn.upper(),
            "dia_chi": str(row.get(col_map.get("dia_chi", ""), "")).replace("\n", " ").strip().replace("nan", ""),
            "tinh_thanh": meta["tinh_thanh"],
            "nguon_du_lieu": meta["loai"],
            "ky_cong_bo": meta["ky_cong_bo"],
            "so_thang_no": int(clean_number(row.get(col_map.get("so_thang_no", ""), 0))),
            "tong_tien_no": clean_number(row.get(col_map.get("tong_tien_no", ""), 0)),
            "tieu_de_nguon": meta["tieu_de"],
            "ngay_cap_nhat": datetime.now().strftime("%Y-%m-%d")
        })
    return records


def parse_excel_bytes(file_bytes, meta):
    records = []
    try:
        excel_file = pd.ExcelFile(BytesIO(file_bytes))
        for sheet in excel_file.sheet_names:
            raw_df = pd.read_excel(excel_file, sheet_name=sheet, header=None)
            if raw_df.empty or len(raw_df) < 3:
                continue
            header_idx = find_header_row(raw_df)
            df = pd.read_excel(excel_file, sheet_name=sheet, header=header_idx)
            records.extend(normalize_dataframe(df, meta))
    except Exception as e:
        print(f"  [!] Lỗi đọc Excel: {e}")
    return records


def parse_pdf_bytes(file_bytes, meta):
    records = []
    try:
        with pdfplumber.open(BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                for table in page.extract_tables():
                    if not table or len(table) < 2:
                        continue
                    df = pd.DataFrame(table[1:], columns=table[0])
                    records.extend(normalize_dataframe(df, meta))
    except Exception as e:
        print(f"  [!] Lỗi đọc PDF: {e}")
    return records


def extract_period(title_text):
    match = re.search(r"tháng\s*(\d{1,2})[/.-](\d{4})", title_text.lower())
    if match:
        return f"{match.group(2)}-{int(match.group(1)):02d}"
    return datetime.now().strftime("%Y-%m")


def crawl_all_sources():
    all_records = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            ignore_https_errors=True,
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        for source in SOURCES:
            print(f"\n>>> Đang mở bằng Chromium: {source['loai']} - {source['tinh_thanh']}")
            try:
                page.goto(source["url"], wait_until="networkidle", timeout=45000)
                page.wait_for_timeout(3000)

                links = page.eval_on_selector_all(
                    "a[href]",
                    "els => els.map(e => ({text: e.innerText.trim(), href: e.href}))"
                )

                article_links = []
                for item in links:
                    text = item["text"]
                    href = item["href"]
                    if len(text) > 20 and any(kw in text.lower() for kw in source["keywords"]):
                        if href not in [x[1] for x in article_links]:
                            article_links.append((text, href))

                print(f"  + Tìm thấy {len(article_links)} bài viết phù hợp.")

                for title, link in article_links[:3]:
                    print(f"  + Đang vào bài: {title[:70]}...")
                    page.goto(link, wait_until="networkidle", timeout=45000)
                    page.wait_for_timeout(2000)

                    meta = {
                        "tinh_thanh": source["tinh_thanh"],
                        "loai": source["loai"],
                        "ky_cong_bo": extract_period(title),
                        "tieu_de": title
                    }

                    # 1. Kiểm tra file đính kèm (.xlsx, .xls, .pdf)
                    sub_links = page.eval_on_selector_all("a[href]", "els => els.map(e => e.href)")
                    found_files = False
                    for file_href in set(sub_links):
                        lower_href = file_href.lower()
                        if any(lower_href.endswith(ext) for ext in [".xlsx", ".xls", ".pdf"]):
                            found_files = True
                            print(f"    -> Tải file đính kèm: {file_href.split('/')[-1]}")
                            api_resp = context.request.get(file_href, timeout=45000)
                            if api_resp.ok:
                                body = api_resp.body()
                                if lower_href.endswith((".xlsx", ".xls")):
                                    all_records.extend(parse_excel_bytes(body, meta))
                                elif lower_href.endswith(".pdf"):
                                    all_records.extend(parse_pdf_bytes(body, meta))

                    # 2. Nếu không có file đính kèm, thử đọc bảng HTML trực tiếp trong bài viết
                    if not found_files:
                        html_content = page.content()
                        try:
                            tables = pd.read_html(BytesIO(html_content.encode("utf-8")))
                            for tbl in tables:
                                if len(tbl) > 5:
                                    all_records.extend(normalize_dataframe(tbl, meta))
                        except Exception:
                            pass

            except Exception as e:
                print(f"  [X] Lỗi nguồn {source['tinh_thanh']}: {e}")

        browser.close()
    return all_records


def update_master_warehouse(new_records):
    existing_data = []
    if os.path.exists(OUTPUT_JSON):
        try:
            with open(OUTPUT_JSON, "r", encoding="utf-8") as f:
                existing_data = json.load(f).get("records", [])
        except Exception:
            existing_data = []

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
        warehouse_map[unique_key] = item

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
    print(f"HOÀN TẤT: Tổng {len(final_records):,} bản ghi (Mới thêm: {added_count:,})")
    print(f"==========================================")


if __name__ == "__main__":
    records = crawl_all_sources()
    update_master_warehouse(records)
