from pathlib import Path

BASE_DIR = Path(r"D:\Ngân hàng số\Data giữa kì")

pdf_files = list(BASE_DIR.rglob("*.pdf"))

print(f"Tìm thấy {len(pdf_files)} file PDF:\n")

for file in pdf_files:
    print(file)


#Bước 2 
import re

def clean_excel_text(text):
    if not isinstance(text, str):
        return text

    # Xóa các ký tự điều khiển không hợp lệ với Excel
    text = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', text)

    # Giới hạn độ dài tối đa của 1 ô Excel
    return text[:32767]

import fitz
import pandas as pd
from pathlib import Path

# =========================
# 1. ĐƯỜNG DẪN PROJECT
# =========================

BASE_DIR = Path(r"D:\Ngân hàng số\Data giữa kì")

OUTPUT_FILE = BASE_DIR / "bank_reports_pages.xlsx"


# =========================
# 2. TÌM TOÀN BỘ FILE PDF
# =========================

pdf_files = list(BASE_DIR.rglob("*.pdf"))

print(f"Tìm thấy {len(pdf_files)} file PDF.\n")


# =========================
# 3. EXTRACT TEXT THEO TỪNG TRANG
# =========================

rows = []

for pdf_path in pdf_files:

    # Tên folder = tên ngân hàng
    bank = pdf_path.parent.name

    filename = pdf_path.name

    # Phân loại báo cáo theo tên file
    filename_lower = filename.lower()

    if "sustainability" in filename_lower:
        report_type = "SustainabilityReport"
    else:
        report_type = "AnnualReport"

    print(f"Đang xử lý: {bank} - {filename}")

    try:
        doc = fitz.open(pdf_path)

        print(f"  Số trang: {len(doc)}")

        for page_number, page in enumerate(doc, start=1):

            text = page.get_text("text").strip()
            text = clean_excel_text(text)

            rows.append({
                "bank": bank,
                "report_type": report_type,
                "filename": filename,
                "page": page_number,
                "text": text
            })

        doc.close()

        print("  ✅ Xong\n")

    except Exception as e:
        print(f"  ❌ Lỗi: {e}\n")


# =========================
# 4. TẠO DATAFRAME
# =========================

df = pd.DataFrame(rows)

print("Đã extract tổng cộng:", len(df), "trang")


# =========================
# 5. XUẤT FILE EXCEL
# =========================

df.to_excel(
    OUTPUT_FILE,
    index=False
)

print("\n============================")
print("✅ HOÀN TẤT BƯỚC 2")
print("File kết quả:")
print(OUTPUT_FILE)
print("============================")