import pandas as pd
import re
import unicodedata
from pathlib import Path


# ============================================================
# 1. ĐƯỜNG DẪN
# ============================================================

BASE_DIR = Path(r"D:\Ngân hàng số\Data giữa kì")

INPUT_FILE = BASE_DIR / "bank_reports_pages.xlsx"

OUTPUT_FILE = BASE_DIR / "step3_tier1_results.xlsx"


# ============================================================
# 2. THÔNG SỐ
# ============================================================

PROXIMITY_WINDOW = 150
CONTEXT_WINDOW = 300


# ============================================================
# 3. ĐỌC DỮ LIỆU
# ============================================================

df = pd.read_excel(INPUT_FILE)
BANK_NAME_MAP = {
    "Vietinbank": "VietinBank",
    "VietinBank": "VietinBank",
    "vietinbank": "VietinBank",
    "VCB": "VCB",
    "BIDV": "BIDV",
    "MB": "MB",
    "OCB": "OCB"
}

df["bank"] = (
    df["bank"]
    .astype(str)
    .str.strip()
    .replace(BANK_NAME_MAP)
)

print("======================================")
print("BƯỚC 3 - LỌC TẦNG 1")
print("======================================")

print("Tổng số trang:", len(df))


# ============================================================
# 4. NORMALIZE
# ============================================================

def normalize_text(text):

    if pd.isna(text):
        return ""

    text = str(text)

    text = unicodedata.normalize(
        "NFC",
        text
    )

    text = text.lower()

    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text.strip()


# ============================================================
# 5. SHORT TERM MATCH
# ============================================================

SHORT_TERMS = {
    "ai",
    "api",
    "rpa",
    "ocr",
    "nlp",
    "soc",
    "siem",
    "stp",
    "bpm",
    "baas",
    "qr"
}


def term_positions(text, term):

    text = normalize_text(text)
    term = normalize_text(term)

    positions = []

    if not term:
        return positions


    # Acronym / từ rất ngắn:
    # phải có word boundary
    if term in SHORT_TERMS:

        pattern = (
            r"(?<![a-zA-Z0-9])"
            + re.escape(term)
            + r"(?![a-zA-Z0-9])"
        )

        for match in re.finditer(
            pattern,
            text,
            flags=re.IGNORECASE
        ):

            positions.append(
                match.start()
            )

        return positions


    # Từ/cụm bình thường
    start = 0

    while True:

        pos = text.find(
            term,
            start
        )

        if pos == -1:
            break

        positions.append(pos)

        start = pos + len(term)

    return positions


# ============================================================
# 6. BỘ TIÊU CHÍ FINAL CHO TẦNG 1
# ============================================================

CRITERIA = {

    # ========================================================
    # A1 - CHIẾN LƯỢC CĐS
    # ========================================================

    "A1": {

        "name": "Chiến lược CĐS",

        "exact": [
            "chiến lược chuyển đổi số",
            "lộ trình chuyển đổi số",
            "kế hoạch chuyển đổi số",
            "định hướng chuyển đổi số",
            "chuyển đổi số",

            "chiến lược ngân hàng số",
            "lộ trình ngân hàng số",
            "ngân hàng số",

            "chiến lược số",
            "lộ trình số hóa",

            "digital transformation strategy",
            "digital transformation",
            "digital strategy",
            "digital roadmap",
            "digital banking strategy"
        ],

        "group1": [
            "chiến lược",
            "lộ trình",
            "kế hoạch",
            "định hướng",

            "strategy",
            "strategic",
            "roadmap",
            "plan"
        ],

        "group2": [
            "chuyển đổi số",
            "ngân hàng số",
            "số hóa",
            "công nghệ số",
            "nền tảng số",
            "doanh nghiệp số",

            "digital transformation",
            "digital banking",
            "digitalization",
            "digitalisation",
            "digital platform"
        ]
    },


    # ========================================================
    # A2 - ĐẦU TƯ CNTT/CĐS
    # ========================================================

    "A2": {

        "name": "Đầu tư CNTT/CĐS",

        "exact": [
            "đầu tư công nghệ",
            "đầu tư cntt",
            "đầu tư công nghệ thông tin",
            "đầu tư chuyển đổi số",

            "ngân sách công nghệ",
            "ngân sách cntt",

            "hạ tầng công nghệ",
            "hạ tầng cntt",

            "technology investment",
            "it investment",
            "digital investment",

            "technology infrastructure",
            "it infrastructure"
        ],

        "group1": [
            "đầu tư",
            "ngân sách",
            "vốn đầu tư",

            "investment",
            "budget"
        ],

        "group2": [
            "công nghệ thông tin",
            "cntt",
            "chuyển đổi số",
            "hạ tầng công nghệ",
            "hạ tầng cntt",

            "technology",
            "digital transformation",
            "it infrastructure"
        ]
    },


    # ========================================================
    # A3 - QUẢN TRỊ CĐS
    # ========================================================

    "A3": {

        "name": "Quản trị CĐS",

        "exact": [
            "quản trị chuyển đổi số",
            "quản trị công nghệ",

            "ban chuyển đổi số",
            "ủy ban chuyển đổi số",
            "hội đồng công nghệ",

            "khối công nghệ",
            "khối công nghệ và chuyển đổi số",

            "trung tâm chuyển đổi số",

            "digital governance",
            "technology governance",
            "digital transformation governance",
            "digital transformation committee",
            "steering committee"
        ],

        "group1": [
            "ủy ban",
            "hội đồng",
            "ban chuyển đổi số",
            "khối công nghệ",
            "trung tâm chuyển đổi số",

            "committee",
            "governance"
        ],

        "group2": [
            "chuyển đổi số",
            "công nghệ số",
            "công nghệ thông tin",

            "digital transformation",
            "technology governance",
            "digital governance"
        ]
    },


    # ========================================================
    # A4 - NHÂN LỰC SỐ
    # ========================================================

    "A4": {

        "name": "Nhân lực số",

        "exact": [
            "nhân lực số",
            "nguồn nhân lực số",

            "nhân lực công nghệ",
            "nhân sự công nghệ",

            "đào tạo chuyển đổi số",
            "đào tạo công nghệ",
            "đào tạo kỹ năng số",

            "kỹ năng số",

            "digital talent",
            "digital workforce",
            "technology workforce",
            "digital skills",
            "digital training"
        ],

        "group1": [
            "nhân lực",
            "nhân sự",
            "nguồn nhân lực",

            "đào tạo",
            "tuyển dụng",
            "kỹ năng",

            "workforce",
            "talent",
            "training",
            "skills"
        ],

        "group2": [
            "chuyển đổi số",
            "công nghệ số",
            "công nghệ thông tin",

            "digital transformation",
            "digital skills",
            "technology"
        ]
    },


    # ========================================================
    # B1 - eKYC
    # ========================================================

    "B1": {

        "name": "eKYC",

        "exact": [
            "ekyc",
            "e-kyc",

            "định danh điện tử",
            "định danh khách hàng điện tử",

            "xác thực điện tử",
            "xác thực trực tuyến",

            "sinh trắc học",
            "xác thực sinh trắc học",

            "nhận diện khuôn mặt",

            "video kyc",

            "electronic kyc",
            "electronic identification",
            "digital identification",

            "biometric authentication",
            "face recognition",
            "liveness detection"
        ],

        "group1": [
            "định danh khách hàng",
            "định danh",

            "xác thực",
            "nhận diện",

            "authentication",
            "identification"
        ],

        "group2": [
            "điện tử",
            "trực tuyến",

            "sinh trắc học",
            "khuôn mặt",

            "electronic",
            "online",
            "biometric",
            "face recognition"
        ]
    },


    # ========================================================
    # B2 - MOBILE / DIGITAL BANKING
    #
    # SỬA LỚN:
    # bỏ "giao dịch + số"
    # ========================================================

    "B2": {

        "name": "Mobile/Digital Banking",

        "exact": [
            "ngân hàng số",
            "ngân hàng điện tử",
            "ngân hàng di động",

            "ứng dụng ngân hàng",
            "ứng dụng ngân hàng số",

            "nền tảng ngân hàng số",

            "mobile banking",
            "digital banking",
            "electronic banking",
            "digital banking platform",

            "ứng dụng mobile banking"
        ],

        "group1": [
            "ứng dụng ngân hàng",
            "nền tảng ngân hàng",
            "dịch vụ ngân hàng",

            "banking application",
            "banking platform",
            "banking service"
        ],

        "group2": [
            "số",
            "điện tử",
            "di động",

            "digital",
            "electronic",
            "mobile"
        ]
    },


    # ========================================================
    # B3 - QR / DIGITAL PAYMENT
    # ========================================================

    "B3": {

        "name": "QR/Digital Payment",

        "exact": [
            "thanh toán qr",
            "mã qr",
            "qr code",
            "vietqr",

            "thanh toán số",
            "thanh toán điện tử",
            "thanh toán không tiền mặt",

            "ví điện tử",

            "napas 247",

            "qr payment",
            "digital payment",
            "electronic payment",
            "cashless payment"
        ],

        "group1": [
            "thanh toán",
            "payment"
        ],

        "group2": [
            "qr",
            "vietqr",

            "điện tử",
            "không tiền mặt",

            "digital payment",
            "electronic payment",
            "cashless"
        ]
    },


    # ========================================================
    # B4 - THẺ SỐ / THẺ ẢO
    # ========================================================

    "B4": {

        "name": "Thẻ số/ảo",

        "exact": [
            "thẻ số",
            "thẻ ảo",
            "thẻ phi vật lý",
            "thẻ điện tử",

            "phát hành thẻ trực tuyến",
            "phát hành thẻ tức thời",

            "virtual card",
            "digital card",
            "non-physical card",

            "virtual debit card",
            "virtual credit card",

            "instant card issuance"
        ],

        "group1": [
            "thẻ",
            "card"
        ],

        "group2": [
            "ảo",
            "phi vật lý",

            "phát hành trực tuyến",
            "phát hành tức thời",

            "virtual",
            "non-physical",
            "digital card",
            "instant issuance"
        ]
    },


    # ========================================================
    # C1 - AI
    # ========================================================

    "C1": {

        "name": "AI",

        "exact": [
            "trí tuệ nhân tạo",
            "artificial intelligence",

            "machine learning",
            "học máy",
            "máy học",

            "generative ai",
            "genai",

            "ai chatbot",
            "chatbot ai",

            "computer vision",
            "natural language processing",
            "nlp",

            "ocr"
        ],

        "group1": [],
        "group2": []
    },


    # ========================================================
    # C2 - CLOUD
    #
    # Không siết thêm vì hiện candidate rất ít.
    # ========================================================

    "C2": {

        "name": "Cloud",

        "exact": [
            "điện toán đám mây",

            "nền tảng đám mây",
            "hạ tầng đám mây",

            "đám mây riêng",
            "đám mây lai",

            "cloud computing",
            "cloud platform",
            "cloud infrastructure",

            "private cloud",
            "hybrid cloud",

            "multi-cloud",
            "multicloud",

            "cloud-native",
            "cloud native"
        ],

        "group1": [],
        "group2": []
    },


    # ========================================================
    # C3 - API / OPEN BANKING
    #
    # Giữ rộng vì hiện chỉ có ~55 candidate.
    # ========================================================

    "C3": {

        "name": "API/Open Banking",

        "exact": [
            "open api",

            "api platform",
            "api gateway",

            "kết nối api",
            "tích hợp api",

            "ngân hàng mở",
            "open banking",

            "api banking",

            "application programming interface",

            "embedded banking",
            "embedded finance",

            "banking as a service",
            "baas"
        ],

        "group1": [
            "kết nối đối tác",
            "tích hợp đối tác",
            "hệ sinh thái mở",

            "partner integration",
            "partner ecosystem"
        ],

        "group2": [
            "api",
            "open api",
            "open banking",
            "banking as a service",
            "baas"
        ]
    },


    # ========================================================
    # C4 - BIG DATA / DATA PLATFORM
    #
    # SỬA:
    # không còn "dữ liệu + quản trị" kiểu quá rộng.
    # ========================================================

    "C4": {

        "name": "Big Data/Data Platform",

        "exact": [
            "dữ liệu lớn",
            "big data",

            "nền tảng dữ liệu",
            "data platform",

            "kho dữ liệu",
            "data warehouse",

            "hồ dữ liệu",
            "data lake",

            "data hub",
            "lakehouse",

            "quản trị dữ liệu",
            "data governance",

            "phân tích dữ liệu",
            "data analytics",

            "advanced analytics",

            "customer data platform",
            "cdp"
        ],

        "group1": [
            "dữ liệu",
            "data"
        ],

        "group2": [
            "warehouse",
            "lake",
            "lakehouse",
            "hub",

            "analytics",

            "governance",

            "platform"
        ]
    },


    # ========================================================
    # D1 - RPA / AUTOMATION
    # ========================================================

    "D1": {

        "name": "RPA/Automation",

        "exact": [
            "tự động hóa quy trình",
            "tự động hóa nghiệp vụ",

            "robotic process automation",
            "process automation",
            "workflow automation",

            "rpa",

            "robot phần mềm",

            "bpm"
        ],

        "group1": [
            "tự động hóa",
            "automation",
            "robot"
        ],

        "group2": [
            "quy trình",
            "nghiệp vụ",
            "workflow",
            "process"
        ]
    },


    # ========================================================
    # D2 - SỐ HÓA QUY TRÌNH
    # ========================================================

    "D2": {

        "name": "Số hóa quy trình",

        "exact": [
            "số hóa quy trình",
            "quy trình số",

            "quy trình nghiệp vụ số",
            "quy trình điện tử",

            "digital process",
            "process digitalization",
            "digital workflow",

            "paperless",
            "e-office",

            "straight-through processing",
            "stp",

            "end-to-end digital"
        ],

        "group1": [
            "quy trình",
            "nghiệp vụ",
            "workflow",
            "process"
        ],

        "group2": [
            "số hóa",
            "điện tử",

            "không giấy tờ",
            "paperless",

            "digital workflow",
            "digital process"
        ]
    },


    # ========================================================
    # D3 - AN NINH MẠNG
    #
    # SỬA LỚN:
    # bỏ các combined quá chung
    # như an toàn + hệ thống,
    # bảo mật + thông tin...
    # ========================================================

    "D3": {

        "name": "An ninh mạng",

        "exact": [
            "an ninh mạng",
            "an ninh thông tin",

            "an toàn thông tin",
            "an toàn hệ thống thông tin",

            "bảo mật thông tin",
            "bảo mật dữ liệu",

            "cybersecurity",
            "cyber security",

            "information security",
            "data security",
            "network security",

            "security operations center",
            "soc",

            "siem",

            "zero trust",

            "iso 27001",
            "iso/iec 27001",

            "pci dss",

            "cyber resilience",

            "ddos",

            "incident response",
            "security monitoring"
        ],

        # D3 giờ chủ yếu dùng exact.
        # Combined chỉ giữ các tổ hợp mạnh.
        "group1": [
            "an ninh mạng",
            "an toàn thông tin",
            "bảo mật thông tin",

            "cyber security",
            "cybersecurity",
            "information security"
        ],

        "group2": [
            "giám sát",
            "ứng phó",
            "phát hiện",
            "quản lý",

            "monitoring",
            "incident",
            "detection",
            "management"
        ]
    },


    # ========================================================
    # D4 - FRAUD
    # ========================================================

    "D4": {

        "name": "Phát hiện gian lận",

        "exact": [
            "phát hiện gian lận",
            "phòng chống gian lận",
            "chống gian lận",
            "giám sát gian lận",

            "giao dịch bất thường",
            "phát hiện bất thường",
            "giám sát giao dịch",

            "fraud detection",
            "fraud monitoring",
            "fraud prevention",
            "anti-fraud",

            "real-time fraud",

            "anomaly detection",
            "transaction anomaly",
            "transaction monitoring"
        ],

        "group1": [
            "gian lận",
            "bất thường",

            "fraud",
            "anomaly"
        ],

        "group2": [
            "phát hiện",
            "phòng chống",
            "giám sát",

            "detection",
            "prevention",
            "monitoring"
        ]
    }
}


# ============================================================
# 7. LẤY CONTEXT
# ============================================================

def get_context(normalized_text, position):

    start = max(
        0,
        position - CONTEXT_WINDOW
    )

    end = min(
        len(normalized_text),
        position + CONTEXT_WINDOW
    )

    return normalized_text[
        start:end
    ].strip()


# ============================================================
# 8. EXACT SEARCH
# ============================================================

def find_exact_matches(text, keywords):

    normalized = normalize_text(text)

    matches = []

    for keyword in keywords:

        positions = term_positions(
            normalized,
            keyword
        )

        for pos in positions:

            matches.append({
                "match_type": "exact",
                "matched_keyword": keyword,
                "position": pos
            })

    return matches


# ============================================================
# 9. COMBINED SEARCH
# ============================================================

def find_combined_matches(
    text,
    group1,
    group2
):

    if not group1 or not group2:
        return []

    normalized = normalize_text(text)

    matches = []

    seen = set()


    for word1 in group1:

        positions1 = term_positions(
            normalized,
            word1
        )

        for pos1 in positions1:

            local_start = max(
                0,
                pos1 - PROXIMITY_WINDOW
            )

            local_end = min(
                len(normalized),
                pos1
                + len(normalize_text(word1))
                + PROXIMITY_WINDOW
            )

            local_text = normalized[
                local_start:local_end
            ]


            for word2 in group2:

                positions2 = term_positions(
                    local_text,
                    word2
                )

                if positions2:

                    key = (
                        word1,
                        word2,
                        pos1
                    )

                    if key not in seen:

                        matches.append({
                            "match_type":
                                "combined",

                            "matched_keyword":
                                f"{word1} + {word2}",

                            "position":
                                pos1
                        })

                        seen.add(key)

    return matches


# ============================================================
# 10. QUÉT TOÀN BỘ
# ============================================================

results = []


for _, row in df.iterrows():

    raw_text = row.get(
        "text",
        ""
    )

    if pd.isna(raw_text):
        continue

    normalized_text = normalize_text(
        raw_text
    )


    for criterion_code, config in CRITERIA.items():

        # ----------------------------------------
        # Exact
        # ----------------------------------------

        exact_matches = find_exact_matches(
            normalized_text,
            config["exact"]
        )


        # ----------------------------------------
        # Combined
        # ----------------------------------------

        combined_matches = find_combined_matches(
            normalized_text,
            config["group1"],
            config["group2"]
        )


        all_matches = (
            exact_matches
            + combined_matches
        )


        for match in all_matches:

            context = get_context(
                normalized_text,
                match["position"]
            )


            results.append({

                "bank":
                    row["bank"],

                "criterion":
                    criterion_code,

                "criterion_name":
                    config["name"],

                "report_type":
                    row["report_type"],

                "filename":
                    row["filename"],

                "page":
                    row["page"],

                "match_type":
                    match["match_type"],

                "matched_keyword":
                    match["matched_keyword"],

                "context":
                    context
            })


# ============================================================
# 11. DATAFRAME
# ============================================================

df_results = pd.DataFrame(
    results
)


# ============================================================
# 12. DROP DUPLICATE KỸ THUẬT
#
# Chỉ loại những dòng thực sự giống nhau hoàn toàn.
# Không gộp evidence ở đây.
# Việc gộp evidence để B5 làm.
# ============================================================

if not df_results.empty:

    df_results = (
        df_results
        .drop_duplicates(
            subset=[
                "bank",
                "criterion",
                "filename",
                "page",
                "match_type",
                "matched_keyword",
                "context"
            ]
        )
    )


# ============================================================
# 13. SORT
# ============================================================

criterion_order = [
    "A1", "A2", "A3", "A4",
    "B1", "B2", "B3", "B4",
    "C1", "C2", "C3", "C4",
    "D1", "D2", "D3", "D4"
]


criterion_rank = {
    criterion: i
    for i, criterion
    in enumerate(criterion_order)
}


if not df_results.empty:

    df_results["_criterion_order"] = (
        df_results["criterion"]
        .map(criterion_rank)
        .fillna(999)
    )


    df_results = (
        df_results
        .sort_values(
            by=[
                "bank",
                "_criterion_order",
                "report_type",
                "filename",
                "page"
            ]
        )
    )


    df_results = df_results.drop(
        columns=[
            "_criterion_order"
        ]
    )


# ============================================================
# 14. THỐNG KÊ
# ============================================================

print("\n======================================")
print("KẾT QUẢ BƯỚC 3 / TẦNG 1")
print("======================================")

print(
    "Tổng candidate Tầng 1:",
    len(df_results)
)


if not df_results.empty:

    print(
        "\n----- Theo ngân hàng -----"
    )

    print(
        df_results
        .groupby(
            "bank",
            observed=True
        )
        .size()
    )


    print(
        "\n----- Theo tiêu chí -----"
    )

    print(
        df_results
        .groupby(
            "criterion",
            observed=True
        )
        .size()
        .reindex(
            criterion_order,
            fill_value=0
        )
    )


    print(
        "\n----- Exact / Combined -----"
    )

    print(
        df_results[
            "match_type"
        ].value_counts()
    )


    print(
        "\n----- Top keyword theo tiêu chí -----"
    )

    for criterion in criterion_order:

        temp = df_results[
            df_results["criterion"]
            == criterion
        ]

        print(
            f"\n{criterion}:"
        )

        print(
            temp[
                "matched_keyword"
            ]
            .value_counts()
            .head(10)
        )


# ============================================================
# 15. EXPORT
# ============================================================

df_results.to_excel(
    OUTPUT_FILE,
    index=False
)


print("\n======================================")
print("✅ HOÀN TẤT BƯỚC 3")
print("======================================")

print(
    "File:",
    OUTPUT_FILE
)