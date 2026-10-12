import pandas as pd
import re
import unicodedata
from pathlib import Path


# ============================================================
# 1. ĐƯỜNG DẪN
# ============================================================

BASE_DIR = Path(r"D:\Ngân hàng số\Data giữa kì")

INPUT_FILE = BASE_DIR / "step3_tier1_results.xlsx"

OUTPUT_FILE = BASE_DIR / "step4_tier2_results.xlsx"


# ============================================================
# 2. ĐỌC DỮ LIỆU TẦNG 1
# ============================================================

df = pd.read_excel(INPUT_FILE)

print("======================================")
print("BƯỚC 4 - LỌC TẦNG 2")
print("======================================")
print(f"Tổng candidate Tầng 1: {len(df)}")


# ============================================================
# 3. CHUẨN HÓA TEXT
# ============================================================

def normalize_text(text):
    if pd.isna(text):
        return ""

    text = str(text)
    text = unicodedata.normalize("NFC", text)
    text = text.lower()
    text = re.sub(r"\s+", " ", text)

    return text.strip()


# ============================================================
# 4. MATCH TERM
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
    "baas"
}


def contains_term(text, term):

    text = normalize_text(text)
    term = normalize_text(term)

    if not term:
        return False

    if term in SHORT_TERMS:

        pattern = (
            r"(?<![a-zA-Z0-9])"
            + re.escape(term)
            + r"(?![a-zA-Z0-9])"
        )

        return bool(
            re.search(
                pattern,
                text,
                flags=re.IGNORECASE
            )
        )

    return term in text


def find_terms(text, terms):

    found = []

    normalized = normalize_text(text)

    for term in terms:

        if term.startswith("REGEX:"):

            pattern = term.replace(
                "REGEX:",
                "",
                1
            )

            if re.search(
                pattern,
                normalized,
                flags=re.IGNORECASE
            ):
                found.append(
                    f"pattern:{pattern}"
                )

        else:

            if contains_term(
                normalized,
                term
            ):
                found.append(term)

    return found


# ============================================================
# 5. MẪU ĐỊNH LƯỢNG DÙNG CHUNG
# ============================================================

QUANT_PATTERNS = [

    r"\b\d+(?:[.,]\d+)?\s*%",

    r"\b\d+(?:[.,]\d+)?\s*"
    r"(?:triệu|nghìn|ngàn)?\s*"
    r"(?:khách hàng|người dùng|nhân viên|nhân sự)",

    r"\b\d+(?:[.,]\d+)?\s*"
    r"(?:triệu|tỷ|nghìn|ngàn)?\s*"
    r"(?:giao dịch|transaction|transactions)",

    r"\b\d+(?:[.,]\d+)?\s*"
    r"(?:tỷ|triệu)\s*"
    r"(?:đồng|vnd|usd)",

    r"\b\d+\s*"
    r"(?:dự án|projects?|"
    r"quy trình|processes?|"
    r"robot|bots?|"
    r"api|apis|"
    r"đối tác|partners?|"
    r"chương trình|programs?|"
    r"thẻ|cards?)"
]


# ============================================================
# 6. CÁC NHÓM DẤU HIỆU TẦNG 2
# ============================================================

GROUPS = {

    # ========================================================
    # A1 - CHIẾN LƯỢC CĐS
    # ========================================================

    "A1": {

        "Thời gian/lộ trình": [
            "giai đoạn",
            "đến năm",
            "định hướng đến",
            "roadmap",
            "milestone",
            "phase",

            "REGEX:\\b20\\d{2}\\s*[-–]\\s*20\\d{2}\\b",
            "REGEX:\\bđến\\s+(?:năm\\s+)?20\\d{2}\\b",
            "REGEX:\\bby\\s+20\\d{2}\\b"
        ],

        "Mục tiêu": [
            "mục tiêu",
            "hướng tới",
            "mục tiêu chiến lược",
            "target",
            "objective",
            "goal"
        ],

        "Trụ cột/ưu tiên": [
            "trụ cột",
            "ưu tiên chiến lược",
            "lĩnh vực trọng tâm",
            "trọng tâm chiến lược",
            "pillar",
            "strategic priority",
            "priority"
        ],

        "Triển khai": [
            "kế hoạch triển khai",
            "kế hoạch hành động",
            "chương trình hành động",
            "triển khai chiến lược",
            "implementation plan",
            "action plan",
            "strategic initiative"
        ]
    },


    # ========================================================
    # A2 - ĐẦU TƯ CNTT/CĐS
    # ========================================================

    "A2": {

        "Định lượng": [
            "REGEX:" + p
            for p in QUANT_PATTERNS
        ],

        "Dự án cụ thể": [
            "dự án",
            "project",
            "initiative",
            "chương trình đầu tư"
        ],

        "Ngân sách/kế hoạch đầu tư": [
            "ngân sách",
            "vốn đầu tư",
            "kế hoạch đầu tư",
            "investment plan",
            "budget",
            "capital expenditure"
        ],

        "Hạ tầng/nền tảng cụ thể": [
            "core banking",
            "data center",
            "trung tâm dữ liệu",
            "platform",
            "infrastructure",
            "hạ tầng công nghệ"
        ]
    },


    # ========================================================
    # A3 - QUẢN TRỊ CĐS
    # ========================================================

    "A3": {

        "Đơn vị quản trị": [
            "ủy ban",
            "hội đồng",
            "ban chuyển đổi số",
            "khối công nghệ",
            "trung tâm chuyển đổi số",
            "committee",
            "council"
        ],

        "Vai trò/trách nhiệm": [
            "vai trò",
            "trách nhiệm",
            "chức năng",
            "nhiệm vụ",
            "phân công",
            "role",
            "responsibility",
            "mandate"
        ],

        "Giám sát/chỉ đạo": [
            "giám sát",
            "chỉ đạo",
            "điều hành",
            "báo cáo",
            "oversight",
            "supervision",
            "reporting"
        ],

        "Cấp lãnh đạo": [
            "hội đồng quản trị",
            "ban điều hành",
            "tổng giám đốc",
            "board of directors",
            "executive board",
            "management board"
        ]
    },


    # ========================================================
    # A4 - NHÂN LỰC SỐ
    # ========================================================

    "A4": {

        "Đào tạo/chương trình": [
            "chương trình đào tạo",
            "khóa đào tạo",
            "chương trình phát triển",
            "training program",
            "learning program",
            "course"
        ],

        "Quy mô": [
            "REGEX:" + p
            for p in QUANT_PATTERNS
        ],

        "Kỹ năng/chứng chỉ": [
            "kỹ năng số",
            "chứng chỉ",
            "chứng nhận",
            "upskill",
            "reskill",
            "certification",
            "certificate"
        ],

        "Nhân sự chuyên môn": [
            "chuyên gia công nghệ",
            "kỹ sư",
            "data scientist",
            "engineer",
            "technology expert",
            "tuyển dụng công nghệ"
        ]
    },


    # ========================================================
    # B1 - eKYC
    # FINAL: siết mạnh các dấu hiệu triển khai
    # ========================================================

    "B1": {

        "Onboarding": [
            "mở tài khoản trực tuyến",
            "mở tài khoản online",
            "mở tài khoản số",
            "đăng ký tài khoản trực tuyến",
            "onboarding",
            "digital onboarding",
            "online onboarding",
            "account opening",
            "online account opening"
        ],

        "Sinh trắc học": [
            "sinh trắc học",
            "nhận diện khuôn mặt",
            "face recognition",
            "biometric",
            "liveness"
        ],

        "Tích hợp sản phẩm": [
            "mở thẻ trực tuyến",
            "đăng ký thẻ trực tuyến",
            "phát hành thẻ trực tuyến",
            "vay trực tuyến",
            "đăng ký khoản vay trực tuyến",
            "digital lending",
            "online lending",
            "online card issuance"
        ],

        "Quy mô": [
            "REGEX:" + p
            for p in QUANT_PATTERNS
        ]
    },


    # ========================================================
    # B2 - DIGITAL BANKING
    # ========================================================

    "B2": {

        "Quy mô": [
            "REGEX:" + p
            for p in QUANT_PATTERNS
        ],

        "Chuyển tiền": [
            "chuyển tiền",
            "transfer"
        ],

        "Thanh toán": [
            "thanh toán",
            "payment"
        ],

        "Tiết kiệm": [
            "tiết kiệm",
            "saving",
            "deposit"
        ],

        "Tín dụng/vay": [
            "vay",
            "khoản vay",
            "tín dụng",
            "loan",
            "credit"
        ],

        "Đầu tư": [
            "đầu tư",
            "investment"
        ],

        "Hệ sinh thái": [
            "hệ sinh thái",
            "ecosystem",
            "digital ecosystem"
        ],

        "Tỷ lệ giao dịch số": [
            "tỷ lệ giao dịch số",
            "tỷ trọng giao dịch số",
            "giao dịch trực tuyến",
            "digital transaction rate",
            "online transaction"
        ]
    },


    # ========================================================
    # B3 - QR / DIGITAL PAYMENT
    # FINAL: bỏ "ứng dụng", "đối tác" quá rộng
    # ========================================================

    "B3": {

        "Tích hợp ứng dụng": [
            "tích hợp trên ứng dụng",
            "tích hợp vào ứng dụng",
            "thanh toán trên ứng dụng",
            "thanh toán qua ứng dụng",
            "quét qr trên ứng dụng",
            "qr trên ứng dụng",
            "in-app payment",
            "payment in app",
            "integrated into app"
        ],

        "Merchant/điểm chấp nhận": [
            "đơn vị chấp nhận thanh toán",
            "điểm chấp nhận thanh toán",
            "điểm chấp nhận qr",
            "merchant",
            "merchant network",
            "acceptance point"
        ],

        "Quy mô": [
            "REGEX:" + p
            for p in QUANT_PATTERNS
        ]
    },


    # ========================================================
    # B4 - THẺ SỐ/ẢO
    # ========================================================

    "B4": {

        "Phát hành online": [
            "phát hành trực tuyến",
            "phát hành tức thời",
            "online issuance",
            "instant issuance"
        ],

        "Tích hợp ứng dụng/ví": [
            "ứng dụng",
            "mobile banking",
            "app",
            "wallet",
            "ví điện tử"
        ],

        "Quy mô": [
            "REGEX:" + p
            for p in QUANT_PATTERNS
        ]
    },


    # ========================================================
    # C1 - AI
    # FINAL: bỏ các từ quá rộng như khách hàng/risk
    # ========================================================

    "C1": {

        "Tín dụng/rủi ro": [
            "chấm điểm tín dụng",
            "chấm điểm khách hàng",
            "xếp hạng tín dụng",
            "credit scoring",
            "credit assessment",
            "risk scoring",
            "risk assessment",
            "mô hình chấm điểm"
        ],

        "Khách hàng/cá nhân hóa": [
            "cá nhân hóa",
            "gợi ý sản phẩm",
            "đề xuất sản phẩm",
            "phân tích hành vi khách hàng",
            "customer personalization",
            "personalization",
            "recommendation engine",
            "customer recommendation"
        ],

        "Fraud/an ninh": [
            "phát hiện gian lận",
            "phòng chống gian lận",
            "fraud detection",
            "fraud prevention",
            "anomaly detection",
            "phát hiện bất thường"
        ],

        "Chatbot/trợ lý": [
            "chăm sóc khách hàng",
            "hỗ trợ khách hàng",
            "tư vấn khách hàng",
            "trợ lý ảo",
            "virtual assistant",
            "customer service",
            "customer support"
        ],

        "Xử lý tài liệu/hình ảnh": [
            "xử lý tài liệu",
            "nhận diện tài liệu",
            "nhận diện hình ảnh",
            "computer vision",
            "document processing",
            "document recognition",
            "image recognition",
            "intelligent document processing"
        ]
    },


    # ========================================================
    # C2 - CLOUD
    # ========================================================

    "C2": {

        "Loại cloud": [
            "private cloud",
            "hybrid cloud",
            "multi-cloud",
            "multicloud",
            "đám mây riêng",
            "đám mây lai"
        ],

        "Cloud-native": [
            "cloud-native",
            "cloud native",
            "microservices",
            "container",
            "kubernetes"
        ],

        "Migration": [
            "migration",
            "migrate",
            "chuyển dịch",
            "triển khai trên cloud",
            "ứng dụng trên cloud"
        ],

        "Hạ tầng/nền tảng": [
            "cloud platform",
            "cloud infrastructure",
            "nền tảng đám mây",
            "hạ tầng đám mây"
        ]
    },


    # ========================================================
    # C3 - API / OPEN BANKING
    # ========================================================

    "C3": {

        "Open API/Open Banking": [
            "open api",
            "open banking",
            "api platform",
            "api gateway"
        ],

        "Đối tác/hệ sinh thái": [
            "đối tác",
            "partner",
            "hệ sinh thái",
            "ecosystem"
        ],

        "Tích hợp": [
            "tích hợp",
            "integration",
            "kết nối",
            "connect"
        ],

        "BaaS/Embedded": [
            "banking as a service",
            "baas",
            "embedded banking",
            "embedded finance"
        ],

        "Quy mô": [
            "REGEX:\\b\\d+\\s*(?:api|apis)\\b",
            "REGEX:\\b\\d+\\s*(?:đối tác|partners?)\\b"
        ]
    },


    # ========================================================
    # C4 - BIG DATA / DATA PLATFORM
    # ========================================================

    "C4": {

        "Kiến trúc dữ liệu": [
            "data lake",
            "data warehouse",
            "lakehouse",
            "data hub",
            "data platform",
            "nền tảng dữ liệu",
            "kho dữ liệu",
            "hồ dữ liệu"
        ],

        "Quản trị dữ liệu": [
            "data governance",
            "quản trị dữ liệu",
            "data quality",
            "chất lượng dữ liệu"
        ],

        "Analytics": [
            "data analytics",
            "advanced analytics",
            "phân tích dữ liệu",
            "phân tích nâng cao"
        ],

        "Customer data": [
            "customer data",
            "customer 360",
            "dữ liệu khách hàng",
            "360 độ khách hàng"
        ]
    },


    # ========================================================
    # D1 - RPA / AUTOMATION
    # FINAL: bỏ "hiệu quả" quá rộng
    # ========================================================

    "D1": {

        "RPA/robot": [
            "rpa",
            "robot",
            "robotic process automation"
        ],

        "Quy trình/nghiệp vụ": [
            "quy trình",
            "nghiệp vụ",
            "process",
            "workflow"
        ],

        "Quy mô": [
            "REGEX:\\b\\d+\\s*(?:quy trình|processes?|robot|bots?)\\b"
        ],

        "Hiệu quả": [
            "giảm thời gian xử lý",
            "rút ngắn thời gian xử lý",
            "tiết kiệm thời gian",
            "tiết kiệm chi phí",
            "giảm chi phí",
            "tăng năng suất",
            "giảm nhân công",
            "giảm thao tác thủ công",
            "productivity improvement",
            "cost saving",
            "time saving",
            "processing time reduction",
            "manual effort reduction"
        ]
    },


    # ========================================================
    # D2 - SỐ HÓA QUY TRÌNH
    # ========================================================

    "D2": {

        "End-to-end/STP": [
            "end-to-end",
            "xuyên suốt",
            "toàn trình",
            "straight-through processing",
            "stp"
        ],

        "Paperless": [
            "paperless",
            "không giấy tờ",
            "e-office",
            "văn phòng điện tử"
        ],

        "Nghiệp vụ cụ thể": [
            "tín dụng",
            "phê duyệt",
            "thanh toán",
            "vận hành",
            "credit",
            "approval",
            "payment",
            "operations"
        ],

        "Quy mô": [
            "REGEX:\\b\\d+\\s*(?:quy trình|processes?|workflows?)\\b"
        ]
    },


    # ========================================================
    # D3 - AN NINH MẠNG
    # ========================================================

    "D3": {

        "SOC/SIEM": [
            "soc",
            "security operations center",
            "siem"
        ],

        "Tiêu chuẩn/chứng nhận": [
            "iso 27001",
            "iso/iec 27001",
            "pci dss",
            "certification",
            "chứng nhận"
        ],

        "Công nghệ phòng thủ": [
            "zero trust",
            "ddos",
            "firewall",
            "ids",
            "ips",
            "cyber resilience"
        ],

        "Giám sát/phản ứng": [
            "giám sát",
            "phát hiện",
            "ứng phó",
            "incident response",
            "monitoring",
            "detection"
        ],

        "Đào tạo/diễn tập": [
            "diễn tập",
            "đào tạo an toàn thông tin",
            "security awareness",
            "cyber drill"
        ]
    },


    # ========================================================
    # D4 - PHÁT HIỆN GIAN LẬN
    # ========================================================

    "D4": {

        "Hệ thống fraud": [
            "fraud detection",
            "fraud monitoring",
            "anti-fraud",
            "fraud prevention",
            "hệ thống chống gian lận",
            "hệ thống phát hiện gian lận"
        ],

        "Realtime": [
            "real-time",
            "real time",
            "thời gian thực",
            "trực thời"
        ],

        "AI/ML": [
            "trí tuệ nhân tạo",
            "artificial intelligence",
            "machine learning",
            "ai"
        ],

        "Anomaly/monitoring": [
            "anomaly detection",
            "phát hiện bất thường",
            "giao dịch bất thường",
            "transaction monitoring",
            "giám sát giao dịch"
        ]
    }
}


# ============================================================
# 7. LẤY CÁC NHÓM MATCH
# ============================================================

def get_matched_groups(
    criterion,
    context
):

    groups = GROUPS.get(
        criterion,
        {}
    )

    matched_groups = []
    details = []

    for group_name, terms in groups.items():

        found = find_terms(
            context,
            terms
        )

        if found:

            matched_groups.append(
                group_name
            )

            details.append(
                f"{group_name}: "
                + ", ".join(found)
            )

    return (
        matched_groups,
        details
    )


# ============================================================
# 8. LOGIC TẦNG 2 FINAL
# ============================================================

def evaluate_tier2(
    criterion,
    context
):

    matched_groups, details = (
        get_matched_groups(
            criterion,
            context
        )
    )

    group_set = set(
        matched_groups
    )

    score = len(
        matched_groups
    )

    passed = False
    rule_reason = ""


    # ========================================================
    # A1
    # ========================================================

    if criterion == "A1":

        passed = (
            score >= 2
        )

        rule_reason = (
            "Cần ít nhất 2 nhóm dấu hiệu "
            "về lộ trình/thời gian, mục tiêu, "
            "trụ cột hoặc kế hoạch triển khai."
        )


    # ========================================================
    # A2
    # ========================================================

    elif criterion == "A2":

        strong_groups = {
            "Định lượng",
            "Dự án cụ thể",
            "Ngân sách/kế hoạch đầu tư"
        }

        passed = bool(
            group_set
            & strong_groups
        )

        rule_reason = (
            "Cần định lượng, dự án cụ thể "
            "hoặc ngân sách/kế hoạch đầu tư."
        )


    # ========================================================
    # A3
    # ========================================================

    elif criterion == "A3":

        has_unit = (
            "Đơn vị quản trị"
            in group_set
        )

        has_function = bool(
            group_set
            & {
                "Vai trò/trách nhiệm",
                "Giám sát/chỉ đạo"
            }
        )

        passed = (
            has_unit
            and has_function
        )

        rule_reason = (
            "Cần có đơn vị quản trị chính thức "
            "và bằng chứng về vai trò/trách nhiệm "
            "hoặc giám sát/chỉ đạo."
        )


    # ========================================================
    # A4
    # ========================================================

    elif criterion == "A4":

        passed = (
            score >= 2
        )

        rule_reason = (
            "Cần ít nhất 2 nhóm về "
            "đào tạo, quy mô, kỹ năng/chứng chỉ "
            "hoặc nhân sự chuyên môn."
        )


    # ========================================================
    # B1 - FINAL
    # ========================================================

    elif criterion == "B1":

        passed = bool(
            group_set
            & {
                "Onboarding",
                "Tích hợp sản phẩm",
                "Quy mô"
            }
        )

        rule_reason = (
            "Tầng 1 đã xác nhận eKYC/sinh trắc học. "
            "Tầng 2 cần chứng minh triển khai thực tế "
            "qua onboarding, tích hợp sản phẩm hoặc quy mô."
        )


    # ========================================================
    # B2
    # ========================================================

    elif criterion == "B2":

        feature_groups = {
            "Chuyển tiền",
            "Thanh toán",
            "Tiết kiệm",
            "Tín dụng/vay",
            "Đầu tư"
        }

        feature_count = len(
            group_set
            & feature_groups
        )

        passed = (
            "Quy mô" in group_set
            or "Tỷ lệ giao dịch số" in group_set
            or "Hệ sinh thái" in group_set
            or feature_count >= 2
        )

        rule_reason = (
            "Pass nếu có quy mô, tỷ lệ giao dịch số, "
            "hệ sinh thái hoặc ít nhất 2 nhóm tính năng."
        )


    # ========================================================
    # B3 - FINAL
    # ========================================================

    elif criterion == "B3":

        passed = bool(
            group_set
            & {
                "Tích hợp ứng dụng",
                "Merchant/điểm chấp nhận",
                "Quy mô"
            }
        )

        rule_reason = (
            "Tầng 2 cần bằng chứng triển khai thực tế "
            "qua tích hợp ứng dụng, merchant/điểm chấp nhận "
            "hoặc quy mô giao dịch."
        )


    # ========================================================
    # B4
    # ========================================================

    elif criterion == "B4":

        passed = bool(
            group_set
            & {
                "Phát hành online",
                "Tích hợp ứng dụng/ví",
                "Quy mô"
            }
        )

        rule_reason = (
            "Cần phát hành online/tức thời, "
            "tích hợp app/ví hoặc quy mô."
        )


    # ========================================================
    # C1 - FINAL
    # ========================================================

    elif criterion == "C1":

        business_groups = {
            "Tín dụng/rủi ro",
            "Khách hàng/cá nhân hóa",
            "Fraud/an ninh",
            "Chatbot/trợ lý",
            "Xử lý tài liệu/hình ảnh"
        }

        passed = bool(
            group_set
            & business_groups
        )

        rule_reason = (
            "AI phải gắn với ít nhất một use case "
            "nghiệp vụ cụ thể như chấm điểm tín dụng, "
            "cá nhân hóa, fraud, trợ lý khách hàng "
            "hoặc xử lý tài liệu/hình ảnh."
        )


    # ========================================================
    # C2
    # ========================================================

    elif criterion == "C2":

        passed = (
            score >= 1
        )

        rule_reason = (
            "Cần ít nhất một dấu hiệu triển khai cloud "
            "cụ thể như private/hybrid cloud, "
            "cloud-native, migration hoặc hạ tầng cloud."
        )


    # ========================================================
    # C3
    # ========================================================

    elif criterion == "C3":

        has_core = bool(
            group_set
            & {
                "Open API/Open Banking",
                "BaaS/Embedded"
            }
        )

        has_application = bool(
            group_set
            & {
                "Đối tác/hệ sinh thái",
                "Tích hợp",
                "Quy mô",
                "BaaS/Embedded"
            }
        )

        passed = (
            has_core
            and has_application
        )

        rule_reason = (
            "Cần Open API/Open Banking/BaaS "
            "và thêm dấu hiệu tích hợp, đối tác hoặc quy mô."
        )


    # ========================================================
    # C4
    # ========================================================

    elif criterion == "C4":

        if (
            "Kiến trúc dữ liệu"
            in group_set
        ):

            passed = True

        else:

            passed = (
                score >= 2
            )

        rule_reason = (
            "Kiến trúc dữ liệu cụ thể tự đủ mạnh; "
            "nếu không cần ít nhất 2 nhóm khác."
        )


    # ========================================================
    # D1 - FINAL
    # ========================================================

    elif criterion == "D1":

        passed = bool(
            group_set
            & {
                "Quy mô",
                "Hiệu quả"
            }
        )

        rule_reason = (
            "Tầng 1 đã xác nhận automation/RPA. "
            "Tầng 2 cần bằng chứng định lượng về quy mô "
            "hoặc hiệu quả vận hành cụ thể."
        )


    # ========================================================
    # D2
    # ========================================================

    elif criterion == "D2":

        has_strong_process = bool(
            group_set
            & {
                "End-to-end/STP",
                "Paperless"
            }
        )

        has_scope = bool(
            group_set
            & {
                "Nghiệp vụ cụ thể",
                "Quy mô"
            }
        )

        passed = (
            has_strong_process
            and has_scope
        )

        rule_reason = (
            "Cần end-to-end/STP hoặc paperless "
            "kèm nghiệp vụ cụ thể hoặc quy mô."
        )


    # ========================================================
    # D3
    # ========================================================

    elif criterion == "D3":

        strong_groups = {
            "SOC/SIEM",
            "Tiêu chuẩn/chứng nhận",
            "Công nghệ phòng thủ"
        }

        if (
            group_set
            & strong_groups
        ):

            passed = True

        else:

            operational_groups = {
                "Giám sát/phản ứng",
                "Đào tạo/diễn tập"
            }

            passed = (
                len(
                    group_set
                    & operational_groups
                )
                >= 2
            )

        rule_reason = (
            "SOC/SIEM, tiêu chuẩn/chứng nhận "
            "hoặc công nghệ phòng thủ tự đủ mạnh; "
            "nếu chỉ là vận hành thì cần >=2 nhóm."
        )


    # ========================================================
    # D4
    # ========================================================

    elif criterion == "D4":

        passed = bool(
            group_set
            & {
                "Hệ thống fraud",
                "Realtime",
                "AI/ML",
                "Anomaly/monitoring"
            }
        )

        rule_reason = (
            "Tầng 1 đã xác nhận fraud; "
            "Tầng 2 cần ít nhất một năng lực nâng cao "
            "như hệ thống fraud, realtime, AI/ML "
            "hoặc anomaly monitoring."
        )


    # ========================================================
    # FALLBACK
    # ========================================================

    else:

        passed = False

        rule_reason = (
            "Không có rule."
        )


    # ========================================================
    # STRENGTH
    # ========================================================

    if not passed:

        if score == 0:
            strength = "None"

        else:
            strength = "Weak"

    else:

        if score >= 3:
            strength = "Strong"

        else:
            strength = "Medium"


    return {

        "tier2_score":
            score,

        "tier2_strength":
            strength,

        "tier2_signals":
            " | ".join(details),

        "tier2_pass":
            passed,

        "tier2_rule":
            rule_reason
    }


# ============================================================
# 9. CHẠY TẤT CẢ CANDIDATE
# ============================================================

tier2_scores = []
tier2_strengths = []
tier2_signals = []
tier2_passes = []
tier2_rules = []


for _, row in df.iterrows():

    criterion = str(
        row["criterion"]
    ).strip()

    context = row.get(
        "context",
        ""
    )

    result = evaluate_tier2(
        criterion,
        context
    )

    tier2_scores.append(
        result["tier2_score"]
    )

    tier2_strengths.append(
        result["tier2_strength"]
    )

    tier2_signals.append(
        result["tier2_signals"]
    )

    tier2_passes.append(
        result["tier2_pass"]
    )

    tier2_rules.append(
        result["tier2_rule"]
    )


df["tier2_score"] = tier2_scores
df["tier2_strength"] = tier2_strengths
df["tier2_signals"] = tier2_signals
df["tier2_pass"] = tier2_passes
df["tier2_rule"] = tier2_rules


# ============================================================
# 10. SORT
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
    in enumerate(
        criterion_order
    )
}


df["_criterion_order"] = (
    df["criterion"]
    .map(
        criterion_rank
    )
    .fillna(999)
)


df = df.sort_values(
    by=[
        "bank",
        "_criterion_order",
        "tier2_pass",
        "tier2_score",
        "page"
    ],

    ascending=[
        True,
        True,
        False,
        False,
        True
    ]
)


df = df.drop(
    columns=[
        "_criterion_order"
    ]
)


# ============================================================
# 11. THỐNG KÊ
# ============================================================

total = len(df)

passed = int(
    df["tier2_pass"]
    .sum()
)

failed = (
    total
    - passed
)


print("\n======================================")
print("KẾT QUẢ BƯỚC 4 / TẦNG 2")
print("======================================")

print(
    "Tổng candidate Tầng 1:",
    total
)

print(
    "Qua Tầng 2:",
    passed
)

print(
    "Không qua Tầng 2:",
    failed
)

if total > 0:

    print(
        "Tỷ lệ qua Tầng 2:",
        f"{passed / total:.2%}"
    )


print(
    "\n----- Qua Tầng 2 theo ngân hàng -----"
)

print(
    df[
        df["tier2_pass"] == True
    ]
    .groupby(
        "bank",
        observed=True
    )
    .size()
)


print(
    "\n----- Qua Tầng 2 theo tiêu chí -----"
)

pass_by_criterion = (
    df[
        df["tier2_pass"] == True
    ]
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
    pass_by_criterion
)


print(
    "\n----- Tổng Tầng 1 theo tiêu chí -----"
)

total_by_criterion = (
    df
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
    total_by_criterion
)


print(
    "\n----- Pass rate theo tiêu chí -----"
)

pass_rate = (
    pass_by_criterion
    / total_by_criterion
    .replace(
        0,
        pd.NA
    )
)

print(
    pass_rate
    .apply(
        lambda x:
        f"{x:.2%}"
        if pd.notna(x)
        else "N/A"
    )
)


print(
    "\n----- Strength -----"
)

print(
    df[
        "tier2_strength"
    ]
    .value_counts()
)


print(
    "\n----- Tier2 pass theo bank + criterion -----"
)

pivot = (
    df[
        df["tier2_pass"]
        == True
    ]
    .groupby(
        [
            "bank",
            "criterion"
        ],
        observed=True
    )
    .size()
    .unstack(
        fill_value=0
    )
)

print(
    pivot
)


# ============================================================
# 12. EXPORT
# ============================================================

df.to_excel(
    OUTPUT_FILE,
    index=False
)


print("\n======================================")
print("✅ HOÀN TẤT BƯỚC 4")
print("======================================")

print(
    "File:",
    OUTPUT_FILE
)