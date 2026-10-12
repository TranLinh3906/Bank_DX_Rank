
# ============================================================
# BANK-DX 2025 | TASK 3
# HE THONG CHAM DIEM VA XEP HANG CHUYEN DOI SO
#
# TASK 1: Framework va rubric 0-1-2
# TASK 2: Evidence_Raw, Tier 1, Tier 2
# TASK 3: Estimated scoring, BDTI, Ranking, Charts
#
# Chay:
# python3 scoring.py all
# python3 scoring.py score
# ============================================================

from pathlib import Path
import argparse
import re
import unicodedata

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt


# ============================================================
# 1. CAU HINH
# ============================================================

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
OUT = BASE / "output"
CHARTS = OUT / "charts"

FRAMEWORK_PATH = DATA / "PCCV-NHS-2.xlsx"
LOCAL_EVIDENCE_PATH = DATA / "evidence.xlsx"
REVIEW_PATH = OUT / "review_scores.xlsx"

BANKS = ["BIDV", "MB", "OCB", "VCB", "VietinBank"]
AUTOMATED_STATUS = "automated"
METHODOLOGY_LABEL = "BDTI ước tính từ phương pháp chấm điểm tự động"


def resolve_evidence_path(base=BASE):
    upstream = (
        base.parent
        / "Digital-banking-scoring-system"
        / "Evidence_Final.xlsx"
    )
    local = base / "data" / "evidence.xlsx"
    return upstream if upstream.is_file() else local


EVIDENCE_PATH = resolve_evidence_path()

PILLARS = {
    "A": "Chien luoc va nguon luc so",
    "B": "San pham so va trai nghiem khach hang",
    "C": "Cong nghe va du lieu",
    "D": "Van hanh so va an toan"
}

CODES = [
    f"{p}{i}"
    for p in "ABCD"
    for i in range(1, 5)
]

# Moi tieu chi toi da 2 diem.
# Moi tru cot co 4 tieu chi, toi da 8 diem.
MAX_SCORE = 2

# Tin hieu muc 1 va dieu kien muc 2 theo rubric goc.
# Ca hai phai duoc tim thay trong cung doan context; Tier 2 khong quyet dinh diem.
RULES = {
    "A1": (
        r"(?:chien luoc|dinh huong|lo trinh|strategy|roadmap).{0,60}"
        r"(?:chuyen doi so|so hoa|ngan hang so|cong nghe so|digital)|"
        r"(?:digital transformation|chuyen doi so|so hoa|ngan hang so)"
        r".{0,60}(?:strategy|roadmap|chien luoc|lo trinh)",
        r"(?:lo trinh|roadmap).{0,50}(?:muc tieu|giai doan|202[5-9]|2030)|"
        r"(?:muc tieu|ke hoach).{0,40}(?:giai doan|202[5-9]|2030)"
    ),
    "A2": (
        r"(?:dau tu|du an|investment).{0,50}"
        r"(?:cntt|cong nghe|chuyen doi so|so hoa|technology|it infrastructure)|"
        r"(?:cntt|cong nghe|chuyen doi so|so hoa|technology|it infrastructure)"
        r".{0,50}(?:dau tu|du an|investment)",
        r"(?:dau tu|du an|investment).{0,50}"
        r"(?:ngan sach|quy mo|von dau tu|"
        r"\b\d+(?:[,.]\d+)?\s*(?:ty|trieu)\s*(?:dong|vnd)\b)|"
        r"(?:ngan sach|quy mo|von dau tu|"
        r"\b\d+(?:[,.]\d+)?\s*(?:ty|trieu)\s*(?:dong|vnd)\b)"
        r".{0,50}(?:dau tu|du an|investment)"
    ),
    "A3": (
        r"(?:quan tri|ban chi dao|uy ban|governance).{0,50}"
        r"(?:chuyen doi so|so hoa|digital|cong nghe)|"
        r"(?:chuyen doi so|so hoa|digital|cong nghe).{0,50}"
        r"(?:quan tri|ban chi dao|uy ban|governance)",
        r"(?:co che.{0,30}(quan tri|dieu hanh|giam sat)|"
        r"phan cong.{0,30}(trach nhiem|nhiem vu|giam sat)|"
        r"ban chi dao.{0,40}(dieu phoi|phan cong|giam sat|co che))"
    ),
    "A4": (
        r"dao tao|nhan luc|nhan vien|digital skill",
        r"chuong trinh dao tao|dao tao.{0,50}khoa hoc|"
        r"khoa hoc.{0,50}dao tao|quy mo.{0,30}(dao tao|nhan luc)|"
        r"(?:dao tao|nhan luc|nhan vien).{0,50}"
        r"\b\d+\s*(?:nhan vien|can bo|nguoi)\b"
    ),
    "B1": (
        r"ekyc|e-kyc|dinh danh dien tu|xac thuc dien tu",
        r"tich hop|trien khai rong|dien rong|toan hang|toan he thong|"
        r"trien khai tren.{0,30}(toan hang|toan he thong|dien rong)"
    ),
    "B2": (
        r"mobile banking|digital banking|smartbanking|"
        r"mb bank|ocb omni|vietinbank ipay|vcb digibank",
        r"\b\d[\d,.]*\s*(?:trieu|nghin|ngan|million)\s*"
        r"(?:khach hang|nguoi dung|giao dich)\b|"
        r"tinh nang.{0,60}(ho tro|cho phep|bao gom|tich hop)|"
        r"so luong.{0,30}(khach hang|nguoi dung|giao dich)"
    ),
    "B3": (
        r"vietqr|qr code|qr payment|thanh toan qr|"
        r"thanh toan so",
        r"(?:qr|thanh toan so).{0,50}"
        r"(?:tich hop|ket noi|mang luoi|he sinh thai|diem chap nhan)|"
        r"(?:tich hop|ket noi|mang luoi|he sinh thai|diem chap nhan)"
        r".{0,50}(?:qr|thanh toan so)"
    ),
    "B4": (
        r"the ao|the phi vat ly|virtual card|digital card|"
        r"contactless",
        r"(?:the ao|the phi vat ly|virtual card|digital card|contactless)"
        r".{0,50}(?:tich hop|he sinh thai|vi dien tu)|"
        r"(?:tich hop|he sinh thai|vi dien tu).{0,50}"
        r"(?:the ao|the phi vat ly|virtual card|digital card|contactless)"
    ),
    "C1": (
        r"tri tue nhan tao|artificial intelligence|"
        r"\bai\b|machine learning",
        r"nhieu nghiep vu|nhieu linh vuc|"
        r"\b\d+\s+nghiep vu\b|"
        r"cac linh vuc.{0,60}(tin dung|gian lan|khach hang|van hanh)"
    ),
    "C2": (
        r"cloud|dien toan dam may",
        r"nen tang.{0,40}(cloud|dam may)|cloud native|"
        r"ha tang.{0,40}(cloud|dam may)|toan hang|quy mo.{0,30}cloud"
    ),
    "C3": (
        r"open api|open banking|\bapi\b",
        r"ket noi.{0,50}(doi tac|he sinh thai)|"
        r"he sinh thai.{0,50}(api|ket noi|tich hop)|"
        r"tich hop.{0,50}(doi tac|he sinh thai)"
    ),
    "C4": (
        r"big data|data platform|nen tang du lieu|du lieu lon",
        r"nen tang du lieu|data platform|kho du lieu|"
        r"phan tich du lieu.{0,50}(khach hang|kinh doanh|van hanh)"
    ),
    "D1": (
        r"\brpa\b|robotic process|automation|tu dong hoa",
        r"nhieu quy trinh|nhieu nghiep vu|toan hang|"
        r"\b\d+\s+quy trinh\b"
    ),
    "D2": (
        r"so hoa quy trinh|quy trinh so|digital workflow|"
        r"so hoa nghiep vu",
        r"(?:so hoa|quy trinh so|digital workflow).{0,50}"
        r"(?:toan hang|dong bo|dien rong|nhieu quy trinh|toan bo)|"
        r"(?:toan hang|dong bo|dien rong|nhieu quy trinh|toan bo)"
        r".{0,50}(?:so hoa|quy trinh so|digital workflow)"
    ),
    "D3": (
        r"an ninh mang|an toan thong tin|cybersecurity|"
        r"bao mat",
        r"iso 27001|soc\s*[12]?|pci dss|"
        r"he thong.{0,25}(?:giam sat an ninh|bao mat|an toan thong tin)|"
        r"trung tam.{0,25}(?:giam sat an ninh|an ninh mang|bao mat)"
    ),
    "D4": (
        r"phat hien gian lan|fraud detection|chong gian lan|"
        r"kiem soat gian lan",
        r"(?:phat hien|kiem soat|chong) gian lan.{0,30}"
        r"(?:tu dong|thoi gian thuc|machine learning|phan tich)|"
        r"(?:tu dong|thoi gian thuc|machine learning|phan tich)"
        r".{0,30}(?:phat hien|kiem soat|chong) gian lan"
    )
}

NEGATED_SIGNAL = re.compile(
    r"\b(?:khong|chua)\s+"
    r"(?:co|duoc|trien khai|su dung|ap dung|tich hop|"
    r"xay dung|van hanh|thanh lap|dau tu)\b.{0,50}$"
)


# ============================================================
# 2. HAM HO TRO
# ============================================================

def create_folders():
    OUT.mkdir(parents=True, exist_ok=True)
    CHARTS.mkdir(parents=True, exist_ok=True)


def save_csv(df, name):
    path = OUT / name
    df.to_csv(path, index=False, encoding="utf-8-sig")
    print(f"[SAVED] {path}")


def normalize_text(value):
    """Chuyen tieng Viet co dau thanh dang de tim kiem."""
    if pd.isna(value):
        return ""

    text = str(value).lower()
    text = text.replace("đ", "d")

    text = unicodedata.normalize("NFKD", text)
    text = "".join(
        c for c in text
        if not unicodedata.combining(c)
    )

    text = re.sub(r"\s+", " ", text)
    return text.strip()


def normalize_bank(value):
    text = normalize_text(value)
    text = re.sub(r"[^a-z0-9]", "", text)

    mapping = {
        "bidv": "BIDV",
        "mb": "MB",
        "mbbank": "MB",
        "ocb": "OCB",
        "vcb": "VCB",
        "vietcombank": "VCB",
        "vietinbank": "VietinBank",
        "ctg": "VietinBank"
    }

    return mapping.get(text, str(value).strip())


def normalize_code(value):
    return str(value).strip().upper().replace(" ", "")


def validate_files():
    for path in [FRAMEWORK_PATH, EVIDENCE_PATH]:
        if not path.exists():
            raise FileNotFoundError(
                f"Khong tim thay file: {path}"
            )


# ============================================================
# 3. DOC FRAMEWORK TASK 1
# ============================================================

def load_framework():
    df = pd.read_excel(
        FRAMEWORK_PATH,
        sheet_name="task1-Khung 16 chỉ tiêu"
    )

    # Doc bang TEN cot thuc te cua file PCCV.
    required = ["Mã", "Chỉ tiêu", "Điểm 0–1–2"]

    for column in required:
        if column not in df.columns:
            raise ValueError(
                f"Framework thieu cot: {column}"
            )

    framework = df.rename(columns={
        "Mã": "criterion",
        "Chỉ tiêu": "criterion_name",
        "Điểm 0–1–2": "rubric"
    })

    framework["criterion"] = (
        framework["criterion"]
        .map(normalize_code)
    )

    framework = framework[
        framework["criterion"].isin(CODES)
    ].copy()

    if (
        len(framework) != 16
        or framework["criterion"].duplicated().any()
        or set(framework["criterion"]) != set(CODES)
    ):
        raise ValueError(
            "Framework phai co dung 16 ma A1-D4."
        )

    framework["pillar"] = (
        framework["criterion"].str[0]
    )

    framework["pillar_name"] = (
        framework["pillar"].map(PILLARS)
    )

    cols = [
        "criterion", "criterion_name",
        "rubric", "pillar", "pillar_name"
    ]

    framework = (
        framework[cols]
        .set_index("criterion")
        .loc[CODES]
        .reset_index()
    )

    print("[OK] Framework: 16 criteria")
    return framework


# ============================================================
# 4. DOC VA KIEM TRA EVIDENCE TASK 2
# ============================================================

def load_evidence():
    print(f"[INFO] Evidence source: {EVIDENCE_PATH}")
    df = pd.read_excel(
        EVIDENCE_PATH,
        sheet_name="Evidence_Raw"
    )

    required = [
        "bank", "criterion", "filename",
        "page", "context", "filter_stage"
    ]

    missing = [
        c for c in required if c not in df.columns
    ]

    if missing:
        raise ValueError(
            f"Evidence thieu cot: {missing}"
        )

    print(f"[INFO] Evidence input: {len(df)} rows")

    df["bank"] = df["bank"].map(normalize_bank)
    df["criterion"] = df[
        "criterion"
    ].map(normalize_code)

    df["context"] = (
        df["context"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df["page"] = pd.to_numeric(
        df["page"], errors="coerce"
    )

    if "tier2_score" not in df.columns:
        df["tier2_score"] = np.nan

    df["tier2_score"] = pd.to_numeric(
        df["tier2_score"], errors="coerce"
    )

    df["filter_stage"] = (
        df["filter_stage"]
        .fillna("")
        .astype(str)
    )

    invalid_mask = (
        ~df["bank"].isin(BANKS)
        | ~df["criterion"].isin(CODES)
        | df["context"].eq("")
        | df["page"].isna()
    )

    save_csv(
        df[invalid_mask],
        "invalid_evidence.csv"
    )

    df = df[~invalid_mask].copy()

    duplicate_cols = [
        c for c in [
            "bank", "criterion", "filename",
            "page", "match_type",
            "matched_keyword", "context",
            "filter_stage"
        ]
        if c in df.columns
    ]

    df = df.drop_duplicates(
        subset=duplicate_cols
    ).copy()

    df["is_tier2"] = (
        df["filter_stage"]
        .str.contains("2", regex=False)
    )

    df["text_normalized"] = (
        df["context"].map(normalize_text)
    )

    print(f"[OK] Valid evidence: {len(df)} rows")

    return df


# ============================================================
# 5. DE XUAT DIEM 0-1-2
# ============================================================

def evaluate_candidate(row, criterion):
    """
    De xuat theo tin hieu muc 1 va dieu kien muc 2 cua rubric.
    Chi la goi y de nguoi dung duyet, khong phai diem chinh thuc.
    """

    text = row["text_normalized"]
    basic_pattern, level2_pattern = RULES[criterion]

    def find_context_signals(pattern, context):
        matches = []
        for match in re.finditer(pattern, context):
            start = max(
                context.rfind(".", 0, match.start()),
                context.rfind(";", 0, match.start()),
                context.rfind("\n", 0, match.start())
            ) + 1
            prefix = context[start:match.start()].rstrip()
            if NEGATED_SIGNAL.search(prefix):
                continue
            matches.append(match)
        return matches

    basic_signal = ""
    level2_signal = ""
    for sentence in re.split(r"(?<=[.!?;])\s+|\n+", text):
        basic_matches = find_context_signals(basic_pattern, sentence)
        if not basic_matches:
            continue
        if not basic_signal:
            basic_signal = basic_matches[0].group(0)

        level2_matches = find_context_signals(
            level2_pattern, sentence
        )
        related_pairs = [
            (basic, strong)
            for basic in basic_matches
            for strong in level2_matches
            if abs(basic.start() - strong.start()) <= 100
        ]
        if related_pairs:
            basic, strong = min(
                related_pairs,
                key=lambda pair: abs(
                    pair[0].start() - pair[1].start()
                )
            )
            basic_signal = basic.group(0)
            level2_signal = strong.group(0)
            break

    if level2_signal:
        proposed = 2
        confidence = "medium"
        reason = (
            "De xuat 2: doan bang chung co tin hieu muc 1 "
            f"('{basic_signal}') va dieu kien muc 2 "
            f"('{level2_signal}') theo rubric; nguoi duyet "
            "can xac nhan y nghia trong ngu canh."
        )

    elif basic_signal:
        proposed = 1
        confidence = "medium"
        reason = (
            "De xuat 1: co tin hieu muc 1 "
            f"('{basic_signal}') nhung doan trich chua co "
            "dieu kien muc 2 ro rang theo rubric."
        )

    else:
        proposed = 0
        confidence = "low"
        reason = (
            "De xuat 0: doan trich khong co tin hieu muc 1 "
            "duoc xac nhan; khong ket luan ngan hang khong "
            "trien khai."
        )

    priority = {
        0: "Cao: xac minh bao cao goc va pham vi tim kiem",
        1: "Thuong: doi chieu lai dieu kien muc 1",
        2: "Cao: xac minh dieu kien muc 2 trong ngu canh"
    }[proposed]
    warning = {
        0: (
            "Chua tim thay bang chung phu hop trong tap du lieu; "
            "khong ket luan ngan hang khong trien khai."
        ),
        1: (
            ""
        ),
        2: (
            "Goi y muc 2 tu heuristic; can doi chieu rubric va "
            "ngu canh bao cao goc."
        )
    }[proposed]

    return (
        proposed, confidence, reason, basic_signal,
        level2_signal, priority, warning
    )


def build_suggestions(framework, evidence):
    all_cases = []
    selected_evidence = []

    for bank in BANKS:
        for criterion in CODES:
            subset = evidence[
                (evidence["bank"] == bank)
                & (evidence["criterion"] == criterion)
            ].copy()

            fw = framework[
                framework["criterion"] == criterion
            ].iloc[0]

            result = {
                "bank": bank,
                "criterion": criterion,
                "criterion_name": fw["criterion_name"],
                "pillar": criterion[0],
                "rubric": fw["rubric"],
                "candidate_count": len(subset),
                "suggested_score": 0,
                "confidence": "low",
                "reason": "",
                "rubric_match": "",
                "review_priority": "",
                "warning": "",
                "filename": "",
                "page": "",
                "context": "",
                "status": AUTOMATED_STATUS
            }

            if subset.empty:
                result["reason"] = (
                    "Khong co candidate trong Evidence_Raw; "
                    "khong co trang trich dan. Can ra soat bao "
                    "cao goc va pham vi tim kiem; chua du co so "
                    "ket luan muc 0."
                )
                result["review_priority"] = (
                    "Cao: ra soat thu cong bao cao goc; "
                    "chua co trang trich dan"
                )
                result["warning"] = (
                    "Khong co candidate hoac trang trich dan; "
                    "can ra soat bao cao goc truoc khi dien giai diem."
                )
                reports = sorted(
                    evidence.loc[
                        evidence["bank"].eq(bank), "filename"
                    ].dropna().astype(str).unique()
                )
                result["filename"] = "; ".join(reports)
                result["page"] = (
                    "N/A - khong co candidate; can ra soat bao cao"
                )
                all_cases.append(result)
                continue

            evaluated = []

            for _, row in subset.iterrows():
                (
                    score, conf, reason, basic_signal,
                    level2_signal, priority, warning
                ) = (
                    evaluate_candidate(row, criterion)
                )

                item = {
                    "bank": bank,
                    "criterion": criterion,
                    "filename": row["filename"],
                    "page": row["page"],
                    "context": row["context"],
                    "filter_stage": row["filter_stage"],
                    "tier2_score": row["tier2_score"],
                    "suggested_score": score,
                    "confidence": conf,
                    "reason": reason,
                    "rubric_match": (
                        f"Muc 1: {basic_signal or 'khong thay'}; "
                        f"Muc 2: {level2_signal or 'khong thay'}"
                    ),
                    "review_priority": priority,
                    "warning": warning
                }

                evaluated.append(item)

            ranked = pd.DataFrame(evaluated)

            ranked = ranked.sort_values(
                "suggested_score",
                ascending=False,
                kind="stable"
            )

            # Luu 3 doan de nguoi duyet xem xet.
            best_three = ranked.head(3)

            selected_evidence.extend(
                best_three.to_dict("records")
            )

            best = ranked.iloc[0]

            result.update({
                "suggested_score": int(
                    best["suggested_score"]
                ),
                "confidence": best["confidence"],
                "reason": best["reason"],
                "rubric_match": best["rubric_match"],
                "review_priority": best["review_priority"],
                "warning": best["warning"],
                "filename": best["filename"],
                "page": best["page"],
                "context": best["context"]
            })

            all_cases.append(result)

    review = pd.DataFrame(all_cases)
    candidates = pd.DataFrame(selected_evidence)

    if len(review) != 80:
        raise ValueError(
            "So truong hop cham diem khong bang 80."
        )

    return review, candidates


# ============================================================
# 6. LUU BANG DUYET EXCEL
# ============================================================

def save_review(review):
    review = review.drop(
        columns=[
            "approved_score", "reviewer", "review_note"
        ],
        errors="ignore"
    )
    review["status"] = AUTOMATED_STATUS

    with pd.ExcelWriter(
        REVIEW_PATH, engine="openpyxl"
    ) as writer:
        review.to_excel(
            writer, sheet_name="Scoring_Review",
            index=False
        )

        sheet = writer.sheets["Scoring_Review"]
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions

        widths = {
            "A": 16, "B": 12, "C": 30,
            "D": 12, "E": 45, "F": 17,
            "G": 18, "H": 16, "I": 50,
            "J": 42, "K": 40, "L": 16,
            "M": 65, "N": 18, "O": 17,
            "P": 20, "Q": 50
        }

        from openpyxl.utils import get_column_letter

        for idx, column in enumerate(
            review.columns, start=1
        ):
            letter = get_column_letter(idx)
            sheet.column_dimensions[letter].width = (
                widths.get(letter, 20)
            )

    print("[SAVED] review_scores.xlsx")


# ============================================================
# 7. TINH DIEM A, B, C, D VA BDTI
# ============================================================

def calculate_scores(df, score_column):
    data = df.copy()

    data["status"] = AUTOMATED_STATUS
    data["methodology"] = METHODOLOGY_LABEL
    if "confidence" not in data.columns:
        data["confidence"] = "low"
    if "warning" not in data.columns:
        data["warning"] = ""

    data["final_score"] = pd.to_numeric(
        data[score_column], errors="coerce"
    )

    if (
        len(data) != 80
        or data.duplicated(
            ["bank", "criterion"]
        ).any()
    ):
        raise ValueError(
            "Can dung 80 cap bank-criterion."
        )

    if not data["final_score"].isin([0, 1, 2]).all():
        raise ValueError(
            "Diem phai thuoc tap {0,1,2}."
        )

    data["pillar"] = data["criterion"].str[0]

    pillars = (
        data.groupby(
            ["bank", "pillar"], as_index=False
        )
        .agg(
            raw_score=("final_score", "sum"),
            count=("criterion", "nunique")
        )
    )

    if (
        len(pillars) != 20
        or not pillars["count"].eq(4).all()
    ):
        raise ValueError(
            "Moi tru cot phai co 4 chi tieu."
        )

    pillars["pillar_score"] = (
        pillars["raw_score"] / 8 * 100
    )
    pillars["status"] = AUTOMATED_STATUS
    pillars["methodology"] = METHODOLOGY_LABEL

    wide = pillars.pivot(
        index="bank",
        columns="pillar",
        values="pillar_score"
    ).reset_index()

    wide["BDTI"] = (
        wide[["A", "B", "C", "D"]]
        .mean(axis=1)
    )
    wide["status"] = AUTOMATED_STATUS
    wide["methodology"] = METHODOLOGY_LABEL
    warning_counts = (
        data.assign(
            _has_warning=data["warning"]
            .fillna("")
            .astype(str)
            .str.strip()
            .ne("")
        )
        .groupby("bank")["_has_warning"]
        .sum()
    )
    low_confidence_counts = (
        data.assign(
            _is_low_confidence=data["confidence"]
            .fillna("")
            .astype(str)
            .str.lower()
            .eq("low")
        )
        .groupby("bank")["_is_low_confidence"]
        .sum()
    )
    wide["warning_count"] = (
        wide["bank"].map(warning_counts).astype(int)
    )
    wide["low_confidence_count"] = (
        wide["bank"].map(low_confidence_counts).astype(int)
    )

    wide["strongest_pillar"] = wide.apply(
        lambda r: ",".join(
            p for p in "ABCD"
            if r[p] == max(r[x] for x in "ABCD")
        ),
        axis=1
    )

    wide["weakest_pillar"] = wide.apply(
        lambda r: ",".join(
            p for p in "ABCD"
            if r[p] == min(r[x] for x in "ABCD")
        ),
        axis=1
    )

    # Dong hang dung competition ranking; display_order chi la thu tu hien thi.
    ranking = wide.sort_values(
        ["BDTI", "bank"],
        ascending=[False, True]
    ).copy()

    ranking["rank"] = (
        ranking["BDTI"]
        .rank(method="min", ascending=False)
        .astype(int)
    )

    ranking.insert(
        0, "display_order",
        range(1, len(ranking) + 1)
    )

    ranking = ranking[
        ["display_order", "rank", "bank", "BDTI",
         "A", "B", "C", "D",
         "strongest_pillar", "weakest_pillar",
         "status", "methodology",
         "warning_count", "low_confidence_count"]
    ]

    # Kiem tra BDTI bang tong 16 diem / 32 * 100.
    manual = (
        data.groupby("bank")["final_score"]
        .sum() / 32 * 100
    )

    automatic = (
        ranking.set_index("bank")["BDTI"]
    )

    if not np.allclose(
        manual.loc[BANKS],
        automatic.loc[BANKS]
    ):
        raise ValueError(
            "Sai lech khi kiem tra BDTI."
        )

    print("[OK] BDTI calculation verified")

    return data, pillars, wide, ranking


# ============================================================
# 8. TRUC QUAN HOA
# ============================================================

def draw_charts(ranking, prefix):
    chart = ranking.sort_values("BDTI")

    fig, ax = plt.subplots(figsize=(9, 5))

    bars = ax.barh(
        chart["bank"],
        chart["BDTI"]
    )

    ax.bar_label(bars, fmt="%.2f")
    ax.set_xlim(0, 110)
    ax.set_xlabel("BDTI / 100")
    ax.set_title(f"Estimated BDTI Ranking ({METHODOLOGY_LABEL})")

    plt.tight_layout()
    plt.savefig(
        CHARTS / f"{prefix}_ranking.png",
        dpi=200
    )
    plt.close()

    fig, ax = plt.subplots(figsize=(10, 6))

    ranking.set_index("bank")[
        ["A", "B", "C", "D"]
    ].plot.bar(ax=ax, rot=0)

    ax.set_ylim(0, 105)
    ax.set_ylabel("Score / 100")
    ax.set_title("Estimated digital transformation pillar scores")

    plt.tight_layout()
    plt.savefig(
        CHARTS / f"{prefix}_pillars.png",
        dpi=200
    )
    plt.close()

    matrix = (
        ranking.set_index("bank")[
            ["A", "B", "C", "D"]
        ]
    )

    fig, ax = plt.subplots(figsize=(8, 5))

    image = ax.imshow(
        matrix.to_numpy(),
        cmap="Blues",
        vmin=0, vmax=100,
        aspect="auto"
    )

    ax.set_xticks(range(4), list("ABCD"))
    ax.set_yticks(
        range(len(matrix)), matrix.index
    )

    for i in range(len(matrix)):
        for j in range(4):
            value = matrix.iloc[i, j]
            ax.text(
                j, i, f"{value:.1f}",
                ha="center", va="center",
                color=(
                    "white" if value > 60
                    else "black"
                )
            )

    fig.colorbar(image, ax=ax)
    ax.set_title(f"Estimated BDTI pillar heatmap ({METHODOLOGY_LABEL})")

    plt.tight_layout()
    plt.savefig(
        CHARTS / f"{prefix}_heatmap.png",
        dpi=200
    )
    plt.close()

    print("[OK] Charts created")


# ============================================================
# 9. CHAY TU DONG - KET QUA SO BO
# ============================================================

def run_all():
    create_folders()
    validate_files()

    print("\n=== STEP 1: FRAMEWORK ===")
    framework = load_framework()

    print("\n=== STEP 2: EVIDENCE ===")
    evidence = load_evidence()

    save_csv(
        evidence.drop(columns=["text_normalized"]),
        "evidence_candidates.csv"
    )

    print("\n=== STEP 3: SCORE SUGGESTIONS ===")

    review, candidates = build_suggestions(
        framework, evidence
    )

    save_review(review)

    save_csv(
        candidates,
        "selected_evidence.csv"
    )

    print("\n=== STEP 4: PROVISIONAL BDTI ===")

    data, pillars, scores, ranking = (
        calculate_scores(
            review, "suggested_score"
        )
    )

    save_csv(
        data, "provisional_detail.csv"
    )
    save_csv(
        pillars, "provisional_pillars.csv"
    )
    save_csv(
        scores, "provisional_scores.csv"
    )
    save_csv(
        ranking, "provisional_ranking.csv"
    )

    draw_charts(ranking, "provisional")

    print("\n=== PROVISIONAL RANKING ===")
    print(
        ranking[
            ["rank", "bank", "BDTI"]
        ].to_string(index=False)
    )

    print(
        "\nWARNING: This is a heuristic "
        "ranking, NOT a human-validated result."
    )


# ============================================================
# 10. TAO BDTI UOC TINH TU PHUONG PHAP TU DONG
# ============================================================

def run_automated():
    create_folders()
    print(
        f"[INFO] Generating estimated BDTI with status "
        f"'{AUTOMATED_STATUS}' using the automated scoring methodology."
    )

    if not REVIEW_PATH.exists():
        raise FileNotFoundError(
            "Hay chay 'all' truoc."
        )

    review = pd.read_excel(
        REVIEW_PATH,
        sheet_name="Scoring_Review"
    )

    required = [
        "bank", "criterion", "suggested_score", "status",
        "filename", "page", "context"
    ]

    missing = [
        c for c in required
        if c not in review.columns
    ]

    if missing:
        raise ValueError(
            f"Review thieu cot: {missing}"
        )

    if (
        len(review) != 80
        or review.duplicated(
            ["bank", "criterion"]
        ).any()
        or set(zip(
            review["bank"], review["criterion"]
        )) != {
            (bank, code)
            for bank in BANKS for code in CODES
        }
    ):
        raise ValueError(
            "Review khong du 80 cap duy nhat."
        )

    score = pd.to_numeric(
        review["suggested_score"],
        errors="coerce"
    )

    status = (
        review["status"]
        .fillna("")
        .astype(str)
        .str.lower()
        .str.strip()
    )

    citation_ok = (
        review["filename"].notna()
        & pd.to_numeric(
            review["page"],
            errors="coerce"
        ).gt(0)
        & review["context"].notna()
        & review["context"]
          .fillna("").astype(str).str.strip().ne("")
    )

    valid_automated_rows = (
        score.isin([0, 1, 2])
        & status.eq(AUTOMATED_STATUS)
        & (score.eq(0) | citation_ok)
    )

    errors = review[~valid_automated_rows].copy()

    if not errors.empty:
        save_csv(errors, "validation_errors.csv")
        raise ValueError(
            f"Con {len(errors)} diem chua "
            "duoc duyet hop le."
        )

    review["suggested_score"] = score.astype(int)

    data, pillars, scores, ranking = (
        calculate_scores(
            review, "suggested_score"
        )
    )

    save_csv(data, "processed_scores.csv")
    save_csv(data, "final_dataset.csv")
    save_csv(pillars, "pillar_scores.csv")
    save_csv(scores, "scores.csv")
    save_csv(ranking, "ranking.csv")

    draw_charts(ranking, "estimated")

    print("\n=== ESTIMATED BDTI RANKING ===")
    print(
        ranking[
            ["rank", "bank", "BDTI"]
        ].to_string(index=False)
    )

    print(
        f"\n[FINISHED] Estimated BDTI generated "
        f"with status '{AUTOMATED_STATUS}'."
    )


# ============================================================
# 11. MAIN
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description="BANK-DX Task 3 Scoring"
    )

    parser.add_argument(
        "mode",
        choices=["all", "score"],
        help="all: auto suggestion; "
             "score: generate estimated automated BDTI"
    )

    args = parser.parse_args()

    if args.mode == "all":
        run_all()
    else:
        run_automated()


if __name__ == "__main__":
    main()
