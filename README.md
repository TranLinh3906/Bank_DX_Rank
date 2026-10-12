# Hệ thống chấm điểm Ngân hàng số

Dự án này xây dựng một quy trình có thể tái lập để trích xuất, lọc và tổng hợp bằng chứng phục vụ đánh giá mức độ chuyển đổi số của một số ngân hàng thương mại Việt Nam dựa trên các tài liệu công khai như Báo cáo thường niên và Báo cáo phát triển bền vững.

Quy trình tập trung vào việc tìm bằng chứng theo 16 tiêu chí ngân hàng số, lọc bằng chứng theo hai tầng và tổng hợp kết quả thành file Excel cuối cùng để phục vụ bước chấm điểm và xếp hạng sau đó.

## Các ngân hàng được phân tích

Dữ liệu hiện tại gồm 5 ngân hàng:

- BIDV
- Vietcombank (VCB)
- OCB
- MB
- VietinBank

## Nguồn báo cáo

Dự án sử dụng Báo cáo thường niên và Báo cáo phát triển bền vững năm 2025 khi có sẵn.

Nguồn dữ liệu hiện tại gồm:

- Vietcombank: Báo cáo thường niên 2025 và Báo cáo phát triển bền vững 2025
- OCB: Báo cáo thường niên 2025 và Báo cáo phát triển bền vững 2025
- MB: Báo cáo thường niên 2025
- VietinBank: Báo cáo thường niên 2025 và Báo cáo phát triển bền vững 2025
- BIDV: Báo cáo thường niên 2025 và Báo cáo phát triển bền vững 2025

## Khung đánh giá

Hệ thống sử dụng 16 tiêu chí, chia thành 4 nhóm chính.

### A. Chiến lược và nguồn lực số

- A1 – Chiến lược chuyển đổi số
- A2 – Đầu tư CNTT / chuyển đổi số
- A3 – Quản trị chuyển đổi số
- A4 – Nhân lực số

### B. Sản phẩm số và trải nghiệm khách hàng

- B1 – eKYC
- B2 – Mobile / Digital Banking
- B3 – QR / Digital Payment
- B4 – Thẻ số / thẻ ảo

### C. Công nghệ và dữ liệu

- C1 – Trí tuệ nhân tạo (AI)
- C2 – Điện toán đám mây (Cloud)
- C3 – API / Open Banking
- C4 – Big Data / Data Platform

### D. Vận hành số và an toàn

- D1 – RPA / Automation
- D2 – Số hóa quy trình
- D3 – An ninh mạng
- D4 – Phát hiện gian lận

## Quy trình xử lý dữ liệu

Quy trình gồm 5 bước chính.

### Bước 1 – Thu thập báo cáo

Thu thập Báo cáo thường niên và Báo cáo phát triển bền vững của từng ngân hàng và lưu theo từng thư mục riêng.

Ví dụ:

```text
Data giữa kì/

Bước 2 – Trích xuất nội dung PDF theo từng trang
File code:
extract_reports.py

Script này đọc toàn bộ file PDF, trích xuất nội dung theo từng trang và tạo dữ liệu ở cấp độ trang.
File đầu ra:
bank_reports_pages.xlsx

Các cột chính:
bank
report_type
filename
page
text

Việc lưu theo từng trang giúp giữ được nguồn báo cáo và số trang để có thể kiểm tra lại bằng chứng sau này.
Bước 3 – Lọc Tầng 1
File code:
search_keywords.py

Tầng 1 dùng để tìm các đoạn có khả năng liên quan đến từng tiêu chí.
Các phương pháp được sử dụng gồm:
- Tìm từ khóa chính xác
- Tìm tổ hợp từ khóa
- Tìm từ khóa theo khoảng cách gần nhau
- Sử dụng bộ từ khóa riêng cho từng tiêu chí
Mục tiêu của Tầng 1 là ưu tiên độ bao quát, tức cố gắng không bỏ sót bằng chứng có liên quan.
File đầu ra:
step3_tier1_results.xlsx

Các cột chính:
bank
criterion
criterion_name
report_type
filename
page
match_type
matched_keyword
context

Bước 4 – Lọc Tầng 2
File code:
filter_tier2.py

Tầng 2 tiếp tục đánh giá các kết quả từ Tầng 1 bằng các dấu hiệu mạnh và cụ thể hơn cho từng tiêu chí.
Ví dụ:
- A1: tìm dấu hiệu về lộ trình, giai đoạn, mục tiêu, trụ cột hoặc kế hoạch triển khai
- A2: tìm quy mô đầu tư, dự án cụ thể hoặc ngân sách đầu tư
- A3: tìm cơ cấu quản trị, vai trò, trách nhiệm và cơ chế giám sát
- A4: tìm chương trình đào tạo, quy mô nhân lực hoặc kỹ năng số
- B1: tìm triển khai eKYC trong onboarding, sản phẩm hoặc quy mô sử dụng
- B2: tìm quy mô người dùng, giao dịch số, tính năng hoặc hệ sinh thái số
- B3: tìm tích hợp QR vào ứng dụng, mạng lưới merchant hoặc quy mô giao dịch
- B4: tìm phát hành thẻ trực tuyến, tích hợp ứng dụng hoặc quy mô
- C1: tìm các use case AI cụ thể trong tín dụng, cá nhân hóa, gian lận, trợ lý ảo hoặc xử lý tài liệu
- C2: tìm dấu hiệu triển khai cloud như private cloud, hybrid cloud, cloud-native hoặc migration
- C3: tìm tích hợp API với đối tác, hệ sinh thái, BaaS hoặc Open Banking
- C4: tìm kiến trúc dữ liệu, quản trị dữ liệu và phân tích dữ liệu
- D1: tìm quy mô hoặc hiệu quả triển khai automation/RPA
- D2: tìm số hóa quy trình end-to-end, STP hoặc paperless
- D3: tìm SOC, SIEM, ISO 27001, PCI DSS và các biện pháp an ninh mạng
- D4: tìm hệ thống phát hiện gian lận, giám sát thời gian thực, AI/ML hoặc anomaly detection
File đầu ra:
step4_tier2_results.xlsx

Các cột bổ sung:
tier2_score
tier2_strength
tier2_signals
tier2_pass
tier2_rule

tier2_pass = True chỉ thể hiện rằng đoạn bằng chứng đủ mạnh để được giữ lại ở Tầng 2, chưa đồng nghĩa với việc ngân hàng tự động được chấm mức điểm cao nhất.
Bước 5 – Tổng hợp evidence
File code:
build_final_evidence.py

Bước này thực hiện:
- Giữ nguyên kết quả Tầng 1
- Giữ nguyên toàn bộ kết quả Tầng 2
- Loại các dòng trùng bằng chứng do match nhiều từ khóa
- Gộp các từ khóa cùng thuộc một evidence
- Chuẩn hóa tên ngân hàng
- Tạo file dữ liệu sạch để phục vụ chấm điểm
File đầu ra:
Evidence_Final.xlsx

File cuối gồm 4 sheet:
1. Step3_Tier1
Giữ nguyên toàn bộ kết quả Tầng 1.
2. Step4_Tier2
Giữ toàn bộ kết quả Bước 4, bao gồm cả các dòng pass và fail.
3. Evidence_Raw
Dữ liệu evidence đã được làm sạch và loại trùng.
4. Evidence_By_Bank
Bằng chứng được trình bày theo cấu trúc:
Ngân hàng
    Tầng 1
        A1
        A2
        ...
        D4

    Tầng 2
        A1
        A2
        ...
        D4

Sheet này phục vụ việc đọc và kiểm tra evidence thủ công.
Các file code chính
extract_reports.py
search_keywords.py
filter_tier2.py
build_final_evidence.py

Các file dữ liệu chính
bank_reports_pages.xlsx
step3_tier1_results.xlsx
step4_tier2_results.xlsx
Evidence_Final.xlsx

Thư viện Python sử dụng
Dự án sử dụng các thư viện chính:
pymupdf
pandas
openpyxl

Cài đặt bằng lệnh:
py -m pip install pymupdf pandas openpyxl

Cách chạy dự án
Chạy các file theo đúng thứ tự:
py extract_reports.py
py search_keywords.py
py filter_tier2.py
py build_final_evidence.py

Mỗi bước sử dụng đầu ra của bước trước nên cần chạy tuần tự.
Ghi chú phương pháp
Tầng 1 và Tầng 2 có mục đích khác nhau.
Tầng 1 ưu tiên độ bao quát và dùng để tìm các đoạn có khả năng liên quan đến tiêu chí.
Tầng 2 dùng các điều kiện chặt hơn để chọn ra các bằng chứng cụ thể và có chất lượng cao hơn.
Một evidence vượt Tầng 2 chưa đồng nghĩa với việc ngân hàng tự động được chấm điểm cao nhất. Điểm cuối cùng cần được xác định ở cấp độ ngân hàng – tiêu chí sau khi tổng hợp toàn bộ evidence liên quan.
Một số tiêu chí cần xem xét nhiều evidence cùng lúc. Ví dụ, để đánh giá mức độ ứng dụng AI, có thể cần tổng hợp nhiều đoạn bằng chứng ở các nghiệp vụ khác nhau thay vì chỉ dựa trên một đoạn duy nhất.
Phạm vi hiện tại
Repository hiện tập trung vào:
- Thu thập dữ liệu
- Trích xuất nội dung báo cáo
- Lọc evidence Tầng 1
- Lọc evidence Tầng 2
- Tổng hợp evidence
Các bước tiếp theo của dự án có thể bao gồm:
- Chấm điểm 0–1–2 cho từng tiêu chí
- Tính điểm theo từng nhóm tiêu chí
- Tính chỉ số tổng hợp
- Xếp hạng các ngân hàng
- Xây dựng dashboard trực quan hóa kết quả
Cấu trúc repository
Digital-banking-scoring-system/
│
├── extract_reports.py
├── search_keywords.py
├── filter_tier2.py
├── build_final_evidence.py
│
├── bank_reports_pages.xlsx
├── step3_tier1_results.xlsx
├── step4_tier2_results.xlsx
├── Evidence_Final.xlsx
│
└── README.md

Mục đích sử dụng
Các báo cáo gốc sử dụng trong dự án là các tài liệu công khai do các ngân hàng công bố.
Repository được xây dựng phục vụ mục đích học tập và nghiên cứu.

Sau khi lưu file, chạy:

```powershell
git add README.md
git commit -m "Add Vietnamese README"
git push
├── BIDV/
├── MB/
├── OCB/
├── VCB/
└── VietinBank/
