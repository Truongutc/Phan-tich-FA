---
name: chung-khoan
description: Use ONLY when analyzing Vietnamese securities/brokerage stocks (CTCK). Front-loaded keywords: chứng khoán, CTCK, môi giới, margin, tự doanh, FVTPL, IB, lưu ký, quản lý quỹ, SSI, VND, HCM, VCI, FTS, SHS, BSI, VIX, MBS, CTS, AGR, BVS, APG, ORS, TVS, VDS.
---

# KĨ NĂNG PHÂN TÍCH NGÀNH CHỨNG KHOÁN (CTCK)

> **Template dùng chung cho MỌI mã CTCK**: `template_securities.py` — hàm `run_securities_analysis(ticker, raw_data)`, được gọi trực tiếp từ `run_analysis.py` khi ticker được nhận diện là CTCK (xem `pipeline-tu-dong` skill). Không cần AI-generated builder riêng cho từng mã.

Doanh thu CTCK được dự phóng theo phương pháp **Bottom-up 5 mảng kinh doanh**, không dùng top-down doanh thu-tổng như ngành khác, vì mỗi mảng có driver kinh tế hoàn toàn khác nhau.

---

## 1. Mô hình doanh thu Bottom-up — 5 mảng

| # | Mảng | Công thức | Driver chính |
|---|---|---|---|
| 1 | **Môi giới** | `DT = GTGD/phiên × Số phiên × Thị phần × Phí(bps) / 10,000` | Thanh khoản thị trường (ADTV), thị phần công ty, phí môi giới |
| 2 | **Cho vay Margin** | `DT = Dư nợ bình quân × NIM` | Dư nợ margin (đầu kỳ+cuối kỳ)/2, biên lãi ròng cho vay |
| 3 | **Tự doanh (FVTPL)** | `DT = Danh mục × (%CDs×R_CDs + %TP×R_TP + %CP×R_VNI)` | Cơ cấu danh mục (CDs/Trái phiếu/Cổ phiếu), lợi suất kỳ vọng mỗi loại |
| 4 | **IB + Lưu ký** | `DT = Pipeline IB × Fee% + AUM lưu ký × Fee lưu ký%` | Giá trị deal M&A/tư vấn niêm yết, AUM lưu ký |
| 5 | **Quản lý quỹ (QLQ)** | `DT = AUM quản lý × Fee rate (~0.75%/năm)` | Quy mô tài sản quản lý (AUM) |

### Chi tiết từng mảng

**(1) Môi giới**: Thị phần môi giới lịch sử được suy ngược từ DT thực tế: `Thị phần = DT môi giới / (ADTV × Số phiên × Phí bình quân giả định)`. Dự phóng lấy trung bình 2 năm gần nhất, điều chỉnh theo xu hướng cạnh tranh phí (một số CTCK đang chạy phí 0% để giành thị phần — cần theo dõi rủi ro biên lợi nhuận).

**(2) Cho vay Margin**: `NIM = (Lãi cho vay margin − Chi phí vốn) / Dư nợ cho vay bình quân`. Dùng dư nợ bình quân (đầu kỳ+cuối kỳ)/2, KHÔNG dùng dư nợ cuối kỳ (dư nợ margin biến động mạnh trong quý theo thị trường). Trong `template_securities.py`, NIM lịch sử được tính bằng NIM gộp quan sát được (Lãi cho vay / Dư nợ bình quân) do BCTC công khai không tách riêng chi phí vốn phân bổ cho từng mảng.

**(3) Tự doanh (FVTPL)**: CHỈ mô hình hóa tài sản **FVTPL** (Fair Value Through Profit or Loss — ghi nhận lãi/lỗ trực tiếp vào P&L mỗi kỳ), KHÔNG mô hình tài sản **AFS** (Available For Sale — ghi nhận qua OCI/vốn chủ sở hữu, chỉ vào P&L khi thực sự bán/impair). Đây là lý do vì sao chỉ FVTPL ảnh hưởng trực tiếp đến dự phóng LNST hàng quý. Cơ cấu danh mục giả định: %CDs (chứng chỉ tiền gửi, rủi ro thấp) + %TP (trái phiếu) + %CP (cổ phiếu, dùng biến động VN-Index kỳ vọng làm proxy lợi suất — R_VNI).

**(4) IB + Lưu ký**: Doanh thu IB (tư vấn niêm yết, M&A, phát hành vốn) rất lồi lõm giữa các quý do deal-based, không đều — pipeline hàng năm được ước tính rồi chia đều/tái niên hóa theo quý (`/4` rồi `×4`) để phản ánh đúng tính chất "annualized run-rate" thay vì suy diễn máy móc từ 1 quý đơn lẻ. Lưu ký tính theo AUM lưu ký × phí lưu ký%, tăng trưởng ổn định hơn IB nhiều.

**(5) Quản lý quỹ (QLQ)**: Nhiều CTCK KHÔNG tách riêng dòng doanh thu QLQ trên BCTC hợp nhất công khai (thường gộp vào "Doanh thu khác" hoặc hạch toán qua công ty con AM riêng). `template_securities.py` dùng ước tính xấp xỉ (40% "Doanh thu khác" làm proxy lịch sử) chỉ để có mốc neo cho driver AUM×Fee rate — cần điều chỉnh thủ công trong `02_Assumptions` nếu công ty có công bố AUM QLQ cụ thể trong BCTN/thuyết minh.

---

## 2. Vietcap IQ Insight — Field code CTCK (khác Bank/thường)

CTCK dùng field prefix **`iss`** (Income Statement Securities) và **`bss`** (Balance Sheet Securities), khác với `isa`/`bsa` (công ty thường) và `isb`/`bsb` (ngân hàng). Field map chính (xem `SEG`, `IS_TOTAL`, `BS` dict đầu file `template_securities.py`):

| Field | Ý nghĩa |
|---|---|
| `iss42` | DT môi giới |
| `iss120` | DT lãi cho vay margin |
| `iss115`/`iss124` | Lãi/lỗ tài sản FVTPL |
| `iss44`/`iss46`/`iss123` | DT bảo lãnh phát hành / tư vấn / tư vấn tài chính (IB) |
| `iss47` | DT lưu ký |
| `bss215` | Dư nợ cho vay Margin — dùng để **tự động nhận diện CTCK** trong `run_analysis.py` |
| `isa16` | LNTT (nhãn Vietcap ghi nhầm "Chi phí thuế TNDN" nhưng vị trí/số liệu ĐÚNG là LNTT — đã đối chiếu `isa20 = isa16 + isa17 + isa18`) |

---

## 3. Phân tích hiệu quả & rủi ro tập trung mảng

- Tính % đóng góp doanh thu mỗi mảng trên tổng doanh thu dự phóng năm gần nhất.
- **Cảnh báo tập trung vốn**: nếu 1 mảng chiếm >45% doanh thu → flag "Cảnh báo" (đặc biệt nguy hiểm nếu đó là mảng Tự doanh, vì LNST khi đó phụ thuộc trực tiếp biến động VN-Index, kém bền vững hơn nguồn thu phí dịch vụ).
- **Độ nhạy VN-Index**: mô phỏng tác động ±10%/±20% VN-Index lên phần cổ phiếu trong danh mục FVTPL → % thay đổi LNST tương ứng (sheet `08_Sensitivity`).
- Mảng có DT tăng nhanh nhưng biên LN thấp (VD Margin tăng dư nợ nhanh nhưng NIM co hẹp do cạnh tranh lãi suất) cần cảnh báo hiệu quả sử dụng vốn kém đi.

---

## 4. Định giá — P/B ưu tiên số 1

| # | Phương pháp | Vai trò | Công thức |
|---|---|---|---|
| 1 | **P/B** | CHÍNH (trọng số 90%, CỐ ĐỊNH cho mọi CTCK) | `Giá mục tiêu = BVPS dự phóng × P/B trung vị lịch sử` |
| 2 | **P/E** | BỘ LỌC CHỐNG NHIỄU (trọng số 10%, CỐ ĐỊNH cho mọi CTCK) | `Giá mục tiêu = EPS dự phóng × P/E trung vị lịch sử` — không bỏ hẳn (vẫn cần đối chiếu chéo khi P/B bị lệch bất thường) nhưng không để cao (LNST CTCK dễ biến động do Tự doanh/IB) |
| 3 | **DCF** | Tùy chọn (chưa triển khai trong template) | Có thể bổ sung nếu cần cross-check, không bắt buộc |

**Vì sao P/B ưu tiên hơn P/E cho CTCK**: Tài sản CTCK (tiền, chứng khoán FVTPL, cho vay margin) có tính thanh khoản rất cao và được đánh giá lại theo giá thị trường thường xuyên → giá trị sổ sách (BVPS) phản ánh sát giá trị thực hơn nhiều so với doanh nghiệp sản xuất. Ngược lại LNST/EPS dễ biến động mạnh theo chu kỳ thị trường (đặc biệt khi tỷ trọng Tự doanh cao) nên P/E kém tin cậy hơn khi dùng làm neo định giá chính.

Trọng số 90/10 (P/B/P/E) CỐ ĐỊNH cho MỌI CTCK trong `template_securities.py::VALUATION_WEIGHTS` (đã chốt với user 2026-07 — không còn thích ứng theo tỷ trọng Tự doanh như bản trước đó). Lý do giữ P/E ở 10% thay vì 0%: vẫn cần vai trò bộ lọc chống nhiễu khi chỉ dựa hoàn toàn vào P/B; lý do không để cao hơn: P/E không phải trọng số lớn khi định giá tài sản CTCK.

---

## 5. CAPM / Beta / COE

Dùng chung logic với `tinh-beta` skill và `template_banking.py` (Blume-adjusted beta, cửa sổ tối đa 500 phiên gần nhất để tránh lệch full-history vs sliced-window). CTCK cộng thêm **phần bù rủi ro đặc thù +2%** (`SPECIFIC_RISK_PREMIUM`) so với mức chuẩn vì đòn bẩy margin + tự doanh khiến rủi ro kinh doanh cao hơn trung bình thị trường.

`COE = Rf + β×ERP + Specific Risk Premium (2%)`

---

## 6. Danh sách mã CTCK đã biết (SECURITIES_TICKERS trong run_analysis.py)

SSI, VND, HCM, VCI, FTS, SHS, BSI, VIX, MBS, CTS, AGR, BVS, APG, ORS, TVS, VDS

Ngoài danh sách cứng này, `run_analysis.py` còn **tự động nhận diện CTCK** cho mã bất kỳ có tài khoản `bss215` (Dư nợ cho vay Margin) khác None trong Bảng cân đối kế toán — nên template vẫn hoạt động đúng cho các mã CTCK mới/ít phổ biến hơn không nằm trong danh sách.

---

## 7. Output chuẩn

`template_securities.py::run_securities_analysis(ticker, raw_data)` xuất đủ 3 định dạng cùng 1 lần chạy, dùng chung 1 nguồn số Python (không tính lại riêng ở Excel/PDF/JSON — theo đúng nguyên tắc "1 luồng tính toán duy nhất" đã áp dụng cho HPG/MWG):

- **Excel** (`Bao cao/<TICKER>/<TICKER>_Model_YYYY-MM.xlsx`) — 15 sheet: `00_Beta`, `00_COE`, `01_Cover`, `02_Assumptions`, `03_Revenue_Model`, `04_PnL`, `05_Balance_Sheet`, `06_Cash_Flow`, `07_Valuation`, `08_Sensitivity`, `09_PESTLE`, `10_Leading_Indicators`, `11_Investment_Thesis`, `12_Summary_Snapshot`, `13_PE_PB_History`, `14_Segment_Quarterly`.
- **PDF** (`Bao cao/<TICKER>/<TICKER>_Phan_Tich_YYYY-MM.pdf`) — 8 trang: Cover, Investment Summary, Mô hình 5 mảng, Margin+Sensitivity, P/E-P/B lịch sử, Bảng tài chính, Định giá, Rủi ro & Kết luận.
- **JSON** (`data/<TICKER>.json`) — schema chuẩn dashboard (data/thesis/risks/moats/pestle/valuation/comments/ratios/pe_quarters/pb_quarters/quarter_labels) + khối `segments` riêng cho 5 mảng kinh doanh (revenueHist/revenueForecast/pctNow theo từng mảng).

**Xác minh bắt buộc trước khi coi là hoàn tất**: mở Excel bằng COM automation (`CalculateFullRebuild()`), kiểm tra 0 formula error và `05_Balance_Sheet`: Tổng Tài sản = Tổng Nguồn Vốn ở mọi năm dự phóng (sai số làm tròn chấp nhận được, không quá vài phần vạn tổng tài sản).
