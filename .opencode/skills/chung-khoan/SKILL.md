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

---

## 8. Cơ cấu Danh mục Tự doanh FVTPL/AFS (bổ sung 2026-07)

Vietcap **không có** dữ liệu cơ cấu danh mục tự doanh theo loại tài sản (chỉ có tổng lãi/lỗ FVTPL/AFS,
đã verify field `nos681`/`nos683` qua metrics NOTE — không có breakdown cổ phiếu/trái phiếu/chứng chỉ
quỹ/CD, không có danh mục cổ phiếu cụ thể đang nắm). Toàn bộ dữ liệu này CHỈ có trong thuyết minh BCTC.

**Vị trí thuyết minh**: luôn nằm ngay sau thuyết minh "Tiền và các khoản tương đương tiền" (mục 3.1),
thường là mục **3.2** (Tài sản tài chính FVTPL) và 3.3/3.4 (AFS) trong "THÔNG TIN BỔ SUNG BÁO CÁO TÌNH
HÌNH TÀI CHÍNH". PDF BCTC gần như luôn là scan không có text layer — **dùng `bctc_pdf_tool.py render`
để chuyển trang thành PNG và đọc trực tiếp bằng ảnh**, KHÔNG dựa vào MD convert (MD hay cắt mất 1/4 đầu
+ 1/4 cuối đúng ngay khu vực note này).

**Kho dữ liệu**: `data/fvtpl_holdings/<TICKER>.json` (xem `data/fvtpl_holdings/_schema_note.md` cho schema
đầy đủ). 5 nhóm tài sản CỐ ĐỊNH (không chi tiết sub-item): `CoPhieuNiemYet` (gộp cả tài sản cơ sở phòng
ngừa rủi ro chứng quyền nếu có), `CoPhieuChuaNiemYet`, `TraiPhieu`, `ChungChiQuy` (gộp ETF+Quỹ mở),
`ChungChiTienGui`. Mỗi kỳ (`yearly`/`quarterly`) lưu `cost` (giá gốc), `fairValue` (giá trị hợp lý),
`gain`/`loss` (chênh lệch đánh giá lại — đây là số dư lũy kế tại thời điểm báo cáo so với giá gốc, KHÔNG
phải lãi/lỗ phát sinh riêng trong kỳ). Khối `holdings` lưu danh mục cổ phiếu niêm yết cụ thể đang nắm
(mã: giá trị hợp lý VND, "Khac" = phần dư gộp mã nhỏ). Đã seed thật cho **HCM** đầy đủ **4 năm (2022-2025)
+ 8 quý (2024 Q1-Q4, 2025 Q1-Q4) + Q1/2026** — tải 8 PDF BCTC hợp nhất (FY2025, FY2023, Q1-Q3/2024,
Q1/2025, Q3/2025, Q1/2026) qua `bctc_pdf_tool.py`, tận dụng cột đối chiếu cùng kỳ năm trước trong mỗi
BCTC để giảm số PDF cần tải (vd BCTC FY2025 cho luôn cả 31/12/2024; BCTC Q1/2026 cho luôn cả 31/12/2025).
Mọi tổng đã đối chiếu khớp 100% với số in trên BCTC trước khi ghi vào kho.

**Dashboard**: `app_securities.js::loadFvtplHoldings(ticker)` tự fetch `data/fvtpl_holdings/<TICKER>.json`
riêng biệt (KHÔNG nhúng vào `data/<TICKER>.json`, giống pattern `peer_benchmark_securities.json`) — nếu
file không tồn tại cho ticker, cả section `#fvtpl-holdings-section` trong `securities.html` tự ẩn. 3
biểu đồ: cơ cấu danh mục theo 5 nhóm (stacked bar, giá trị hợp lý, gộp chung mốc năm+quý theo trục thời
gian — năm nào đã có Q4 trùng thời điểm 31/12 thì tự bỏ điểm yearly trùng, xem `_fvtplMergedPeriods`),
lãi/lỗ đánh giá lại theo nhóm (diverging stacked bar), danh mục cổ phiếu cụ thể đang nắm giữ (stacked bar
theo mã). Mở rộng sang CTCK khác: tạo `data/fvtpl_holdings/<TICKER>.json` theo đúng schema là dashboard
tự nhận, không cần sửa code JS.

**Excel + PDF**: `template_securities.py::build_fvtpl_holdings_sheet()` tạo sheet `16_FVTPL_Composition`
(3 bảng: cơ cấu theo nhóm, lãi/lỗ ròng theo nhóm, danh mục cổ phiếu cụ thể) — chỉ tạo khi
`load_fvtpl_holdings(ticker)` trả về dữ liệu, bỏ qua nếu ticker chưa có. PDF thêm trang "Trang 5c" (2
chart `make_fvtpl_composition_chart`/`make_fvtpl_holdings_chart`, cùng logic gộp kỳ năm/quý với web) chỉ
chèn khi có ảnh — không ảnh hưởng CTCK chưa có dữ liệu này.

---

## 9. Driver dự phóng doanh thu Tự doanh — ưu tiên cơ cấu THẬT + lợi suất CAPM nhất quán (bổ sung 2026-07)

Công thức dự phóng doanh thu Tự doanh (mục 1.3) không đổi: `DT = Danh mục × (%CDs×R_CDs + %TP×R_TP +
%CP×R_VNI)`. Điều đã cải thiện là NGUỒN của tỷ trọng (%CDs/%TP/%CP) và lợi suất kỳ vọng R_VNI:

- **`template_securities.py::derive_real_fvtpl_mix(ticker)`** — nếu ticker có `data/fvtpl_holdings/
  <TICKER>.json` (xem §8), hàm lấy kỳ GẦN NHẤT và tính tỷ trọng THẬT từ giá trị hợp lý (ánh xạ 5 nhóm
  thuyết minh → 3 nhóm gốc: `CP = CoPhieuNiemYet + CoPhieuChuaNiemYet + ChungChiQuy`, `TP = TraiPhieu`,
  `CDs = ChungChiTienGui` — Chứng chỉ quỹ gộp vào CP vì tại HCM toàn bộ là ETF/quỹ mở theo dõi chỉ số cổ
  phiếu; nếu CTCK khác có quỹ trái phiếu cần tách riêng khi thêm dữ liệu ticker đó).
- Nếu ticker CHƯA có dữ liệu này → tự fallback về bảng giả định `_FVTPL_MIX_OVERRIDE` cũ (không regression
  cho các mã chưa thu thập dữ liệu FVTPL).
- **R_VNI (lợi suất kỳ vọng phần CP+Chứng chỉ quỹ)**: khi dùng cơ cấu thật, gán `R_VNI = COE` (đúng CAPM
  Rf+β×ERP+α đã tính cho chính ticker đó ở mục 5) thay vì hằng số 10% chủ quan trước đây — nhất quán với
  suất chiết khấu dùng trong toàn bộ mô hình định giá. R_CDs/R_TP giữ nguyên ~7% (lãi suất CDs/coupon trái
  phiếu quan sát thị trường — xem ghi chú đối chiếu trong `build_assumptions_sheet`); KHÔNG dùng "chênh
  lệch đánh giá lại" trong thuyết minh FVTPL làm lợi suất trái phiếu vì con số đó chỉ là lãi/lỗ giá thị
  trường (mark-to-market), không gồm lãi coupon — dùng sẽ làm lợi suất bị hiểu sai/thấp hơn thực tế.
- Sheet `02_Assumptions` tự đổi ghi chú theo nguồn (✓ CƠ CẤU THẬT vs ⚠ GIẢ ĐỊNH THỦ CÔNG) tùy ticker có
  dữ liệu hay không — xem log console `[FVTPL Mix]`/`[FVTPL Rate]` khi chạy `run_analysis.py` để biết
  đang dùng nguồn nào.
- **Đã verify thực tế cho HCM (2026-07)**: cơ cấu giả định cũ (CDs 35%/TP 30%/CP 35%) sai lệch lớn so với
  thực tế tại Q1/2026 (CDs 21%/TP 72%/CP 7% — danh mục đã chuyển dịch mạnh sang trái phiếu) — xác nhận
  đúng lý do user yêu cầu cải thiện ("model định giá chưa rõ ràng"). Chạy `python run_analysis.py HCM`
  thành công end-to-end (Excel/PDF/JSON), `data/HCM.json::macro_liquidity.fvtpl_rates.R_VNI` khớp
  `valuation.coe` (14.56%).
- **Phạm vi đã chốt với user**: CHỈ cải thiện driver dự phóng doanh thu Tự doanh — KHÔNG đổi cách định giá
  tổng thể công ty (vẫn P/B 90% + P/E 10% cố định ở mục 4, không chuyển sang Sum-of-Parts).
