---
name: fa
description: Use ONLY when analyzing Vietnamese stocks or building financial models. Front-loaded keywords: phân tích cổ phiếu, định giá, financial model, stock analysis, ticker VN (HPG, MWG, PNJ, TCB, SSI, KSV, NLG, FPT, VNM, DHC...), PESTLE, Excel model, định giá doanh nghiệp, dự báo tài chính.
---

# SKILLALL — FRAMEWORK PHÂN TÍCH CỔ PHIẾU VIỆT NAM TOÀN DIỆN

> **Triết lý cốt lõi**: Phân tích cổ phiếu là tìm kiếm sự **chênh lệch giữa giá trị nội tại và giá thị trường**, không phải đoán hướng thị trường. Mỗi cổ phiếu là một câu chuyện riêng, nhưng công cụ phân tích là chung.

Phiên bản: 2.3 | Cập nhật: 06/2026

---

## MỤC LỤC

1. [TỔNG QUAN QUY TRÌNH — 4 BƯỚC LỚN](#1-tổng-quan-quy-trình--4-bước-lớn)
2. [BƯỚC 1: THU THẬP DỮ LIỆU REAL-TIME](#2-bước-1-thu-thập-dữ-liệu-real-time)
   - 2.1 Nguyên tắc cốt lõi
   - 2.2 4 nhóm dữ liệu
   - 2.3 Nguồn dữ liệu
   - 2.4 Output bắt buộc
   - 2.5 Quy tắc bắt buộc phân kỳ Actual (A) vs Estimated (E)
3. [BƯỚC 2: PHÂN TÍCH 6 TẦNG](#3-bước-2-phân-tích-6-tầng)
   - 3.1 Tầng 1 — Chuỗi giá trị & Mô hình kinh doanh
   - 3.2 Tầng 2 — Thị trường & Vị trí cạnh tranh (Porter 5F + PESTLE)
   - 3.3 Tầng 3 — Lợi thế cạnh tranh & Quản trị
   - 3.4 Tầng 4 — Phân tích tài chính & Dự báo
   - 3.5 Tầng 5 — Rủi ro & Catalyst
   - 3.6 Tầng 6 — Định giá
4. [BƯỚC 3: BUILD EXCEL MODEL](#4-bước-3-build-excel-model)
   - 4.1 Cấu trúc 12+1 sheets
   - 4.2 Quy trình build từng sheet
   - 4.3 Sheet 13 — Chấm điểm Leading Indicators
   - 4.4 Quy trình sau khi hoàn thành Excel
5. [BƯỚC 4: VIẾT DOCX RESEARCH REPORT](#5-bước-4-viết-docx-research-report)
   - 5.1 Điều kiện tiên quyết
   - 5.2 Thứ tự build bắt buộc
   - 5.3 Assembly guide từng trang
   - 5.4 Investment Summary template
   - 5.5 Yêu cầu chất lượng
6. [PHƯƠNG PHÁP DỰ BÁO DOANH THU THEO NGÀNH](#6-phương-pháp-dự-báo-doanh-thu-theo-ngành)
   - 6.1–6.14: 14 ngành
7. [PHƯƠNG PHÁP ĐỊNH GIÁ THEO NGÀNH](#7-phương-pháp-định-giá-theo-ngành)
8. [QUY CHUẨN EXCEL MODEL](#8-quy-chuẩn-excel-model)
9. [QUY CHUẨN DOCX REPORT](#9-quy-chuẩn-docx-report)
10. [DANH SÁCH CỔ PHIẾU ĐÃ PHÂN TÍCH](#10-danh-sách-cổ-phiếu-đã-phân-tích)
11. [CHECKLIST KIỂM TRA TRƯỚC KHI GIAO](#11-checklist-kiểm-tra-trước-khi-giao)

---

## 1. TỔNG QUAN QUY TRÌNH — 4 BƯỚC LỚN

```
NHẬN TICKER
    │
    ▼
┌─────────────────────────────────────────────────────────────┐
│ B1: THU THẬP DỮ LIỆU REAL-TIME  (song song 4 nhóm)         │
│ → Giá CP, BCTC 4-5 năm, định tính, ngành & peer            │
│ → Xác định ngày phân tích → tính đúng nhãn A/E             │
└─────────────────────────────────────────────────────────────┘
    │
    ▼
┌─────────────────────────────────────────────────────────────┐
│ B2: PHÂN TÍCH 6 TẦNG  (mỗi tầng → 1 artifact output)       │
│                                                             │
│ Tầng 1: Chuỗi giá trị → Bảng chuỗi giá trị (≥5 bước)       │
│ Tầng 2: Thị trường + PESTLE → Porter 5F + TAM + Chu kỳ     │
│ Tầng 3: Moat & Quản trị → Bảng moat 5 loại + ROIC vs WACC  │
│ Tầng 4: Tài chính & Dự báo → Driver table + P&L bridge     │
│ Tầng 5: Rủi ro & Catalyst → Risk matrix + Catalyst table   │
│ Tầng 6: Định giá → Bear/Base/Bull + Sensitivity 5×5        │
└─────────────────────────────────────────────────────────────┘
    │
    ▼
┌─────────────────────────────────────────────────────────────┐
│ B3: BUILD EXCEL MODEL  (13 sheets, tiếng Việt có dấu)       │
│ → 01_Cover đến 13_Theo_Doi                                  │
│ → Tất cả giả định trong sheet Giả định                      │
│ → Recalc = 0 lỗi                                            │
└─────────────────────────────────────────────────────────────┘
    │
    ▼
┌─────────────────────────────────────────────────────────────┐
│ B4: VIẾT DOCX REPORT  (Assembly — từ artifacts có sẵn)      │
│ → Pre-flight checklist trước khi mở file                    │
│ → Build theo thứ tự: Chuỗi GT → Thị trường → Moat →         │
│   Tài chính → Định giá → Rủi ro → Indicators → Phụ lục     │
│ → Investment Summary VIẾT CUỐI                              │
└─────────────────────────────────────────────────────────────┘
    │
    ▼
GIAO: [TICKER]_Model_[YYYY-MM].xlsx + [TICKER]_Phan_Tich_[YYYY-MM].docx
```

**Output**: 1 file Excel `[TICKER]_Model_[YYYY-MM].xlsx` + 1 file DOCX `[TICKER]_Phan_Tich_[YYYY-MM].docx`

---

## 2. BƯỚC 1: THU THẬP DỮ LIỆU REAL-TIME

### 2.1 Nguyên tắc cốt lõi

```
1. LUÔN lấy dữ liệu ngày hiện tại — không dùng số cũ từ lần phân tích trước
2. KHÔNG suy luận — nếu chưa có số thì ghi "[chưa xác nhận]", không ước đoán
3. BCTC lịch sử: lấy BCTC gốc từ IR công ty hoặc HSX, không copy từ nguồn tổng hợp chưa kiểm chứng
4. Luôn ghi nguồn + ngày cho mọi dữ liệu
```

### 2.2 4 nhóm dữ liệu (thu thập song song)

```
NHÓM A — THỊ TRƯỜNG (cập nhật real-time):
├─ Giá cổ phiếu hiện tại
├─ Vốn hóa thị trường
├─ Số CP lưu hành (không phải vốn điều lệ / 10.000)
├─ Khối lượng giao dịch bình quân 20 phiên
├─ Tỷ lệ sở hữu nước ngoài & room ngoại còn lại
├─ 52-week high / low

NHÓM B — TÀI CHÍNH (4-5 năm lịch sử):
├─ KQKD: Doanh thu, LN gộp, LNST, EPS — từng năm
├─ Bảng CĐKT: TS, Nợ, VCSH — cuối mỗi năm
├─ Báo cáo LCTT: CFO, CFI, CFF
├─ Các chỉ số: ROE, ROA, ROIC, D/E, Net Debt/EBITDA, DSO, DPO
├─ Tỷ lệ cover: ICR, DSCR

NHÓM C — ĐỊNH TÍNH:
├─ Mô hình kinh doanh: doanh thu từ đâu, khách hàng ai, đối thủ nào
├─ Kế hoạch ĐHĐCĐ năm hiện tại
├─ Đánh giá chất lượng quản lý (capital allocation, track record)
├─ Báo cáo phân tích từ CTCK (tối thiểu 2 nguồn)
├─ Tin tức gần đây, sự kiện sắp tới

NHÓM D — NGÀNH & PEER:
├─ 3-5 công ty cùng ngành — vốn hóa, P/E, EV/EBITDA, P/B
├─ Quy mô thị trường (TAM) và tăng trưởng
├─ Giá hàng hóa / chỉ số ngành (nếu commodity)
├─ Chỉ số vĩ mô ảnh hưởng: GDP, CPI, PMI, lãi suất, tỷ giá
```

### 2.3 Nguồn dữ liệu

```
BCTC gốc (ưu tiên 1):   ir.[company].com.vn / hsx.vn công bố thông tin
BCTC tổng hợp:           finance.vietstock.vn/[TICKER]/financials.htm
Giá CP + KLGD:           simplize.vn/co-phieu/[TICKER]
Tin tức + ĐHĐCĐ:         cafef.vn/[ticker].chn, dnse.com.vn, tinnhanhchungkhoan.vn
Kế hoạch công ty:        Search "[TICKER] ĐHĐCĐ [năm] kế hoạch tài liệu"

**API dữ liệu lịch sử P/E, P/B, EV/EBITDA (Vietcap)**:
  Base URL: `https://trading.vietcap.com.vn/api/iq-insight-service/v1`
  - **Endpoint ratios lịch sử**: `GET /company/{ticker}/statistics-financial`
    → P/E, P/B, EV/EBITDA, P/S, ROE, ROA, ROIC, Market Cap theo từng quý (TTM) và năm
    → Cần handshake: GET `https://trading.vietcap.com.vn/iq/company?ticker={ticker}` trước
    → Header: User-Agent chuẩn browser
  - **Endpoint BCTC**: `GET /company/{ticker}/financial-statement?section={INCOME_STATEMENT|BALANCE_SHEET|CASH_FLOW}`
  - **Endpoint field mapping**: `GET /company/{ticker}/financial-statement/metrics`
  - **Endpoint company details**: `GET /company/details?ticker={ticker}`

Báo cáo CTCK (≥ 2 nguồn bắt buộc):
  SSI:    research.ssi.com.vn / "[TICKER] SSI Research [năm]"
  VCSC:   vcsc.com.vn/research / "[TICKER] VCSC báo cáo [năm]"
  VCBS:   vcbs.com.vn / "[TICKER] VCBS [năm]"
  HSC:    hsc.com.vn/research / "[TICKER] HSC Research [năm]"
  MBS:    mbs.com.vn/research / "[TICKER] MBS Research [năm]"
  VDSC:   vdsc.com.vn / "[TICKER] VDSC Rồng Việt [năm]"
  BSC:    "[TICKER] BSC BIDV Securities [năm]"
  KIS:    "[TICKER] KIS Vietnam [năm]"
  Yuanta: "[TICKER] Yuanta báo cáo [năm]"

Giá hàng hóa:    LME, tradingeconomics.com, metal.com
SCFI/cước tàu:   "[SCFI index tháng năm]", drewry.co.uk
FDI/XNK VN:      gso.gov.vn, mpi.gov.vn
THC/Hàng hải:    vinamarine.gov.vn
```

### 2.4 Output bắt buộc

Sau Bước 1, phải có:

```
☐ [ ] Bảng giá CP + thông số thị trường (ngày hiện tại)
☐ [ ] Bảng BCTC tóm tắt ≥4 năm (DT, LNST, EPS, BVPS, ROE, ROIC, Net Debt/EBITDA)
☐ [ ] Bảng cơ cấu doanh thu + % mảng
☐ [ ] Bảng peer 3-5 công ty (vốn hóa, P/E, EV/EBITDA, ROE)
☐ [ ] Ghi rõ: ngày phân tích, năm A gần nhất, năm E đầu tiên
☐ [ ] **Fetch lịch sử P/E, P/B, EV/EBITDA từ Vietcap API** → tính median, vẽ biểu đồ
```

> **Nếu Bước 1 thiếu dữ liệu, các bước sau sẽ sai. Dành đủ thời gian cho bước này.**

### 2.5 Quy tắc bắt buộc phân kỳ Actual (A) vs Estimated (E)

> **Nguyên tắc vàng**: Năm đã kết thúc VÀ BCTC đã được công bố → gán A. Tất cả các năm còn lại → gán E.

#### Định nghĩa

| Nhãn | Ý nghĩa | Điều kiện |
|---|---|---|
| **A** | Actual — số thực từ BCTC đã kiểm toán/soát xét | Năm tài chính đã kết thúc + BCTC đã công bố |
| **A(SXX)** | Actual semi — số thực từ BCTC soát xét bán niên | BCTC bán niên soát xét đã công bố |
| **A(QX)** | Actual quarterly — số thực từ BCTC quý | BCTC quý đã công bố |
| **E** | Estimated — số dự báo | Năm hiện tại chưa kết thúc HOẶC đã kết thúc nhưng BCTC chưa công bố |
| **E(RE)** | Estimated revised — số dự báo đã điều chỉnh | Khi có thông tin mới làm thay đổi dự báo cũ |

#### Cách xác định

```
1. Xác định ngày phân tích hôm nay: DD/MM/YYYY
2. Nếu ngày phân tích > 31/03/N và BCTC năm N-1 đã công bố:
   → Năm N-1 = A, Năm N và các năm sau = E
3. Ví dụ: Phân tích ngày 18/06/2026
   → 2024A, 2025A (BCTC đã công bố), 2026E, 2027E, 2028E
4. Nếu phân tích tháng 2, BCTC năm N-1 chưa công bố:
   → Năm N-2 = A, Năm N-1 và các năm sau = E
   → Ghi chú: "[Năm N-1]E — chờ BCTC quý 4 hoặc năm N-1"
```

#### Bảng tra nhanh (phân tích năm 2026)

| Thời điểm phân tích | Năm A gần nhất | Năm E đầu tiên | Ghi chú |
|---|---|---|---|
| Tháng 1-3/2026 | 2024A | 2025E + 2026E | BCTC 2025 chưa công bố |
| Tháng 4/2026 (nếu BCTC 2025 chưa CB) | 2024A | 2025E + 2026E | Chờ CB, ghi chú rõ |
| Tháng 4-12/2026 (BCTC 2025 đã CB) | 2025A | 2026E + 2027E | Trường hợp phổ biến |
| Tháng 2/2027 | 2025A | 2026E + 2027E | BCTC 2026 chưa CB |

#### Quy tắc trong Excel & DOCX

```
1. Hàng năm: [N-3]A | [N-2]A | [N-1]A | [N]E | [N+1]E | [N+2]E
2. Nếu có BCTC quý gần nhất: thêm cột phụ hoặc ghi chú "[N]A(Q1/Q2/Q3)"
3. Mọi ô gán A PHẢI chứa số thực từ BCTC — không ước tính
4. Năm E không được có số giống y hệt năm A — phải có thay đổi (tăng/giảm)
```

#### Checklist kiểm tra A/E

```
☐ Đã xác định ngày phân tích → tính đúng năm A gần nhất?
☐ Năm A: số liệu khớp với BCTC gốc (không copy từ nguồn tổng hợp)?
☐ Năm E: có dự báo chủ động (không copy năm A)?
☐ DOCX và Excel dùng cùng nhãn A/E?
☐ Năm có one-off: đã tách riêng và chú thích?
☐ Nếu BCTC chưa tìm được: ghi "[chưa xác nhận]", không ước tính?
```

---

## 3. BƯỚC 2: PHÂN TÍCH 6 TẦNG

> **Nguyên tắc**: Mỗi tầng là một artifact độc lập. Artifact là bảng/có dạng bảng, KHÔNG phải văn xuôi dài dòng. Artifact của tầng này là INPUT cho tầng sau và cho DOCX.

```
Sơ đồ luồng:
Tầng 1 → Tầng 2 → Tầng 3 → Tầng 4 → Tầng 5 → Tầng 6
  │         │         │         │         │         │
  ▼         ▼         ▼         ▼         ▼         ▼
DOCX      DOCX      DOCX      DOCX      DOCX      DOCX
Tr 3-4    Tr 5-6    Tr 7-8    Tr 9-11   Tr 14-15  Tr 12-13
```

---

### 3.1 Tầng 1 — Chuỗi giá trị & Mô hình kinh doanh

**Mục tiêu**: Xác định vị trí của công ty trong chuỗi giá trị ngành, ai capture value nhiều nhất, tại sao.

#### Các bước thực hiện

```
Bước 1: Vẽ chuỗi giá trị từ nguyên liệu → sản xuất → phân phối → bán lẻ → người dùng cuối
Bước 2: Với mỗi bước, ước tính biên lợi nhuận đặc thù (VD: thượng nguồn 5-10%, trung nguồn 15-25%, hạ nguồn 3-8%)
Bước 3: Xác định vị trí công ty trong chuỗi — capture value ở đâu?
Bước 4: Nếu công ty có nhiều mảng → vẽ riêng từng mảng
```

#### Chuỗi giá trị mẫu — Thép (HPG)

```
Nguyên liệu (quặng/than) → Luyện cốc → SX phôi thép → Cán thép → Phân phối → Người dùng cuối
          Biên: 0-5%         6-12%     15-25%      8-15%      2-4%
                                     ▲
                                     HPG ở đây (tích hợp dọc từ luyện cốc đến cán thép)
```

#### Bảng chuỗi giá trị bắt buộc

| Bước trong chuỗi | Mô tả | Biên ước tính | Ai chi phối | Vị trí công ty |
|---|---|---|---|---|
| [Bước 1] | [mô tả] | [X-Y%] | [ai nắm quyền] | [Thượng/Trung/Hạ nguồn] |
| ... | ... | ... | ... | ... |
| Tổng hợp | | | | |

#### Output bắt buộc (artifact Tầng 1)

```
☐ Bảng chuỗi giá trị ≥5 bước, mỗi bước có biên ước tính
☐ Xác định rõ: công ty đang capture value ở bước nào, biên bao nhiêu
☐ Nếu đa mảng: bảng riêng cho từng mảng
→ INSERT vào DOCX Trang 3-4
```

**Ví dụ artifact Tầng 1 — GMD:**

```
Bảng chuỗi giá trị Cảng biển & Logistics:
| Bước | Mô tả | Biên ước tính | Ai chi phối | Vị trí GMD |
|---|---|---|---|---|
| 1 | Hàng hóa từ nhà máy → ICD | 1-3% | Các ICD tư nhân | — |
| 2 | Vận chuyển ICD → Cảng | 3-5% | Hãng xe tải/container | — |
| 3 | Xếp dỡ container tại cảng | 35-50% | Cảng biển (GMD, HAH, VSC) | ✅ Trung nguồn |
| 4 | Lưu bãi, kho bãi | 20-30% | Cảng biển | ✅ Trung nguồn |
| 5 | Vận tải biển quốc tế | 5-15% | MSC, Maersk, CMA CGM | — |
| 6 | Bốc dỡ + phân phối đầu cuối | 2-5% | Forwarder, 3PL | — |

Vị trí GMD: Chiếm đoạn có biên cao nhất (xếp dỡ + lưu bãi) nhờ vị trí cảng độc quyền khu vực.
```

---

### 3.2 Tầng 2 — Thị trường & Vị trí cạnh tranh (Porter 5F + PESTLE)

**Mục tiêu**: Đánh giá mức độ hấp dẫn của ngành, vị trí cạnh tranh của công ty, bối cảnh vĩ mô.

#### Bước 2A — Porter 5 Forces

Phân tích 5 lực và kết luận mức độ hấp dẫn.

| Lực | Mức độ (Cao/TB/Thấp) | Giải thích | Chiều tác động (+/−) |
|---|---|---|---|
| 1. Đối thủ hiện tại | [Cao/TB/Thấp] | [giải thích] | [+] nếu hấp dẫn |
| 2. Đối thủ tiềm năng | [Cao/TB/Thấp] | [giải thích] | [−] nếu đe dọa |
| 3. Hàng thay thế | [Cao/TB/Thấp] | [giải thích] | [−] nếu có |
| 4. Nhà cung cấp | [Cao/TB/Thấp] | [giải thích] | [−] nếu mạnh |
| 5. Người mua | [Cao/TB/Thấp] | [giải thích] | [−] nếu mạnh |

**Kết luận tổng hợp**: Ngành [tên] mức hấp dẫn [Cao/TB/Thấp] — [lý do 1 câu].

#### Bước 2B — PESTLE

Phân tích 6 yếu tố vĩ mô tác động đến ngành/công ty.

| Yếu tố | Tác động đến ngành | Tác động đến công ty | Chiều (+/−) |
|---|---|---|---|
| **P** — Chính trị | | | |
| **E** — Kinh tế | | | |
| **S** — Xã hội | | | |
| **T** — Công nghệ | | | |
| **L** — Pháp lý | | | |
| **E** — Môi trường | | | |

PESTLE dùng làm INPUT cho Tầng 5 (Risk) và Tầng 6 (điều chỉnh discount rate).

#### Bước 2C — Chu kỳ ngành

```
Xác định vị trí hiện tại của ngành trong chu kỳ:

┌──────────┐     ┌──────────┐     ┌──────────┐     ┌──────────┐
│ Suy thoái│────▶│ Phục hồi │────▶│ Tăng trưởng│───▶│ Bão hòa  │
│ (trough) │     │ (recovery)│     │(growth)  │     │(maturity)│
└──────────┘     └──────────┘     └──────────┘     └──────────┘
      ▲                                                      │
      └──────────────────────────────────────────────────────┘
```

#### Bước 2D — TAM (Total Addressable Market)

| Chỉ tiêu | Giá trị | Nguồn |
|---|---|---|
| TAM (tỷ VNĐ) | | |
| CAGR 3-5 năm | | |
| Thị phần công ty hiện tại | | |
| Thị phần tiềm năng (3 năm) | | |
| Tăng trưởng ngành YoY | | |

#### Bảng xác định chu kỳ

| Tiêu chí | Trạng thái | Dấu hiệu nhận biết |
|---|---|---|
| Tăng trưởng doanh thu ngành | Cao/TB/Thấp/Âm | YoY% |
| Số lượng đối thủ | Tăng/Ổn định/Giảm | Số DN mới gia nhập |
| Công suất ngành | Thiếu hụt/Cân bằng/Dư thừa | Utilization rate |
| Giá bán bình quân | Tăng/Ổn định/Giảm | ASP trend |
| Đầu tư CAPEX | Tăng/Ổn định/Giảm | Tổng CAPEX ngành |

#### Output bắt buộc (artifact Tầng 2)

```
☐ Bảng Porter 5F đầy đủ 5 lực + kết luận mức hấp dẫn ngành
☐ Bảng PESTLE 6 yếu tố
☐ Vị trí chu kỳ ngành + giải thích 2-3 câu
☐ Bảng TAM/SAM/SOM
→ INSERT vào DOCX Trang 5-6
```

---

### 3.3 Tầng 3 — Lợi thế cạnh tranh & Quản trị

**Mục tiêu**: Công ty có moat không? ROIC > WACC? Quản lý có tốt không?

#### Phân tích Moat (5 loại)

| Loại moat | Mức độ (Mạnh/TB/Yếu/Không) | Bằng chứng định lượng |
|---|---|---|
| 1. Lợi thế chi phí (Cost) | | ROA, biên GP vs peer |
| 2. Chuyển đổi (Switching) | | Tỷ lệ giữ chân KH |
| 3. Tài sản vô hình (Intangible) | | Bằng sáng chế, giấy phép |
| 4. Hiệu ứng mạng (Network) | | Số user, platform effects |
| 5. Quy mô hiệu quả (Efficient Scale) | | Thị phần, rào cản gia nhập |

**Overall Moat Rating**: **Wide / Narrow / No Moat** — [lý do 1-2 câu].

#### Test ROIC vs WACC

```
ROIC = NOPAT / Invested Capital
     = EBIT × (1 - Thuế suất) / (Tổng TS - Tiền mặt - Nợ ngắn hạn không lãi)

WACC = E/(D+E) × Ke + D/(D+E) × Kd × (1-t)
```

| Năm | [N-3]A | [N-2]A | [N-1]A | [N]A/E |
|---|---|---|---|---|
| ROIC | | | | |
| WACC | | | | |
| Chênh lệch | | | | |

**Kết luận**: ROIC > WACC → ✅ Tạo giá trị. ROIC < WACC → ❌ Hủy giá trị.

- ROIC > WACC ít nhất 3 năm gần nhất → Wide moat
- ROIC > WACC 1-2 năm → Narrow moat
- ROIC ≤ WACC → No moat (hoặc turnaround)

#### Bảng so sánh ROIC với peers

| Công ty | Ticker | ROIC (T) | ROIC (T-1) | Biên GP | Vốn hóa |
|---|---|---|---|---|---|
| [DN phân tích] | | | | | |
| Peer 1 | | | | | |
| Peer 2 | | | | | |
| Peer 3 | | | | | |

#### Đánh giá chất lượng quản lý (3 tiêu chí)

| Tiêu chí | Đánh giá | Bằng chứng |
|---|---|---|
| **Capital allocation** | Tốt/TB/Kém | ROIC, track record M&A, chia cổ tức, mua lại CP |
| **Communication** | Minh bạch/Mờ/Tệ | Chất lượng IR, ĐHĐCĐ trả lời câu hỏi |
| **Alignment** | Cao/TB/Thấp | Tỷ lệ sở hữu board, stock-based comp |

**Tổng hợp quản lý**: [Tốt/Khá/Trung bình/Kém] — [lý do ngắn gọn].

#### Output bắt buộc (artifact Tầng 3)

```
☐ Bảng moat 5 loại — có bằng chứng định lượng cho mỗi loại
☐ Overall moat rating: Wide / Narrow / No Moat
☐ Bảng ROIC vs WACC (ít nhất 4 năm)
☐ Bảng ROIC so sánh peers (ít nhất 3 peers)
☐ Đánh giá quản lý 3 tiêu chí
→ INSERT vào DOCX Trang 7-8
```

---

### 3.4 Tầng 4 — Phân tích tài chính & Dự báo

**Mục tiêu**: Driver-based forecasting — không dùng % cố định.

#### Revenue Driver Table

| Mảng | Driver 1 | Driver 2 | Công thức | Nguồn dữ liệu |
|---|---|---|---|---|
| [Mảng A] | | | | |
| [Mảng B] | | | | |

Mỗi mảng chọn phương pháp dự báo phù hợp từ **Mục 6 — Phương pháp dự báo doanh thu theo ngành**.

#### Bảng P&L bắt buộc

| Chỉ tiêu | [N-3]A | [N-2]A | [N-1]A | [N]E | [N+1]E | [N+2]E |
|---|---|---|---|---|---|---|
| Doanh thu thuần | | | | | | |
| - Mảng A | | | | | | |
| - Mảng B | | | | | | |
| GVHB | ( ) | ( ) | ( ) | ( ) | ( ) | ( ) |
| LN gộp | | | | | | |
| Biên GP % | | | | | | |
| CP BH & QLDN | | | | | | |
| LN từ HĐKD | | | | | | |
| Doanh thu tài chính | | | | | | |
| Chi phí tài chính | | | | | | |
| LN khác | | | | | | |
| LNTT | | | | | | |
| Thuế | | | | | | |
| LNST CĐCTM | | | | | | |
| EPS (VNĐ) | | | | | | |

> **Ràng buộc A/E**: Tất cả nhãn A/E phải theo Mục 2.5. Năm có one-off phải có chú thích rõ: "* Bao gồm [X] tỷ one-off từ [sự kiện]".

#### Bridge Analysis

Phân tích từ năm [N-1]A sang [N]E — giải thích từng dòng thay đổi lớn:

```
[N-1]A → [N]E Bridge:
DT: +X% nhờ [driver cụ thể]
Biên GP: [tăng/giảm] Y% điểm nhờ [lý do]
CPBH: [tăng/giảm] Z% vì [lý do]
LNST: +X tỷ → +X% YoY
```

#### Phân tích chất lượng lợi nhuận

```
1. One-off items: liệt kê, tách riêng, tính EPS core
2. Cash conversion: CFO / LNST — nếu < 0.8 liên tục → cảnh báo
3. Accruals ratio: (CFO - LNST) / Tổng tài sản
4. CAPEX maintenance vs growth: bao nhiêu % CAPEX để duy trì hiện tại?
```

#### Output bắt buộc (artifact Tầng 4)

```
☐ Revenue driver table — mỗi mảng 1 phương pháp cụ thể
☐ Bảng P&L ≥4A + 3E — có tách one-off nếu có
☐ Bridge analysis từ năm A cuối sang năm E đầu tiên
☐ Phân tích chất lượng LN: one-off, cash conversion, accruals
→ INSERT vào DOCX Trang 9-11
```

---

### 3.5 Tầng 5 — Rủi ro & Catalyst

**Mục tiêu**: Xác định rủi ro có thể làm sai thesis và catalyst có thể đẩy nhanh upside.

#### Risk Matrix (tối thiểu 5 rủi ro)

| Rủi ro | Loại | Xác suất | Tác động (định lượng) | Leading Indicator | Ngưỡng kích hoạt |
|---|---|---|---|---|---|
| [Tên rủi ro] | VM/Ngành/DN/Cạnh tranh | Cao/TB/Thấp | [−X tỷ LNST] | [chỉ số theo dõi] | [khi nào cần hành động] |

- **Xác suất**: Cao > 50% / TB 25-50% / Thấp < 25%
- **Tác động**: Định lượng bằng số tỷ đồng LNST hoặc % EPS impact
- **Leading indicator**: Chỉ số có thể theo dõi hàng tuần/tháng để biết rủi ro có xảy ra không

#### Catalyst Timeline (tối thiểu 3 catalyst)

| Catalyst | Mô tả | Thời điểm dự kiến | Tác động giá ước tính | Xác suất |
|---|---|---|---|---|
| [Sự kiện] | [mô tả] | [Quý/Tháng/Năm] | [+X% đến giá] | Cao/TB/Thấp |

#### Output bắt buộc (artifact Tầng 5)

```
☐ Risk matrix ≥5 rủi ro, mỗi rủi ro có leading indicator
☐ Catalyst table ≥3 catalyst, mỗi catalyst có timeline
→ INSERT vào DOCX Trang 14-15
```

---

### 3.6 Tầng 6 — Định giá

**Mục tiêu**: Tính target price với nhiều kịch bản, có sensitivity.

**Nguyên tắc vàng — multiple từ lịch sử chính cổ phiếu, không trung bình ngành:**
  - Dùng **median P/E, P/B, EV/EBITDA lịch sử của CHÍNH cổ phiếu đó** làm multiple mục tiêu
  - Trung bình ngành chỉ để tham khảo chéo (cross-check), không phải primary input
  - Mỗi doanh nghiệp có vị thế cạnh tranh khác nhau → lịch sử riêng phản ánh đúng hơn
  - **Cách lấy dữ liệu**: Fetch từ Vietcap API (`statistics-financial`), lọc TTM (quarter != 5)
  - Tính **median** (không mean) của tất cả TTM quarters để loại outlier
  - Vẽ biểu đồ phân phối: histogram P/E/PB/EV + median line + percentile bands (25-75%)

#### Chọn phương pháp (xem Mục 7)

```
Căn cứ vào:
1. Ngành của công ty → method chính (bắt buộc)
2. Giai đoạn chu kỳ → method phụ (khuyến nghị)
3. Công ty đa ngành → SOTP
```

#### Bảng định giá Bear/Base/Bull

| Kịch bản | Giá mục tiêu | P/E | EV/EBITDA | Giả định chính | Xác suất |
|---|---|---|---|---|---|
| Bear | XX,XXX | X.Xx | X.Xx | [giả định xấu nhất] | X% |
| Base | XX,XXX | X.Xx | X.Xx | [giả định cơ sở] | X% |
| Bull | XX,XXX | X.Xx | X.Xx | [giả định tốt nhất] | X% |
| **Weighted avg** | **XX,XXX** | | | | **100%** |

> **Tổng xác suất Bear + Base + Bull = 100%.**

#### Sensitivity Table 5×5

Chọn 2 biến quan trọng nhất ảnh hưởng đến định giá (VD: WACC × Growth, P/E × EPS, EBITDA margin × EV/EBITDA multiple):

```
                    ┌──────────┬──────────┬──────────┬──────────┬──────────┐
                    │ Biến 1 — │ [Thấp]   │ [TB-]    │ [Base]   │ [TB+]    │ [Cao]   │
├──────────┼──────────┼──────────┼──────────┼──────────┤
│ Biến 2 — Thấp     │          │          │          │          │          │
│ Biến 2 — TB-      │          │          │          │          │          │
│ Biến 2 — Base     │          │    ✅    │          │          │          │
│ Biến 2 — TB+      │          │          │          │          │          │
│ Biến 2 — Cao      │          │          │          │          │          │
└──────────┴──────────┴──────────┴──────────┴──────────┴──────────┘
```

> Heatmap: giá cao = xanh, giá thấp = đỏ (conditional format trong Excel).

#### So sánh với consensus CTCK

| Nguồn | Khuyến nghị | Giá mục tiêu | P/E target | Ngày |
|---|---|---|---|---|
| Báo cáo này | [tự đánh giá] | [TP] | [PE] | [hiện tại] |
| SSI Research | | | | |
| VCSC | | | | |
| ... | | | | |

#### Output bắt buộc (artifact Tầng 6)

```
☐ Bảng Bear/Base/Bull — xác suất cộng = 100%
☐ Weighted average target price tính được
☐ Sensitivity table 5×5
☐ Bảng consensus ≥2 CTCK (nếu có)
☐ So sánh multiple với peers
→ INSERT vào DOCX Trang 12-13
```

---

## 4. BƯỚC 3: BUILD EXCEL MODEL

### 4.1 Cấu trúc 12+1 sheets

```
┌─────────┬────────────────────────────────────────────────────┐
│ Sheet   │ Nội dung                                           │
├─────────┼────────────────────────────────────────────────────┤
│ 01_Cover│ Thông tin công ty, thesis 1 câu, key metrics       │
│ 02_GD   │ Giả định — TẤT CẢ giả định tập trung tại đây      │
│ 03_DT   │ Doanh thu — Driver-based, tách phân khúc           │
│ 04_PnL  │ P&L lịch sử + dự báo (≥4A + 3E)                   │
│ 05_CDKT │ Bảng CĐKT — tài sản, nợ, VCSH                     │
│ 06_LCTT │ LCTT — CFO, CFI, CFF, FCFF                        │
│ 07_DG   │ Định giá — Tham chiếu từ sheet nguồn, tổng hợp weighted │
│ 08_NS   │ Nhạy cảm — Sensitivity 5×5                         │
│ 09_PEST │ PESTLE — 6 yếu tố vĩ mô                           │
│ 10_CT   │ Cạnh tranh — Porter 5F + Peer comparison           │
│ 11_ND   │ Moat & Quản trị (nếu chưa có sheet riêng)         │
│ 12_RR   │ Rủi ro — Risk Matrix + Catalyst                    │
│ 13_TD   │ Theo dõi — Leading Indicators scoring              │
└─────────┴────────────────────────────────────────────────────┘
```

### 4.2 Quy trình build từng sheet

#### Sheet 01 — Cover

| Element | Nội dung |
|---|---|
| Header | Logo, tên công ty, mã CK, sàn, ngày |
| Key metrics | Giá, vốn hóa, P/E trailing, P/E forward, EPS, BVPS |
| Khuyến nghị | MUA/MUA MẠNH/TÍCH LŨY/TRUNG LẬP/GIẢM/BÁN |
| Thesis | 1 câu — tại sao mua/bán? |

#### Sheet 02 — Giả định (QUAN TRỌNG NHẤT)

TẤT CẢ giả định phải nằm trong sheet này, không hardcode ở sheet khác.

```
┌──────────┬──────────┬──────────┬──────────┬──────────┬──────────┬──────────┐
│ Giả định │ [N-1]A   │ [N]E     │ [N+1]E   │ [N+2]E   │ Nguồn    │ Ghi chú  │
├──────────┼──────────┼──────────┼──────────┼──────────┼──────────┼──────────┤
│ GDP VN % │          │          │          │          │ IMF/GSO  │          │
│ Lãi suất  │          │          │          │          │ SBV      │          │
│ Tỷ giá    │          │          │          │          │          │          │
│ Giá hàng  │          │          │          │          │ LME      │          │
│ ...       │          │          │          │          │          │          │
└──────────┴──────────┴──────────┴──────────┴──────────┴──────────┴──────────┘
```

**Màu sắc**: Input giả định = ô nền vàng chữ xanh. Ý nghĩa: "đây là số do user nhập — đừng sửa nếu không biết".

#### Sheet 03 — Doanh thu (Driver-based)

Áp dụng phương pháp dự báo theo ngành từ Mục 6.

```
┌──────────┬──────────┬──────────┬──────────┬──────────┬──────────┐
│ Driver   │ [N-1]A   │ [N]E     │ [N+1]E   │ [N+2]E   │ Công thức │
├──────────┼──────────┼──────────┼──────────┼──────────┼──────────┤
│ Driver 1 │          │          │          │          │          │
│ Driver 2 │          │          │          │          │          │
│ DT tính   │ =D1×D2  │          │          │          │          │
└──────────┴──────────┴──────────┴──────────┴──────────┴──────────┘
```

**Nguyên tắc**:
- Mỗi dòng driver PHẢI có công thức hoặc reference đến Sheet 02
- KHÔNG hardcode số — mọi số gõ tay đều đặt trong Sheet 02
- Nếu có nhiều mảng → tách dòng riêng, tổng hợp ở cuối

#### Sheet 04 — P&L

```
┌────────────────────┬──────────┬──────────┬──────────┬──────────┬──────────┬──────────┐
│                    │ [N-3]A   │ [N-2]A   │ [N-1]A   │ [N]E     │ [N+1]E   │ [N+2]E   │
├────────────────────┼──────────┼──────────┼──────────┼──────────┼──────────┼──────────┤
│ Doanh thu          │          │          │          │          │          │          │
│ (Chi phí)          │          │          │          │          │          │          │
│ EBITDA             │          │          │          │          │          │          │
│ Khấu hao           │          │          │          │          │          │          │
│ EBIT               │          │          │          │          │          │          │
│ Lãi/vay            │          │          │          │          │          │          │
│ LNTT               │          │          │          │          │          │          │
│ Thuế               │          │          │          │          │          │          │
│ LNST               │          │          │          │          │          │          │
│ LNST CĐCTM         │          │          │          │          │          │          │
└────────────────────┴──────────┴──────────┴──────────┴──────────┴──────────┴──────────┘
```

- Tách 1 dòng phụ: One-off items (nếu có)
- EPS = LNST CĐCTM / Số CP lưu hành
- **Giá theo P/E** = EPS × P/E mục tiêu (tham chiếu từ Assumptions). Ô này là nguồn cho Sheet 07.

#### Sheet 05 — Bảng CĐKT

Cấu trúc chuẩn, đảm bảo Tổng TS = Tổng NV (Balance).

**Các chỉ số bắt buộc**:
- Net Debt = Tổng vay - Tiền
- Invested Capital = Total Equity + Net Debt + Minority
- ROIC, ROE, D/E, Net Debt/EBITDA
- **Giá theo P/B** = BVPS × P/B mục tiêu (tham chiếu từ Assumptions). Ô này là nguồn cho Sheet 07.

#### Sheet 06 — LCTT

```
CFO = LNST + Khấu hao - Δ Working Capital
CFI = -CAPEX (mở rộng + duy trì)
CFF = Cổ tức + Vay mới - Trả nợ + Phát hành CP

FCFF = CFO - CAPEX (chỉ maintenance)
FCFE = FCFF + Vay mới - Trả nợ

- **EV tính** = EBITDA × EV/EBITDA mục tiêu
- **Giá theo EV/EBITDA** = (EV - Nợ ròng + Tiền) / Số CP lưu hành. Ô này là nguồn cho Sheet 07.
- **Giá theo DCF** = FCFF / (WACC - growth) nếu cash flow dương. Ô này là nguồn cho Sheet 07.

#### Sheet 07 — Định giá

**Chỉ là sheet tổng hợp** — không tính toán gì từ đầu. Từng phương pháp định giá đã có ô kết quả tại sheet nguồn:
- Giá theo P/E  → từ Sheet 04 (P&L)
- Giá theo P/B  → từ Sheet 05 (Bảng CĐKT)
- Giá theo EV/EBITDA → từ Sheet 06 (LCTT)
- Giá theo DCF  → từ Sheet 06 (LCTT)

**Cấu trúc bắt buộc**:
```
┌──────────────────────┬──────────┬──────────────────────┬──────────────────────┐
│ Phương pháp          │ Trọng số │ Giá mục tiêu (VND)   │ Nguồn                │
├──────────────────────┼──────────┼──────────────────────┼──────────────────────┤
│ EV/EBITDA            │ XX%      │ =Sheet06!$Z$Z        │ Primary method       │
│ P/B                  │ XX%      │ =Sheet05!$Z$Z        │ Cyclical check       │
│ P/E                  │ XX%      │ =Sheet04!$Z$Z        │ Tham khảo            │
│ DCF                  │ XX%      │ =Sheet06!$Z$Z        │ Nếu CF ổn định       │
├──────────────────────┼──────────┼──────────────────────┼──────────────────────┤
│ Giá mục tiêu weighted│ 100%     │ =SUMPRODUCT(...)     │                     │
│ Giá hiện tại         │          │ [giá]                │                     │
│ Upside               │          │ =(TP/Price)-1        │                     │
└──────────────────────┴──────────┴──────────────────────┴──────────────────────┘
```

Cũng chứa 3 kịch bản Bear/Base/Bull — mỗi kịch bản là 1 bảng riêng, thay đổi multiple mục tiêu ở Assumptions.

**Ưu điểm**: Bấm vào ô nào cũng trace được công thức gốc — không phải đoán "số này từ đâu ra".

#### Sheet 08 — Nhạy cảm

Sensitivity table 5×5 với conditional formatting heatmap.

#### Sheet 09 — PESTLE

#### Sheet 10 — Cạnh tranh (Porter 5F + Peer comparison)

#### Sheet 11 — Moat & Quản trị

#### Sheet 12 — Rủi ro (Risk Matrix)

#### Sheet 13 — Theo dõi (Leading Indicators Scoring)

**Framework 7 bước** — xem Mục 4.3 bên dưới.

### 4.3 Sheet 13 — Chấm điểm Leading Indicators

#### BƯỚC 1 — XÁC ĐỊNH TRIGGER — 10-20 yếu tố có khả năng tác động đến P&L

```
┌──────────┬──────────────┬────────────────┬─────────────────┬──────────────┬───────────────────┐
│ Yếu tố   │ Loại trigger │ Kênh tác động  │ Chiều tác động  │ Mức độ 1-5   │ Nguồn theo dõi    │
│          │ (VM/Ngành/DN)│                │ (+/−)           │              │                   │
├──────────┼──────────────┼────────────────┼─────────────────┼──────────────┼───────────────────┤
│ [Trigger]│              │ Giá → DT       │ +               │ 4            │ [URL/dataset]     │
│ ...      │              │                │                 │              │                   │
└──────────┴──────────────┴────────────────┴─────────────────┴──────────────┴───────────────────┘
```

Mỗi trigger ghi rõ: tên, % COGS ước tính, nguồn dữ liệu theo dõi.

#### BƯỚC 2 — BẢN ĐỒ NHÂN QUẢ

Với mỗi trigger quan trọng, vẽ chuỗi: **Macro/Ngành trigger → Tác động ngành → Tác động công ty → Kết quả tài chính cụ thể**.

Bắt buộc ước tính lag time từng bước. Ví dụ:
```
Giá đầu vào tăng → Chi phí/tấn tăng → Biên GP giảm → LNST giảm | Lag tổng: 1-2 tháng
Cầu ngành tăng   → Đơn hàng tăng   → DT tăng      → EBITDA tăng | Lag tổng: 2-3 quý
Catalyst chiến lược hoàn thành → Công suất mới → Chi phí cố định hấp thụ tốt hơn | Lag: 12-18 tháng
```

Chuỗi nào lag ngắn hơn → có thể dùng làm leading indicator.

#### BƯỚC 3 — LEADING INDICATORS (ưu tiên 5-7 indicators)

Chỉ lấy những indicator thỏa mãn: **(a)** có dữ liệu công khai, **(b)** xuất hiện trước kết quả tài chính ít nhất 1 quý, **(c)** đặc thù cho công ty/ngành này.

Phân tầng:
- **Tầng 1-2 (vĩ mô):** Chỉ đưa vào nếu có chuỗi nhân quả rõ ràng đến P&L
- **Tầng 3 (ngành):** Giá nguyên liệu đặc thù, chỉ số cung/cầu ngành, hành động cạnh tranh
- **Tầng 4 (doanh nghiệp):** Tiến độ dự án chiến lược, đơn hàng tồn đọng, mở rộng công suất

Cột bắt buộc cho mỗi indicator:

```
┌──────────────────┬──────────────┬───────────────────┬──────────┬────────────────────────┬────────────────────┬──────────────┐
│ Indicator        │ Giá trị HT   │ Ngưỡng tốt/TL/xấu │ Trọng số │ Nguồn dữ liệu          │ Lag time ước tính  │ Tần suất     │
└──────────────────┴──────────────┴───────────────────┴──────────┴────────────────────────┴────────────────────┴──────────────┘
```

Ngưỡng 3 vùng tô màu QUAL(): `'pos'` (xanh) / `'neu'` (vàng) / `'neg'` (đỏ).

**Lưu ý hướng ngưỡng**:
- Chi phí đầu vào, rủi ro, đòn bẩy: **Thấp = tốt** → ngưỡng tốt ở vùng thấp
- Doanh thu, biên lợi nhuận, thị phần: **Cao = tốt** → ngưỡng tốt ở vùng cao
- Tỷ lệ công suất, LDR ngân hàng: **Dải tối ưu** → tốt ở giữa, xấu ở cả hai đầu

#### BƯỚC 4 — COINCIDENT INDICATORS (3-5 indicators)

Phản ánh trạng thái hiện tại. Dùng để xác nhận chiều hướng, không để dự báo.

```
┌──────────────────┬──────────────┬────────────────┬──────────────┐
│ Indicator        │ Giá trị HT   │ Nguồn          │ Tần suất     │
└──────────────────┴──────────────┴────────────────┴──────────────┘
```

Ví dụ: Doanh thu quý YoY%, biên GP thực tế, volume sản lượng, số đơn hàng mới.

#### BƯỚC 5 — LAGGING INDICATORS (3-4 indicators)

Xác nhận sau khi kết quả đã xảy ra — **KHÔNG dùng để ra quyết định mua/bán**.

Ví dụ: EPS trailing 12T, ROE năm, P/E trailing, tỷ lệ cổ tức thực trả.

#### BƯỚC 6 — TRẠNG THÁI HIỆN TẠI & ĐIỂM TỔNG HỢP

Chấm điểm từng leading indicator: **+2 / +1 / 0 / -1 / -2**

```
┌──────────────────┬──────────┬──────────────┬─────────────────────┐
│ Indicator        │ Trọng số │ Điểm (-2→+2) │ Điểm có trọng số    │
├──────────────────┼──────────┼──────────────┼─────────────────────┤
│ [Indicator 1]    │ XX%      │ [điền]       │ [tự tính]           │
│ ...              │ ...      │ ...          │ ...                 │
├──────────────────┼──────────┼──────────────┼─────────────────────┤
│ TỔNG             │ 100%     │              │ [SUM]               │
└──────────────────┴──────────┴──────────────┴─────────────────────┘
```

Mapping điểm tổng → khuyến nghị:
```
> +1.2   : MUA MẠNH   (CGL_BG)
+0.6→+1.2: MUA        (CGL_BG nhạt)
+0.2→+0.6: TÍCH LŨY   (CYL_BG)
-0.2→+0.2: TRUNG LẬP  (CYL_BG)
-0.6→-0.2: GIẢM       (CRL_BG nhạt)
< -0.6   : BÁN        (CRL_BG)
```

#### BƯỚC 7 — RỦI RO & CATALYST

```
RỦI RO ĐẢO CHIỀU (3 rủi ro có thể làm thesis sai hoàn toàn):
• [Rủi ro 1: mô tả cụ thể + ngưỡng kích hoạt + tác động ước tính đến LNST]
• [Rủi ro 2]
• [Rủi ro 3]

CATALYST TĂNG TỐC THESIS (3 catalyst có thể đẩy nhanh upside):
• [Catalyst 1: sự kiện cụ thể + thời điểm dự kiến + tác động giá ước tính]
• [Catalyst 2]
• [Catalyst 3]
```

**Quy tắc bắt buộc khi build Sheet 13:**
- Phải hoàn thành Bước 1 & 2 (phân tích ngành + bản đồ nhân quả) trước khi chọn indicator
- Không đưa indicator nào vào nếu không xác định được lag time và nguồn dữ liệu
- Trọng số các leading indicators phải cộng lại = 100%
- Phải cập nhật lại mỗi khi phân tích mới — không copy nguyên từ ticker khác

### 4.4 Quy trình sau khi hoàn thành Excel

```bash
# Bắt buộc: chạy recalc và kiểm tra 0 lỗi
python3 scripts/recalc.py "<path_to_xlsx>"

# Kết quả phải đạt:
{"status": "success", "total_errors": 0}
```

---

## 5. BƯỚC 4: VIẾT DOCX RESEARCH REPORT

> **Nguyên tắc cốt lõi**: DOCX là sản phẩm ASSEMBLY — lắp ghép các artifacts đã phân tích từ 6 tầng, KHÔNG phải viết lại từ đầu. Nếu tầng nào chưa có artifact → PHẢI hoàn thành trước, không được viết DOCX trước để bù sau.

### 5.1 ĐIỀU KIỆN TIÊN QUYẾT — Pre-flight checklist

```
PRE-FLIGHT CHECKLIST — Kiểm tra trước khi gõ dòng đầu tiên của DOCX:

☐ Tầng 1: Bảng chuỗi giá trị hoàn chỉnh (≥5 bước, có biên ước tính)
☐ Tầng 2: Bảng Porter 5F hoàn chỉnh (5 lực, có kết luận mức độ hấp dẫn ngành)
☐ Tầng 2: Vị trí chu kỳ ngành + TAM đã xác định
☐ Tầng 3: Bảng moat rating hoàn chỉnh (5 loại, có overall rating Wide/Narrow/No)
☐ Tầng 3: ROIC vs WACC đã tính; đánh giá chất lượng quản lý đã có
☐ Tầng 4: Revenue driver table hoàn chỉnh (A/E đúng theo Mục 2.5)
☐ Tầng 4: Bảng P&L lịch sử ≥4A + 3E, one-off đã tách riêng
☐ Tầng 5: Risk matrix ≥5 rủi ro, có xác suất + tác động + leading indicator
☐ Tầng 6: Valuation Bear/Base/Bull hoàn chỉnh, xác suất cộng = 100%
☐ Tầng 6: Sensitivity table 5×5 đã có
☐ Sheet 13 Excel: Scoring leading indicators đã hoàn chỉnh, điểm tổng đã tính
☐ Consensus ≥2 CTCK đã tìm và ghi lại

→ Nếu còn ô chưa tick: DỪNG — quay lại hoàn thiện artifact đó trước.
```

### 5.2 THỨ TỰ BUILD BẮT BUỘC

```
Thứ tự viết DOCX (bắt buộc theo trình tự này):

BƯỚC 1: Trang 3–4   ← Chuỗi giá trị & Mô hình KD (Tầng 1)
BƯỚC 2: Trang 5–6   ← Thị trường & Cạnh tranh (Tầng 2)
BƯỚC 3: Trang 7–8   ← Moat & Quản trị (Tầng 3)
BƯỚC 4: Trang 9–11  ← Phân tích tài chính (Tầng 4)
BƯỚC 5: Trang 12–13 ← Định giá (Tầng 6) — viết TRƯỚC rủi ro
BƯỚC 6: Trang 14–15 ← Rủi ro & Catalyst (Tầng 5)
BƯỚC 7: Trang 16    ← Leading Indicators (Sheet 13 Excel)
BƯỚC 8: Trang 17+   ← Phụ lục (BCTC tóm tắt, peer table, nguồn)
BƯỚC 9: Trang 2     ← Investment Summary (VIẾT CUỐI CÙNG — tổng hợp tất cả)
BƯỚC 10: Trang 1    ← Cover Page
```

> **Tại sao Investment Summary viết cuối?** Vì nó là **distillation** của toàn bộ phân tích — không thể viết đúng nếu chưa có đủ con số từ các bước trước.

### 5.3 ASSEMBLY GUIDE — Từng trang DOCX

#### Trang 1 — Cover Page

| Element | Nội dung bắt buộc |
|---|---|
| Tên công ty (tiếng Việt đầy đủ) | VD: "CÔNG TY CỔ PHẦN GEMADEPT" |
| Mã + Sàn | "GMD \| HOSE" |
| Loại báo cáo | "Báo cáo Phân tích Đầu tư" |
| Ngày | "Tháng MM/YYYY" |
| Bảng key metrics (4 ô) | Giá hiện tại \| Giá mục tiêu \| Upside \| Khuyến nghị |
| Bảng thông số (4 ô) | Vốn hóa \| P/E forward \| EV/EBITDA \| Net Cash/(Debt) |

#### Trang 2 — Investment Summary (VIẾT SAU CÙNG)

Xem template đầy đủ tại Mục 5.4 bên dưới.

**Quy tắc cứng:**
- Phải đứng độc lập — người đọc chỉ trang này đủ hiểu thesis
- Không được dài hơn 1 trang A4
- Không được dùng câu mơ hồ kiểu "có thể tích cực hoặc tiêu cực"
- Phải có BUY / HOLD / SELL / TÍCH LŨY / QUAN SÁT rõ ràng

#### Trang 3–4 — Chuỗi giá trị & Mô hình kinh doanh

**Nguồn**: Artifact Tầng 1 (bảng chuỗi giá trị)

**Phải có:**
- Bảng chuỗi giá trị (từ artifact Tầng 1) — KHÔNG viết lại, INSERT trực tiếp
- Đoạn giải thích: công ty capture giá trị ở đâu, tại sao, biên như thế nào
- Mô hình doanh thu: công thức và drivers chính (từ Mục 6 phương pháp ngành)
- Bảng cơ cấu doanh thu theo mảng (nếu đa mảng)

**Không được:**
- ❌ Mô tả chung chung về ngành không liên quan đến vị trí của công ty
- ❌ Copy Wikipedia/tóm tắt Wikipedia về ngành

#### Trang 5–6 — Phân tích thị trường & Cạnh tranh

**Nguồn**: Artifact Tầng 2 (bảng Porter 5F + chu kỳ + TAM)

**Phải có:**
- Bảng Porter 5F đầy đủ 5 lực — INSERT từ artifact Tầng 2
- Kết luận Porter: "Ngành [tên] ở mức hấp dẫn [Cao/TB/Thấp] vì..."
- TAM + vị trí thị phần công ty
- Vị trí chu kỳ ngành với giải thích 2–3 câu
- Bối cảnh cạnh tranh hiện tại: ai là đối thủ chính, xu hướng gần đây

**Không được:**
- ❌ Chỉ liệt kê tên đối thủ mà không có phân tích so sánh
- ❌ Bỏ qua kết luận — mỗi lực phải có chiều tác động rõ ràng

#### Trang 7–8 — Lợi thế cạnh tranh & Quản trị

**Nguồn**: Artifact Tầng 3 (bảng moat + ROIC test + quản lý)

**Phải có:**
- Bảng moat 5 loại — INSERT từ artifact Tầng 3
- Overall moat rating: **Wide / Narrow / No Moat** với 1–2 câu lý do
- ROIC vs WACC: số liệu + kết luận (có tạo giá trị thật không?)
- Đánh giá chất lượng quản lý: 3 tiêu chí (capital allocation, communication, alignment)
- Bảng ROIC so sánh peers (ít nhất 3 peers)

**Không được:**
- ❌ Kết luận moat mà không có bằng chứng định lượng
- ❌ Bỏ qua ROIC vs WACC — đây là test khách quan nhất

#### Trang 9–11 — Phân tích tài chính

**Nguồn**: Artifact Tầng 4 (driver table + P&L table + bridge)

**Phải có (theo thứ tự):**
1. Bảng Revenue Driver Table — INSERT từ artifact Tầng 4
2. Bảng P&L lịch sử + dự báo (≥4A + 3E) — INSERT từ artifact Tầng 4
3. Phân tích chất lượng lợi nhuận: one-off, cash conversion, accruals
4. Bridge analysis: từ [N-1]A → [N]E, giải thích từng dòng thay đổi lớn
5. Phân tích BCTC: B/S strength (Net Debt/EBITDA, ICR), Cash Flow quality

**Ràng buộc A/E:**
- Tất cả nhãn A/E phải theo Mục 2.5
- Năm có one-off phải có chú thích rõ: "* Bao gồm [X] tỷ one-off từ [sự kiện]"

**Không được:**
- ❌ Điền số ước tính vào cột A
- ❌ Bỏ qua phân tích một-off — làm sai lệch EPS core

#### Trang 12–13 — Định giá

**Nguồn**: Artifact Tầng 6 (valuation table + sensitivity + consensus)

**Phải có:**
- Bảng Valuation Bear/Base/Bull — INSERT từ artifact Tầng 6
- Tổng xác suất = 100% (viết rõ: "Bear X% + Base X% + Bull X% = 100%")
- Weighted average target price tính được
- Sensitivity table 5×5 — INSERT từ artifact Tầng 6
- Bảng consensus CTCK (nếu có ≥2 báo cáo)
- So sánh multiple với peers (P/E, EV/EBITDA)

**Không được:**
- ❌ Viết range giá mục tiêu rộng hơn 15%
- ❌ Bỏ qua sensitivity — người đọc cần biết giả định nào quan trọng nhất
- ❌ Xác suất kịch bản không cộng đủ 100%

#### Trang 14–15 — Rủi ro & Catalyst

**Nguồn**: Artifact Tầng 5 (risk matrix + catalyst table)

**Phải có:**
- Bảng Risk Matrix — INSERT từ artifact Tầng 5 (ít nhất 5 rủi ro)
- Bảng Catalyst Timeline — INSERT từ artifact Tầng 5 (ít nhất 3 catalyst)
- Với mỗi rủi ro: mô tả + ngưỡng kích hoạt + xác suất + tác động định lượng + leading indicator
- Với mỗi catalyst: sự kiện cụ thể + thời điểm + tác động giá ước tính + xác suất

**Không được:**
- ❌ Rủi ro mơ hồ kiểu "tình hình kinh tế xấu đi" — phải cụ thể và đo được
- ❌ Thiếu leading indicator cho mỗi rủi ro

#### Trang 16 — Leading Indicators

**Nguồn**: Sheet 13 Excel (Framework 7 bước)

**Phải có:**
- Điểm scoring tổng hợp từ Sheet 13 (copy kết quả, không tóm tắt lại)
- Bảng 6 leading indicators với giá trị hiện tại + ngưỡng + điểm
- Điểm tổng có trọng số → mapping sang khuyến nghị (theo thang đã quy định)
- Hướng dẫn khi nào điều chỉnh khuyến nghị (ngưỡng downgrade/upgrade)

#### Trang 17+ — Phụ lục (BẮT BUỘC, không được bỏ qua)

**Phụ lục A — BCTC tóm tắt 5 năm:**

| Chỉ tiêu | [N-4]A | [N-3]A | [N-2]A | [N-1]A | [N]A |
|---|---|---|---|---|---|
| Doanh thu | | | | | |
| EBITDA | | | | | |
| LNST CĐCTM | | | | | |
| EPS (VNĐ) | | | | | |
| Tổng tài sản | | | | | |
| Nợ vay | | | | | |
| VCSH | | | | | |
| CFO | | | | | |
| CAPEX | | | | | |

**Phụ lục B — Peer Comparison:**

| Công ty | Ticker | Vốn hóa | P/E | EV/EBITDA | ROE | Net Debt/EBITDA |
|---|---|---|---|---|---|---|
| [Công ty phân tích] | | | | | | |
| [Peer 1] | | | | | | |
| [Peer 2–4] | | | | | | |
| **Median peers** | | | | | | |

**Phụ lục C — Nguồn dữ liệu:**
Liệt kê tất cả nguồn đã dùng: BCTC (nguồn + ngày), giá CP (ngày), báo cáo CTCK (tên + ngày), tin tức tham khảo.

### 5.4 INVESTMENT SUMMARY — Template bắt buộc (Trang 2)

```
╔══════════════════════════════════════════════════════════════════════╗
║  [MÃ CK] — [TÊN CÔNG TY ĐẦY ĐỦ TIẾNG VIỆT]                        ║
║  [Ngành] | [Sàn] | Báo cáo: DD/MM/YYYY                              ║
╠══════════════════════════════════════════════════════════════════════╣
║  Khuyến nghị: [BUY/HOLD/SELL/TÍCH LŨY/QUAN SÁT]                    ║
║  Giá hiện tại: XX.XXX VNĐ  |  Giá mục tiêu: XX.XXX–XX.XXX VNĐ     ║
║  Upside: XX,X%  |  Vốn hóa: XX.XXX tỷ VNĐ                          ║
╠══════════════════════════════════════════════════════════════════════╣
║  LUẬN ĐIỂM ĐẦU TƯ (2–3 câu, súc tích, không mơ hồ):               ║
║  [Câu 1: Tại sao mua — lý do cốt lõi nhất]                         ║
║  [Câu 2: Catalyst gần nhất xác nhận thesis]                         ║
║  [Câu 3: Rủi ro chính cần theo dõi và tại sao vẫn mua]             ║
╠══════════════════════════════════════════════════════════════════════╣
║  KỊCH BẢN ĐỊNH GIÁ:                                                 ║
║                    Bear (X%)     Base (X%)     Bull (X%)            ║
║  Giá mục tiêu:    XX.XXX VNĐ   XX.XXX VNĐ   XX.XXX VNĐ            ║
║  P/E:             X,Xx          X,Xx          X,Xx                  ║
║  EV/EBITDA:       X,Xx          X,Xx          X,Xx                  ║
║  Giả định chính:  [1 từ]        [1 từ]        [1 từ]               ║
╠══════════════════════════════════════════════════════════════════════╣
║  3 LÝ DO ĐỂ MUA:              3 RỦI RO CHÍNH:                      ║
║  1. [cụ thể, định lượng]      1. [cụ thể, có ngưỡng kích hoạt]     ║
║  2. [cụ thể, định lượng]      2. [cụ thể, có ngưỡng kích hoạt]     ║
║  3. [cụ thể, định lượng]      3. [cụ thể, có ngưỡng kích hoạt]     ║
╠══════════════════════════════════════════════════════════════════════╣
║  SNAPSHOT TÀI CHÍNH:                                                 ║
║           [N-3]A  [N-2]A  [N-1]A   [N]A   [N+1]E  [N+2]E          ║
║  DT(tỷ):                                                             ║
║  LNST(tỷ):                                                           ║
║  EPS(VNĐ):                                                           ║
║  P/E:                                                                ║
║  EV/EBIT:                                                            ║
╠══════════════════════════════════════════════════════════════════════╣
║  MOAT: Wide/Narrow/No  |  Chu kỳ: [giai đoạn]  |  ROIC vs WACC:✅❌║
║  LEADING SCORE: [+X,XX] → [KHUYẾN NGHỊ]                             ║
╚══════════════════════════════════════════════════════════════════════╝
```

**Ví dụ điền — GMD (tháng 6/2026):**

```
╔══════════════════════════════════════════════════════════════════════╗
║  GMD — CÔNG TY CỔ PHẦN GEMADEPT                                     ║
║  Cảng biển & Logistics | HOSE | Báo cáo: 18/06/2026                 ║
╠══════════════════════════════════════════════════════════════════════╣
║  Khuyến nghị: TÍCH LŨY                                              ║
║  Giá hiện tại: 73.400 VNĐ  |  Giá mục tiêu: 79.000–83.000 VNĐ    ║
║  Upside: +7,6%–13,2%  |  Vốn hóa: 31.305 tỷ VNĐ                  ║
╠══════════════════════════════════════════════════════════════════════╣
║  LUẬN ĐIỂM ĐẦU TƯ:                                                  ║
║  GMD sở hữu 2 cảng chiến lược không thể nhân bản (NĐV tại Hải      ║
║  Phòng, Gemalink tại Cái Mép) với Narrow Moat từ vị trí địa lý và  ║
║  giấy phép 50 năm. THC tăng +10% từ 2/2026 là catalyst tức thì     ║
║  đang phản ánh vào P&L (Q1/2026: LNTT +23% YoY). Rủi ro gần nhất  ║
║  là MSC rút cargo NĐV — cần theo dõi số tàu cập bến tuần/tuần.     ║
╠══════════════════════════════════════════════════════════════════════╣
║  KỊCH BẢN ĐỊNH GIÁ:                                                 ║
║                    Bear (25%)    Base (50%)    Bull (25%)           ║
║  Giá mục tiêu:    52.397 VNĐ   80.394 VNĐ   108.391 VNĐ          ║
║  P/E:             12,0x         17,0x         22,0x                ║
║  EV/EBITDA:       9,0x          12,0x         15,0x                ║
║  Giả định chính:  MSC không bù  Base case     THC tăng lần 2       ║
╠══════════════════════════════════════════════════════════════════════╣
║  3 LÝ DO ĐỂ MUA:              3 RỦI RO CHÍNH:                      ║
║  1. THC +10% từ 2/2026 →      1. MSC rút NĐV: mỗi tháng delay     ║
║     +840 tỷ DT incremental,      = −30 tỷ LNST. Indicator: số tàu  ║
║     margin tác động ngay        MSC/tuần tại NĐV                   ║
║  2. Gemalink Phase 2A khởi    2. Dư cung HP: Lạch Huyện berths     ║
║     công 4/2026 → growth         3–8 vận hành → áp lực THC dài hạn ║
║     option +600K TEU/năm       3. Phase 2A delay: EIA phức tạp →   ║
║  3. Net cash 2.894 tỷ →          delay 6–12T = de-rate múltiple    ║
║     CAPEX tự tài trợ không                                          ║
║     cần phát hành thêm CP                                           ║
╠══════════════════════════════════════════════════════════════════════╣
║  SNAPSHOT TÀI CHÍNH:                                                 ║
║         2022A   2023A*  2024A   2025A   2026E   2027E              ║
║  DT:    3.916   3.846   4.832   5.200   7.410   8.884              ║
║  LNST:   995     830*  1.470   1.720   2.057   2.582              ║
║  EPS:   2.333  1.946*  3.448   4.032   4.824   6.054              ║
║  P/E:   31,5x  37,7x   21,3x   18,2x   15,2x   12,1x             ║
║  EV/EB: 22,0x  20,4x   17,1x   14,5x   12,0x    9,8x             ║
║  * 2023 loại one-off 1.772 tỷ thoái vốn                            ║
╠══════════════════════════════════════════════════════════════════════╣
║  MOAT: Narrow Moat  |  Chu kỳ: Growth→Maturity  |  ROIC>WACC: ✅   ║
║  LEADING SCORE: +0,40 → TÍCH LŨY                                    ║
╚══════════════════════════════════════════════════════════════════════╝
```

### 5.5 Yêu cầu chất lượng DOCX

- **Ngôn ngữ**: Tiếng Việt hoàn toàn (giữ tiếng Anh cho thuật ngữ chuyên môn không có từ VN tương đương: EV, EBITDA, NIM, ROIC, WACC, moat, DCF...)
- **Số liệu**: Luôn ghi nguồn và ngày lấy dữ liệu — không có số "trên trời"
- **Kết luận**: Mỗi section phải kết thúc bằng 1 câu kết luận rõ ràng (positive/negative/neutral)
- **Không dùng cụm mơ hồ**: "có thể", "nhiều khả năng", "tùy thuộc vào" → phải quantify
- **Tối thiểu 15 trang**, không tính cover và phụ lục
- **Tất cả bảng**: số liệu căn phải, mô tả căn trái, header căn giữa (xem Mục 9)
- **Không có ô trống không giải thích** — nếu chưa có dữ liệu, ghi "[chưa xác nhận]"

---

## 6. PHƯƠNG PHÁP DỰ BÁO DOANH THU THEO NGÀNH

> Đây là menu — chọn phương pháp phù hợp với bản chất kinh doanh, không phải với tên ngành. Một công ty có thể kết hợp nhiều phương pháp cho các mảng khác nhau.

---

### 6.1 Sản xuất Commodity (Thép, Kim loại, Hóa chất, Vật liệu xây dựng)
**Ví dụ mẫu**: HPG, BMP, TLG, KSV

```
Doanh thu = Sản lượng × Giá bán bình quân
Biên lợi nhuận = Giá bán - Chi phí nguyên liệu - Chi phí chuyển đổi

Drivers cần dự báo riêng:
- Sản lượng: Công suất × Utilization rate
- Giá bán: Follow giá hàng hóa thế giới (LME, HRC index...)
- Chi phí nguyên liệu: Giá quặng, than, năng lượng
- Spread/EBITDA per tấn: Metric quan trọng nhất
```

**Leading indicators**: Giá LME, giá HRC Trung Quốc, tồn kho ngành, KLGD thị trường.

---

### 6.2 Bán lẻ Chuỗi (Multi-store Retail)
**Ví dụ mẫu**: MWG, PNJ, FRT

```
Doanh thu = Số cửa hàng × Doanh thu/CH bình quân
Doanh thu/CH = Lưu lượng khách × Tỷ lệ chuyển đổi × Giá trị giỏ hàng

SSS Growth = Thay đổi DT của CH đã mở >12 tháng

Cần tách riêng:
- New store contribution (mở thêm)
- SSS của existing stores (cải thiện hiệu quả)
- Mix effect (đóng cửa hàng yếu, mở lại CH mới)
```

**Trường hợp đặc biệt**: Nếu công ty có nhiều chuỗi ở các chu kỳ khác nhau (như MWG: TGDĐ vs BHX vs AVA) → tách riêng từng chuỗi, dùng SOTP.

---

### 6.3 Phân phối (Distribution)
**Ví dụ mẫu**: DGW, PET, PHR

```
Doanh thu = Thị trường tổng × Thị phần
Biên = Gross margin phân phối (thường 5-8%, mỏng)

Drivers:
- Doanh số thị trường của hãng principal (Apple, Xiaomi, HP...)
- Số lượng hãng/brand được phân phối
- Mix hãng cao cấp vs. phổ thông (ASP)
- Thị phần tại channel phân phối
```

---

### 6.4 Ngân hàng Thương mại
**Ví dụ mẫu**: TCB, ACB, MBB, VCB

```
Thu nhập lãi thuần (NII) = Tổng TSS lãi × NIM
Total Income = NII + Non-interest income (phí, ngoại hối, bảo hiểm bancas)
Chi phí tín dụng = Dư nợ × Credit cost
LNTT = Total Income × (1 - CIR%) - Chi phí tín dụng

KPIs theo dõi:
- NIM (Net Interest Margin): trend tăng/giảm
- CIR (Cost-to-Income): hiệu quả vận hành
- NPL ratio: chất lượng tài sản
- LLR/NPL: mức độ phòng thủ
- Tăng trưởng tín dụng vs. room NHNN
- CASA ratio: chi phí huy động vốn
```

**Định giá**: Residual Income model + P/B — không dùng EV (ngân hàng không có EV theo nghĩa thông thường).

---

### 6.5 Chứng khoán
**Ví dụ mẫu**: SSI, VND, HCM, MBS

```
Doanh thu = Phí môi giới + Lãi cho vay margin + Lãi tự doanh + Phí IB

Môi giới = KLGD thị trường × Thị phần × Phí bình quân
Margin = Dư nợ margin × Lãi suất bình quân

Drivers theo VN-Index:
- VN-Index tăng → KLGD tăng → phí môi giới tăng
- VN-Index giảm → tự doanh lỗ → LNST volatile
```

**Định giá**: Cần scenario analysis (thị trường bullish/neutral/bearish).

---

### 6.6 Khu công nghiệp
**Ví dụ mẫu**: NTC, SZC, BCM, IDC

```
Doanh thu = Diện tích cho thuê mới × Giá thuê (USD/m²) × Tỷ giá
          + Phí dịch vụ từ diện tích đã cho thuê
          + Doanh thu điện, nước, hạ tầng

Drivers:
- Quỹ đất sạch còn lại (ha) và tiến độ giải phóng mặt bằng
- Tốc độ hấp thụ hàng năm (ha signed)
- Giá thuê USD/m²/kỳ thuê (50 năm)
- FDI inflow vào VN và vùng địa lý của KCN
```

**Định giá chính**: RNAV — Quy đổi toàn bộ quỹ đất về NAV hiện tại.

---

### 6.7 Bất động sản Nhà ở
**Ví dụ mẫu**: NLG, KDH, DXG, Nam Long

```
Doanh thu = Σ (Số căn bàn giao × Giá bán bình quân) theo từng dự án

Đặc điểm: Doanh thu rất lumpy — phụ thuộc lịch bàn giao
Không thể extrapolate — phải model từng dự án riêng

Pipeline phân tích:
- Dự án đang bán/bàn giao: số căn × giá × tỷ lệ sell-through
- Dự án sắp mở bán: tiến độ pháp lý, giá dự kiến
- Quỹ đất dự phòng: giá đất, tiềm năng phát triển
```

---

### 6.8 EPC / Xây lắp
**Ví dụ mẫu**: TV2, PC1, HHV, VCG

```
Doanh thu = Backlog đầu kỳ × % ghi nhận + Hợp đồng mới ký × % ghi nhận
Backlog cuối kỳ = Backlog đầu + Hợp đồng mới ký - Doanh thu đã ghi nhận

Drivers:
- Backlog hiện tại (tổng giá trị chưa ghi nhận)
- Win rate đấu thầu + pipeline dự án mới
- Tỷ lệ ghi nhận (%) theo tiến độ thi công
- Biên gộp theo loại công trình (điện, giao thông, dân dụng)
```

---

### 6.9 Tập đoàn Đa ngành
**Ví dụ mẫu**: MSN, VIC, REE, HAX

```
Không có công thức chung — phân tách từng mảng:
- Mảng A: Revenue method phù hợp với ngành A
- Mảng B: Revenue method phù hợp với ngành B
- Phần holding/corporate: quản lý danh mục

Sau đó SOTP để tổng hợp.
Conglomerate discount: 15-25% vào NAV tổng (phản ánh complexity).
```

---

### 6.10 Logistics / Cảng biển
**Ví dụ mẫu**: GMD, HAH, VSC, PHP

```
Cảng biển:
Doanh thu = Sản lượng container (TEU) × Giá xếp dỡ (USD/TEU) × Tỷ giá
          + Doanh thu lưu bãi + Dịch vụ kho, logistics

Drivers:
- Tăng trưởng xuất nhập khẩu VN (XNK GDP)
- Market share tại cảng/khu vực
- Cước phí container quốc tế (SCFI index)
- Công suất kho, diện tích bãi

Vận tải biển (HAH):
Doanh thu = Số tàu × Capacity (TEU/tàu) × Utilization × Cước phí/TEU
```

**Leading indicators**: SCFI (Shanghai Containerized Freight Index), XNK VN hàng tháng.

---

### 6.11 Dược phẩm / Y tế
**Ví dụ mẫu**: DHG, IMP, VMD, DBD

```
Doanh thu = Kênh ETC (đấu thầu BV) + Kênh OTC (nhà thuốc) + Xuất khẩu

Kênh ETC:
- Dự báo theo ngân sách mua thuốc Nhà nước
- Tỷ lệ win thầu, số gói thầu
- Giá đấu thầu (thường áp lực giảm giá)

Kênh OTC:
- Số điểm bán × Doanh thu/điểm
- Nhận diện thương hiệu, marketing spend
- Xu hướng tự mua thuốc

Drivers:
- Tỷ lệ tăng dân số + già hóa dân số
- Số bệnh viện, giường bệnh mới
- Quy định đấu thầu thuốc (Thông tư 15, 08...)
- Pipeline sản phẩm mới (Generic vs. Branded)
```

---

### 6.12 Công nghệ / Phần mềm / IT Services
**Ví dụ mẫu**: FPT, CMG, ELC

```
FPT (đa ngành tech):
SOTP: IT Services (offshore) + FPT Telecom + FPT Education + FPT Retail

IT Services / Offshore:
Doanh thu = Số nhân sự billable × Utilization × Rate/người/tháng × Tỷ giá
Hoặc: Revenue per headcount × Headcount growth

Telecom:
Doanh thu = Số thuê bao × ARPU (Average Revenue Per User) × 12
           + Doanh thu B2B (bandwidth, data center)

Education:
Doanh thu = Số học viên/sinh viên × Học phí bình quân

Drivers:
- Tỷ giá USD/VND (IT Services thu USD)
- Headcount tăng trưởng và attrition rate
- Thị trường offshore IT (Nhật, Mỹ, EU)
- ARPU telecom và penetration rate
```

---

### 6.13 Tiêu dùng / FMCG
**Ví dụ mẫu**: VNM, SAB, MCH, MSN Consumer

```
Doanh thu = Volume × Giá bán bình quân (ASP)
Volume = Thị trường tổng × Thị phần

Hoặc: Distribution-based model
Doanh thu = Số điểm phân phối × Doanh thu/điểm bán

Drivers:
- GDP per capita → sức tiêu dùng
- Thị phần trong danh mục
- ASP: pricing power, mix shift sang premium
- Điểm phân phối: MT (Modern Trade) vs. GT (General Trade)
- Marketing spend / A&P

Leading indicators:
- Nielsen retail sales data (nếu có)
- PMI sản xuất hàng tiêu dùng
- Tỷ lệ lạm phát (ảnh hưởng đến ASP và demand)
```

---

### 6.14 Khoáng sản / Mining
**Ví dụ mẫu**: KSV (Vimico), MSR (Masan HTM)

```
Doanh thu = Sản lượng × Giá hàng hóa quốc tế × Tỷ giá

Sản lượng: Công suất mỏ × Utilization rate
Giá: Follow LME / giá spot quốc tế — công ty là price taker

Option Value assets: Nếu có mỏ chưa khai thác (đất hiếm, vàng...)
→ Định giá riêng bằng NAV/option value method

Leading indicators:
- LME prices (Cu, Zn, Au, Ag...)
- Inventory tại sàn LME
- Trung Quốc demand signals (PMI, construction activity)
- USD index (nghịch chiều với commodity giá)
```

---

## 7. PHƯƠNG PHÁP ĐỊNH GIÁ THEO NGÀNH

### Bảng tổng hợp phương pháp

| Ngành | Method chính | Method phụ | Multiple reference |
|---|---|---|---|
| Sản xuất commodity | EV/EBITDA | DCF (mid-cycle) | 5-10x |
| Bán lẻ ổn định | DCF / FCFF | EV/EBIT | P/E 12-20x |
| Bán lẻ đa chuỗi | SOTP | — | Theo từng chuỗi |
| Ngân hàng | P/B (Residual Income) | Gordon Growth | P/B 1.0-3.0x |
| Chứng khoán | P/B + ROE | P/E scenario | P/B 1.0-2.5x |
| KCN | RNAV | EV/EBITDA | P/RNAV 0.7-1.2x |
| BĐS nhà ở | RNAV / P/NAV | — | Discount 20-40% |
| EPC | EV/EBIT | P/E | P/E 8-14x |
| Đa ngành | SOTP | NAV | Conglomerate disc. 15-25% |
| Logistics/Cảng | EV/EBITDA | DCF | EV/EBITDA 8-14x |
| Dược phẩm | P/E | EV/EBITDA | P/E 14-25x |
| Công nghệ/IT | EV/Revenue hoặc P/E | DCF | P/E 15-30x |
| FMCG | EV/EBITDA | DCF | P/E 15-25x |
| Khoáng sản | EV/EBITDA + NAV | — | EV/EBITDA 6-12x |
| Pre-profit | EV/Revenue hoặc GMV | — | Context-driven |

### Quy trình định giá

```
NGUYÊN TẮC: Mỗi phương pháp định giá tự tính tại sheet nguồn, Sheet 07 chỉ tham chiếu.

1. Xác định giai đoạn chu kỳ → chọn method phù hợp

2. Tại từng sheet nguồn, tạo ô tính giá:
   - P/E:  Sheet 04 (PnL) → Giá = EPS × P/E target
   - P/B:  Sheet 05 (CDKT) → Giá = BVPS × P/B target
   - EV/EBITDA: Sheet 06 (LCTT) → EV = EBITDA × multiple; Giá = (EV - Nợ ròng)/CP
   - DCF:  Sheet 06 (LCTT) → Giá = FCFF / (WACC - g) / Số CP

3. Sheet 07 tham chiếu các ô đó, nhân trọng số → Weighted Target Price:
   Weighted TP = Σ(Method Price × Weight) / Σ Weights

4. So sánh với peer multiples:
   - Nếu target price từ DCF cho multiple cao hơn peer → đắt
   - Nếu lower → rẻ

5. Sensitivity 2 biến (table 5×5):
   - Biến 1: Giả định quan trọng nhất (WACC, growth, margin)
   - Biến 2: Giả định quan trọng thứ 2

6. Kịch bản Bear/Base/Bull:
   Bear: worst case assumptions (giá hàng hóa thấp, biên thu hẹp)
   Base: consensus/mid-cycle assumptions
   Bull: best case (giá cao, biên mở rộng, catalyst xảy ra)

7. Weighted average target price = Bear×P + Base×P + Bull×P
```

---

## 8. QUY CHUẨN EXCEL MODEL

### NGÔN NGỮ — BẮT BUỘC

> **Tất cả text trong Excel PHẢI dùng tiếng Việt đầy đủ dấu.** Không được dùng ASCII không dấu.

```
✅ ĐÚNG: "Doanh thu thuần", "Lợi nhuận gộp", "Phải thu khách hàng"
❌ SAI:  "Doanh thu thuan", "Loi nhuan gop", "Phai thu khach hang"

✅ ĐÚNG: "Giả định", "Tầng phân tích", "Kịch bản cơ sở"
❌ SAI:  "Gia dinh", "Tang phan tich", "Kich ban co so"
```

### MÀU SẮC — 2 HỆ THỐNG

#### Hệ 1: Màu cấu trúc (dùng cho tất cả sheet)

| Màu nền | Màu chữ | Ý nghĩa | Hex bg / hex font |
|---|---|---|---|
| Xanh navy đậm | Trắng | Tiêu đề sheet, banner chính | `1F3864` / `FFFFFF` |
| Xanh medium | Trắng | Section header, sub-banner | `2E75B6` / `FFFFFF` |
| Xanh nhạt | Đen đậm | Dòng subtotal / tổng hợp | `D6E4F0` / `000000` |
| Vàng | Xanh dương | Input cứng (hardcoded — user nhập) | `FFFF00` / `0000FF` |
| Trắng | Đen | Formula tính toán | `FFFFFF` / `000000` |
| Xám nhạt | Đen | Dòng xen kẽ (zebra stripe) | `F2F2F2` / `000000` |

#### Hệ 2: Màu đánh giá định tính (dùng cho cells chứa nhận định)

**Áp dụng bắt buộc** cho: cột đánh giá trong Risk Matrix, PESTLE, Moat, Porter 5F, Leading Indicators, bất kỳ ô nào chứa từ đánh giá định tính.

| Đánh giá | Màu nền | Màu chữ | Hex bg / hex font | Ví dụ từ |
|---|---|---|---|---|
| **Tích cực / Tốt / Mạnh** | Xanh lá nhạt | Xanh lá đậm | `E2EFDA` / `00B050` | TÍCH CỰC, TỐT, MẠNH, THẤP (rủi ro), CAO (lợi thế) |
| **Trung lập / Trung bình** | Vàng nhạt | Nâu vàng | `FFEB9C` / `9C6500` | TRUNG LẬP, TRUNG BÌNH, VỪA |
| **Tiêu cực / Xấu / Yếu** | Đỏ nhạt | Đỏ đậm | `FFC7CE` / `9C0006` | TIÊU CỰC, XẤU, YẾU, CAO (rủi ro), THẤP (lợi thế) |
| **Thông tin / Ghi chú** | Xanh nhạt | Xanh | `DDEEFF` / `0070C0` | Ghi chú, chú thích, nguồn |

```python
# Code mẫu Python/openpyxl cho màu đánh giá:
CGL_BG = "FFE2EFDA";  CGL_FT = "FF00B050"  # Tích cực
CYL_BG = "FFFFEB9C";  CYL_FT = "FF9C6500"  # Trung lập  
CRL_BG = "FFFFC7CE";  CRL_FT = "FF9C0006"  # Tiêu cực
CBL_BG = "FFDDEEFF";  CBL_FT = "FF0070C0"  # Ghi chú

def QUAL(ws, r, col, text, sentiment):
    """sentiment: 'pos' | 'neu' | 'neg' | 'info'"""
    colors = {
        'pos': (CGL_BG, CGL_FT),
        'neu': (CYL_BG, CYL_FT),
        'neg': (CRL_BG, CRL_FT),
        'info': (CBL_BG, CBL_FT),
    }
    bg, ft = colors[sentiment]
    C(ws, r, col, text, bg=bg, fc=ft, bold=True, ha="center", bdr=True)
```

#### Áp dụng màu đánh giá theo sheet

| Sheet | Nơi áp dụng màu đánh giá |
|---|---|
| `02_Giả định` | Cột "Kịch bản": Bear=đỏ, Base=vàng, Bull=xanh |
| `03_Thị trường` | Cột "Đánh giá" Porter 5F: Cao/TB/Thấp |
| `04_Moat` | Cột "Mức độ moat": Mạnh/TB/Yếu |
| `07_Định giá` | Cột kịch bản: Bear=đỏ, Base=vàng, Bull=xanh |
| `09_Nhạy cảm` | Heatmap: ô giá cao=xanh, thấp=đỏ (conditional format) |
| `11_PESTLE` | Cột "Chiều tác động": Tích cực/Trung lập/Tiêu cực |
| `12_Rủi ro` | Cột "Xác suất" và "Tác động": Cao=đỏ, TB=vàng, Thấp=xanh |
| `13_Theo dõi` | 3 cột ngưỡng: Tốt=xanh, Trung lập=vàng, Xấu=đỏ; Downgrade header=đỏ, Upgrade header=xanh |

### Quy tắc formula bắt buộc

```python
# ❌ SAI — hardcode số vào formula
=B7*1.15

# ✅ ĐÚNG — reference Assumptions sheet
=B7*(1+Assumptions!$B$12)

# ❌ SAI — số âm dùng dấu trừ
-1,234

# ✅ ĐÚNG — số âm dùng ngoặc đơn
(1,234)

# ❌ SAI — zero hiển thị 0
0

# ✅ ĐÚNG — zero hiển thị dấu gạch
- (dùng format #,##0;(#,##0);-)
```

### Kiểm tra sau khi build

```bash
cd /sessions/.../mnt/.claude/skills/xlsx
python3 scripts/recalc.py "<path>"
# Target: {"status": "success", "total_errors": 0}
```

**Lỗi thường gặp**:
- `#REF!` → range bị vỡ khi copy formula
- `#DIV/0!` → mẫu số = 0, dùng `IFERROR` hoặc `IF(denom=0,0,...)`
- `#VALUE!` → ô chứa text nhưng formula dùng như số
- `#NAME?` → ô text bắt đầu bằng `=` (bỏ dấu `=`)

---

## 9. QUY CHUẨN DOCX REPORT

### Tên file
```
[TICKER]_Phan_Tich_[YYYY-MM].docx
Ví dụ: KSV_Phan_Tich_2026-06.docx
```

### Cấu trúc bắt buộc (15-20 trang)

| Trang | Nội dung | Yêu cầu |
|---|---|---|
| 1 | Cover | Logo, ticker, tên, ngày, analyst |
| 2 | Investment Summary | Standalone — đọc 1 trang đủ hiểu thesis |
| 3-4 | Chuỗi giá trị & Mô hình KD | Sơ đồ + text giải thích |
| 5-6 | Thị trường & Vị trí cạnh tranh | Porter 5F bảng, chu kỳ ngành |
| 7-8 | Moat & Quản trị | Rating moat, ROIC vs peer |
| 9-11 | Phân tích tài chính | Bảng lịch sử 5 năm, biểu đồ trend |
| 12-14 | Định giá | Method, kịch bản, sensitivity chart |
| 15-16 | Rủi ro & Catalysts | Bảng risk matrix, timeline catalyst |
| 17 | Kết luận & Leading indicators | BUY/HOLD/SELL + checklist theo dõi |
| 18+ | Phụ lục | BCTC tóm tắt, nguồn dữ liệu |

### Nguyên tắc viết

- **Tiếng Việt hoàn toàn** (giữ tiếng Anh cho thuật ngữ chuyên môn không có từ VN tương đương)
- **Mỗi section = 1 kết luận rõ ràng** — không viết "điều này có thể ảnh hưởng tích cực hoặc tiêu cực"
- **Số liệu phải có nguồn và ngày** — ghi cuối trang hoặc phụ lục
- **Tối thiểu 3 biểu đồ**: Revenue/LNST trend, Margin trend, Sensitivity heatmap
- **Kết luận phải có rating**: BUY / HOLD / SELL — không được mơ hồ

### Quy chuẩn bảng biểu (BẮT BUỘC — áp dụng cho mọi bảng trong DOCX)

**Vấn đề cần tránh tuyệt đối**: Chữ và số không được chồng lấn nhau, text không được bị cắt.

#### Canh lề trong bảng
```
Cột mô tả / tên chỉ tiêu (cột đầu tiên):   Căn TRÁI
Cột số liệu tài chính (%, tỷ đồng, VNĐ):   Căn PHẢI
Cột tiêu đề (header row):                   Căn GIỮA
Cột đánh giá định tính (Cao/TB/Thấp...):    Căn GIỮA
```

#### Độ rộng cột
```
□ Không dùng độ rộng cố định cho tất cả cột — mỗi cột có width riêng phù hợp nội dung
□ Cột chứa text dài (mô tả rủi ro, chiến lược): width ≥ 5 cm
□ Cột số liệu năm (2022A, 2023A...): width 2-2.5 cm mỗi cột
□ Cột "Chỉ tiêu" / "Hạng mục": width 4-5 cm
□ Tổng chiều rộng bảng ≤ chiều rộng trang (tối đa ~16 cm cho A4 margin chuẩn)
```

#### Xử lý text dài
```
□ Bật wrap text (tự động xuống dòng) cho tất cả ô — KHÔNG cắt ngắn nội dung
□ Nếu bảng có nhiều cột → chia thành 2 bảng nhỏ, mỗi bảng một chủ đề
□ Tiêu đề cột dài → dùng viết tắt + chú thích bên dưới bảng
□ Không dùng font size < 8pt trong bảng
```

#### Khoảng cách
```
□ Cell padding: top/bottom ≥ 2pt, left/right ≥ 4pt
□ Khoảng cách giữa bảng và đoạn văn trên/dưới: 6pt
□ Không để bảng bị vỡ trang ở giữa một hàng dữ liệu quan trọng
```

#### Định dạng số trong bảng
```
□ Số tỷ đồng: dùng dấu phẩy ngăn hàng nghìn (3,521 tỷ)
□ Phần trăm: 1 chữ số thập phân (15.5%), không phải 15.4567%
□ EPS, BVPS: định dạng #,##0 VNĐ
□ Tỷ số (P/E, EV/EBITDA): 1 chữ số thập phân (8.3x)
□ Số âm: dùng dấu ngoặc (1,234) hoặc dấu trừ tùy context — nhất quán trong toàn file
```

#### Công cụ build DOCX
- Dùng **python-docx** (library) để tạo file .docx
- Thiết lập `table.autofit = False` và set width tuyệt đối cho từng cột
- Dùng `WD_TABLE_ALIGNMENT.CENTER` cho toàn bảng, alignment riêng từng cell
- Font chữ: Times New Roman hoặc Arial 10-11pt cho body, 9pt cho bảng
- Tiêu đề section dùng Heading 1/2/3 style của Word (không dùng bold text thông thường)

---

## 10. DANH SÁCH CỔ PHIẾU ĐÃ PHÂN TÍCH

| Ticker | Tên | Ngành | File Excel | File PDF | Ngày phân tích |
|---|---|---|---|---|---|
| HPG | Hòa Phát Group | Thép | `01_HPG_Thep_Model.xlsx` | — | Template |
| MWG | Thế Giới Di Động | Bán lẻ | `02_BanLe_MWG_PNJ_Model.xlsx` | — | Template |
| MWG | Thế Giới Di Động | SOTP 3 chuỗi | `09_MWG_SOTP_Model.xlsx` | — | Template |
| TCB | Techcombank | Ngân hàng | `03_NganHang_TCB_ACB_Model.xlsx` | — | Template |
| ACB | ACB | Ngân hàng | `03_NganHang_TCB_ACB_Model.xlsx` | — | Template |
| SSI | SSI Securities | Chứng khoán | `04_ChungKhoan_SSI_Model.xlsx` | — | Template |
| NTC | Nam Tân Uyên | KCN | `05_KhuCongNghiep_NTC_SZC_Model.xlsx` | — | Template |
| SZC | Sonadezi Châu Đức | KCN | `05_KhuCongNghiep_NTC_SZC_Model.xlsx` | — | Template |
| NLG | Nam Long | BĐS | `06_BDS_NLG_Model.xlsx` | — | Template |
| MSN | Masan Group | Đa ngành | `07_MSN_Conglomerate_Model.xlsx` | — | Template |
| BMP | Bình Minh Plastic | Sản xuất | `08_SanXuat_BMP_TLG_TV2_PC1_Model.xlsx` | — | Template |
| KSV | Vimico | Khoáng sản | `KSV_Model_2026-06.xlsx` | `KSV_Phan_Tich_2026-06.pdf` | 2026-06 |
| PNJ | PNJ | Trang sức bán lẻ | `PNJ_Model_2026-06.xlsx` | `PNJ_Phan_Tich_2026-06.pdf` | 2026-06 |
| DHC | Đông Hải Bến Tre | Giấy & Bao bì | `DHC_Model_2026-06.xlsx` | `DHC_Phan_Tich_2026-06.pdf` | 2026-06 |

> **Note**: Từ DHC trở đi, output báo cáo chuyển sang DOCX (`.docx`). KSV và PNJ vẫn là PDF legacy.
> Các file Template chưa có dữ liệu thực — cần điền BCTC thực tế khi phân tích sâu từng cổ phiếu.

---

## 11. CHECKLIST KIỂM TRA TRƯỚC KHI GIAO

### Dữ liệu
```
□ Giá cổ phiếu là ngày phân tích (không phải training data cũ)
□ Số CP lưu hành chính xác (không phải vốn điều lệ / 10,000)
□ Vốn hóa = Giá × Số CP lưu hành (tính lại, không copy)
□ BCTC là năm tài chính gần nhất đã kiểm toán

[KIỂM TRA A/E — BẮT BUỘC]
□ Đã xác định ngày phân tích → tính đúng năm Actual gần nhất (xem Mục 2.5)
□ Mọi năm đã kết thúc VÀ đã công bố BCTC → gán nhãn A, KHÔNG phải E
□ Snapshot DOCX và Excel dùng cùng nhãn A/E (không được lệch nhau)
□ Không có ô nào gán A nhưng dùng số ước tính (phải là số thực từ BCTC)
□ Nếu BCTC chưa tìm được: ghi "[chưa xác nhận]", KHÔNG điền ước tính

□ Kế hoạch công ty lấy từ ĐHĐCĐ năm hiện tại (không phải năm trước)
□ Giá hàng hóa / chỉ số ngành là hiện tại (nếu relevant)
□ Đã tìm báo cáo từ ≥ 2 CTCK để cross-check số liệu và consensus
```

### Excel model
```
□ Tính lại (Recalc) = 0 lỗi
□ Tất cả giả định nằm trong sheet Giả định
□ Bảng cân đối kế toán cân bằng (Tổng TS = Tổng NV)
□ FCFF > 0 ở kịch bản cơ sở (hoặc có giải thích nếu âm)
□ Bảng nhạy cảm chạy đúng chiều
□ So sánh cùng ngành có ít nhất 3-5 công ty
```

### DOCX report
```
□ Investment Summary đứng độc lập được (đọc 1 trang hiểu thesis)
□ Có BUY / HOLD / SELL rõ ràng với target price
□ Bear/Base/Bull có xác suất cộng lại = 100%
□ Có leading indicators cụ thể (không chung chung)
□ Nguồn dữ liệu được ghi
□ Ngày phân tích ở cover page
□ Tất cả bảng: số liệu căn phải, mô tả căn trái, header căn giữa
□ Không có text/số bị chồng lấn hoặc bị cắt trong bảng
□ Tiếng Việt đầy đủ dấu (không thiếu dấu thanh, dấu mũ)
```

### Tên file
```
Excel: [TICKER]_Model_[YYYY-MM].xlsx
DOCX:  [TICKER]_Phan_Tich_[YYYY-MM].docx
```

---

## PHỤ LỤC: GHI CHÚ KỸ THUẬT

### Môi trường
```bash
# Python packages
pip install openpyxl python-docx matplotlib pandas --break-system-packages

# Recalc formula
cd /sessions/stoic-intelligent-ramanujan/mnt/.claude/skills/xlsx
python3 scripts/recalc.py "/path/to/file.xlsx"

# Skills path
/sessions/stoic-intelligent-ramanujan/mnt/.claude/skills/
├── xlsx/   — Excel skill + recalc script
├── pdf/    — PDF skill
└── docx/   — Word skill
```

### Nguồn dữ liệu VN — Tham chiếu nhanh

```
BCTC gốc (ưu tiên 1):   ir.[company].com.vn / hsx.vn công bố thông tin
BCTC tổng hợp:           finance.vietstock.vn/[TICKER]/financials.htm
Giá CP + KLGD:           simplize.vn/co-phieu/[TICKER]
Tin tức + ĐHĐCĐ:         cafef.vn/[ticker].chn, dnse.com.vn, tinnhanhchungkhoan.vn
Kế hoạch công ty:        Search "[TICKER] ĐHĐCĐ [năm] kế hoạch tài liệu"

Báo cáo CTCK (≥ 2 nguồn bắt buộc):
  SSI:    research.ssi.com.vn / "[TICKER] SSI Research [năm]"
  VCSC:   vcsc.com.vn/research / "[TICKER] VCSC báo cáo [năm]"
  VCBS:   vcbs.com.vn / "[TICKER] VCBS [năm]"
  HSC:    hsc.com.vn/research / "[TICKER] HSC Research [năm]"
  MBS:    mbs.com.vn/research / "[TICKER] MBS Research [năm]"
  VDSC:   vdsc.com.vn / "[TICKER] VDSC Rồng Việt [năm]"
  BSC:    "[TICKER] BSC BIDV Securities [năm]"
  KIS:    "[TICKER] KIS Vietnam [năm]"
  Yuanta: "[TICKER] Yuanta báo cáo [năm]"

Giá hàng hóa:    LME, tradingeconomics.com, metal.com
SCFI/cước tàu:   "[SCFI index tháng năm]", drewry.co.uk
FDI/XNK VN:      gso.gov.vn, mpi.gov.vn
THC/Hàng hải:    vinamarine.gov.vn
```

---

## 12. Bổ sung: Median Valuation & Negative Earnings Handling

### 12.1 Median cho định giá (bắt buộc — KHÔNG dùng average)

```
Mọi chỉ số định giá dùng MEDIAN thay vì AVERAGE:
  • P/B median toàn bộ lịch sử → target PB multiple
  • P/E median all-time → tham chiếu (banking dùng P/B là chính)
  • EV/EBITDA median all-time → cho ngành phi ngân hàng
  
Lý do: median resistant với outlier (quý LNST âm, quý đặc biệt)
```

### 12.2 Median theo năm

```
Với mỗi năm tài chính:
  PE_nam = median(PE_Q1, PE_Q2, PE_Q3, PE_Q4)
  PB_nam = median(PB_Q1, PB_Q2, PB_Q3, PB_Q4)
  
Dùng để so sánh YoY trong bảng chỉ tiêu chính.
```

### 12.3 Xử lý LNST âm (bắt buộc)

```
Quý nào LNST âm:
  1. P/E quý đó = P/E của quý trước đó (carry forward)
  2. Đánh dấu là "P/E điều chỉnh — LNST âm"
  3. Trên biểu đồ: vẽ bằng marker khác (hollow circle, màu đỏ)
  4. PB và EV/EBITDA không thay đổi (vẫn dùng giá trị gốc)
  
Lý do: P/E khi LNST âm là vô nghĩa (âm), carry giúp giữ tính liên tục
```

### 12.4 Công thức Excel cross-sheet ref

```
Sử dụng hệ thống reg_row() + ref() để đảm bảo mọi ô là formula:

  def reg_row(sheet, key, row):    # đăng ký row cho 1 chỉ tiêu
      _SHEET_ROWS[f"{sheet}|{key}"] = row
      
  def ref(sheet, key, col):        # tạo cross-sheet ref
      return f"'{sheet}'!{col_letter}{row}"
      
  def write_formula_row(ws, r, c, label, vals, fmt):
      # Viết row với formula strings (bắt đầu bằng '=')
      # vals có thể là: số (hardcode), string bắt đầu '=' (formula)
```

Tất cả forecast cells trong Excel PHẢI dùng `ref()` để tham chiếu chéo sheet. Không hardcode số forecast.

---

*Tài liệu này được cập nhật liên tục theo kinh nghiệm phân tích thực tế. Version 2.4 — 06/2026.*
*Thay đổi v2.1: (1) Chuyển output báo cáo từ PDF → DOCX; (2) Thêm quy chuẩn bảng biểu (căn lề, độ rộng cột, wrap text); (3) Thêm Sheet 13 — Các yếu tố cần theo dõi vào cấu trúc Excel.*
*Thay đổi v2.2: (4) Thêm Mục 2.5 — Quy tắc bắt buộc phân kỳ Actual (A) vs Estimated (E); (5) Mở rộng Mục 2.3 nguồn dữ liệu; (6) Cập nhật checklist Section 11 — bổ sung kiểm tra A/E và yêu cầu ≥ 2 CTCK.*
*Thay đổi v2.3: (7) Thêm OUTPUT BẮT BUỘC box sau mỗi trong 6 tầng phân tích; (8) Assembly Guide — Pre-flight checklist, Build Order bắt buộc, Investment Summary template có ví dụ GMD, Phụ lục bắt buộc.*
