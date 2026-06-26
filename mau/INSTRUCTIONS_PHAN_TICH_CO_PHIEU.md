# INSTRUCTIONS: HỆ THỐNG PHÂN TÍCH CỔ PHIẾU VIỆT NAM
> Tài liệu vận hành nội bộ — ITE GROUP / Vnstockmarket.com  
> Phiên bản: 2.0 | Cập nhật: 2026-06

---

## ĐỌC TRƯỚC: TRIẾT LÝ CỐT LÕI

**Tài liệu này là bộ QUY TẮC VÀ QUY TRÌNH BẮT BUỘC TUÂN THỦ 100% — KHÔNG phải hướng dẫn tham khảo, tuyệt đối KHÔNG được tự ý bỏ qua, bỏ bước hay thay đổi.**

Các cổ phiếu mẫu đã phân tích (HPG, MWG, KSV, TCB, SSI, NLG, MSN...) là **ví dụ minh họa và tiêu chuẩn kỹ thuật** bắt buộc phải đạt được. Khi nhận một cổ phiếu mới, AI phải:

1. **Tuân thủ đúng và đủ toàn bộ các bước trong tài liệu này** — không được bỏ sót bất kỳ bước nào trong quy trình phân tích.
2. **Gọi API Vietcap lấy đầy đủ dữ liệu tài chính** — bắt buộc lấy toàn bộ dữ liệu lịch sử theo cả Năm (tối thiểu 5 năm gần nhất) và Quý (tối thiểu 20 quý gần nhất) để thực hiện tính toán.
3. **Thu thập dữ liệu mới nhất** — luôn dùng ngày hiện tại, không dùng số liệu cũ từ training.

> **Quy tắc vàng**: "Hiểu cách công ty kiếm tiền, tại sao họ kiếm được nhiều hơn đối thủ, và điều gì có thể phá vỡ điều đó" — tuân thủ nghiêm ngặt các bước phân tích để làm rõ điều này.

---

## MỤC LỤC

1. [Quy trình tổng thể](#1-quy-trình-tổng-thể)
2. [Bước 1: Thu thập dữ liệu real-time](#2-bước-1-thu-thập-dữ-liệu-real-time)
3. [Bước 2: Phân tích framework 6 tầng](#3-bước-2-phân-tích-framework-6-tầng)
4. [Phân tích PESTLE — Môi trường vĩ mô](#4-phân-tích-pestle--môi-trường-vĩ-mô)
5. [Bước 3: Build Excel model](#5-bước-3-build-excel-model)
6. [Bước 4: Viết DOCX research report](#6-bước-4-viết-docx-research-report)
7. [Phương pháp dự báo doanh thu theo ngành](#7-phương-pháp-dự-báo-doanh-thu-theo-ngành)
8. [Phương pháp định giá theo ngành](#8-phương-pháp-định-giá-theo-ngành)
9. [Quy chuẩn Excel model](#9-quy-chuẩn-excel-model)
10. [Quy chuẩn DOCX report](#10-quy-chuẩn-docx-report)
11. [Danh sách cổ phiếu đã phân tích](#11-danh-sách-cổ-phiếu-đã-phân-tích)
12. [Checklist kiểm tra trước khi giao](#12-checklist-kiểm-tra-trước-khi-giao)

---

## 1. QUY TRÌNH TỔNG THỂ

Mỗi lần nhận yêu cầu phân tích một cổ phiếu, AI thực hiện đúng thứ tự sau:

```
NHẬN TICKER
    │
    ▼
BƯỚC 1: THU THẬP DỮ LIỆU REAL-TIME
│  ├─ Agent A: Dữ liệu thị trường (giá, vốn hóa, số CP, thanh khoản)
│  ├─ Agent B: KQKD lịch sử (4-5 năm: DT, LN, biên, B/S, CF)
│  ├─ Agent C: Tin tức, chiến lược, kế hoạch, rủi ro gần nhất
│  └─ Agent D: Peer comparison và ngành
│
▼
XÁC ĐỊNH BẢN CHẤT KINH DOANH
│  ├─ Ngành → chọn phương pháp dự báo doanh thu
│  └─ Chu kỳ → chọn phương pháp định giá
│
▼
BƯỚC 2: PHÂN TÍCH 6 TẦNG
│  Tầng 1 → 6 (xem chi tiết Mục 3)
│
▼
BƯỚC 3: BUILD EXCEL MODEL
│  ├─ Quyết định: template có sẵn hay build mới?
│  ├─ Điền dữ liệu lịch sử đã thu thập
│  ├─ Build revenue model từ drivers
│  ├─ Chạy recalc → 0 errors
│  └─ Sensitivity analysis
│
▼
BƯỚC 4: VIẾT DOCX REPORT (~15-20 trang)
│  Theo cấu trúc chuẩn (Mục 10)
│
▼
GIAO 2 FILE:
├─ [TICKER]_Phan_Tich_[YYYY-MM].docx
└─ [TICKER]_Model_[YYYY-MM].xlsx
```

**Output bắt buộc**: Mỗi phân tích PHẢI cho ra **đúng 2 file**:
1. **DOCX** — research report 15-20 trang, tiếng Việt, chuyên nghiệp
2. **Excel** — financial model đầy đủ sheets, 0 formula errors

---

## 2. BƯỚC 1: THU THẬP DỮ LIỆU REAL-TIME

### 2.1 Nguyên tắc bắt buộc

- **Bắt buộc gọi API Vietcap để lấy đầy đủ dữ liệu**: Khi phân tích bất kỳ cổ phiếu nào, bước đầu tiên bắt buộc phải sử dụng các API/hàm Vietcap có sẵn để tải toàn bộ dữ liệu tài chính lịch sử. Không được phép tự ước lượng hoặc bỏ qua.
- **Tải đủ dữ liệu theo năm và quý**: Dữ liệu tài chính tải về phải đầy đủ theo cả Năm (tối thiểu 5 năm gần nhất) và Quý (tối thiểu 20 quý gần nhất).
- **Bắt buộc lấy Thị giá, Số lượng CP lưu hành và Vốn hóa động từ API Vietcap**: Khi phân tích, bắt buộc phải gọi API chi tiết của Vietcap (`https://trading.vietcap.com.vn/api/iq-insight-service/v1/company/details?ticker={ticker}`) để lấy chính xác thị giá hiện tại (`currentPrice`), số lượng cổ phiếu đang lưu hành (`numberOfSharesMktCap`), và vốn hóa thị trường (`marketCap`). Tuyệt đối KHÔNG được hardcode các thông số này trong mã nguồn phân tích vì số lượng cổ phiếu lưu hành và giá có thể thay đổi liên tục (chia tách, phát hành thêm, biến động thị trường). Nếu hardcode sẽ dẫn đến tính toán P/E và P/B hiện tại bị sai lệch nghiêm trọng.
- **Luôn dùng ngày hiện tại** — không dùng số liệu training của AI
- **Thu thập song song** bằng nhiều agent (hoặc nhiều search queries đồng thời) để tiết kiệm thời gian
- **Kết hợp 3 nguồn** theo thứ tự ưu tiên: API Vietcap / BCTC gốc > Cafef/Vietstock > Tin tức/báo cáo
- **Kiểm tra `references/` trước** — nếu ngành đang phân tích có file trong `references/[ten_nganh]/`, đọc file đó TRƯỚC KHI search online. File do Huy cung cấp có độ tin cậy cao hơn nguồn tìm kiếm.

### 2.2 Danh sách dữ liệu cần thu thập

#### Nhóm A — Dữ liệu thị trường (real-time)
```
□ Giá cổ phiếu hiện tại (ngày phân tích)
□ Giá 52 tuần cao / thấp
□ Số lượng cổ phiếu đang lưu hành (Outstanding shares)
□ Vốn hóa thị trường = Giá × Số CP
□ KLGD bình quân 30 ngày (thanh khoản)
□ Tỷ lệ sở hữu nước ngoài (foreign ownership)
□ Cơ cấu cổ đông lớn (>5%)
□ Room ngoại còn lại
```

#### Nhóm B — Dữ liệu tài chính lịch sử (4-5 năm)
```
Kết quả kinh doanh:
□ Doanh thu thuần
□ Lợi nhuận gộp + biên GP
□ Chi phí bán hàng, QLDN
□ EBIT + biên EBIT
□ EBITDA (EBIT + D&A)
□ Lợi nhuận trước thuế
□ Lợi nhuận sau thuế (LNST)
□ LNST thuộc về cổ đông công ty mẹ
□ EPS (pha loãng)
□ DPS (cổ tức đã trả)

Cân đối kế toán:
□ Tổng tài sản
□ Tiền và tương đương tiền
□ Hàng tồn kho
□ Khoản phải thu
□ Tài sản cố định (TSCĐ)
□ Tổng nợ vay (ngắn hạn + dài hạn)
□ Nợ ròng (Net Debt = Nợ vay - Tiền)
□ Vốn chủ sở hữu (VCSH)
□ Book value per share

Dòng tiền:
□ Dòng tiền hoạt động (Operating CF)
□ Capex
□ FCFF / FCFE
□ Dòng tiền tài chính

Các chỉ số:
□ ROE, ROA, ROIC
□ D/E ratio, Net Debt/EBITDA
□ DSO (Days Sales Outstanding)
□ Inventory Turnover
```

#### Nhóm C — Dữ liệu định tính
```
□ Mô hình kinh doanh, sản phẩm/dịch vụ chính
□ Kế hoạch năm hiện tại (từ ĐHĐCĐ gần nhất)
□ Chiến lược trung hạn (3-5 năm)
□ Rủi ro đặc thù được CEO/CFO đề cập
□ Tin tức 6 tháng gần nhất có tác động đến thesis
□ Analyst reports từ CTCK VN nếu có
```

#### Nhóm D — Dữ liệu ngành và peer
```
□ Danh sách 3-5 công ty cùng ngành (peers)
□ Multiples của peers (P/E, EV/EBITDA, P/B)
□ Tăng trưởng ngành (TAM growth)
□ Giá hàng hóa liên quan (nếu là ngành commodity)
□ Dữ liệu macro liên quan (lãi suất, tỷ giá, GDP ngành)
```

### 2.3 Nguồn dữ liệu và cách truy cập

#### Nhóm 1 — Dữ liệu tài chính và thị trường

| Nguồn | Dữ liệu | URL / Cách truy cập | Ưu tiên |
|---|---|---|---|
| **Cafef.vn** | KQKD, BCTC tóm tắt, tin tức, ĐHĐCĐ | `cafef.vn/[ticker].chn` / Search | ⭐⭐⭐⭐⭐ |
| **Vietstock Finance** | BCTC chi tiết 5 năm, chỉ số tài chính | `finance.vietstock.vn/[TICKER]/financials.htm` | ⭐⭐⭐⭐⭐ |
| **Simplize.vn** | Giá CP, KLGD, chỉ số kỹ thuật | `simplize.vn/co-phieu/[TICKER]` | ⭐⭐⭐⭐ |
| **HOSE / HNX** | BCTC PDF gốc kiểm toán, nghị quyết ĐHĐCĐ | `hsx.vn` → Tra cứu → Công bố thông tin | ⭐⭐⭐⭐⭐ |
| **IR website doanh nghiệp** | BCTC gốc, thuyết minh, kế hoạch chiến lược | `ir.[tên công ty].com.vn` hoặc `[tên công ty].com.vn/quan-he-co-dong` | ⭐⭐⭐⭐⭐ |
| **DNSE / Fili** | Dữ liệu lịch sử giá, dividend history | `dnse.com.vn`, `fili.vn` | ⭐⭐⭐ |
| **CafeF Doanh nghiệp** | Hồ sơ doanh nghiệp, cổ đông lớn | Search `[TICKER] cổ đông lớn site:cafef.vn` | ⭐⭐⭐⭐ |

#### Nhóm 2 — Báo cáo phân tích CTCK (BẮT BUỘC tìm ≥ 2 CTCK)

Mỗi phân tích PHẢI tìm báo cáo từ ít nhất 2 công ty chứng khoán để có consensus và cross-check số liệu.

| CTCK | Đặc điểm | URL / Search query |
|---|---|---|
| **SSI Research** | Phổ rộng, uy tín cao, cập nhật thường xuyên | `research.ssi.com.vn` / Search `[TICKER] SSI Research [năm]` |
| **VCI / VCSC** | Phân tích sâu mid-cap, consumer | `vcsc.com.vn/research` / Search `[TICKER] VCSC báo cáo [năm]` |
| **VCBS** | Ngân hàng và tài chính tốt | `vcbs.com.vn/phan-tich` / Search `[TICKER] VCBS [năm]` |
| **HSC (HoSE Securities)** | Phân tích kỹ thuật + cơ bản | `hsc.com.vn/research` / Search `[TICKER] HSC Research [năm]` |
| **MBS Research** | Cảng biển, công nghiệp, BĐS tốt | `mbs.com.vn/research` / Search `[TICKER] MBS [năm]` |
| **VDSC (Rồng Việt)** | Thép, khoáng sản, ngân hàng | `vdsc.com.vn` / Search `[TICKER] VDSC báo cáo phân tích` |
| **KIS Vietnam** | Logistics, công nghiệp | Search `[TICKER] KIS Research [năm]` |
| **Mirae Asset** | Bán lẻ, tiêu dùng | Search `[TICKER] Mirae Asset [năm]` |
| **Yuanta Vietnam** | Đa ngành | Search `[TICKER] Yuanta báo cáo [năm]` |
| **BSC (BIDV Securities)** | SOE, ngân hàng, hạ tầng | Search `[TICKER] BSC báo cáo phân tích [năm]` |
| **VPS Securities** | Đa ngành, small-cap | Search `[TICKER] VPS Research [năm]` |

> **Search query template chuẩn để tìm báo cáo CTCK**:
> ```
> "[TICKER] báo cáo phân tích [SSI/VCSC/MBS/HSC/VDSC] [năm hiện tại]"
> "[TICKER] research report [CTCK] [năm]"
> "[TICKER] target price [tên CTCK]"
> "[TICKER] khuyến nghị [tên CTCK] [năm]"
> ```

#### Nhóm 3 — Tin tức và dữ liệu định tính

| Nguồn | Dữ liệu | Cách truy cập |
|---|---|---|
| **Tinnhanhchungkhoan.vn** | Tin tức cổ phiếu, phân tích ngắn | Search `[TICKER] site:tinnhanhchungkhoan.vn` |
| **Nhịp cầu đầu tư** | Phân tích chiến lược, phỏng vấn CEO | Search `[TICKER] site:nhipcaudautu.vn` |
| **VnExpress Kinh tế** | Tin tức vĩ mô, ngành | Search `[TICKER] OR [tên công ty] site:vnexpress.net` |
| **Thanh Niên / Tuổi Trẻ Kinh tế** | Tin tức doanh nghiệp lớn | Search news |
| **ĐHĐCĐ materials** | Kế hoạch năm, chiến lược | Search `[TICKER] ĐHĐCĐ [năm] tài liệu` |

#### Nhóm 4 — Giá hàng hóa và dữ liệu quốc tế (nếu relevant)

| Nguồn | Dữ liệu | URL |
|---|---|---|
| **TradingEconomics** | LME, giá dầu, tỷ giá, GDP | `tradingeconomics.com` |
| **Metal.com / LME** | Giá kim loại thực tế | `metal.com`, `lme.com` |
| **Drewry / Freightos** | Chỉ số cước container (SCFI, WCI) | Search `SCFI index [tháng/năm]` |
| **GSO (Tổng cục thống kê)** | GDP, CPI, XNK, FDI VN | `gso.gov.vn` |
| **Cục Hàng hải VN** | THC, quy định cảng biển | `vinamarine.gov.vn` |
| **MPI (Bộ KHĐT)** | Số liệu FDI, KCN | `mpi.gov.vn` |

#### Thứ tự ưu tiên khi tìm BCTC năm Actual

```
Ưu tiên 1: IR website của chính doanh nghiệp → BCTC PDF gốc, thuyết minh đầy đủ
Ưu tiên 2: HOSE/HNX công bố thông tin → BCTC đã kiểm toán nộp sở
Ưu tiên 3: Vietstock Finance → BCTC tổng hợp, đã parse sẵn
Ưu tiên 4: Cafef → tóm tắt KQKD (ít chi tiết hơn, dùng để verify)
Ưu tiên 5: Báo cáo CTCK → thường có bảng tóm tắt lịch sử trong phần appendix

❌ Không dùng dữ liệu từ training của AI → luôn fetch/search real-time
```

### 2.4 Xử lý sau khi thu thập

1. AI tổng hợp dữ liệu → báo cáo những con số đã tìm được
2. **AI ghi chú rõ các ô chưa tìm được** hoặc không chắc chắn
3. Huy xem qua và xác nhận hoặc cung cấp bổ sung
4. Sau khi verify → chạy model

> **Lưu ý**: Nếu không tìm được số liệu chính xác, AI phải ghi rõ "Chưa có dữ liệu — cần xác nhận" thay vì tự ước tính và không nói.

---

### 2.5 QUY TẮC BẮT BUỘC: Phân kỳ Actual (A) vs Estimated/Forecast (E)

> **Đây là quy tắc không được bỏ qua.** Sai nhãn A/E là lỗi nghiêm trọng — làm sai lệch toàn bộ chất lượng báo cáo.

#### Định nghĩa

```
Năm tài chính X gán nhãn ACTUAL (A) khi ĐỒNG THỜI thỏa mãn:
  ① Năm X đã kết thúc hoàn toàn (31/12/X < ngày phân tích), VÀ
  ② BCTC năm X đã được công bố (thông thường tháng 3–4 của năm X+1)

Năm tài chính X gán nhãn ESTIMATED / FORECAST (E) khi:
  → Năm X chưa kết thúc, HOẶC
  → Năm X kết thúc nhưng BCTC chưa được công bố

Năm X đang diễn ra (hiện tại) luôn là E, trừ khi có BCTC bán niên
đã kiểm toán — khi đó ghi rõ "2026H1A" / "2026H2E".
```

#### Ví dụ áp dụng theo ngày phân tích

| Ngày phân tích | 2022 | 2023 | 2024 | 2025 | 2026 | 2027 |
|---|---|---|---|---|---|---|
| Tháng 1/2025 | A | A | A* | E | E | E |
| Tháng 6/2025 | A | A | A | E | E | E |
| Tháng 1/2026 | A | A | A | A* | E | E |
| **Tháng 6/2026** | **A** | **A** | **A** | **A** | **E** | **E** |

> *Tháng 1: BCTC năm trước có thể chưa kiểm toán xong → phải xác nhận trước khi gán A. Nếu chưa có → gán "E (chưa KT)".

#### Hành động bắt buộc khi bắt đầu phân tích

```
BƯỚC 0 — Trước khi làm bất cứ điều gì:

1. Xác định ngày phân tích (= ngày hôm nay)
2. Tính: Năm Actual gần nhất = năm hiện tại − 1 (nếu BCTC đã có)
3. Kiểm tra: BCTC năm (hiện tại − 1) đã công bố chưa?
   → Chưa: gán "E (chưa KT)", tiếp tục tìm kiếm
   → Rồi: gán A và BẮT BUỘC thu thập số liệu thực tế
4. Snapshot tài chính phải có đúng số cột A theo công thức trên

KHÔNG ĐƯỢC:
  ❌ Gán năm đã kết thúc là E vì không tìm được dữ liệu
  ❌ Dùng số ước tính cho năm A mà không ghi rõ "ước tính"
  ❌ Bỏ qua năm A gần nhất vì ngại tìm kiếm

PHẢI:
  ✅ Tìm BCTC năm A ít nhất 3 nguồn trước khi kết luận "chưa có"
  ✅ Nếu thực sự chưa công bố: ghi "Chưa có BCTC [năm] — cần bổ sung"
  ✅ Snapshot DOCX và Excel PHẢI khớp nhãn A/E với nhau
```

#### Format snapshot chuẩn (ví dụ phân tích tháng 6/2026)

```
           2022A    2023A    2024A    2025A    2026E    2027E
DT (tỷ):   ...      ...      ...      ...      ...      ...
LNST (tỷ): ...      ...      ...      ...      ...      ...
EPS (VNĐ): ...      ...      ...      ...      ...      ...
P/E:       ...      ...      ...      ...      ...      ...
```

> **Nếu không tìm được 2025A**: ghi `[chưa xác nhận]` vào ô — KHÔNG điền số ước tính và gán nhãn A.

---

## 3. BƯỚC 2: PHÂN TÍCH FRAMEWORK 6 TẦNG

### Tầng 1: Chuỗi giá trị ngành

**Mục tiêu**: Vẽ được sơ đồ từ nguyên liệu đầu vào đến khách hàng cuối, xác định công ty đang capture giá trị ở đâu trong chuỗi.

Câu hỏi:
- Ai kiểm soát đầu vào? Chi phí chuyển đổi nhà cung cấp?
- Quy trình tạo ra sự khác biệt ở bước nào?
- Kênh phân phối: trực tiếp hay gián tiếp? Margin từng kênh?
- Khách hàng cuối: cá nhân hay doanh nghiệp? Độ nhạy giá?

> **📄 OUTPUT BẮT BUỘC → DOCX Trang 3–4**
>
> Artifact phải hoàn thành TRƯỚC KHI viết bất kỳ trang nào của DOCX.
>
> **Bảng chuỗi giá trị** (5–7 bước, từ nguyên liệu/đầu vào đến khách hàng cuối):
>
> | Bước | Hoạt động | Công ty có mặt? | Biên ước tính | Ai kiểm soát? |
> |---|---|---|---|---|
> | 1 | [tên bước] | ✅ Có / ❌ Không | ~XX% | [tên bên] |
> | ... | | | | |
>
> **3 câu hỏi bắt buộc phải trả lời được trước khi sang Tầng 2:**
> 1. Công ty đang capture giá trị ở bước nào? Tại sao bước đó có lợi thế hơn bước khác?
> 2. Bước nào trong chuỗi có biên cao nhất — công ty đã/chưa/không thể tiếp cận? Vì sao?
> 3. Rào cản dịch chuyển lên/xuống chuỗi là gì? (Vốn? Giấy phép? Quan hệ khách hàng?)
>
> **Ví dụ — GMD (Cảng biển):**
>
> | Bước | Hoạt động | GMD? | Biên | Ai kiểm soát |
> |---|---|---|---|---|
> | 1 | Vận chuyển biển quốc tế | ❌ | ~3–5% | Hãng tàu (MSC, CMA CGM) |
> | 2 | Xếp dỡ container tại cảng | ✅ | ~35–40% | **GMD** (NĐV, Gemalink) |
> | 3 | Logistics nội địa | ✅ một phần | ~8–12% | GMD Logistics + 3PLs |
> | 4 | Kho bãi, phân phối | ✅ một phần | ~10–15% | GMD + cạnh tranh |
> | 5 | Thông quan, forwarder | ❌ | ~5–8% | CTCP forwarder phân mảnh |
>
> → **Kết luận Tầng 1:** GMD capture bước có biên cao nhất (xếp dỡ ~38%). Không kiểm soát đầu vào (hãng tàu — rủi ro MSC rút cargo) và không có mặt ở đầu ra (forwarder). Đây là moat candidate vì vị trí địa lý và giấy phép cảng không thể nhân bản.


### Tầng 2: Phân tích thị trường

**Công cụ**: Porter 5 Forces + TAM/SAM/SOM + Vị trí chu kỳ ngành

| Lực | Đánh giá | Ý nghĩa với lợi nhuận |
|---|---|---|
| Đối thủ hiện tại | HHI, số lượng player, tăng trưởng ngành | Ngành tập trung = lợi nhuận cao hơn |
| Đối thủ tiềm năng | Rào cản vốn, giấy phép, quy mô | Barriers cao = bảo vệ biên |
| Sản phẩm thay thế | Công nghệ mới, hành vi người dùng | Disruption risk |
| NCC | Số lượng, chi phí chuyển đổi | Ảnh hưởng COGS |
| Khách hàng | Fragmented vs. concentrated | Ảnh hưởng pricing power |

**Chu kỳ ngành**: Emerging / Growth / Maturity / Decline — xác định để chọn valuation method.

> **📄 OUTPUT BẮT BUỘC → DOCX Trang 5–6**
>
> Phải hoàn thành **2 artifacts** trước khi viết trang 5–6:
>
> **Artifact 1 — Bảng Porter 5 Forces:**
>
> | Lực cạnh tranh | Mức độ | Lý do (1–2 câu) | Tác động đến biên LN |
> |---|---|---|---|
> | Đối thủ hiện tại | Cao/TB/Thấp | [dẫn chứng cụ thể] | [↑↓ biên bao nhiêu %] |
> | Đối thủ tiềm năng | | | |
> | Sản phẩm thay thế | | | |
> | Nhà cung cấp (NCC) | | | |
> | Khách hàng | | | |
> | **Kết luận Porter** | **Ngành hấp dẫn / TB / Kém** | | |
>
> **Artifact 2 — Vị trí chu kỳ và TAM:**
> - Chu kỳ ngành: Emerging / Growth / **Maturity** / Decline (khoanh tròn)
> - TAM hiện tại: ~X tỷ USD/tỷ VNĐ; CAGR dự báo: X%/năm
> - Công ty đang ở: Top X / Y players; thị phần: Z%
>
> **Ví dụ — GMD (Cảng biển VN):**
>
> | Lực | Mức độ | Lý do | Tác động biên |
> |---|---|---|---|
> | Đối thủ hiện tại | **Cao** | Lạch Huyện mở berths 3–8 (+40% công suất HP), PHP tại CMTV | Áp lực giá THC dài hạn, nguy cơ mất thị phần NĐV |
> | Đối thủ tiềm năng | **Thấp** | Đầu tư cảng nước sâu >500M USD, mất 10–15 năm, cần giấy phép Nhà nước | Barrier to entry cực cao → bảo vệ biên |
> | Sản phẩm thay thế | **Thấp** | Không có alternative cho vận tải container quy mô lớn | Không đáng kể |
> | NCC (hãng tàu) | **Cao** | MSC/CMA/COSCO có thể rút cargo bất cứ lúc nào; GMD phụ thuộc routing | Rủi ro volume giảm đột ngột (−17% NĐV do MSC) |
> | Khách hàng (shipper) | **Trung bình** | Fragmented shipper nhưng hãng tàu là khách hàng concentrated | Hãng tàu có negotiating power về THC |
> | **Kết luận** | **Trung bình** | Barrier cao bù đắp áp lực từ hãng tàu và dư cung | Biên ổn định nhưng khó mở rộng |
>
> - Chu kỳ: **Growth → Maturity** (VN đang tăng trưởng 12%/năm nhưng Hải Phòng dư cung)
> - TAM: Throughput VN ~30M TEU (2024); GMD chiếm ~14,8% thị phần quốc gia


### Tầng 3: Lợi thế cạnh tranh (Moat)

| Loại Moat | Dấu hiệu | Test thực tế |
|---|---|---|
| Network Effect | Giá trị tăng theo người dùng | Churn rate thấp, CAC giảm dần |
| Cost Advantage | Chi phí/đơn vị thấp hơn đối thủ | Gross margin cao hơn peer dù giá tương đương |
| Switching Cost | Khách hàng không thể dời dễ dàng | Revenue retention >90%, hợp đồng dài hạn |
| Intangible Assets | Thương hiệu, giấy phép, bằng sáng chế | Premium pricing được chấp nhận |
| Efficient Scale | Thị trường đủ nhỏ cho 1-2 player | ROIC > WACC liên tục >5 năm |

**Nguyên tắc kiểm tra moat**: Nếu đối thủ mới có vốn lớn vào thị trường, moat có giúp công ty này giữ được thị phần và biên lợi nhuận không?

> **📄 OUTPUT BẮT BUỘC → DOCX Trang 7–8**
>
> Phải hoàn thành **3 artifacts** trước khi viết trang 7–8:
>
> **Artifact 1 — Bảng đánh giá 5 loại Moat:**
>
> | Loại Moat | Rating | Bằng chứng cụ thể | Độ bền (năm) |
> |---|---|---|---|
> | Network Effect | Mạnh/TB/Yếu/Không có | [dẫn chứng đo được] | [X–Y năm] |
> | Cost Advantage | | | |
> | Switching Cost | | | |
> | Intangible Assets | | | |
> | Efficient Scale | | | |
> | **Moat tổng thể** | **Wide/Narrow/No Moat** | | |
>
> **Artifact 2 — Test ROIC vs WACC:**
> - ROIC trung bình 3 năm gần nhất: X%
> - WACC ước tính: Y% (rf + beta × ERP + spread nợ)
> - ROIC > WACC liên tục? → Có moat thật / Không → Không có moat
>
> **Artifact 3 — Đánh giá chất lượng quản lý (3 tiêu chí):**
> - Allocation of capital: ROIC tăng/giảm theo thời gian? M&A có tạo giá trị?
> - Communication: Guidance accuracy (thực tế vs kế hoạch ĐHĐCĐ)?
> - Alignment: Cổ đông lớn có cùng lợi ích với cổ đông nhỏ?
>
> **Ví dụ — GMD:**
>
> | Loại Moat | Rating | Bằng chứng |
> |---|---|---|
> | Network Effect | **Không có** | Throughput tăng không kéo theo giá tăng; hãng tàu không bị lock-in |
> | Cost Advantage | **Trung bình** | EBITDA margin 37–40% vs peers VN 25–30%; nhưng không thấp hơn peers quốc tế |
> | Switching Cost | **Trung bình** | Alliance routing mất 1–2 quý để dịch chuyển; nhưng MSC đã rút → switching xảy ra |
> | Intangible Assets | **Mạnh** | Vị trí đất cảng tại Cái Mép không thể nhân bản; giấy phép khai thác 50 năm; quan hệ CMA CGM (25% GML) |
> | Efficient Scale | **Mạnh** | CMTV chỉ cần 2–3 cảng lớn; GMD đã chiếm 20% → thị trường không đủ lớn cho thêm player |
> | **Tổng thể** | **Narrow Moat** | Moat từ Intangible + Efficient Scale; bị hạn chế bởi Switching Cost thấp |
>
> - ROIC 2022–2024: 8% → 12% → 15%; WACC ~11% → ROIC > WACC từ 2024 → **Moat thật**
> - Quản lý: Guidance accuracy tốt (2024 thực tế vs kế hoạch ±5%); CĐ lớn nắm >51% → aligned


### Tầng 4: Dự báo doanh thu và lợi nhuận

*(Chi tiết tại Mục 6)*

Nguyên tắc bất biến:
- **KHÔNG dùng % tăng trưởng cố định** — phân tích driver cụ thể
- **KHÔNG extrapolate đỉnh/đáy chu kỳ** — dùng mid-cycle assumptions
- Mỗi driver phải có logic kinh doanh giải thích được

> **📄 OUTPUT BẮT BUỘC → DOCX Trang 9–11**
>
> Phải hoàn thành **3 artifacts** trước khi viết trang 9–11. Nhãn A/E theo Mục 2.5.
>
> **Artifact 1 — Bảng Revenue Drivers (từ phương pháp ngành ở Mục 7):**
>
> | Driver | 2022A | 2023A | 2024A | 2025A | 2026E | 2027E | Nguồn/Logic |
> |---|---|---|---|---|---|---|---|
> | [Driver 1 — VD: TEU volume (triệu)] | | | | | | | [Nguồn thực tế] |
> | [Driver 2 — VD: THC bq (tỷ/M TEU)] | | | | | | | |
> | **Doanh thu tính được** | | | | | | | =D1×D2 |
>
> **Artifact 2 — Bảng P&L lịch sử + dự báo:**
>
> | Chỉ tiêu (tỷ VNĐ) | [N-4]A | [N-3]A | [N-2]A | [N-1]A | NA | [N+1]E | [N+2]E |
> |---|---|---|---|---|---|---|---|
> | Doanh thu thuần | | | | | | | |
> | Tăng trưởng DT (%) | | | | | | | |
> | Lợi nhuận gộp | | | | | | | |
> | Biên GP (%) | | | | | | | |
> | EBITDA | | | | | | | |
> | Biên EBITDA (%) | | | | | | | |
> | LNTT | | | | | | | |
> | LNST CĐCTM | | | | | | | |
> | EPS (VNĐ) | | | | | | | |
> | Net Cash/(Debt) | | | | | | | |
>
> ⚠️ **Bắt buộc**: Năm có one-off (thoái vốn, bất thường) → ghi chú rõ và tách core/reported.
>
> **Artifact 3 — Bridge analysis** (năm A gần nhất → năm E tiếp theo):
> Giải thích từng dòng thay đổi lớn: "DT tăng X% nhờ A (+Y tỷ), bù bởi B (−Z tỷ)"
>
> **Ví dụ — GMD Revenue Driver Table:**
>
> | Driver | 2022A | 2023A | 2024A | 2025A | 2026E | 2027E |
> |---|---|---|---|---|---|---|
> | TEU NĐV (triệu) | 1,10 | 1,20 | 1,80 | 2,10 | 2,50 | 2,70 |
> | TEU GML (triệu) | 1,40 | 1,60 | 1,90 | 2,10 | 2,40 | 2,80 |
> | THC bq (tỷ/M TEU) | 0,860 | 0,880 | 0,946 | 0,960 | 1,056 | 1,080 |
> | DT cảng (tỷ) | 2.150 | 2.464 | 3.496 | 3.994 | 5.202 | 5.939 |
> | DT logistics (tỷ) | 530 | 560 | 632 | 680 | 780 | 860 |
> | **Tổng DT** | **3.680** | **3.024** | **4.128** | **4.674** | **5.982** | **6.799** |
>
> ⚠️ 2023A one-off: +1.772 tỷ từ thoái vốn cảng → core LNTT 2023 = 1.372 tỷ, không phải 3.144 tỷ.


### Tầng 5: Phân tích rủi ro

**5 nhóm rủi ro cần cover**:
1. **Vĩ mô**: Lãi suất, tỷ giá, GDP, chính sách
2. **Ngành**: Chu kỳ, quy định, disruption
3. **Doanh nghiệp**: Quản trị, đòn bẩy, thực thi
4. **Cạnh tranh**: Đối thủ mới, pricing war, thay thế
5. **Sự kiện**: M&A, pháp lý, nhân sự lãnh đạo

**Mỗi rủi ro cần**: Mô tả + Xác suất (%) + Tác động (tỷ đồng hoặc % LNST) + Chỉ báo leading indicator

> **📄 OUTPUT BẮT BUỘC → DOCX Trang 14–15**
>
> Phải hoàn thành **Risk Matrix đầy đủ** trước khi viết trang 14–15:
>
> **Bảng Risk Matrix (tối thiểu 5 rủi ro, tối đa 10):**
>
> | Rủi ro | Nhóm | Mô tả + ngưỡng kích hoạt | Xác suất 12T | Tác động LNST (tỷ/%) | Leading indicator theo dõi |
> |---|---|---|---|---|---|
> | [Tên rủi ro 1] | Vĩ mô/Ngành/DN/CT/Sự kiện | [mô tả cụ thể; khi nào rủi ro xảy ra] | Thấp/TB/Cao (X%) | −Y tỷ hoặc −Z% LNST | [chỉ số cụ thể để theo dõi] |
> | ... | | | | | |
>
> **Bảng Catalyst Timeline:**
>
> | Catalyst | Mô tả | Thời điểm dự kiến | Upside giá ước tính | Xác suất |
> |---|---|---|---|---|
> | [Catalyst 1] | [cụ thể] | Q[X]/[năm] | +X% hoặc +Y VNĐ/CP | X% |
>
> **Ví dụ — GMD Risk Matrix:**
>
> | Rủi ro | Nhóm | Mô tả | Xác suất | Tác động | Indicator |
> |---|---|---|---|---|---|
> | MSC rút cargo NĐV | Cạnh tranh | MSC tái cơ cấu Alliance, rút ~17% throughput NĐV. Kích hoạt: MSC ký routing với cảng khác | **Cao (đang xảy ra)** | −350 tỷ LNST/năm (−24%) | Số tàu MSC cập NĐV mỗi tuần |
> | Dư cung Hải Phòng | Ngành | Lạch Huyện berths 3–8 vận hành, tổng công suất HP +40%. Kích hoạt: Utilization NĐV <75% | Cao (2026–2027) | Áp lực THC, −5–10% biên GP | Tốc độ xây dựng Lạch Huyện; KLGD NĐV/tháng |
> | Phase 2A chậm tiến độ | DN | EIA, chuyển đổi đất phức tạp → delay 6–12T. Kích hoạt: EIA chưa được phê duyệt T9/2026 | Trung bình (35%) | −growth story, P/E re-rate xuống 12x | Cập nhật pháp lý từ IR GMD |
> | Tariff Mỹ–VN leo thang | Vĩ mô | Thuế >35% → XK VN giảm → TEU giảm 8–12% | Thấp (15%) | −600–900 tỷ DT | XK VN sang Mỹ tháng/tháng (GSO) |
> | Lãi suất tăng | Vĩ mô | SBV tăng lãi suất >50bps → chi phí Phase 2A CAPEX tăng | Thấp (20%) | +100–200 tỷ chi phí lãi/năm | SBV policy rate |


### Tầng 6: Định giá

*(Chi tiết tại Mục 8)*

Nguyên tắc:
- Dùng **≥ 2 phương pháp** và so sánh chéo
- Luôn làm **sensitivity analysis** với 2 biến quan trọng nhất
- **Bear / Base / Bull** với xác suất rõ ràng
- Kết luận: **BUY / HOLD / SELL** với margin of safety

> **📄 OUTPUT BẮT BUỘC → DOCX Trang 12–13**
>
> Phải hoàn thành **3 artifacts** trước khi viết trang 12–13:
>
> **Artifact 1 — Bảng Valuation Bear/Base/Bull:**
>
> | Phương pháp | Multiple | Bear (X%) | Base (X%) | Bull (X%) |
> |---|---|---|---|---|
> | P/E × EPS [N+1]E | Bear: Xx / Base: Xx / Bull: Xx | XXX VNĐ | XXX VNĐ | XXX VNĐ |
> | EV/EBITDA [N+1]E | Bear: Xx / Base: Xx / Bull: Xx | XXX VNĐ | XXX VNĐ | XXX VNĐ |
> | DCF (WACC=X%, g=X%) | — | — | XXX VNĐ | — |
> | **Blended Target** | | **XXX VNĐ** | **XXX VNĐ** | **XXX VNĐ** |
> | Upside/(Downside) | | X% | X% | X% |
> | **Xác suất kịch bản** | | **X%** | **X%** | **X%** |
>
> ⚠️ **Bắt buộc**: Tổng xác suất Bear + Base + Bull = **100%**. Thiếu điều này là lỗi nghiêm trọng.
>
> **Artifact 2 — Sensitivity Table 5×5:**
> - Biến 1: [giả định quan trọng nhất — VD: TEU volume, EPS, WACC]
> - Biến 2: [giả định quan trọng thứ 2 — VD: THC rate, P/E multiple]
>
> **Artifact 3 — Consensus CTCK** (nếu tìm được ≥ 2 báo cáo):
>
> | CTCK | Giá mục tiêu | Khuyến nghị | P/E target | Ngày |
> |---|---|---|---|---|
>
> **Ví dụ — GMD Valuation:**
>
> | Phương pháp | Multiple | Bear (25%) | Base (50%) | Bull (25%) |
> |---|---|---|---|---|
> | P/E × EPS 2026E | 12x / 17x / 22x | 57.888 | 82.008 | 106.128 |
> | EV/EBITDA 2026E | 9x / 12x / 15x | 46.906 | 78.780 | 110.654 |
> | DCF (WACC=11%, g=3%) | — | — | 80.570 | — |
> | **Blended** | | **52.397** | **80.394** | **108.391** |
> | Upside | | −28,6% | **+9,5%** | +47,7% |
> | **Xác suất** | | **25%** | **50%** | **25%** |
>
> → Weighted average target = 52.397×25% + 80.394×50% + 108.391×25% = **80.394 VNĐ**
> → Khuyến nghị: **TÍCH LŨY** (upside +9,5%, không đủ để MUA MẠNH nhưng > 0)


---

## 4. PHÂN TÍCH PESTLE — MÔI TRƯỜNG VĨ MÔ

### Vị trí trong framework

PESTLE là **tầng phân tích ngoài cùng** — môi trường vĩ mô mà công ty không kiểm soát được nhưng bị ảnh hưởng trực tiếp. Nó thực hiện song song với Tầng 2 (Thị trường) và kết quả đưa vào Tầng 5 (Rủi ro) và Tầng 6 (Định giá — điều chỉnh discount rate).

```
PESTLE (vĩ mô) ─────────────────────────────────┐
                                                  ▼
Tầng 1 → Tầng 2 (Porter) → Tầng 3 → Tầng 4 → Tầng 5 (Risk) → Tầng 6 (Valuation)
```

**Quy tắc sử dụng**:
- Mỗi PESTLE factor phải trả lời được: *"Điều này ảnh hưởng đến EPS / dòng tiền / multiple của công ty này bao nhiêu và theo hướng nào?"*
- Nếu một factor không ảnh hưởng đến công ty đang phân tích → bỏ qua, không liệt kê cho có
- Nếu ảnh hưởng lớn → đưa vào Risk Matrix (Tầng 5) với xác suất và tác động cụ thể

---

### P — POLITICAL (Chính trị)

#### Bối cảnh Việt Nam

| Yếu tố | Nội dung cần phân tích | Tác động |
|---|---|---|
| **Nghị quyết / QĐ Chính phủ** | Quy hoạch ngành (điện, khoáng sản, đất đai, logistics) | Tạo hoặc đóng cửa cơ hội tăng trưởng |
| **Thay đổi lãnh đạo** | Thay Thủ tướng, Bộ trưởng ngành → thay đổi ưu tiên chính sách | Short-term uncertainty, thường không ảnh hưởng dài hạn |
| **Quan hệ VN–Mỹ** | Thuế quan, FTA, chip export controls | Ảnh hưởng FDI, xuất khẩu điện tử, logistics |
| **Quan hệ VN–Trung** | Chuỗi cung ứng, nguyên liệu đầu vào, thương mại biên giới | Thép, khoáng sản, FMCG bị ảnh hưởng mạnh |
| **Cổ phần hóa / Thoái vốn Nhà nước** | SCIC bán bớt, room ngoại mở | Catalyst giá cổ phiếu với cổ phiếu SOE |
| **FDI Policy** | Ưu đãi thuế KCN, ngành ưu tiên | KCN, logistics, EPC hưởng lợi |

#### Tác động theo ngành

| Ngành | Mức độ nhạy cảm chính trị | Lý do |
|---|---|---|
| KCN / Hạ tầng | ⭐⭐⭐⭐⭐ | Phụ thuộc hoàn toàn vào quy hoạch và FDI policy |
| Khoáng sản | ⭐⭐⭐⭐⭐ | Giấy phép khai thác do Nhà nước cấp, xuất khẩu controlled |
| BĐS | ⭐⭐⭐⭐⭐ | Luật Đất đai, cấp phép dự án, tín dụng bất động sản |
| Ngân hàng | ⭐⭐⭐⭐ | Hạn mức tín dụng NHNN, lãi suất điều hành |
| EPC / Điện | ⭐⭐⭐⭐ | Quy hoạch điện VIII, đầu tư công |
| Thép | ⭐⭐⭐ | Thuế CBPG, quota nhập khẩu, quy hoạch xây dựng |
| Bán lẻ / FMCG | ⭐⭐ | Quy định an toàn thực phẩm, thuế tiêu thụ đặc biệt |
| Dược phẩm | ⭐⭐⭐ | Thông tư đấu thầu thuốc (15, 08...), phê duyệt sản phẩm |
| Công nghệ | ⭐⭐ | Quy định dữ liệu, cyber security |

#### Cách đưa vào investment thesis

```
Câu hỏi cần trả lời:
1. Công ty này có phụ thuộc vào giấy phép/quyết định chính phủ để tạo ra doanh thu không?
   → CÓ: Đây là moat (nếu đã có) hoặc rủi ro (nếu chưa có/sắp hết hạn)

2. Chính sách hiện tại đang hỗ trợ hay cản trở ngành này?
   → Đang hỗ trợ (quy hoạch điện, FDI...) = tailwind → nâng multiple
   → Đang siết (tín dụng BĐS, room khoáng sản...) = headwind → giảm multiple

3. Có event chính trị nào trong 12 tháng tới có thể là catalyst?
   → ĐHĐCĐ, Nghị quyết mới, thay đổi quy định → đưa vào Catalyst section
```

---

### E — ECONOMIC (Kinh tế)

#### Bối cảnh Việt Nam

| Chỉ số | Dữ liệu cần thu thập | Tác động đến cổ phiếu |
|---|---|---|
| **GDP Growth** | GDP VN YoY %, breakdown theo khu vực | Nền kinh tế tăng trưởng mạnh → tiêu dùng, đầu tư tốt |
| **Lãi suất** | SBV policy rate, lãi suất huy động/cho vay bình quân | ↑ lãi suất → tăng discount rate → giảm P/E chấp nhận được; ↑ chi phí nợ |
| **Tỷ giá USD/VND** | Spot rate, NHNN band | Ảnh hưởng công ty xuất khẩu, nhập khẩu nguyên liệu, nợ ngoại tệ |
| **Lạm phát (CPI)** | CPI headline và core | Ảnh hưởng sức mua, chi phí đầu vào FMCG/bán lẻ |
| **Tín dụng** | Tăng trưởng tín dụng toàn hệ thống, room ngân hàng | Ngân hàng, BĐS, bán lẻ (consumer credit) |
| **FDI** | Vốn đăng ký và giải ngân FDI | KCN, logistics, thép, vật liệu xây dựng |
| **Xuất nhập khẩu** | Kim ngạch XNK, cán cân thương mại | Logistics, thép, điện tử (DGW) |
| **Giá hàng hóa** | LME (Cu, Zn, Al, Ni), giá HRC, giá dầu | Thép, khoáng sản, hóa chất, vận tải |

#### Ma trận tác động kinh tế lên từng ngành

| | ↑ Lãi suất | ↓ Lãi suất | ↑ Tỷ giá (VND yếu) | ↓ Tỷ giá | ↑ GDP | ↑ Giá HH |
|---|---|---|---|---|---|---|
| **Ngân hàng** | NIM co lại ban đầu, NIM tăng dài hạn | NIM tăng ngay | Trung tính | Trung tính | Tín dụng ↑ | — |
| **BĐS** | ❌ Nhu cầu giảm | ✅ Nhu cầu tăng | Trung tính | Trung tính | ✅ | — |
| **Thép** | ❌ Chi phí vốn | ✅ | ✅ Lợi cho xuất khẩu | ❌ | ✅ Xây dựng | ✅ Giá thép |
| **KCN** | Trung tính | ✅ FDI ↑ | ✅ Giá thuê USD ↑ giá VND | ❌ | ✅ FDI | — |
| **Bán lẻ** | ❌ Sức mua, chi phí | ✅ | ❌ Nhập khẩu đắt | ✅ | ✅ | — |
| **Khoáng sản** | Trung tính | Trung tính | ✅ Revenue USD → VND nhiều hơn | ❌ | Trung tính | ✅ Giá kim loại |
| **Logistics** | Trung tính | ✅ | ✅ Cước thu USD | ❌ | ✅ XNK ↑ | ❌ Giá dầu ↑ |
| **IT/Tech** | Trung tính | ✅ | ✅ Revenue USD | ❌ | ✅ | — |
| **FMCG** | ❌ Nhẹ | ✅ | ❌ Nguyên liệu nhập đắt | ✅ | ✅ | ❌ Raw mat |
| **Dược phẩm** | Trung tính | ✅ | ❌ API nhập đắt | ✅ | ✅ | — |

> **Đọc bảng**: ✅ = hưởng lợi | ❌ = bất lợi | Trung tính = ảnh hưởng không đáng kể

#### Cách đưa vào model

```
Tỷ giá: Với công ty có doanh thu USD và chi phí VND (KCN, Khoáng sản, IT offshore):
→ Sensitivity: thay đổi tỷ giá ±5% → tác động doanh thu bao nhiêu %?

Lãi suất: Với công ty có nợ vay lớn:
→ Net debt × ΔLãi suất = thay đổi chi phí lãi hàng năm → ảnh hưởng EPS

GDP: Với bán lẻ/FMCG:
→ GDP elasticity: GDP tăng 1% → doanh thu tăng bao nhiêu %? (thường 1.5-2.5x)
```

---

### S — SOCIAL (Xã hội)

#### Các xu hướng xã hội tác động đến cổ phiếu VN

| Xu hướng | Nội dung | Ngành hưởng lợi | Ngành bị ảnh hưởng |
|---|---|---|---|
| **Dân số trẻ & đô thị hóa** | 70% dân số <40 tuổi, đô thị hóa 40%→60% | Bán lẻ, FMCG, BĐS đô thị, giáo dục, logistics | — |
| **Tầng lớp trung lưu tăng** | 30M người trung lưu 2025 → 50M 2030 | Premium retail (PNJ, MWG), thực phẩm chế biến, du lịch, bảo hiểm | Hàng giá rẻ/bình dân |
| **Già hóa dân số** | Tỷ lệ người >60 tuổi tăng dần | Dược phẩm, bảo hiểm nhân thọ, y tế | FMCG trẻ em |
| **E-commerce & Digital** | Penetration rate e-commerce tăng 25%/năm | Logistics, payment, digital retail | Bán lẻ offline thuần túy |
| **Health consciousness** | Hậu COVID, quan tâm sức khỏe tăng | Thực phẩm sạch, dược phẩm OTC, thể thao | Rượu bia (nhẹ) |
| **Migration lao động** | Lao động từ nông thôn → KCN phía Nam/Bắc | KCN, logistics khu công nghiệp, nhà ở công nhân | Nông nghiệp truyền thống |
| **Tiêu dùng trải nghiệm** | Ưu tiên du lịch, ăn ngoài, giải trí | F&B, du lịch, AVA Sport | Tiêu dùng hàng hóa cứng |

#### Cách ứng dụng

```
Không phải mọi xu hướng xã hội đều liên quan đến mọi công ty.
Chỉ chọn 1-2 xu hướng CÓ TÁC ĐỘNG RÕ RÀNG đến công ty đang phân tích.

Ví dụ MWG (BHX):
→ Đô thị hóa + tầng lớp trung lưu → nhu cầu siêu thị tiện lợi tăng
→ Quantify: Dân số đô thị tăng X% → tệp khách hàng tiềm năng BHX tăng Y triệu người
→ So sánh với mật độ cửa hàng hiện tại → room tăng trưởng Z CH

Ví dụ Dược phẩm (DHG):
→ Già hóa dân số → nhu cầu thuốc mãn tính tăng
→ Quantify: Số người >60 tăng X triệu trong 10 năm → market size tăng Y%
```

---

### T — TECHNOLOGICAL (Công nghệ)

#### Phân tích 2 chiều: Disruption Risk vs. Enabler

```
Công nghệ với doanh nghiệp = ĐỒNG TIỀN 2 MẶT

Mặt 1 — ENABLER: Công nghệ giúp công ty làm gì tốt hơn?
→ Tự động hóa → giảm COGS/đơn vị
→ Digital marketing → giảm CAC
→ Data analytics → tăng SSS, giảm hàng tồn kho

Mặt 2 — DISRUPTOR: Công nghệ có đang phá vỡ mô hình của họ không?
→ E-commerce → threat với bán lẻ offline
→ Fintech → threat với ngân hàng truyền thống
→ Tự động hóa → threat với EPC/lao động thủ công
→ AI → threat với IT outsourcing đơn giản
```

| Ngành | Disruption Risk | Công nghệ Enabler |
|---|---|---|
| **Bán lẻ truyền thống** | ⭐⭐⭐⭐ E-commerce cannibalization | Omnichannel, inventory AI, loyalty app |
| **Ngân hàng** | ⭐⭐⭐ Fintech, neo-bank, BNPL | Mobile banking, AI credit scoring, CASA app |
| **Dược phẩm** | ⭐⭐ Telemedicine thay kê đơn | E-pharmacy, biotech pipeline |
| **Logistics** | ⭐⭐ Drone delivery (dài hạn) | Route optimization AI, tracking, WMS |
| **Thép/Khoáng sản** | ⭐ Vật liệu thay thế (dài hạn) | Smart factory, IoT mining, automation |
| **KCN** | ⭐ Không gian ảo thay thế (rất dài hạn) | Smart industrial zones, energy management |
| **IT/FPT** | ⭐⭐⭐ AI thay thế junior developers | AI augmentation, nâng cấp lên AI consulting |
| **FMCG** | ⭐⭐ D2C bypass distribution | Precision marketing, supply chain AI |

#### Checklist công nghệ khi phân tích

```
□ Công ty đang ứng dụng công nghệ gì để cải thiện biên lợi nhuận?
  → Nếu có: quantify được không? (COGS giảm X%, headcount giảm Y%)

□ Đối thủ công nghệ (tech natives) có đang tấn công thị phần không?
  → Ai? Market share trend? Tốc độ mất thị phần?

□ Capex công nghệ/IT chiếm bao nhiêu % tổng Capex?
  → Thấp (<5%) trong ngành công nghệ = risk
  → Cao trong ngành truyền thống = đang chuyển đổi tốt

□ AI ảnh hưởng như thế nào? (Tích cực: giảm chi phí vận hành; Tiêu cực: disrupt sản phẩm)
```

---

### L — LEGAL (Pháp lý)

#### Hệ thống pháp lý ảnh hưởng trực tiếp đến cổ phiếu VN

| Luật / Quy định | Năm hiệu lực | Ngành ảnh hưởng | Tác động |
|---|---|---|---|
| **Luật Đất đai 2024** | 01/2025 | BĐS, KCN | Cởi mở hơn, tháo gỡ pháp lý dự án |
| **Luật Chứng khoán sửa đổi** | 2024 | Tất cả | Room ngoại, công ty đại chúng, chào bán thêm |
| **Quy định NHNN về room tín dụng** | Hàng năm | Ngân hàng | Cap tăng trưởng tín dụng |
| **Thông tư 08, 15 đấu thầu thuốc** | Ongoing | Dược phẩm | Áp lực giảm giá ETC |
| **Thuế tối thiểu toàn cầu (BEPS)** | 2024 | KCN, FDI | Giảm ưu đãi thuế → FDI có thể re-direct |
| **Luật Khoáng sản** | Đang sửa | Khoáng sản | Điều kiện cấp phép, thuế tài nguyên |
| **Quy định Basel II/III** | Ongoing | Ngân hàng | CAR requirements, provisioning |
| **Thuế CBPG thép nhập khẩu** | Theo vụ việc | Thép | Bảo hộ nội địa HPG |
| **Net Zero 2050 (COP26 commitments)** | 2050 | Điện, than, khoáng sản | Chuyển dịch năng lượng dài hạn |

#### Checklist pháp lý khi phân tích

```
□ Luật nào đang thay đổi trong 12-24 tháng tới có ảnh hưởng trực tiếp?
  → Xác định: favorable hay unfavorable

□ Công ty có đang vi phạm hoặc cận ngưỡng vi phạm quy định nào không?
  → KSV: cơ cấu cổ đông vs. Luật Chứng khoán
  → BĐS: pháp lý dự án chưa hoàn chỉnh

□ Giấy phép / license quan trọng sắp hết hạn?
  → Gia hạn được không? Chi phí bao nhiêu? Rủi ro không gia hạn được?

□ Thuế quan / CBPG (chống bán phá giá) tác động thế nào?
  → Thép: Mỹ, EU áp thuế CBPG với thép VN?
  → Giày dép, điện tử: trade war impact?
```

---

### E — ENVIRONMENTAL (Môi trường)

#### Tại sao môi trường quan trọng với cổ phiếu VN hiện nay

VN cam kết Net Zero 2050 (COP26) + xu hướng ESG investing từ nước ngoài đang tạo ra **2 lực song song**:
- **Regulatory push**: Quy định môi trường ngày càng khắt khe hơn → tăng chi phí compliance
- **ESG premium**: Công ty ESG tốt được định giá cao hơn từ nhà đầu tư nước ngoài

| Yếu tố Môi trường | Ngành bị tác động | Chiều tác động |
|---|---|---|
| **Chuyển dịch năng lượng (RE)** | EPC điện gió/mặt trời (PC1, TV2) | ✅ Cơ hội lớn backlog mới |
| | Than (Vinacomin, nhiệt điện) | ❌ Stranded asset dài hạn |
| **Quy định khí thải sản xuất** | Thép (HPG lò cao) | ❌ Capex xanh bắt buộc tăng |
| | Xi măng, hóa chất | ❌ |
| **Tiêu chuẩn môi trường mỏ** | Khoáng sản (KSV, Đông Pao) | ❌ Chi phí compliance, delay khai thác |
| **Nước và chất thải công nghiệp** | KCN, sản xuất | ❌ Yêu cầu xử lý nước thải tăng |
| **ESG scoring** | Tất cả (đặc biệt với room ngoại) | Foreign investor ưu tiên ESG → tác động valuation |
| **Tín chỉ carbon** | Năng lượng tái tạo, trồng rừng | ✅ Doanh thu mới tiềm năng |
| **Biến đổi khí hậu** | Bảo hiểm, nông nghiệp, logistics | ❌ Rủi ro thời tiết cực đoan |

#### Cách đưa vào phân tích

```
Cấp độ 1 — Compliance cost:
→ Capex môi trường hiện tại chiếm bao nhiêu % tổng Capex?
→ Quy định mới sẽ tăng thêm bao nhiêu? (ảnh hưởng EBITDA bao nhiêu %)

Cấp độ 2 — Stranded asset risk:
→ Tài sản của công ty có nguy cơ mất giá vì chuyển dịch năng lượng không?
→ Ví dụ: nhà máy nhiệt điện than vs. timeline RE

Cấp độ 3 — ESG premium/discount:
→ Foreign ownership thấp vì ESG score kém? → Room ngoại, trading discount
→ Ngược lại: ESG tốt → attract sustainable funds → multiple premium

Cấp độ 4 — New opportunity:
→ Tín chỉ carbon, RE certification → doanh thu mới?
→ ESG-linked financing → lãi suất vay thấp hơn?
```

---

### TỔNG HỢP PESTLE — ĐÁNH GIÁ NET IMPACT

Sau khi phân tích 6 yếu tố, tổng hợp lại thành:

**Bảng PESTLE Impact Summary**

| Yếu tố | Tác động với [TICKER] | Chiều | Mức độ | Thời hạn |
|---|---|---|---|---|
| Political | [Mô tả cụ thể] | ✅/❌/➡️ | Cao/TB/Thấp | Ngắn/Trung/Dài |
| Economic | | | | |
| Social | | | | |
| Technological | | | | |
| Legal | | | | |
| Environmental | | | | |
| **Net PESTLE** | | **Tích cực / Trung tính / Tiêu cực** | | |

**Net PESTLE → Điều chỉnh định giá**:
- Net tích cực rõ ràng → có thể nâng multiple lên 5-10% so với peer median
- Net tiêu cực → áp dụng discount 5-15% vào target price
- Neutral → dùng peer median làm anchor

---

### VÍ DỤ THỰC TẾ: PESTLE của KSV (Vimico)

| Yếu tố | Nội dung cụ thể | Chiều | Mức độ |
|---|---|---|---|
| **Political** | Giấy phép khai thác do Nhà nước cấp → barrier to entry; TKV >98% → nguy cơ hủy niêm yết | ✅❌ | Cao |
| **Economic** | Giá đồng LME $9,500-12,000 (2025-2026); Tỷ giá USD/VND ~25,500 → doanh thu VND cao | ✅ | Rất cao |
| **Social** | Đô thị hóa → nhu cầu dây điện, ống đồng; Đất hiếm cho EV, điện tử | ✅ | Trung bình |
| **Technological** | Mining automation giảm chi phí khai thác; Đất hiếm critical cho AI/chip supply chain | ✅ | Trung bình |
| **Legal** | Luật Khoáng sản đang sửa đổi; KSV cận ngưỡng vi phạm Luật Chứng khoán (cơ cấu CP) | ❌ | Cao |
| **Environmental** | Đông Pao: phê duyệt môi trường phức tạp, delay tiến độ; Chi phí xử lý chất thải mỏ tăng | ❌ | Trung bình |
| **Net PESTLE** | Kinh tế rất tích cực (giá hàng hóa); Pháp lý tiêu cực (niêm yết + khoáng sản) | Trung tính | — |

---

## 5. BƯỚC 3: BUILD EXCEL MODEL

### 5.1 Quyết định template hay build mới

```
Nhận ticker mới
    │
    ├─ Ngành đã có template phù hợp?
    │   ├─ CÓ → Mở template, review structure, điền data mới
    │   │        (Không sửa format nếu không cần thiết)
    │   └─ KHÔNG → Build file mới theo cấu trúc chuẩn 10-12 sheets
    │
    └─ Công ty có cấu trúc đặc biệt (SOTP, đa ngành, pre-profit)?
        ├─ CÓ → Build custom model (ví dụ: MWG SOTP)
        └─ KHÔNG → Model single-entity chuẩn
```

**Danh sách templates hiện có**:

| Template | Phù hợp ngành | Không phù hợp |
|---|---|---|
| `01_HPG_Thep_Model.xlsx` | Sản xuất commodity (thép, hóa chất, vật liệu) | |
| `02_BanLe_MWG_PNJ_Model.xlsx` | Chuỗi bán lẻ | Đa chuỗi nhiều chu kỳ |
| `03_NganHang_TCB_ACB_Model.xlsx` | Ngân hàng thương mại | Fintech, bảo hiểm |
| `04_ChungKhoan_SSI_Model.xlsx` | CTCK | |
| `05_KhuCongNghiep_NTC_SZC_Model.xlsx` | KCN, hạ tầng cho thuê | |
| `06_BDS_NLG_Model.xlsx` | BĐS nhà ở, dự án | BĐS thương mại (khác) |
| `07_MSN_Conglomerate_Model.xlsx` | Tập đoàn đa ngành | |
| `08_SanXuat_BMP_TLG_TV2_PC1_Model.xlsx` | Sản xuất công nghiệp, EPC | |
| `09_MWG_SOTP_Model.xlsx` | Đa chuỗi bán lẻ | |

### 5.2 Cấu trúc sheet chuẩn (Bắt buộc)

```
Sheet 1:  Cover                  — Thông tin, thesis 1 trang, snapshot tài chính
Sheet 2:  Value Chain            — Sơ đồ chuỗi giá trị ngành
Sheet 3:  Market                 — Porter 5F, TAM, chu kỳ ngành
Sheet 4:  Moat                   — Đánh giá lợi thế, peer ROIC comparison
Sheet 5:  Assumptions            — TẤT CẢ giả định trung tâm
Sheet 6:  Revenue Model          — Driver-based revenue forecast
Sheet 7:  P&L (Yearly)           — Báo cáo kết quả kinh doanh theo Năm (5 năm lịch sử + 3 năm dự báo)
Sheet 8:  P&L (Quarterly)        — Báo cáo kết quả kinh doanh theo Quý (tối thiểu 20 quý gần nhất)
Sheet 9:  Balance Sheet (Yearly)  — Bảng cân đối kế toán theo Năm (5 năm lịch sử + 3 năm dự báo)
Sheet 10: Balance Sheet (Quarterly)— Bảng cân đối kế toán theo Quý (tối thiểu 20 quý gần nhất)
Sheet 11: Ratios                 — Các chỉ số tài chính (NIM, ROE, LDR, CASA, CIR, NPL...) tính bằng công thức từ dữ liệu Quý và Năm
Sheet 12: Valuation              — Phương pháp định giá chính
Sheet 13: Sensitivity            — Bảng nhạy cảm 2 biến × 3 kịch bản
Sheet 14: Risk Matrix            — Ma trận rủi ro
Sheet 15: Leading Indicators     — Các yếu tố cần theo dõi (xem chi tiết bên dưới)
```

> **YÊU CẦU BẮT BUỘC VỀ DỮ LIỆU NĂM VÀ QUÝ:**
> 1. **Tải dữ liệu đầy đủ**: Khi load dữ liệu từ API Vietcap, bắt buộc phải tải toàn bộ dữ liệu Cân đối kế toán, Kết quả kinh doanh, Lưu chuyển tiền tệ, và Thuyết minh cho cả Năm (tối thiểu 5 năm gần nhất) và Quý (tối thiểu 20 quý gần nhất).
> 2. **Phân tách các sheet dữ liệu**: Dữ liệu Quý và Năm phải được import đầy đủ và có hệ thống vào các sheet tương ứng (`P&L (Yearly)`, `P&L (Quarterly)`, `Balance Sheet (Yearly)`, `Balance Sheet (Quarterly)`).
> 3. **Tính toán bằng công thức Excel**: Các chỉ số tài chính ở sheet `Ratios` (như NIM, ROE, LDR, CASA, CIR, NPL, tăng trưởng tín dụng, tăng trưởng huy động...) bắt buộc phải được tính toán bằng **công thức Excel liên kết trực tiếp** đến các ô dữ liệu trong sheet Quý và Năm. Tuyệt đối không được tính toán trước ở Python rồi điền số cứng (hardcode) vào Excel.

> **Nếu ngành đặc thù** (ngân hàng, BĐS, SOTP): có thể thêm/bớt sheet nhưng phải giữ cover + assumptions + valuation + sensitivity + risk + **leading indicators**.

### 5.3 Cấu trúc Sheet 13: Theo dõi đầu tư — Framework 7 bước

Sheet này không phải danh sách chỉ số tài chính — mà là **bản đồ nhân quả** giữa môi trường kinh doanh và kết quả P&L của công ty. Phải được xây từng bước dưới đây, áp dụng cụ thể cho từng ticker.

**Nguyên tắc lọc indicator:** Chỉ đưa vào indicator nào mà khi nó thay đổi 10%, EBITDA của công ty này thay đổi đo được được — và khác biệt so với đối thủ cùng ngành. Chỉ số ảnh hưởng đều như nhau cho toàn thị trường (VD: GDP tổng hợp) thường không đủ đặc thù.

---

#### BƯỚC 1 — PHÂN TÍCH NGÀNH: Xác định driver trực tiếp

Trả lời 3 câu hỏi trước khi điền bất kỳ indicator nào:

```
1. Giá bán phụ thuộc vào gì? (3 yếu tố chính)
   → VD sản xuất: cung/cầu sản phẩm ngành, giá benchmark quốc tế, cạnh tranh hàng nhập khẩu
   → VD bán lẻ: sức mua người tiêu dùng, giá đối thủ, mix sản phẩm
   → VD ngân hàng: lãi suất điều hành, cạnh tranh NIM, CASA ratio

2. Sản lượng/tăng trưởng phụ thuộc vào gì? (3 yếu tố chính)
   → VD sản xuất: cầu từ ngành hạ nguồn, công suất hiện tại, thị phần vs đối thủ
   → VD ngân hàng: tăng trưởng tín dụng, chất lượng tài sản, vốn hóa

3. Top 3 chi phí đầu vào là gì và % trong giá vốn?
   → Mỗi chi phí ghi rõ: tên, % COGS ước tính, nguồn dữ liệu theo dõi
```

---

#### BƯỚC 2 — BẢN ĐỒ NHÂN QUẢ

Với mỗi trigger quan trọng, vẽ chuỗi: **Macro/Ngành trigger → Tác động ngành → Tác động công ty → Kết quả tài chính cụ thể**

Bắt buộc ước tính lag time từng bước. Ví dụ:
```
Giá đầu vào tăng → Chi phí/tấn tăng → Biên GP giảm → LNST giảm | Lag tổng: 1-2 tháng
Cầu ngành tăng   → Đơn hàng tăng   → DT tăng      → EBITDA tăng | Lag tổng: 2-3 quý
Catalyst chiến lược hoàn thành → Công suất mới → Chi phí cố định hấp thụ tốt hơn | Lag: 12-18 tháng
```

Chuỗi nào lag ngắn hơn → có thể dùng làm leading indicator.

---

#### BƯỚC 3 — LEADING INDICATORS (ưu tiên 5-7 indicators)

Chỉ lấy những indicator thỏa mãn: **(a)** có dữ liệu công khai, **(b)** xuất hiện trước kết quả tài chính ít nhất 1 quý, **(c)** đặc thù cho công ty/ngành này.

Phân tầng:
- **Tầng 1-2 (vĩ mô):** Chỉ đưa vào nếu có chuỗi nhân quả rõ ràng đến P&L — không đưa GDP chung chung nếu không chứng minh được kênh tác động
- **Tầng 3 (ngành):** Giá nguyên liệu đặc thù, chỉ số cung/cầu ngành, hành động cạnh tranh
- **Tầng 4 (doanh nghiệp):** Tiến độ dự án chiến lược, đơn hàng tồn đọng, mở rộng công suất

Cột bắt buộc cho mỗi indicator:

```
┌──────────────────┬──────────────┬───────────────────┬──────────┬────────────────────────┬────────────────────┬──────────────┐
│ Indicator        │ Giá trị HT   │ Ngưỡng tốt/TL/xấu │ Trọng số │ Nguồn dữ liệu          │ Lag time ước tính  │ Tần suất     │
└──────────────────┴──────────────┴───────────────────┴──────────┴────────────────────────┴────────────────────┴──────────────┘
```

Ngưỡng 3 vùng tô màu QUAL(): 'pos' (xanh) / 'neu' (vàng) / 'neg' (đỏ).

**Lưu ý hướng ngưỡng** — phải xác định trước khi điền:
- Chi phí đầu vào, rủi ro, đòn bẩy: **Thấp = tốt** → ngưỡng tốt ở vùng thấp
- Doanh thu, biên lợi nhuận, thị phần: **Cao = tốt** → ngưỡng tốt ở vùng cao
- Tỷ lệ công suất, LDR ngân hàng: **Dải tối ưu** → tốt ở giữa, xấu ở cả hai đầu

---

#### BƯỚC 4 — COINCIDENT INDICATORS (3-5 indicators)

Phản ánh trạng thái hiện tại — di chuyển cùng chu kỳ kinh doanh. Dùng để xác nhận chiều hướng, không để dự báo.

```
┌──────────────────┬──────────────┬────────────────┬──────────────┐
│ Indicator        │ Giá trị HT   │ Nguồn          │ Tần suất     │
└──────────────────┴──────────────┴────────────────┴──────────────┘
```

Ví dụ điển hình: Doanh thu quý YoY%, biên GP thực tế, volume sản lượng, số đơn hàng mới.

---

#### BƯỚC 5 — LAGGING INDICATORS (3-4 indicators)

Xác nhận sau khi kết quả đã xảy ra — **KHÔNG dùng để ra quyết định mua/bán**.

```
┌──────────────────┬──────────────┬────────────────┬──────────────────────────────────────────────┐
│ Indicator        │ Giá trị HT   │ Nguồn          │ Lưu ý                                        │
└──────────────────┴──────────────┴────────────────┴──────────────────────────────────────────────┘
```

Ví dụ điển hình: EPS trailing 12T, ROE năm, P/E trailing, tỷ lệ cổ tức thực trả.

---

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

---

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

---

**Quy tắc bắt buộc khi build Sheet 13:**
- Phải hoàn thành Bước 1 & 2 (phân tích ngành + bản đồ nhân quả) trước khi chọn indicator
- Không đưa indicator nào vào nếu không xác định được lag time và nguồn dữ liệu
- Trọng số các leading indicators phải cộng lại = 100%
- Phải cập nhật lại mỗi khi phân tích mới — không copy nguyên từ ticker khác

### 5.3 Quy trình sau khi hoàn thành Excel

```bash
# Bắt buộc: chạy recalc và kiểm tra 0 lỗi
cd /sessions/.../mnt/.claude/skills/xlsx
python3 scripts/recalc.py "<path_to_xlsx>"

# Kết quả phải đạt:
{"status": "success", "total_errors": 0}
```

---

## 6. BƯỚC 4: VIẾT DOCX RESEARCH REPORT

> **Nguyên tắc cốt lõi**: DOCX là sản phẩm ASSEMBLY — lắp ghép các artifacts đã phân tích từ 6 tầng, KHÔNG phải viết lại từ đầu. Nếu tầng nào chưa có artifact → PHẢI hoàn thành trước, không được viết DOCX trước để bù sau.

---

### 6.0 ĐIỀU KIỆN TIÊN QUYẾT — Không được mở file DOCX nếu chưa có đủ

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

---

### 6.1 THỨ TỰ BUILD BẮT BUỘC

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

---

### 6.2 ASSEMBLY GUIDE — Từng trang DOCX

#### 🔵 Trang 1 — Cover Page

| Element | Nội dung bắt buộc |
|---|---|
| Tên công ty (tiếng Việt đầy đủ) | VD: "CÔNG TY CỔ PHẦN GEMADEPT" |
| Mã + Sàn | "GMD \| HOSE" |
| Loại báo cáo | "Báo cáo Phân tích Đầu tư" |
| Ngày | "Tháng MM/YYYY" |
| Bảng key metrics (4 ô) | Giá hiện tại \| Giá mục tiêu \| Upside \| Khuyến nghị |
| Bảng thông số (4 ô) | Vốn hóa \| P/E forward \| EV/EBITDA \| Net Cash/(Debt) |

---

#### 🔴 Trang 2 — Investment Summary (VIẾT SAU CÙNG)

Xem template đầy đủ tại **Mục 6.3** bên dưới.

**Quy tắc cứng:**
- Phải đứng độc lập — người đọc chỉ trang này đủ hiểu thesis
- Không được dài hơn 1 trang A4
- Không được dùng câu mơ hồ kiểu "có thể tích cực hoặc tiêu cực"
- Phải có BUY / HOLD / SELL / TÍCH LŨY / QUAN SÁT rõ ràng

---

#### 🟢 Trang 3–4 — Chuỗi giá trị & Mô hình kinh doanh

**Nguồn**: Artifact Tầng 1 (bảng chuỗi giá trị)

**Phải có:**
- Bảng chuỗi giá trị (từ artifact Tầng 1) — KHÔNG viết lại, INSERT trực tiếp
- Đoạn giải thích: công ty capture giá trị ở đâu, tại sao, biên như thế nào
- Mô hình doanh thu: công thức và drivers chính (từ Mục 7 phương pháp ngành)
- Bảng cơ cấu doanh thu theo mảng (nếu đa mảng)

**Không được:**
- ❌ Mô tả chung chung về ngành không liên quan đến vị trí của công ty
- ❌ Copy Wikipedia/tóm tắt Wikipedia về ngành

---

#### 🟢 Trang 5–6 — Phân tích thị trường & Cạnh tranh

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

---

#### 🟢 Trang 7–8 — Lợi thế cạnh tranh & Quản trị

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

---

#### 🟢 Trang 9–11 — Phân tích tài chính

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

---

#### 🟢 Trang 12–13 — Định giá

**Nguồn**: Artifact Tầng 6 (valuation table + sensitivity + consensus)

**Phải có:**
- Bảng Valuation Bear/Base/Bull — INSERT từ artifact Tầng 6
- Tổng xác suất = 100% (viết rõ: "Bear 25% + Base 50% + Bull 25% = 100%")
- Weighted average target price tính được
- Sensitivity table 5×5 — INSERT từ artifact Tầng 6
- Bảng consensus CTCK (nếu có ≥2 báo cáo)
- So sánh multiple với peers (P/E, EV/EBITDA)

**Không được:**
- ❌ Viết range giá mục tiêu rộng hơn 15% (VD: "60.000–90.000" là quá rộng)
- ❌ Bỏ qua sensitivity — người đọc cần biết giả định nào quan trọng nhất
- ❌ Xác suất kịch bản không cộng đủ 100%

---

#### 🟢 Trang 14–15 — Rủi ro & Catalyst

**Nguồn**: Artifact Tầng 5 (risk matrix + catalyst table)

**Phải có:**
- Bảng Risk Matrix — INSERT từ artifact Tầng 5 (ít nhất 5 rủi ro)
- Bảng Catalyst Timeline — INSERT từ artifact Tầng 5 (ít nhất 3 catalyst)
- Với mỗi rủi ro: mô tả + ngưỡng kích hoạt + xác suất + tác động định lượng + leading indicator
- Với mỗi catalyst: sự kiện cụ thể + thời điểm + tác động giá ước tính + xác suất

**Không được:**
- ❌ Rủi ro mơ hồ kiểu "tình hình kinh tế xấu đi" — phải cụ thể và đo được
- ❌ Thiếu leading indicator cho mỗi rủi ro

---

#### 🟢 Trang 16 — Leading Indicators

**Nguồn**: Sheet 13 Excel (Framework 7 bước)

**Phải có:**
- Điểm scoring tổng hợp từ Sheet 13 (copy kết quả, không tóm tắt lại)
- Bảng 6 leading indicators với giá trị hiện tại + ngưỡng + điểm
- Điểm tổng có trọng số → mapping sang khuyến nghị (theo thang Mục 5.3 Bước 6)
- Hướng dẫn khi nào điều chỉnh khuyến nghị (ngưỡng downgrade/upgrade)

---

#### 🟢 Trang 17+ — Phụ lục (BẮT BUỘC, không được bỏ qua)

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

---

### 6.3 INVESTMENT SUMMARY — Template bắt buộc (Trang 2)

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

---

### 6.4 Yêu cầu chất lượng DOCX

- **Ngôn ngữ**: Tiếng Việt hoàn toàn (giữ tiếng Anh cho thuật ngữ chuyên môn không có từ VN tương đương: EV, EBITDA, NIM, ROIC, WACC, moat, DCF...)
- **Số liệu**: Luôn ghi nguồn và ngày lấy dữ liệu — không có số "trên trời"
- **Kết luận**: Mỗi section phải kết thúc bằng 1 câu kết luận rõ ràng (positive/negative/neutral)
- **Không dùng cụm mơ hồ**: "có thể", "nhiều khả năng", "tùy thuộc vào" → phải quantify
- **Tối thiểu 15 trang**, không tính cover và phụ lục
- **Tất cả bảng**: số liệu căn phải, mô tả căn trái, header căn giữa (xem Mục 10)
- **Không có ô trống không giải thích** — nếu chưa có dữ liệu, ghi "[chưa xác nhận]"

---


## 7. PHƯƠNG PHÁP DỰ BÁO DOANH THU THEO NGÀNH

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
Doanh thu = Thị trường tổng × Thị phần DGW
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

Môi giới = KLGD thị trường × Thị phần KSV × Phí bình quân
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

## 8. PHƯƠNG PHÁP ĐỊNH GIÁ THEO NGÀNH

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
1. Xác định giai đoạn chu kỳ → chọn method phù hợp

2. Tính EV / Equity Value:
   EV = Vốn hóa + Nợ ròng + Minority interest
   Equity Value = EV - Nợ ròng

3. Tính Target Price:
   Target Price = Equity Value / Số CP lưu hành

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

## 9. QUY CHUẨN EXCEL MODEL

### NGÔN NGỮ — BẮT BUỘC

> **Tất cả text trong Excel PHẢI dùng tiếng Việt đầy đủ dấu.** Không được dùng ASCII không dấu (VD: "BANG CAN DOI KE TOAN" là SAI — phải là "BẢNG CÂN ĐỐI KẾ TOÁN").

```
✅ ĐÚNG: "Doanh thu thuần", "Lợi nhuận gộp", "Phải thu khách hàng"
❌ SAI:  "Doanh thu thuan", "Loi nhuan gop", "Phai thu khach hang"

✅ ĐÚNG: "Giả định", "Tầng phân tích", "Kịch bản cơ sở"
❌ SAI:  "Gia dinh", "Tang phan tich", "Kich ban co so"
```

Python/openpyxl hỗ trợ Unicode đầy đủ — không cần font đặc biệt như PDF. Chỉ cần gõ thẳng tiếng Việt vào string Python.

---

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
| `13_Theo dõi` | 3 cột ngưỡng: Tốt=xanh (`'pos'`), Trung lập=vàng (`'neu'`), Xấu=đỏ (`'neg'`); Downgrade header=đỏ, Upgrade header=xanh |

### Màu sắc cấu trúc (tham chiếu nhanh)

### Quy tắc formula và logic tính toán bắt buộc

#### 1. Định dạng công thức Excel (Tránh lỗi `#REF!`)
- **Nghiêm cấm** hardcode tên sheet không có số thứ tự hoặc thiếu dấu nháy đơn (ví dụ: `Assumptions!E12` hoặc `02_Assumptions!E12` là SAI và sẽ gây lỗi `#REF!`).
- **Bắt buộc** sử dụng đúng tên sheet có số thứ tự dạng `'02_Assumptions'!` hoặc `'05_Assumptions'!`, và **bắt buộc bọc trong nháy đơn** để công thức Excel không bị lỗi tham chiếu.

```python
# ❌ SAI — hardcode số vào formula hoặc thiếu nháy đơn/số thứ tự sheet
=B7*1.15
=B7*(1+Assumptions!B12)
=B7*(1+02_Assumptions!B12)

# ✅ ĐÚNG — tham chiếu có số thứ tự sheet và bọc trong nháy đơn
=B7*(1+'02_Assumptions'!$B$12)
=B7*(1+'05_Assumptions'!$B$12)

# ❌ SAI — số âm dùng dấu trừ
-1,234

# ✅ ĐÚNG — số âm dùng ngoặc đơn
(1,234)

# ❌ SAI — zero hiển thị 0
0

# ✅ ĐÚNG — zero hiển thị dấu gạch
- (dùng format #,##0;(#,##0);-)
```

#### 2. Logic tính toán P/E và P/B
- **Lấy dữ liệu quý**: Lấy P/E, P/B hàng quý của từng năm.
- **Xử lý LNST âm (PE Carry Forward)**: Đối với quý nào có Lợi nhuận sau thuế (LNST) bị âm (lỗ), P/E của quý đó bắt buộc phải lấy bằng P/E của quý trước liền kề có LNST dương (PE carry forward). Không được chia cho số âm hoặc bằng 0, không được bỏ trống.
- **Màu sắc trên biểu đồ lịch sử PE/PB**: Trên biểu đồ, các điểm dữ liệu PE bị carry forward này bắt buộc phải được vẽ bằng màu sắc khác biệt rõ rệt (ví dụ màu đỏ hoặc cam) để người dùng dễ nhận biết.
- **Median năm**: P/E và P/B của mỗi năm phải là trung vị (median) của 4 quý thuộc năm đó.
- **Đầy đủ dữ liệu median lịch sử**: Bắt buộc phải điền đầy đủ chuỗi P/E và P/B median lịch sử cho toàn bộ các năm (tối thiểu 5 năm gần nhất) vào cả sheet `02_Assumptions` (dòng P/B mục tiêu và P/E median) và sheet `06_Ratios` để người dùng theo dõi tổng quan. Không được bỏ trống bất kỳ năm nào trong quá khứ.
- **Median dùng cho định giá (all-time median)**: P/E, P/B và EV/EBITDA dùng cho định giá chung phải là trung vị (median) của toàn bộ chuỗi thời gian lịch sử được thống kê.
- **Xuất biểu đồ PE PB**: Bắt buộc phải xuất biểu đồ line chart lịch sử P/E & P/B trực tiếp vào sheet `Charts` trong Excel và chèn hình ảnh biểu đồ vào PDF báo cáo.

### CÁC QUY TẮC CỐT LÕI BẮT BUỘC KHÁC
- **Bắt buộc dùng công thức Excel liên kết động**: Không được hardcode bất kỳ ô tính toán tỷ số tài chính (NIM, ROE, LDR, CASA, CIR...) hay kết quả định giá, bảng độ nhạy (COE x g) nào trong file Excel.
- **Bắt buộc dùng công thức Excel liên kết động cho toàn bộ các năm dự báo (forecast)**: Tại các sheet báo cáo tài chính dự báo như P&L (ví dụ: `04_PnL`), Balance Sheet (ví dụ: `05_Balance_Sheet`), và Income Model (ví dụ: `03_Income_Model`), tuyệt đối KHÔNG được phép sử dụng các con số ước tính cứng (hardcoded values) cho các năm tương lai (ví dụ: 2026F, 2027F, 2028F). Toàn bộ các ô này phải được viết bằng công thức Excel để liên kết chặt chẽ với sheet Assumptions (`'02_Assumptions'`) và liên kết giữa các chỉ tiêu với nhau. Người dùng phải có thể kích vào từng ô để kiểm chứng logic tính toán (ví dụ: công thức tính chi phí dự phòng = dư nợ bình quân x Credit Cost, công thức tính thuế TNDN = MAX(LNTT x Thuế suất, 0), công thức tính Vốn chủ sở hữu = Vốn chủ sở hữu năm trước + LNST năm nay, và công thức tính Nợ khác làm số dư để tự cân đối Bảng cân đối kế toán...). Toàn bộ các ô này phải sử dụng công thức Excel liên kết động (ví dụ: công thức tính P/B hiện tại, PV của RI, Continuing Value và bảng độ nhạy phải tự động tính toán dựa trên các ô tham chiếu).
- **Bắt buộc đệm đầy đủ số liệu lịch sử (Padding Historical Data) khi tạo các sheet báo cáo**: Khi ghi mảng dữ liệu ra file Excel bằng Python (đặc biệt là mảng `inc_model` cho sheet `03_Income_Model`), toàn bộ các biến chứa chuỗi dự phóng (ví dụ: `iea_fc`, `nii_fc`, `opex_fc`) PHẢI được nối (concatenate) sau chuỗi dữ liệu lịch sử tương ứng (ví dụ: `nii_hist + nii_fc`, hoặc `[None]*5 + iea_growth_fc`). Tuyệt đối không đẩy trực tiếp mảng dự phóng thiếu phần đệm lịch sử vào dòng ghi dữ liệu, dẫn đến việc cột dự phóng bị lệch lùi sang các năm lịch sử. Đồng thời, chỉ tiêu 'IEA bình quân' bắt buộc phải dùng trung bình cộng đầu năm và cuối năm, thay vì số dư cuối năm.
- **Dùng biểu đồ miền chồng cho Cơ cấu tài sản**: Đối với việc thể hiện cơ cấu tài sản sinh lời qua các năm (Earning Assets Structure), bắt buộc phải dùng **Biểu đồ miền chồng 100% (100% Stacked Area Chart)** thay vì vẽ nhiều hình tròn (pie charts) riêng biệt để tiết kiệm không gian và tránh bị méo mó hình ảnh khi hiển thị. Biểu đồ phải có chú giải (legend) đầy đủ và hiển thị nhãn % tại các mốc năm.
- **Định dạng Font chữ tiếng Việt trong bảng ReportLab**: Khi tạo các bảng biểu (Table) bằng thư viện ReportLab, bắt buộc phải thiết lập style `('FONTNAME', (0,0), (-1,-1), FONT_REG)` ở đầu danh sách kiểu dáng của bảng để áp dụng font hỗ trợ tiếng Việt (Arial, Calibri...) cho toàn bộ các ô. Nếu không thiết lập, ReportLab sẽ dùng font mặc định Helvetica không hỗ trợ tiếng Việt, dẫn đến các ký tự có dấu bị lỗi hiển thị thành ô vuông đen `■`.

### QUY CHUẨN XỬ LÝ FONT TIẾNG VIỆT (TRÁNH LỖI FONT PDF VÀ BIỂU ĐỒ)
Khi tạo báo cáo PDF hoặc vẽ biểu đồ, **bắt buộc** phải tuân thủ việc xử lý font chữ tiếng Việt như sau để tránh lỗi hiển thị ô vuông hoặc mất chữ có dấu:

1. **Đối với biểu đồ Matplotlib**:
   - Trước khi vẽ biểu đồ, bắt buộc phải cài đặt cấu hình font hỗ trợ tiếng Việt:
     ```python
     import matplotlib.pyplot as plt
     import matplotlib.font_manager as fm
     
     # Đặt font hệ thống hỗ trợ tiếng Việt
     plt.rcParams['font.family'] = 'sans-serif'
     plt.rcParams['font.sans-serif'] = ['Arial', 'Segoe UI', 'Tahoma', 'DejaVu Sans', 'sans-serif']
     ```
   - Tuyệt đối không để mặc định của matplotlib vì matplotlib không hỗ trợ font tiếng Việt mặc định.

2. **Đối với ReportLab (Xuất PDF)**:
   - Nghiêm cấm sử dụng các font mặc định của ReportLab như `Helvetica`, `Times-Roman`, `Courier` vì chúng không hỗ trợ ký tự Unicode tiếng Việt.
   - Bắt buộc phải đăng ký font chữ TrueType (.ttf) hỗ trợ tiếng Việt từ thư mục font hệ thống (ví dụ: `C:/Windows/Fonts/arial.ttf` trên Windows hoặc `/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf` trên Linux):
     ```python
     from reportlab.pdfbase import pdfmetrics
     from reportlab.pdfbase.ttfonts import TTFont
     
     # Đăng ký font
     font_path = "C:/Windows/Fonts/arial.ttf" # Hoặc đường dẫn font tương ứng trên hệ điều hành
     pdfmetrics.registerFont(TTFont('Arial', font_path))
     pdfmetrics.registerFont(TTFont('Arial-Bold', "C:/Windows/Fonts/arialbd.ttf"))
     ```
   - Sử dụng font đã đăng ký này (`Arial`, `Arial-Bold`) trong toàn bộ các định nghĩa Style, Paragraph và Table trong ReportLab.

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

## 10. QUY CHUẨN DOCX REPORT

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

## 11. DANH SÁCH CỔ PHIẾU ĐÃ PHÂN TÍCH

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

## 12. CHECKLIST KIỂM TRA TRƯỚC KHI GIAO

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

*Tài liệu này được cập nhật liên tục theo kinh nghiệm phân tích thực tế. Version 2.2 — 06/2026.*
*Thay đổi v2.1: (1) Chuyển output báo cáo từ PDF → DOCX; (2) Thêm quy chuẩn bảng biểu (căn lề, độ rộng cột, wrap text); (3) Thêm Sheet 13 — Các yếu tố cần theo dõi vào cấu trúc Excel.*
*Thay đổi v2.3: (7) Thêm OUTPUT BẮT BUỘC box sau mỗi trong 6 tầng phân tích — gắn kết tầng với section DOCX cụ thể, có ví dụ GMD điền đầy đủ; (8) Thay thế Mục 6 bằng kiến trúc Assembly Guide — Pre-flight checklist, Build Order bắt buộc, Assembly guide từng trang với source/must-have/forbidden, Investment Summary template có ví dụ GMD, Phụ lục bắt buộc.
*Thay đổi v2.2: (4) Thêm Mục 2.5 — Quy tắc bắt buộc phân kỳ Actual (A) vs Estimated (E) với bảng ví dụ và checklist; (5) Mở rộng Mục 2.3 nguồn dữ liệu: thêm IR website doanh nghiệp, 10 CTCK với URL cụ thể (SSI, VCSC, VCBS, HSC, MBS, VDSC, KIS, Mirae, Yuanta, BSC), nguồn hàng hóa quốc tế và dữ liệu vĩ mô VN; (6) Cập nhật checklist Section 12 — bổ sung kiểm tra A/E và yêu cầu ≥ 2 CTCK.*
