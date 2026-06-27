---
name: ngan-hang
description: Use ONLY when analyzing Vietnamese banking stocks. Front-loaded keywords: ngân hàng, bank, banking, NIM, NPL, CASA, tín dụng, huy động, TCB, ACB, VCB, MBB, VPB, BID, CTG, HDB, VIB, LPB, OCB, MSB, SSB, SHB, EIB, STB, NAB, BAB.
---

# KĨ NĂNG PHÂN TÍCH NGÀNH NGÂN HÀNG

> **CẤU TRÚC BÁO CÁO CHUẨN (MBB format)**
> Tham khảo báo cáo mẫu: `Hoc mau phan tich/Mau phan tich MBB.md` và `Mau phan tich mbb 2.md`

1. **Thông tin CP & Khuyến nghị**: Giá, vốn hóa, KLGD, 52w high/low, P/E, P/B, %NN, beta, biến động 1T/3T/12T vs VNINDEX
2. **Luận điểm đầu tư**: 3-4 lý do chính, kèm định lượng
3. **Tín dụng & Huy động**: Credit growth, loan mix (retail/corporate/BĐS), deposit growth, LDR
4. **NIM Analysis**: YOEA, COF, NIM trend, earning assets structure
5. **CASA & Funding**: CASA ratio, absolute CASA, funding breakdown, cost of funds
6. **Chất lượng tài sản**: NPL, Group 2 combined, LLR, credit cost, tái cơ cấu, CIC debt
7. **Định giá**: P/B + Residual Income weighted, so sánh consensus CTCK
8. **Bảng dữ liệu TC chi tiết**: KQKD NH (NII/Non-II/Provision/LNST) + Bảng chỉ tiêu chính + Dự phóng 3A+2E

---

## 1. Chất lượng tài sản (30%) — Quan trọng nhất

**Nợ xấu (NPL)**: <1% xuất sắc | 1-2% tốt | 3% cần chú ý

**Tỷ lệ bao phủ nợ xấu (LLR = Dự phòng / Nợ xấu)**: >200% rất mạnh | 100-200% tốt | <100% theo dõi

**Nợ nhóm 2** — quan trọng hơn NPL. Nợ nhóm 2 tăng mạnh báo hiệu 6-12 tháng sau NPL tăng.

**NPL + Nhóm 2 combined**: Tỷ lệ (Nợ nhóm 2 + Nợ xấu) / Tổng dư nợ. Chỉ số toàn diện hơn NPL đơn thuần. MBB mẫu luôn track cả 2.

**Nợ liên đới CIC**: Khoản vay bị hạ nhóm do khách hàng liên quan (CIC connection) — thường là hiệu ứng đám đông, có thể hồi phục sau 1-2 quý khi đàm phán với ngân hàng liên quan.

**Nợ tái cơ cấu**: Các khoản vay được cơ cấu lại thời hạn trả nợ. Theo dõi riêng biệt, có thể chuyển thành nợ xấu nếu tái cơ cấu thất bại.

**Công thức**:
- NPL = Nợ xấu (Nhóm 3+4+5) / Tổng dư nợ cho vay KH
- Nợ nhóm 2 / Tổng dư nợ cho vay — nếu NPL giảm nhưng nợ nhóm 2 tăng → cảnh báo
- NPL Combined = (Nợ nhóm 2 + Nợ nhóm 3+4+5) / Tổng dư nợ
- Lãi & phí phải thu / Tổng tài sản: >2% cảnh báo rủi ro nợ xấu tiềm ẩn

**Bộ 3 chỉ số cần theo dõi song song**: Nợ nhóm 2, NPL, LLR.

---

## 2. Tăng trưởng quy mô (25%)

### Tăng trưởng tín dụng

Phân biệt 2 khái niệm:
1. "Tín dụng toàn nền kinh tế" (NHNN) — rộng, nhiều hình thức cấp tín dụng
2. "Dư nợ cấp tín dụng" (Luật TCTD) — bao gồm dư nợ cho vay + trái phiếu doanh nghiệp

→ **Total earning credit** = Loan book + Corporate bond book. Luôn tách riêng.
→ CoC, NPL, nợ nhóm 2 đều tính trên **tổng tín dụng** (cho vay + trái phiếu TCKT).

**Loan mix** (cơ cấu cho vay) — phân tích 3 nhóm:
- **Cho vay cá nhân (Retail)**: Biên cao hơn, rủi ro phân tán, CASA từ cá nhân thường tốt
- **Cho vay doanh nghiệp (Corporate)**: Quy mô lớn, biên thấp hơn, rủi ro tập trung
- **Cho vay BĐS & Xây dựng**: Rủi ro chu kỳ, cần theo dõi tỷ trọng và tăng trưởng

**CAR (Capital Adequacy Ratio)**: Theo dõi để đánh giá dư địa tăng trưởng tín dụng. >11% an toàn, <10% cần theo dõi.

### Tăng trưởng huy động

| Khái niệm | Công thức | Ghi chú |
|---|---|---|
| **Tiền gửi KH** | Trực tiếp từ BCTC | Đơn giản nhất |
| **CASA** | TG không kỳ hạn VNĐ / Tổng TG KH | Chi phí vốn thấp, NIM cao |
| **CASA2** | (TG không kỳ hạn VNĐ + TG không kỳ hạn USD + TG có kỳ hạn USD) / Tổng TG KH | Phản ánh sâu hơn chất lượng vốn giá rẻ (VD: VCB có USD lớn) |
| **Market 1 Funding** | Tiền gửi KH + Giấy tờ có giá (CD, trái phiếu, kỳ phiếu) | Đầy đủ nhất |
| **LDR Thuần túy (Simple LDR)** | Cho vay khách hàng / Tiền gửi khách hàng | LDR đơn giản từ API (thường bị vọt lên >100% như TCB ~130% do chưa tính Giấy tờ có giá CD/Trái phiếu ở mẫu số) |
| **LDR điều chỉnh (NHNN)** | Cho vay khách hàng / (Tiền gửi KH + Giấy tờ có giá CD + Tiền gửi KBNN theo lộ trình) | LDR pháp lý theo Thông tư 26, NHNN quy định trần <85% |


### Earning Assets Structure (Tài sản sinh lãi)
- Cho vay KH (chiếm tỷ trọng lớn nhất)
- Tiền gửi NHNN có lãi
- Tài sản sinh lãi liên NH
- Trái phiếu đầu tư (bao gồm TPDN — cần tách riêng để đánh giá rủi ro)

### TPDN (Trái phiếu doanh nghiệp) Tracking
- Quy mô TPDN / Tổng tài sản — nếu >5% cần phân tích chi tiết danh mục
- Tăng trưởng TPDN YoY — xu hướng tăng/giảm
- Rủi ro tập trung vào ngành BĐS — cần theo dõi chất lượng trái phiếu nắm giữ

### Chênh lệch tín dụng vs huy động
Tín dụng tăng > Huy động tăng → áp lực thanh khoản, COF tăng, NH phải huy động vốn liên NH/GTCG chi phí cao → NIM giảm.

---

## 3. NIM Framework (Biên lãi ròng)

### Công thức cốt lõi

NIM = YOEA − COF

Trong đó:
- **YOEA** (Yield on Earning Assets) = Thu nhập lãi / Earning Assets bình quân
- **COF** (Cost of Funds) = Chi phí lãi / Funding bình quân
- **NIM** (Net Interest Margin) = Thu nhập lãi thuần / Earning Assets bình quân

### Earning Assets (Tài sản sinh lãi)
Gồm: Cho vay KH + Tiền gửi NHNN có lãi + TS sinh lãi liên NH + Trái phiếu đầu tư sinh lãi (trừ VAMC)
Không gồm: Tiền mặt, TSCĐ, BĐS, tài sản thuế hoãn lại

### Funding (Nguồn vốn)
- Tiền gửi KH (CASA + TG có kỳ hạn)
- Giấy tờ có giá (CD, trái phiếu, kỳ phiếu)
- Vay liên NH
- Vốn huy động khác

CASA là nguồn vốn rẻ nhất → CASA cao = COF thấp = NIM cao.

### 3 kịch bản NIM

| YOEA | COF | NIM | Nguyên nhân |
|---|---|---|---|
| ↑ | ↓ | ↑ **NIM mở rộng** | Lãi cho vay tăng, LS huy động giảm — chu kỳ lý tưởng |
| ↑ | ↑ | → **NIM ổn định** | Cả 2 tăng tương ứng — thị trường bình thường |
| ↓ | ↑ | ↓ **NIM thu hẹp** | Lãi cho vay giảm nhanh hơn LS huy động — áp lực cạnh tranh |

### Cơ chế NIM compression
- Khi LS cho vay giảm nhanh hơn chi phí huy động (do cạnh tranh, NHNN giảm lãi)
- Khi nợ chậm trả gia tăng (khách hàng không trả lãi → giảm thu nhập lãi)
- Khi huy động chuyển dịch từ CASA sang TG có kỳ hạn (CASA↓ → COF↑)

### NIM theo dõi
- >4% rất tốt | 3-4% khá | <3% yếu
- Luôn so sánh YoY và QoQ để thấy trend

---

## 4. PPOP & Earnings Quality

### Cấu trúc thu nhập ngân hàng

```
NII (Thu nhập lãi thuần)
+ Non-II (Thu nhập ngoài lãi)
  + NFI (Thu nhập dịch vụ — phí)
  + Lãi ngoại hối & vàng
  + Lãi đầu tư chứng khoán
  + Thu nhập khác (thu hồi nợ xấu, bán tài sản...)
= TOI (Tổng thu nhập hoạt động)
- Chi phí hoạt động (OPEX)
= PPOP (Pre-Provision Operating Profit)
- Chi phí dự phòng rủi ro
= LNTT
- Thuế
= LNST
```

### Phân biệt tăng trưởng thực vs. ảo

- **PPOP** (Pre-Provision Operating Profit) = Thu nhập hoạt động - Chi phí hoạt động
- PPOP↑ 10%, LNST↑ 30% → nhờ giảm dự phòng (không bền vững)
- PPOP↑ 20%, LNST↑ 10% → chủ động trích lập mạnh (chất lượng hơn)
- PPOP tăng + LNST tăng tương ứng → tăng trưởng thực chất

**Cost of Credit (CoC)** = Dự phòng tín dụng / Dư nợ bình quân. Nếu LN bùng nổ nhưng CoC giảm mạnh → đặt câu hỏi.

**So sánh LNST vs tăng trưởng tín dụng**: LNST dài hạn ≈ Tăng trưởng tín dụng × Hiệu quả sinh lời.

**Loại bỏ thu nhập bất thường**: Bán công ty con, hoàn nhập dự phòng, bancassurance trả trước, thu hồi nợ đã xử lý.

**Ưu tiên ROE hơn lợi nhuận tuyệt đối**.

### Bộ tiêu chí 4 tầng đánh giá chất lượng tăng trưởng

| Tầng | Chỉ tiêu |
|---|---|
| **1. Tăng trưởng** | LNST, PPOP quý/lũy kế YoY; CAGR LNST 3-5 năm; CAGR PPOP 3-5 năm |
| **2. Chất lượng tài sản** | CoC, LLR, Nợ xấu/tổng tín dụng, Nợ nhóm 2/tổng tín dụng |
| **3. Hiệu quả** | ROE, ROA, NIM |
| **4. Nguồn tăng trưởng** | Tín dụng↑ / CASA↑ / NIM↑ / Dịch vụ↑ (tốt) vs Dự phòng↓ / Hoàn nhập↑ / Bất thường↑ (xấu) |

---

## 5. CASA Analysis

### 2 chiều phân tích CASA

1. **CASA ratio** (%) = TG không kỳ hạn / Tổng TG KH — tỷ lệ, so sánh với peer
2. **CASA absolute** (tỷ đồng) — quy mô tuyệt đối, biến động QoQ/YoY

CASA giảm cả 2 chiều → COF tăng → NIM giảm.
CASA giảm ratio nhưng tăng absolute → vẫn ổn (tăng trưởng TG có kỳ hạn nhanh hơn).

### Phân tích CASA chi tiết

- **CASA từ cá nhân vs doanh nghiệp**: Cá nhân thường ổn định hơn, DN nhạy lãi suất hơn
- **CASA2**: Mở rộng CASA bao gồm cả TGKH USD — phản ánh đúng chi phí vốn thực (VD: VCB có USD lớn)
- **Digital penetration**: Số lượng KH cá nhân, tỷ lệ KH giao dịch online → leading indicator cho CASA

### CASA Benchmark VN

| Mức | Ý nghĩa |
|---|---|
| >35% | Rất tốt — top ngành (MBB, TCB, VCB) |
| 25-35% | Khá |
| 15-25% | Trung bình |
| <15% | Yếu — phụ thuộc vốn liên NH |

---

## 6. Hiệu quả hoạt động (15%)

### ROE
- >20% xuất sắc | 15-20% tốt | <10% yếu
- Công thức: LNST / VCSH bình quân (đầu+cuối năm)/2 hoặc tổng LNST 4 quý / VCSH BQ 4 quý

### NIM (Biên lãi ròng) — xem chi tiết §3

### CIR (Cost-to-Income)
- <35% rất mạnh | 50% khá tệ
- Công thức: Chi phí hoạt động / Thu nhập hoạt động

**Lưu ý quan trọng**:
- CIR thấp chưa chắc tốt (có thể do cắt giảm đầu tư, nhân sự)
- Không so sánh bank bán lẻ (VPB, HDB, VIB — CIR cao) với bank doanh nghiệp (TCB, MBB, ACB — CIR thấp)
- Kiểm tra nguồn giảm CIR: thu nhập↑ > chi phí↑ (tốt) vs cắt giảm chi phí (xấu)
- Cẩn thận với thu nhập bất thường (bán công ty con, bancassurance trả trước)
- Nên xem CIR BQ 3-5 năm, kết hợp với tăng trưởng TOI và PPOP

### CAR (Capital Adequacy Ratio)
- >12% an toàn cho tăng trưởng tín dụng
- <10% NH có thể bị hạn chế room tín dụng
- Theo dõi CAR khi NH tăng trưởng nóng → CAR giảm

---

## 7. Định giá (20%)

**P/B là chính, không phải P/E**:
- <1x: rẻ
- 1-1.5x: hợp lý
- >2x: phải có tăng trưởng cao

**P/B lịch sử** quan trọng hơn P/B hiện tại. So sánh với trung bình/trung vị 5 năm.

**Phương pháp định giá kết hợp**: Residual Income + P/B. Không dùng EV (ngân hàng không có EV thông thường).

### Residual Income (Thu nhập thặng dư) — Step-by-step

**Công thức**:
```
Equity Value = Book Value + PV của Residual Income (RI)
RI(t) = (ROE(t) - Ke) × BV(t-1)
PV của Equity = BV(0) + Σ RI(t)/(1+Ke)^t + Terminal Value/(1+Ke)^n
```

**Các bước**:
1. Dự phóng ROE, BVPS cho 5 năm (giai đoạn tăng trưởng)
2. Xác định Ke (chi phí vốn CSH — thường 12-15% cho NH VN)
3. Tính RI mỗi năm = (ROE - Ke) × BV đầu kỳ
4. PV các RI
5. Tính Terminal Value: RI(n) × (1+g) / (Ke - g) với g ~3-5%
6. PV của Equity = BV hiện tại + PV(RI) + PV(TV)
7. Giá mục tiêu = PV của Equity / Số CP lưu hành

**Trọng số khuyến nghị**:
- 40-60% Residual Income
- 60-40% P/B (median 5 năm)
- P/B dùng P/B median lịch sử × BVPS forward

### So sánh consensus CTCK

| Nguồn | Khuyến nghị | Giá mục tiêu | P/B target |
|---|---|---|---|
| Báo cáo này | [tự đánh giá] | [TP] | [P/B] |
| SSI Research | | | |
| VCBS | | | |
| MBS | | | |

---

## 8. Harvest Signs — Dấu hiệu hái quả ngọt

1. **NIM bottom + expansion**: COF đạt đỉnh, bắt đầu giảm → NIM mở rộng trở lại. Dấu hiệu: lãi suất huy động ngừng tăng, CASA hồi phục.

2. **NPL cycle peak + LLR high**: Nợ xấu đạt đỉnh, LLR bắt đầu tăng trở lại >100%. Dấu hiệu: Nợ nhóm 2 giảm, tái cơ cấu thành công, xóa nợ xấu giảm.

3. **CASA recovery**: CASA ratio và absolute CASA cùng tăng trở lại. Dấu hiệu: lãi suất TG có kỳ hạn giảm, KH cá nhân mới tăng mạnh.

4. **Credit cost normalizing**: CoC giảm từ đỉnh, chi phí dự phòng ổn định. Dấu hiệu: NPL mới giảm, LLR duy trì >100%.

**Khi đủ 3/4 harvest signs → NH bước vào pha thu hoạch (harvest phase).**

---

## 9. Bảng dữ liệu chuẩn NH

### Bảng 1: Kết quả kinh doanh NH (tỷ đồng)

| Kết quả kinh doanh | N-2A | N-1A | N E | N+1E | N+2E |
|---|---|---|---|---|---|
| Thu nhập lãi thuần (NII) | | | | | |
| Thu nhập ngoài lãi (Non-II) | | | | | |
| - Thu nhập dịch vụ (NFI) | | | | | |
| - Lãi ngoại hối | | | | | |
| - Lãi chứng khoán | | | | | |
| - Thu nhập khác | | | | | |
| **Tổng thu nhập hoạt động (TOI)** | | | | | |
| Chi phí hoạt động | ( ) | ( ) | ( ) | ( ) | ( ) |
| **PPOP** | | | | | |
| Chi phí dự phòng | ( ) | ( ) | ( ) | ( ) | ( ) |
| **LNTT** | | | | | |
| Thuế | ( ) | ( ) | ( ) | ( ) | ( ) |
| **LNST** | | | | | |
| *Tăng trưởng LNST* | | | | | |

### Bảng 2: Các chỉ tiêu chính đo lường hiệu quả NH

| Chỉ tiêu | N-2A | N-1A | N E | N+1E | N+2E |
|---|---|---|---|---|---|
| Tăng trưởng tín dụng (%YTD) | | | | | |
| Tăng trưởng huy động (%YTD) | | | | | |
| NIM | | | | | |
| YOEA | | | | | |
| COF | | | | | |
| CASA | | | | | |
| Tỷ lệ nợ xấu (NPL) | | | | | |
| Nợ xấu + Nhóm 2 combined | | | | | |
| LLR (Dự phòng / Nợ xấu) | | | | | |
| Credit cost (CoC) | | | | | |
| CIR | | | | | |
| ROE | | | | | |
| ROA | | | | | |
| CAR | | | | | |

### Bảng 3: Dự phóng KQKD 3A + 2E

Template tương tự Bảng 1, thêm các KPIs (credit growth, NIM, CASA, NPL, LLR, CoC, CIR).
Driver chính:
- **NII** = Earning Assets × NIM — forecast từ credit growth + NIM assumption
- **Non-II** = Fee growth (số KH × phí/KH) + FX + Securities
- **Provision** = Dư nợ BQ × CoC assumption
- **LNTT** = (NII + Non-II) × (1 - CIR) - Provision

---

## 10. Đánh giá tổng quát

**Theo dõi**: Khối ngoại, Tự doanh, Quỹ ETF, Tổ chức trong nước. Tỷ lệ sở hữu NN và room ngoại còn lại.

**Dấu hiệu chu kỳ tăng giá lớn**: NPL↓ + CASA↑ + ROE >20% + P/B dưới TB lịch sử + Tổ chức gom mạnh + NIM mở rộng.

---

## 11. Lưu ý đặc thù khi phân tích ngân hàng

- **LNST điều chỉnh**: Dùng Adjusted Earnings (loại bỏ thu nhập bất thường) thay vì LNST báo cáo khi tính CAGR
- **Dư nợ cho vay vs Tín dụng**: Trái phiếu doanh nghiệp là một dạng cấp tín dụng — NHNN đã yêu cầu đưa vào giới hạn cấp tín dụng
- **BCTC ngân hàng**: Luôn xem thêm mục Chứng khoán đầu tư / Trái phiếu doanh nghiệp / Thuyết minh chứng khoán nợ
- **Mặt bằng NIM VN hiện nay**: ~3-4% là phổ biến cho ngân hàng tốt
- **Provisioning strategy**: Theo dõi số dư quỹ dự phòng và xóa nợ xấu hàng quý — NH trích mạnh khi LN cao và ngược lại
- **Rủi ro tập trung tín dụng**: Theo dõi danh sách khách hàng lớn (BĐS, TPDN) — các tên Novaland, Trung Nam, SunGroup, Vạn Thịnh Phát là rủi ro điển hình
- **Tác động của room tín dụng**: NHNN cấp room hàng năm — NH nào có M&A (nhận chuyển giao bắt buộc) thường được ưu tiên room cao hơn
- **Chênh lệch credit-deposit growth**: Credit > Deposit → áp lực thanh khoản + COF

---

## 12. Biểu đồ & Trực quan hóa (12 Chart Types — MBS format)

Mỗi báo cáo ngân hàng cần tối thiểu **10-12 biểu đồ**. Template từ mẫu ACB (MBS).

### Nhóm A: NIM & Thu nhập (Profitability)

#### A — NIM Decomposition
| Thuộc tính | Mô tả |
|---|---|
| **Mục đích** | Phân tích spread: YOEA - COF = NIM, xác định động lực thay đổi NIM |
| **Cách đọc** | YOEA↑ + COF↓ = NIM mở rộng (tốt). YOEA↓ + COF↑ = NIM compression (xấu) |
| **Dữ liệu** | `isb25/iea` → YOEA, `isb26/funding` → COF, `isb27/iea` → NIM |
| **Format** | 3 lines (YOEA, COF, NIM) trên cùng axis, NIM overlay đậm hơn |
| **Màu sắc** | YOEA=#4472C4, COF=#ED7D31, NIM=#70AD47 (đậm, d=2) |
| **Ví dụ ACB** | H8: YOEA 8.63%, COF 3.14%, NIM 5.49% |

#### B — Peer NPL Comparison
| Thuộc tính | Mô tả |
|---|---|
| **Mục đích** | So sánh NPL của bank vs peer cùng quy mô |
| **Cách đọc** | Bank có NPL thấp hơn peer → quản lý rủi ro tốt hơn |
| **Dữ liệu** | NPL của bank + 6-8 peer banks (3 quý gần nhất) |
| **Format** | Grouped bars: mỗi bank 1 cụm 3 bar (các quý) |
| **Ví dụ ACB** | H11: ACB 1.07% vs VCB 1.14%, MBB 1.43%, TCB 1.53%... |

#### C — Peer Credit Growth
| Thuộc tính | Mô tả |
|---|---|
| **Mục đích** | So sánh tốc độ tăng trưởng tín dụng với peer |
| **Cách đọc** | Bank tăng trưởng > peer → chiếm thị phần |
| **Dữ liệu** | Credit growth % của bank + 6-8 peer (YTD %) |
| **Format** | Horizontal bars, bank highlighted |
| **Ví dụ ACB** | H4: ACB 15.6% vs VCB 9.5%, MBB 13.2% |

#### D — Bank vs Industry Average
| Thuộc tính | Mô tả |
|---|---|
| **Mục đích** | Định vị bank so với trung bình ngành |
| **Cách đọc** | Bank ở trên hay dưới đường trung bình? Khoảng cách bao nhiêu? |
| **Dữ liệu** | Chỉ số của bank + trung bình ngành cho 4-5 metrics |
| **Format** | 2 lines (bank + industry) theo thời gian |
| **Ví dụ ACB** | H5: NIM, CIR, ROE vs ngành |

### Nhóm B: Chất lượng tài sản (Asset Quality)

#### E — Loan Structure (Retail/Corp/SME)
| Thuộc tính | Mô tả |
|---|---|
| **Mục đích** | Phân tích cơ cấu cho vay theo đối tượng KH |
| **Cách đọc** | Retail > 50% → NIM cao hơn, rủi ro phân tán hơn |
| **Dữ liệu** | `nob_loans_by_type` hoặc thuyết minh |
| **Format** | Stacked bar 100% |
| **Ví dụ ACB** | H6: Retail ~55%, Corp ~25%, SME ~20% |

#### F — Earning Assets Structure
| Thuộc tính | Mô tả |
|---|---|
| **Mục đích** | Cơ cấu tài sản sinh lãi (cho vay, TP đầu tư, TG NHNN, liên NH) |
| **Cách đọc** | Cho vay chiếm >70% là chuẩn. TP đầu tư >15% → có thể đang thận trọng |
| **Dữ liệu** | `bsb*` earning assets components |
| **Format** | Stacked bar % |
| **Ví dụ ACB** | H10: Loans 72%, Bonds 15%, Interbank 8% |

#### G — NPL + Group 2 Combo
| Thuộc tính | Mô tả |
|---|---|
| **Mục đích** | Theo dõi toàn diện chất lượng nợ qua NPL và nợ nhóm 2 |
| **Cách đọc** | NPL↓ + Gr2↓ = sạch. NPL↓ nhưng Gr2↑ = cảnh báo trước 6-12 tháng |
| **Dữ liệu** | NPL (tỷ), Nợ nhóm 2 (tỷ), NPL%, Gr2% |
| **Format** | Dual axis: bars (absolute) + lines (ratio) |
| **Màu sắc** | Bar NPL=#C00000, Bar Gr2=#ED7D31, Line=dark |
| **Ví dụ ACB** | H13: NPL giảm từ 1.5% → 1.07%, Gr2 ổn định |

#### H — LLR + CoC Dual
| Thuộc tính | Mô tả |
|---|---|
| **Mục đích** | Đánh giá mức độ trích lập dự phòng và xu hướng chi phí tín dụng |
| **Cách đọc** | LLR >100% + CoC ổn định → trích lập đủ. LLR <100% + CoC tăng → rủi ro |
| **Dữ liệu** | LLR% (dự phòng/NPL), CoC% (dự phòng/dư nợ BQ) |
| **Format** | 2 lines (LLR + CoC) |
| **Ví dụ ACB** | H14: LLR 150% → 138%, CoC 0.4% → 0.6% |

### Nhóm C: Tăng trưởng & Thị trường (Growth & Market)

#### I — NIM Delta Peer
| Thuộc tính | Mô tả |
|---|---|
| **Mục đích** | So sánh thay đổi NIM (bps) giữa các bank trong cùng kỳ |
| **Cách đọc** | Bank có NIM delta > 0 = đang mở rộng NIM so với peer |
| **Dữ liệu** | ΔNIM (bps) YoY hoặc QoQ của các peer |
| **Format** | Horizontal bars, bank highlighted |
| **Ví dụ ACB** | H7: ACB +16bps, VCB +5bps, MBB -22bps |

#### J — Non-II Breakdown
| Thuộc tính | Mô tả |
|---|---|
| **Mục đích** | Phân tích cấu trúc thu nhập ngoài lãi |
| **Cách đọc** | NFI (phí) chiếm tỷ trọng lớn là bền vững. Lãi CK/ngoại hối biến động mạnh |
| **Dữ liệu** | `isb*` non-interest components |
| **Format** | Stacked bar (fee, FX, securities, other) |
| **Ví dụ ACB** | H3: NFI 45%, FX 20%, CK 15%, Khác 20% |

#### K — TOI + NPAT Stacked
| Thuộc tính | Mô tả |
|---|---|
| **Mục đích** | Tổng quan thu nhập và lợi nhuận qua các năm |
| **Cách đọc** | TOI↑ + NPAT↑ cùng tốc độ → tăng trưởng chất lượng |
| **Dữ liệu** | TOI, NPAT, growth % |
| **Format** | Stacked bars TOI + NPAT overlay + growth line |
| **Ví dụ ACB** | H2: TOI CAGR 12%, NPAT CAGR 15% |

#### L — Peer Comparison Table
| Thuộc tính | Mô tả |
|---|---|
| **Mục đích** | So sánh 15 chỉ số chính giữa các bank trong ngành |
| **Cách đọc** | Bank tốt hơn peer ở chỉ số nào? Yếu hơn chỉ số nào? |
| **Dữ liệu** | 15 metrics × 8-12 banks (NIM, NPL, CIR, ROE, ROA, CAR, CASA, LDR, credit growth...) |
| **Format** | Table với conditional formatting (xanh = tốt hơn, đỏ = kém hơn peer median) |
| **Ví dụ ACB** | H16: 15 metrics × 12 banks |

#### M — PE/PB/EVEBITDA Historical
| Thuộc tính | Mô tả |
|---|---|
| **Mục đích** | Theo dõi biến động định giá theo thời gian — P/E, P/B, EV/EBITDA qua các quý |
| **Cách đọc** | P/E < 10x = rẻ, > 20x = đắt. P/B < 1x = dưới giá trị sổ sách. So sánh với trung vị 5 năm |
| **Dữ liệu** | PE, PB từ `statistics-financial` API (12+ quý gần nhất). EV/EBITDA = (MCap + Net Debt) / EBITDA |
| **Format** | Dual axis: P/E (line, #4472C4) + P/B (line, #ED7D31) + EV/EBITDA (bar, #70AD47) |
| **Màu sắc** | P/E=#4472C4, P/B=#ED7D31, EV/EBITDA=#70AD47 |
| **Yêu cầu** | **Historical only** — KHÔNG có forecast. Thêm đường reference: trung vị 5 năm (đứt nét) |
| **Ví dụ TCB** | P/E 5.8x (rẻ), P/B 0.84x (dưới 1x), EV/EBITDA ~4.5x |

---

## 13. Mẫu báo cáo chuẩn (ACB/MBS format)

Báo cáo ngân hàng chuẩn MBS có cấu trúc sau:

### 13.1 Stock Info & Recommendation
- Giá, vốn hóa, KLGD, 52w high/low, P/E, P/B, %NN, beta
- Khuyến nghị (MUA/THEO DÕI/NẮM GIỮ/BÁN) + target price

### 13.2 Investment Thesis
- 3-4 luận điểm chính, mỗi luận điểm kèm dữ liệu định lượng
- 3 rủi ro chính (giải thích tại sao, không chỉ liệt kê)

### 13.3 Detailed KQBD Table (QoQ + YoY + YTD)
Bảng KQBD với 3 chiều phân tích:
| Khoản mục | Q1/2025 | Q2/2025 | ... | QoQ | YoY | YTD | Nhận xét |
|---|---|---|---|---|---|---|---|
| NII | | | | ±% | ±% | ±% | trend |
| Non-II | ... | | | | | | |
| TOI | | | | | | | |
| PPOP | ... | | | | | | |
| Provision | ... | | | | | | |
| LNST | ... | | | | | | |

Mỗi dòng có **nhận xét ngắn** (ví dụ: "NII tăng 5.1% QoQ nhờ credit growth +18.4% YoY")

### 13.4 Asset Quality (3 charts)
- NPL + Group 2 combo (Chart G)
- LLR + CoC dual (Chart H)
- Peer NPL comparison (Chart B)

### 13.5 Credit & Deposit Analysis
- Credit growth vs peer (Chart C)
- Loan structure (Chart E)
- Credit vs deposit growth + LDR

### 13.6 NIM Deep Dive
- NIM decomposition (Chart A)
- NIM delta peer (Chart I)
- Earning assets structure (Chart F)
- Funding structure (CASA, term deposits, bonds, interbank)

### 13.7 Profitability & Efficiency
- Non-interest income breakdown (Chart J)
- PPOP analysis (TOI vs PPOP vs NPAT growth)
- CIR, ROE, ROA trend
- Bank vs industry average (Chart D)

### 13.8 Valuation
- **Residual Income model** (7-year projection):
  - ROE forecast → BVPS → RI → PV
  - Terminal value với giả định cụ thể
  - Dải nhạy cảm COE × g (5×5 ma trận)
- **P/B multiple**:
  - P/B median 5 năm
  - P/B forward so với peer
- **Weighted target** (50% RI + 50% P/B)
- Upside từ giá hiện tại

### 13.9 Forecast Assumptions (explicit table)
| Driver | 2025A | 2026F | 2027F | 2028F | Giải thích |
|---|---|---|---|---|---|
| Credit growth | 15.4% | 14% | 13% | 12% | Giảm dần do nền cao + NHNN room |
| NIM | 3.2% | 3.3% | 3.3% | 3.2% | Ổn định nhờ CASA hồi phục |
| CASA | 24% | 25% | 26% | 27% | Tăng nhờ số hóa |
| NPL | 1.1% | 1.1% | 1.2% | 1.3% | Tăng nhẹ do credit mở rộng |
| CoC | 0.7% | 0.7% | 0.8% | 0.8% | Ổn định |
| CIR | 37% | 36% | 35% | 35% | Cải thiện nhờ TOI↑ |
| D/E | 2.5x | 2.4x | 2.3x | 2.3x | Giảm dần |

### 13.10 Risks
3 rủi ro chính, mỗi rủi ro gồm:
- **Mô tả**: rủi ro gì, xảy ra khi nào
- **Tác động**: ảnh hưởng đến chỉ số nào (NIM, NPL, CASA...)
- **Kịch bản**: nếu xảy ra, LNST giảm bao nhiêu %, target price affected thế nào

### 13.11 Industry Context
- Định vị bank so với trung bình ngành (Chart D)
- So sánh peer chi tiết (Chart L - 15 metrics)
- Xu hướng ngành: room tín dụng, NHNN policy, cạnh tranh

### 13.12 SWOT & Recommendation
- SWOT 4 ô (Sáng/Tối/Cơ hội/Thách thức)
- Kết luận: khuyến nghị + target price + upside

---

## 14. Hướng dẫn Python Implementation cho Ngân hàng

### 14.1 Fetch dữ liệu Vietcap API

```python
# BẮT BUỘC: Lấy statistics-financial (PE/PB/EVEBITDA theo quý)
def fetch_vietcap_ratios(ticker):
    s = requests.Session()
    s.get(f"https://trading.vietcap.com.vn/iq/company?ticker={ticker}")
    r = s.get(f"{BASE}/company/{ticker}/statistics-financial", timeout=15)
    data = r.json()["data"]  # list of dicts: year, quarter, pe, pb, evToEbitda...
    # KHÔNG giới hạn số quý — lấy tất cả
    return sorted(data, key=lambda x: (x["year"], x["quarter"]))
```

### 14.2 Xử lý PE/PB theo quý (bắt buộc)

```python
# Với mỗi quý trong lịch sử:
pe_all = []
pb_all = []
prev_pe = None
for r in ttms:  # ttms = sorted(VAB_RATIOS, key=year+quarter)
    pe = r.get("pe")
    pb = r.get("pb", 0)
    npat = r.get("npat")  # LNST quý đó — cần từ BCTC
    # Nếu LNST âm → P/E = quý trước
    if npat is not None and npat < 0 and prev_pe is not None:
        pe = prev_pe
        pe_is_carried = True  # đánh dấu để vẽ màu khác
    pe_all.append(pe)
    pb_all.append(pb)
    if pe is not None: prev_pe = pe
```

### 14.3 Median năm + Median all-time

```python
from statistics import median

# Median từng năm
pe_by_year = {}
for r in ttms:
    y = r["year"]
    pe = r.get("pe")
    if pe is not None:
        pe_by_year.setdefault(y, []).append(pe)
pe_median_per_year = {y: median(vals) for y, vals in pe_by_year.items()}

# Median all-time cho định giá
pe_all_median = median([r["pe"] for r in ttms if r.get("pe")])
pb_all_median = median([r["pb"] for r in ttms if r.get("pb")])
evebitda_all_median = median([r["evToEbitda"] for r in ttms if r.get("evToEbitda")])

# P/B 3 mức theo phân phối lịch sử:
_pb_below = [p for p in pb_all_vals if p <= pb_all_median]
_pb_above = [p for p in pb_all_vals if p >= pb_all_median]
pb_attractive = median(_pb_below) if _pb_below else pb_all_median * 0.85
pb_target     = median(_pb_above) if _pb_above else pb_all_median * 1.15
```

### 14.4 Excel Formula Rules cho Banking

| Chỉ tiêu | Công thức Excel | Giải thích |
|---|---|---|
| CIR (%) | `='04_PnL'!E5/'04_PnL'!E4*100` | OPEX/TOI×100 |
| NIM (%) | `=YOEA-COF` hoặc =NII/Earning_Assets_BQ | Link Assumptions |
| ROE (%) | `=NPAT/Average_Equity*100` | Dùng VCSH bình quân |
| NPL (%) | `=NPL_Absolute/Loans*100` | Link Balance Sheet |
| LDR (%) | `=Loans/Deposits*100` | Link Balance Sheet |
| P/B (x) | `=Assumptions!E15` | P/B target từ assumptions |
| BV/share | `=05_Balance_Sheet!E15*1e9/Assumptions!B3` | VCSH×1e9/số CP |

### 14.5 Chart Rules

- Chart N (PE/PB Historical): Dual axis, median all-time line đứt nét
- Quý LNST âm: P/E carry từ quý trước, marker hollow circle màu đỏ
- Luôn có ít nhất 10 charts cho báo cáo ngân hàng
- Median tính trên TOÀN BỘ dữ liệu, không chỉ 12 quý

---

## 15. Execution Checklist (Step-by-step, KHÔNG SKIP)

### Bước 1: Thu thập dữ liệu
```
☐ Fetch BCTC 3-5 năm (Income Statement + Balance Sheet)
☐ Fetch statistics-financial (PE/PB/EVEBITDA tất cả quý)
☐ Fetch company details (giá, MCap, CP lưu hành)
☐ Fetch peer data (18 banks, NIM/NPL/CASA/ROE/CIR/PB)
☐ Xác định ngày phân tích → đúng nhãn A/E
```

### Bước 2: Tính toán chỉ số
```
☐ NPL, Nợ nhóm 2, LLR, CoC
☐ NIM = YOEA - COF
☐ CASA ratio + absolute
☐ CIR = OPEX/TOI
☐ ROE, ROA, LDR, CAR
☐ P/E, P/B, EV/EBITDA từng quý
☐ LNST âm → P/E carry
☐ Median năm + median all-time
```

### Bước 3: Forecast
```
☐ Credit growth (3 năm)
☐ NIM forecast (3 kịch bản)
☐ CASA forecast
☐ NPL, CoC forecast
☐ P&L forecast (NII, Non-II, OPEX, Provision, NPAT)
☐ Balance Sheet forecast (Loans, Deposits, Equity)
```

### Bước 4: Định giá
```
☐ Residual Income model (3 năm + terminal)
☐ P/B median all-time × BVPS forward
☐ Weighted: 50% RI + 50% P/B
☐ Sensitivity: COE × g (5×5)
```

### Bước 5: Excel
```
☐ 14+ sheets (01_Cover → 13_PE_PB_History + 14_...)
☐ Mọi ô forecast là formula (cross-sheet ref)
☐ Sheet 13_PE_PB_History: đủ quý + MEDIAN cuối bảng
☐ Chart N: PE/PB dual axis + median line + carry markers
☐ Công thức CIR đúng: ='04_PnL'!E5/'04_PnL'!E4*100
☐ Recalc = 0 lỗi
```

### Bước 6: PDF
```
☐ Đủ section (Cover, Thesis, Asset Quality, NIM, CASA, Valuation, Historical Valuation, SWOT)
☐ Chart images hiển thị đủ (≥10 chart)
☐ Bảng PE/PB: current vs median all-time
☐ Khuyến nghị cuối cùng + target price
```
- Nên có **sensitivity table** (COE × g 5×5 ma trận giá trị RI)

---

## 16. QUY TẮC BẮT BUỘC — TRÁNH LỖI KHI BUILD MODEL NGÂN HÀNG

> **Các lỗi này đã được phát hiện và sửa trong thực tế. BẮT BUỘC tuân thủ.**

### 16.1 Công thức NIM chuẩn (IEA bình quân đầu+cuối kỳ)

**❌ SAI — NIM lịch sử tính trên IEA cuối kỳ:**
```python
nim_hist = [nii_hist[i] / iea_end_hist[i] * 100 for i in ...]
```

**✅ ĐÚNG — NIM lịch sử tính trên IEA bình quân đầu kỳ + cuối kỳ:**
```python
# Bước 1: tính IEA cuối kỳ từng năm
iea_end_hist = [loans_hist[i] + bank_dep_hist[i] + inv_sec_bs_hist[i] + cash_hist[i] + sbv_dep_hist[i]
                for i in range(len(years_hist))]

# Bước 2: NIM = NII / IEA bình quân (đầu kỳ + cuối kỳ) / 2
nim_hist = [round(nii_hist[i] / ((iea_end_hist[i-1] + iea_end_hist[i]) / 2 if i > 0 else iea_end_hist[i]) * 100, 2)
            for i in range(len(years_hist))]
```

**Tương tự cho NII dự phóng — dùng IEA bình quân:**
```python
# IEA bình quân cho từng năm forecast
iea_end_hist_last = iea_end_hist[-1]  # IEA cuối năm lịch sử gần nhất
iea_avg_fc = []
for i in range(3):
    prev_iea_end = iea_end_hist_last if i == 0 else iea_fc[i-1]
    iea_avg_fc.append((prev_iea_end + iea_fc[i]) / 2)

# NII = IEA bình quân × NIM
nii_fc = [iea_avg_fc[i] * nim_fc[i] / 100 for i in range(3)]
```

**Lý do:** IEA là tài sản sinh lời trung bình cả năm, không phải tại một thời điểm. Dùng IEA cuối kỳ cho kết quả NIM bị bóp méo khi ngân hàng tăng trưởng mạnh.

---

### 16.2 NIM dự phóng phải căn cứ vào thực tế quý gần nhất

**❌ SAI — dùng NIM forecast thấp hơn nhiều so với thực tế:**
```python
nim_fc = [2.80, 2.90, 3.00]  # Trong khi NIM thực tế Q1/2026 = 3.44%
```

**✅ ĐÚNG — căn cứ NIM quý gần nhất (annualized) + trend:**
```python
# Kiểm tra NIM annualized gần nhất trước khi đặt nim_fc
# NIM annualized Q1/2026 TCB = NII_Q1 * 4 / IEA_binh_quan * 100 = 3.44%
# → nim_fc phải nằm trong khoảng hợp lý ~3.40-3.60%
nim_fc = [3.40, 3.45, 3.50]  # TCB
```

**Cách kiểm tra NIM annualized từ quý thực tế:**
```python
# NIM annualized = NII_quý × 4 / IEA_bình_quân × 100
# IEA bình quân = (IEA cuối quý trước + IEA cuối quý này) / 2
nii_q = (q1_2026_is.get('isb27') or 0) / 1e9
iea_avg_q = (iea_q4_prev + iea_q1_curr) / 2
nim_annualized = nii_q * 4 / iea_avg_q * 100
```

---

### 16.3 Tỷ lệ phần trăm trong Assumptions phải lưu dạng thập phân

**❌ SAI — lưu dạng số lớn rồi nhân 100:**
```python
("NIM (%)", nim_hist + [n * 100 for n in nim_fc])   # → ghi 280 vào Excel
("CIR (%)", cir_hist + [c * 100 for c in cir_fc])   # → ghi 33 vào Excel
("Thuế suất (%)", [None]*5 + [20, 20, 20])            # → ghi 20 vào Excel
```

**✅ ĐÚNG — lưu dạng thập phân, format FMT_PCT='0.00%' sẽ hiển thị đúng:**
```python
("NIM (%)",        [n/100 for n in nim_hist] + [n/100 for n in nim_fc])  # 0.028 → hiển thị 2.80%
("CIR (%)",        [c/100 for c in cir_hist] + cir_fc)                   # 0.33 → hiển thị 33.00%
("Thuế suất (%)",  [None]*5 + [tax_rate, tax_rate, tax_rate])             # 0.20 → hiển thị 20.00%
("CASA ratio (%)", [c/100 for c in casa_ratio_hist] + casa_target_fc)
("NPL ratio (%)",  [n/100 for n in npl_ratio_hist] + npl_fc)
("CoC (%)",        [c/100 for c in coc_hist] + coc_fc)
```

**Quy tắc định dạng ô trong Assumptions:**
```python
FMT_PCT = '0.00%'  # Format chuẩn cho mọi tỷ lệ phần trăm
# Các dòng từ 4-14 (NIM, CIR, CoC, NPL, CASA, Non-TOI growth, COE, g, Tax) → FMT_PCT
# Dòng 2-3 (Giá CP, Số CP) → FMT_NUM
# Dòng 15 (P/B mục tiêu) → FMT_NUM1
fmt_to_use = FMT_PCT if r in [4,5,6,7,8,9,10,11,12,13,14] else (FMT_NUM1 if r == 15 else FMT_NUM)
```

**Hậu quả nếu sai:** Excel nhân thêm 100 lần → NIM hiển thị 280% → công thức NII = IEA × 280 → PPOP âm hàng nghìn tỷ.

---

### 16.4 Công thức liên kết giữa Assumptions và các sheet tính toán

Sau khi Assumptions lưu dạng thập phân, **công thức Excel KHÔNG cần chia 100**:

| Chỉ tiêu | Công thức Excel đúng | Ghi chú |
|---|---|---|
| NII 2026F | `=G2*G4` | G2=IEA bq, G4=NIM=0.028 |
| OPEX 2026F | `=G11*'02_Assumptions'!G7` | G7=CIR=0.33 |
| Provision | `=(avg_loans)*'02_Assumptions'!G8` | G8=CoC=0.013 |
| Thuế TNDN | `=MAX(G8*'02_Assumptions'!G14,0)` | G14=tax_rate=0.20 |
| CASA | `=G9*'02_Assumptions'!G10` | G10=CASA=0.055 |
| NPL amount | `=G5*'02_Assumptions'!G9` | G9=NPL=0.012 |
| Cho vay FC | `=F5*(1+'02_Assumptions'!G4)` | G4=credit_growth=0.14 |
| Tiền gửi FC | `=F9*(1+'02_Assumptions'!G5)` | G5=dep_growth=0.12 |

---

## 17. QUY TẮC ĐỊNH GIÁ RESIDUAL INCOME — BẮT BUỘC

### 17.1 Công thức chuẩn Residual Income Model

```
RI(t) = EPS(t) - BVPS(t-1) × COE
BVPS(t) = BVPS(t-1) + EPS(t)         ← BV tăng bằng EPS, KHÔNG tăng bằng RI
RI Value = BVPS(0) + Σ PV[RI(t)] + PV[Continuing Value]
CV = RI(3) × (1+g) / (COE - g)
```

**❌ SAI — cộng RI vào BV:**
```python
bv = bv_start + ri   # SAI: RI = EPS - Capital Charge ≠ EPS
```

**✅ ĐÚNG — cộng EPS vào BV:**
```python
eps_i = eps_fc_calc[i]           # EPS từ dự phóng lợi nhuận
capital_charge = bv_start * COE  # Chi phí vốn trên BVPS đầu kỳ
ri = eps_i - capital_charge      # RI = EPS - Capital Charge
bv = bv_start + eps_i            # BVPS(t) = BVPS(t-1) + EPS(t)
```

**Giải thích:** BVPS phản ánh vốn chủ sở hữu. Vốn CSH tăng bằng lợi nhuận giữ lại (EPS). RI là khái niệm lợi nhuận kinh tế vượt trội (economic profit), không làm thay đổi BV trực tiếp.

---

### 17.2 RI âm là tín hiệu phân tích, không phải lỗi code

Khi **RI < 0** (ROE < COE):
- **Ý nghĩa:** Công ty đang tạo ra lợi nhuận thấp hơn chi phí vốn → cổ phiếu overvalued về mặt intrinsic value
- **Không** che giấu hoặc điều chỉnh để RI dương một cách nhân tạo
- **Nên** ghi nhận trong báo cáo: "ROE dự phóng (X%) thấp hơn COE (Y%) trong giai đoạn 2026-2027 → RI âm phản ánh áp lực NIM compression; kỳ vọng phục hồi từ 2028F"
- **Kiểm tra:** Nếu RI âm cả 3 năm và CV âm → RI Value < BVPS → cổ phiếu tradeing at P/B > intrinsic → khuyến nghị NẮM GIỮ/BÁN

---

### 17.3 P/B — 3 mức theo phân phối lịch sử (BẮT BUỘC)

Không dùng một P/B median duy nhất. Phân phối P/B lịch sử chia thành 3 mức:

| Mức | Tên | Tính | Ý nghĩa đầu tư |
|---|---|---|---|
| **P/B hấp dẫn** | `pb_attractive` | Median của P/B ≤ median all-time | **Vùng MUA** — giá rẻ so lịch sử |
| **P/B fair value** | `pb_all_median` | Median toàn bộ lịch sử | Điểm cân bằng — nắm giữ |
| **P/B mục tiêu** | `pb_target` | Median của P/B ≥ median all-time | **Vùng BÁN / chốt lời** |

**✅ ĐÚNG — code Python:**
```python
pb_all_median = stats.median(pb_all_vals)

# Chia phân phối tại median
_pb_below = [p for p in pb_all_vals if p <= pb_all_median]
_pb_above = [p for p in pb_all_vals if p >= pb_all_median]

pb_attractive = stats.median(_pb_below)  # Median nửa dưới → MUA
pb_target     = stats.median(_pb_above)  # Median nửa trên → BÁN / mục tiêu

# Định giá dùng pb_target (mức mục tiêu bán)
bvps_forward = bvps_base + eps_fc_calc[0]
pb_value = pb_target * bvps_forward
```

**❌ SAI — inflate nhân tạo:**
```python
pb_target = max(pb_all_median, pb_current * 1.1)  # Cố đẩy target lên, không khách quan
pb_target = pb_all_median                          # Dùng fair value làm mục tiêu, quá thận trọng
```

**Hiển thị trong Excel (sheet 07_Valuation):**
- Row 13: P/B hấp dẫn = `pb_attractive` (màu xanh lá, nhãn "MUA")
- Row 14: P/B median all-time = `pb_all_median` (màu xám, nhãn "FAIR VALUE")
- Row 15: P/B mục tiêu = `pb_target` (màu đỏ, nhãn "BÁN / CHỐT LỜI")
- Row 16: BVPS tương lai = `=B6+B26` (BVPS hiện tại + EPS 2026F)
- Row 17: Giá trị P/B = `=B15×B16` (dùng P/B mục tiêu)

**Ứng dụng trong phân tích:**
- Giá hiện tại ≤ P/B hấp dẫn × BVPS → **MUA**
- P/B hấp dẫn < giá < P/B mục tiêu → **NẮM GIỮ**
- Giá ≥ P/B mục tiêu × BVPS → **CHỐT LỜI / BÁN**

---

### 17.4 P/B Value dùng BVPS tương lai (forward BVPS)

```python
# BVPS tương lai = BVPS hiện tại + EPS năm dự phóng đầu tiên
bvps_forward = bvps_hist[-1] + eps_fc_calc[0]

# P/B Value = P/B mục tiêu × BVPS tương lai
pb_value = pb_target * bvps_forward
```

Tương đương trong Excel:
```excel
=B14 * (B6 + B25)
# B14 = P/B mục tiêu (median)
# B6  = BVPS hiện tại (cuối 2025)
# B25 = EPS 2026F (VND/share)
```

---

### 17.5 Weighted Target Price

```
Target = 50% × RI Value + 50% × P/B Value
```
- Trọng số 50/50 là mặc định cho ngân hàng Việt Nam
- Có thể điều chỉnh nếu COE rất không chắc chắn (tăng trọng P/B) hoặc forecast ổn định (tăng trọng RI)

---

### 17.6 Checklist định giá ngân hàng bắt buộc

- [ ] RI model dùng `bv = bv_start + eps_i` (không phải `+ ri`)
- [ ] P/B target = `pb_target` (median nửa trên P/B >= median all-time, dùng làm giá mục tiêu bán)
- [ ] P/B value = P/B target × BVPS forward (BV hiện tại + EPS dự phóng)
- [ ] Nếu RI âm → ghi nhận trong báo cáo, giải thích nguyên nhân
- [ ] Sensitivity table: COE × g (5×5) để hiển thị range hợp lý
- [ ] RI Value và P/B Value phải hiển thị trong Excel dưới dạng công thức liên kết động

---

## 18. QUY TẮC BẮT BUỘC: TỰ TÍNH TOÁN CHỈ SỐ (RATIOS) — CẤM LẤY SỐ CÓ SẴN TỪ VIETCAP

> **Lý do**: Vietcap API cung cấp một số tỷ lệ tính toán sẵn (như LDR, NIM, CASA) bị sai lệch lớn do công thức tính toán đơn giản (ví dụ LDR TCB bị vọt lên 132.9% do Vietcap không tính Giấy tờ có giá vào mẫu số).

### Quy định tự tính toán cho mọi năm (Lịch sử + Dự báo)

Bắt buộc tự tính các chỉ số ngân hàng theo công thức chuẩn từ dữ liệu gốc BCTC:

1. **LDR (Tỷ lệ Dư nợ / Huy động)**:
   - **Mã tài khoản sử dụng**: Cho vay khách hàng (`bsb103`), Tiền gửi khách hàng (`bsb113`), Giấy tờ có giá phát hành (`bsb116`).
   - **Công thức**: `LDR = bsb103 / (bsb113 + bsb116) * 100` (được bảo vệ mẫu số khác 0).
   - *Kết quả thực tế*: LDR của TCB phải nằm trong khoảng hợp lý **80-87%**, không được phép là 132.9%.

2. **NIM (Biên lãi ròng)**:
   - **Công thức**: `NIM = Thu nhập lãi thuần (isb27 hoặc isb22) / Tổng tài sản sinh lãi bình quân (Average Earning Assets)`.
   - Cấm lấy giá trị `netInterestMargin` tính sẵn.

3. **CASA Ratio**:
   - **Công thức**: `CASA = Tiền gửi không kỳ hạn của khách hàng (bsb114) / Tổng tiền gửi khách hàng (bsb113) * 100`.

4. **NPL Ratio (Tỷ lệ nợ xấu)**:
   - **Công thức**: `NPL = Tổng nợ xấu bsb105 (nhóm 3+4+5) / Tổng dư nợ bsb103 (Gross Loans) * 100`.

5. **ROE & ROA**:
   - **Công thức**: `ROE = LNST / Average Equity`, `ROA = LNST / Average Assets`.

### Checklist Ratios
- [ ] Cấm đọc trực tiếp trường ratios có sẵn của Vietcap để nạp vào Dashboard.
- [ ] Mọi chỉ số của ngân hàng ở Dashboard (app.js) và file xuất JSON (data/<ticker>.json) phải khớp hoàn toàn với số liệu tự tính toán chuẩn xác từ BCTC.
- [ ] Logic tự tính toán phải bao phủ đầy đủ tất cả các năm lịch sử trong quá khứ lẫn các năm dự báo.


