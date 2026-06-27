---
name: xuat-bao-cao
description: Use ONLY when needing to generate the final Excel model and PDF report files. Front-loaded keywords: xuất báo cáo, export report, tạo file Excel, tạo file PDF, generate model, build Excel, render PDF, in báo cáo.
---

# SKILL XUẤT BÁO CÁO

## Mục tiêu

Tạo 2 file đầu ra sau khi hoàn thành phân tích:
1. **Excel model** — Tính toán, định giá, biểu đồ
2. **PDF report** — Báo cáo phân tích chi tiết

Lưu vào: `Bao cao/<TICKER>/<TICKER>_<Loại>_<YYYY-MM>.<ext>`

---

## Cấu trúc thư mục đầu ra

```
Bao cao/
└── <TICKER>/
    ├── <TICKER>_Model_<YYYY-MM>.xlsx
    └── <TICKER>_Phan_Tich_<YYYY-MM>.pdf
```

Ví dụ:
```
Bao cao/PNJ/
├── PNJ_Model_2026-06.xlsx
└── PNJ_Phan_Tich_2026-06.pdf
```

---

## Yêu cầu Excel Model

**12 sheets** theo đúng thứ tự, đúng tên:

| Sheet | Nội dung | Yêu cầu |
|---|---|---|
| `01_Cover` | Thông tin doanh nghiệp, snapshot, khuyến nghị | Giá CP, vốn hóa, số CP, kế hoạch, P/E, rating, target price |
| `02_Assumptions` | Tất cả giả định đầu vào | Giá CP, giá HH, tỷ giá, số CH, biên LN, nợ, CAPEX, D&A — mỗi dòng có ghi chú |
| `03_Revenue_Model` | Driver-based theo từng phân khúc | Công thức rõ ràng, % YoY tự động, mix % |
| `04_PnL` | P&L lịch sử 5 năm + dự báo 3-5 năm | DT → GP → EBIT → LNST → EPS. Biên % tự động |
| `05_Balance_Sheet` | Cân đối kế toán | Tổng TS = Tổng NV. Net Debt. D/E |
| `06_Cash_Flow` | LCTT + FCFF | CFO, CFI, CFF, FCFF, FCF Yield |
| `07_Valuation` | Bear/Base/Bull weighted avg | P/E, EV/EBITDA, DCF. Weighted target price. Upside |
| `08_Sensitivity` | Ma trận nhạy cảm 2 biến | P/E×EPS (5×5), EV/EBITDA×EBITDA (5×5). Tô màu: đỏ < giá HT, xanh > target |
| `09_PESTLE` | 6 yếu tố vĩ mô | Mỗi yếu tố: nội dung, tác động, mức độ (Tích cực/Tiêu cực/Trung tính) |
| `10_Leading_Indicators` | Chỉ báo theo dõi | Tên chỉ báo, ngưỡng tích cực, ngưỡng tiêu cực, giá trị hiện tại, trạng thái |
| `11_Investment_Thesis` | 6 tầng đánh giá | Mỗi tầng 1-2 dòng kết luận + rating |
| `12_Summary_Snapshot` | Tóm tắt tài chính | DT, GP, EBITDA, LNST, EPS, cổ tức, CAPEX, Net Debt, P/E, EV/EBITDA — kèm CAGR |

### Quy tắc Excel — FORMULA BẮT BUỘC

**Mọi ô số liệu PHẢI là công thức Excel, KHÔNG hardcode số.** Người dùng cần mở file và thấy công thức để kiểm tra.

#### Phân loại ô:
| Loại | Màu | Nội dung | Ví dụ |
|---|---|---|---|
| **Input cứng** | Xanh dương | Dữ liệu lịch sử (BCTC thật), giả định chính | `34,480` (LNST 2021), `23,600` (giá CP) |
| **Formula** | Đen | Tính toán từ input | `=F2*(1+Assumptions!G5/100)` |
| **Link sheet khác** | Xanh lá | Tham chiếu chéo sheet | `=Revenue!G2` |
| **Giả định quan trọng** | Nền vàng | Biên %, tăng trưởng %, thuế suất | `17.5%` (GP Margin) |

#### Quy tắc cụ thể:

**Sheet `02_Assumptions`:**
- Là nơi DUY NHẤT chứa số hardcode (input)
- Mỗi dòng một biến, có ghi chú
- Cột B-F = lịch sử (hardcode), Cột G-I = dự báo (có thể hardcode hoặc formula)

**Sheet `03_Revenue_Model`:**
- Doanh thu forecast: `=F<row>*(1+Assumptions!G<growth_row>/100)` — KHÔNG ghi số 210,000
- YoY Growth: `=G2/F2-1` — KHÔNG ghi "34.5%"
- Sản lượng forecast: hardcode được (giả định riêng)

**Sheet `04_PnL`:**
- Doanh thu: `=Revenue!G2` (link Revenue sheet)
- Giá vốn: `=G<rev_row>*(1-Assumptions!G<gp_margin_row>/100)`
- Lợi nhuận gộp: `=G<rev_row>-G<cogs_row>`
- Biên GP: `=G<gp_row>/G<rev_row>*100`
- CP BH&QLDN: `=G<rev_row>*Assumptions!$B$<sgka_row>` (dùng tỷ lệ từ assumptions)
- EBIT: `=G<gp_row>-G<sgka_row>`
- Biên EBIT: `=G<ebit_row>/G<rev_row>*100`
- EBT: `=G<ebit_row>+G<fin_i_row>-G<fin_c_row>`
- Thuế: `=G<ebt_row>*Assumptions!G<tax_row>/100`
- LNST: `=G<ebt_row>-G<tax_row>`
- Biên LNST: `=G<ni_row>/G<rev_row>*100`
- EPS: `=G<ni_row>*1e9/(Assumptions!G<shares_row>*1e6)`
- EBITDA: `=G<ebit_row>+Assumptions!G<da_row>`

**Sheet `05_Balance_Sheet`:**
- Tổng TS: `=SUM(B<start>:B<end>)` — KHÔNG ghi số 280,000
- Tổng nợ: `=SUM(...)`
- Tổng VCSH: `=SUM(...)`
- Tổng nợ+VCSH: `=B<total_liab>+B<total_eq>`
- Net Debt: `=B<debt>-B<cash>`
- D/E: `=B<net_debt>/B<equity>`

**Sheet `06_Cash_Flow`:**
- CFO ước: `=B<ni>+B<da>+B<wcip>`
- FCFF: `=B<cfo>+B<cfi>`
- Tiền cuối kỳ: `=B<opening>+B<cfo>+B<cfi>+B<cff>`
- Tiền đầu kỳ sau: `=B<closing>`

**Sheet `07_Valuation`:**
- EV/EBITDA target: `=(PnL!G<ebitda>*Assumptions!$B$<ev_ebitda_multiple> - BS!G<net_debt>)*1e9 / (Assumptions!$B$<shares>*1e6)`
- P/B target: `=Assumptions!$B$<target_pb> * (BS!G<equity>*1e9/(Assumptions!$B$<shares>*1e6))`
- P/E target: `=Assumptions!$B$<target_pe> * PnL!G<eps>`
- Weighted target: `=G<ev_row>*0.5 + G<pb_row>*0.3 + G<pe_row>*0.2`
- Upside: `=G<target>/Assumptions!$B$<price_row> - 1`

**Sheet `08_Sensitivity`:**
- Mỗi ô trong matrix: `=($A<ebitda_row>*B$<ev_ebitda_header> - BS!G<net_debt>)*1e9/(Assumptions!$B$<shares>*1e6)`

**Sheet `12_Summary_Snapshot`:**
- CAGR: `=(I<h>/F<h>)^(1/3)-1` — KHÔNG ghi "20.0%"
- ROE: `=B<ni>/B<equity>*100`
- ROA: `=B<ni>/B<total_assets>*100`
- D/E: `=(B<debt>-B<cash>)/B<equity>`
- P/E: `=Assumptions!$B$<price>*Assumptions!$B$<shares>/(B<ni>*1e3)`
- P/B: `=(Assumptions!$B$<price>*Assumptions!$B$<shares>)/B<equity>/1e3`
- EV/EBITDA: `=((Assumptions!$B$<price>*Assumptions!$B$<shares>/1000)+(B<debt>-B<cash>))/B<ebitda>`

### Formula patterns đặc thù Ngân hàng

**Sheet `03_Income_Model` (banking):**
```
NII forecast:  =F<iea_row> * (1 + Assumptions!$E$<nim_growth>/100) * Assumptions!$E$<nim_rate>/100
IEA forecast:  =F<iea_row> * (1 + Assumptions!$E$<iea_growth>/100)
Non-II forecast: =SUM(F<fee_row>:F<other_row>)
TOI:           =SUM(F<nii_row>:F<nonii_row>)
OPEX:          =F<toi_row> * Assumptions!$E$<cir_rate>/100
PPOP:          =F<toi_row> - F<opex_row>
Provision:     =AVERAGE(F<loans_row>,F<prev_loans_row>) * Assumptions!$E$<coc_rate>/100
LNST:          =F<ppop_row> - F<prov_row>
```

**Sheet `04_PnL` (banking) — historical + forecast overlap:**
```
NII (hist):    hardcode từ BCTC
NII (fc):     ='03_Income_Model'!G<nii_row>
TOI:          ='03_Income_Model'!G<toi_row>
EPS:          =G<ni_row>*1e9/Assumptions!$B$<shares_row>
```

**Sheet `05_Balance_Sheet` (banking):**
```
Loans forecast:  =F<loans_row> * (1 + Assumptions!$E$<credit_growth>/100)
Dep forecast:    =F<dep_row> * (1 + Assumptions!$E$<dep_growth>/100)
CASA forecast:   =G<dep_row> * Assumptions!$E$<casa_rate>/100
Equity forecast: =F<equity_row> + '03_Income_Model'!G<ni_row> - Dividends
```

**Sheet `06_Ratios` (tất cả là formula, KHÔNG hardcode):**
```
NIM:     ='04_PnL'!G<nii_row> / '05_BS'!G<iea_row> * 100
NPL:     ='05_BS'!G<npl_row> / '05_BS'!G<loans_row> * 100
CASA:    ='05_BS'!G<casa_row> / '05_BS'!G<dep_row> * 100
CIR:     ='04_PnL'!G<opex_row> / '04_PnL'!G<toi_row> * 100
ROE:     ='04_PnL'!G<ni_row> / AVERAGE('05_BS'!F<equity_row>,'05_BS'!G<equity_row>) * 100
LDR:     ='05_BS'!G<loans_row> / '05_BS'!G<dep_row> * 100
CoC:     ='04_PnL'!G<prov_row> / AVERAGE('05_BS'!F<loans_row>,'05_BS'!G<loans_row>) * 100
LLR:     ='05_BS'!G<llr_row> / '05_BS'!G<npl_row> * 100
```

**Sheet `07_Valuation` (banking — RI Model trong Excel):**
```
BV/share:        ='05_BS'!G<equity_row>*1e9 / Assumptions!$B$<shares_row>
ROE forecast:    ='04_PnL'!G<ni_row> / '05_BS'!F<equity_row> * 100
RI:              =G<bv_ps_row> * (G<roe_row>/100 - Assumptions!$G$<coe_row>)
PV of RI:        =G<ri_row> / (1 + Assumptions!$G$<coe_row>)^(COLUMN()-2)
PV of CV:        =(H<ri_row>*(1+Assumptions!$G$<terminal_growth>)/(Assumptions!$G$<coe_row>-Assumptions!$G$<terminal_growth>)) / (1+Assumptions!$G$<coe_row>)^3
RI Value:        =G<bv_ps_row> + SUM(G<pv_ri_row>:H<pv_ri_row>) + G<pv_cv_row>
P/B Target:      =Assumptions!$G$<target_pb> * G<bv_ps_row>
Weighted Target: =G<ri_value_row>*0.5 + G<pb_target_row>*0.5
Upside:          =G<target_row> / Assumptions!$B$<price_row> - 1
```

**Sheet `13_PE_PB_History` (thêm mới cho banking):**
```
Cột layout: Quý | P/E | P/B | EV/EBITDA | Giá | VNI
Dữ liệu: hardcode từ API (historical only, KHÔNG forecast)
Chart: Line chart P/E + P/B dual axis
```

**Kiểm tra bắt buộc:**
- [ ] Mở file Excel, bấm vào ô → phải thấy formula bar hiện công thức (không phải số)
- [ ] Recalc 0 lỗi (#REF!, #DIV/0!, #VALUE!, #NAME?)
- [ ] Balance sheet cân (Tổng TS = Tổng NV)
- [ ] FCFF > 0 base case
- [ ] Số âm dùng ngoặc đơn `(1,234)`
- [ ] Zero hiển thị `-` (format: `#,##0;(#,##0);-`)
- [ ] Hàng năm = bold, phân cách hàng nghìn, 2 số thập phân cho tỷ lệ %
- [ ] Sheet 13_PE_PB_History tồn tại và chỉ chứa dữ liệu lịch sử (không forecast)

### Biểu đồ trong Excel

Tối thiểu **3 biểu đồ**:
1. **Revenue & LNST trend** — Cột chồng/kết hợp đường (Sheet `04_PnL`)
2. **Margin trend** — GP%, EBIT%, LNST% theo năm (Sheet `04_PnL` hoặc `12_Summary_Snapshot`)
3. **Sensitivity heatmap** — Conditional formatting trên matrix (Sheet `08_Sensitivity`)

---

## Yêu cầu PDF Report

### Cấu trúc (15-20 trang)

| Trang | Nội dung | Ghi chú |
|---|---|---|
| 1 | **Cover** | Logo, ticker, tên công ty, ngày, analyst, khuyến nghị, target price |
| 2 | **Investment Summary** | Standalone — 1 trang đủ hiểu thesis. Rating, target, upside, 3 lý do mua, 3 rủi ro, snapshot TC |
| 3-4 | Chuỗi giá trị & Mô hình KD | Sơ đồ + giải thích |
| 5-6 | Thị trường & Vị trí cạnh tranh | Porter 5F, chu kỳ ngành |
| 7-8 | Moat & Quản trị | Rating moat, ROIC vs peer |
| 9-11 | Phân tích tài chính | Bảng số 5 năm, biểu đồ trend |
| 12-14 | Định giá | Phương pháp, Bear/Base/Bull, sensitivity chart |
| 15-16 | Rủi ro & Catalysts | Risk matrix, timeline |
| 17 | Kết luận | BUY/HOLD/SELL, leading indicators |
| 18+ | Phụ lục | BCTC tóm tắt, peer table, nguồn |

### Biểu đồ trong PDF

Tối thiểu **3 biểu đồ** (khác biệt so với Excel):
1. **Doanh thu & LNST** — Biểu đồ cột/nến (matplotlib/seaborn)
2. **Biên lợi nhuận** — Đường xu hướng GP%, EBIT%, LNST%
3. **Sensitivity** — Heatmap hoặc tornado chart

Định dạng: Font tiếng Việt (Arial/Time New Roman), màu sắc hài hòa, legend rõ, trục có nhãn.

### Yêu cầu nội dung

- **Ngôn ngữ**: Tiếng Việt, giữ thuật ngữ TA không có từ tương đương (EV, EBITDA, NIM...)
- **Số liệu**: Ghi nguồn & ngày lấy
- **Kết luận**: BUY/HOLD/SELL rõ ràng, không mơ hồ
- **Investment Summary**: Đọc 1 trang hiểu toàn bộ thesis

---

## Công nghệ triển khai

### Excel
- **Thư viện**: openpyxl
- Tạo workbook, định dạng cell (font, màu, border, number_format)
- Vẽ chart bằng openpyxl chart module
- Lưu file

### PDF
- **Thư viện**: reportlab (hoặc fpdf2)
- Tạo document với Paragraph, Table, Image
- Chèn hình biểu đồ (xuất từ matplotlib ra PNG, nhúng vào PDF)
- Font hỗ trợ tiếng Việt: register TTF (Arial, Times New Roman)

### Biểu đồ
- **Thư viện**: matplotlib
- Vẽ chart, xuất PNG tạm → nhúng vào Excel/PDF
- Font: hỗ trợ tiếng Việt (rcParams['font.family'])

---

## Quy trình xuất

```
1. Chuẩn bị dữ liệu:
   - Map các giả định từ phân tích vào Assumptions dict
   - Tính toán các sheet P&L, B/S, CF, Valuation

2. Tạo Excel:
   - Dùng openpyxl build workbook theo 12 sheets
   - Điền công thức, kẻ bảng, tô màu
   - Vẽ chart
   - Lưu: Bao cao/<TICKER>/<TICKER>_Model_<YYYY-MM>.xlsx

3. Tạo PDF:
   - Dùng reportlab/pdf fpdf2 tạo document
   - Viết nội dung các trang
   - Vẽ/và nhúng biểu đồ
   - Lưu: Bao cao/<TICKER>/<TICKER>_Phan_Tich_<YYYY-MM>.pdf

4. Kiểm tra:
   - Excel: mở file, check 0 lỗi formula
   - PDF: check font, layout, đủ nội dung
```

---

## Checklist trước khi bàn giao

- [ ] Excel: Đủ 13+ sheets, đúng tên, đúng thứ tự
- [ ] Excel: 0 formula errors (#REF!, #DIV/0!, #VALUE!, #NAME?)
- [ ] Excel: Balance sheet cân (Tổng TS = Tổng NV)
- [ ] Excel: Mọi ô forecast là FORMULA (không hardcode số)
- [ ] Excel: Sheet 13_PE_PB_History có MEDIAN cuối bảng
- [ ] Excel: Có ít nhất 10 biểu đồ (ngân hàng) / 6 biểu đồ (ngành khác)
- [ ] Excel: Sensitivity có conditional formatting
- [ ] Check: P/E quý LNST âm = P/E carry từ quý trước
- [ ] Check: P/B target = MEDIAN all-time P/B (không phải average)
- [ ] Check: CIR formula = ='04_PnL'!E5/'04_PnL'!E4*100 (đúng PnL ref)
- [ ] PDF: Đủ 15-20 trang
- [ ] PDF: Investment Summary standalone
- [ ] PDF: BUY/HOLD/SELL + target price rõ ràng
- [ ] PDF: Bear/Base/Bull có xác suất = 100%
- [ ] PDF: Có nguồn dữ liệu
- [ ] PDF: Chart PE/PB historical: có median line + carry markers
- [ ] File name đúng format
- [ ] Lưu đúng thư mục `Bao cao/<TICKER>/`

---

## Quy tắc bắt buộc — Assumptions và tỷ lệ phần trăm (BẮT BUỘC ĐỌC)

### Lưu tỷ lệ % dạng thập phân trong Assumptions

Khi ghi dữ liệu vào sheet `02_Assumptions`, **tất cả tỷ lệ phần trăm phải lưu dạng số thập phân** (0.028 cho 2.80%), **KHÔNG nhân × 100**. Định dạng ô `FMT_PCT = '0.00%'` sẽ tự động hiển thị đúng.

```python
# ❌ SAI — nhân 100, Excel nhân thêm 100 → hiển thị 280%
("NIM (%)", nim_hist + [n * 100 for n in nim_fc])

# ✅ ĐÚNG — thập phân, FMT_PCT hiển thị 2.80%
("NIM (%)", [n/100 for n in nim_hist] + [n/100 for n in nim_fc])
("CIR (%)", [c/100 for c in cir_hist] + cir_fc)          # cir_fc đã là 0.33
("CoC (%)", [c/100 for c in coc_hist] + coc_fc)           # coc_fc đã là 0.013
("NPL (%)", [n/100 for n in npl_ratio_hist] + npl_fc)     # npl_fc đã là 0.012
("CASA (%)", [c/100 for c in casa_ratio_hist] + casa_target_fc)
("Non-TOI growth (%)", [None]*5 + non_int_growth_fc)       # đã là 0.15
("Thuế suất (%)", [None]*5 + [tax_rate, tax_rate, tax_rate])  # tax_rate=0.20
```

### Công thức Excel không cần chia 100

Sau khi Assumptions đã lưu thập phân, công thức Excel dùng trực tiếp:
- **NII** = `=G2*G4` (IEA bq × NIM thập phân)
- **OPEX** = `=G11*'02_Assumptions'!G7` (TOI × CIR thập phân)
- **Thuế** = `=MAX(G8*'02_Assumptions'!G14,0)` (PBT × tax_rate thập phân)
- **CASA** = `=G9*'02_Assumptions'!G10` (Tiền gửi × CASA ratio thập phân)

### NIM chuẩn — dùng IEA bình quân

NIM = NII / **IEA bình quân (đầu kỳ + cuối kỳ) / 2**, KHÔNG phải IEA cuối kỳ.

```python
# IEA bình quân cho dự phóng
iea_avg_fc = [(prev_iea + iea_fc[i]) / 2 for i in range(3)]
nii_fc = [iea_avg_fc[i] * nim_fc[i] / 100 for i in range(3)]
```

### NIM dự phóng phải căn cứ thực tế

NIM forecast phải gần với NIM annualized của quý thực tế gần nhất:
```
NIM annualized = NII_quý × 4 / IEA_bình_quân_quý × 100
```
Nếu Q1/2026 NIM annualized = 3.44% → nim_fc[0] không được thấp hơn 3.20% mà không có lý do rõ ràng.

### Checklist bổ sung cho banking model

- [ ] Assumptions: NIM hiển thị `2.80%` (không phải `280%`)
- [ ] Assumptions: CIR hiển thị `33.00%` (không phải `3300%`)
- [ ] Assumptions: Tax hiển thị `20.00%` (không phải `2000%`)
- [ ] Income Model G5: `=G2*G4` với G4=0.028 → NII hợp lý
- [ ] PPOP dương và hợp lý so với năm trước (không âm nghìn tỷ)
- [ ] nim_fc đặt sát với NIM quý gần nhất (annualized)

---

## Quy tắc Residual Income — Sheet 07_Valuation (BẮT BUỘC)

### Công thức Excel đúng trong 07_Valuation

| Row | Label | Formula |
|---|---|---|
| B6 | BVPS hiện tại | `='05_Balance_Sheet'!F15*1e9/'02_Assumptions'!$B$3` |
| B25/C25/D25 | EPS 2026/27/28F | `='04_PnL'!G10*1e9/'02_Assumptions'!$B$3` |
| B26/C26/D26 | BVPS đầu kỳ | `='05_Balance_Sheet'!F15*1e9/'02_Assumptions'!$B$3` (2026) |
| B27/C27/D27 | Capital Charge | `=B26*'07_Valuation'!$B$2` (BVPS_đầu × COE) |
| B28/C28/D28 | Residual Income | `=B25-B27` (EPS - Capital Charge) |
| B29/C29/D29 | Discount Factor | `=1/(1+$B$2)^1` |
| B30/C30/D30 | PV of RI | `=B28*B29` |
| B7 | PV tổng RI | `=SUM(B30:D30)` |
| B8 | PV Continuing Value | `=(D28*(1+B3)/(B2-B3))*D29` |
| B9 | RI Value | `=B6+B7+B8` |
| B13 | P/B hấp dẫn | giá trị số `pb_attractive` từ code |
| B14 | P/B median all-time | giá trị số `pb_all_median` từ code |
| B15 | P/B mục tiêu | `='02_Assumptions'!G15` = pb_target |
| B16 | BVPS forward | `=B6+B26` (BVPS hiện tại + EPS 2026F) |
| B17 | P/B Value | `=B15*B16` |
| B21 | Target Price | `=B9*0.5+B17*0.5` |

### Lỗi phổ biến cần tránh

- **SAI:** `pb_target = max(pb_median, pb_current * 1.1)` → inflate target nhân tạo
- **SAI:** `pb_target = pb_all_median` → dùng fair value làm mục tiêu, không phân biệt vùng mua/bán
- **ĐÚNG:** 3 mức P/B theo phân phối lịch sử:
  ```python
  _pb_below     = [p for p in pb_all_vals if p <= pb_all_median]
  _pb_above     = [p for p in pb_all_vals if p >= pb_all_median]
  pb_attractive = stats.median(_pb_below)  # median nửa dưới → MUA
  pb_target     = stats.median(_pb_above)  # median nửa trên → BÁN / mục tiêu
  # pb_value = pb_target × BVPS_forward (dùng P/B mục tiêu để định giá)
  ```
- **SAI:** `bv = bv_start + ri` trong vòng lặp RI → BVPS tăng sai
- **ĐÚNG:** `bv = bv_start + eps_fc_calc[i]` → BVPS tăng bằng EPS
- **ĐÚNG:** Để RI âm hiển thị, giải thích trong báo cáo (ROE < COE = overvalued)

---

## JSON Dashboard Export — BẮT BUỘC (Gọi tại Phase 3 của Master Skill)

> **Mục đích**: File `data/<TICKER>.json` là nguồn dữ liệu duy nhất cho dashboard web (index.html + app.js). Mọi builder script PHẢI tạo file này với đầy đủ fields. File này được `run_analysis.py` patch thêm `gdriveExcelUrl`/`gdrivePdfUrl` sau khi upload.

### Schema đầy đủ

```json
{
  "ticker": "TCB",
  "companyName": "Ngân hàng TMCP Kỹ thương Việt Nam",
  "sector": "Ngân hàng",
  "currentPrice": 25000,
  "marketCap": 89000000000000,
  "shares": 3560000000,
  "gdriveExcelUrl": null,
  "gdrivePdfUrl": null,

  "data": {
    "years":   [2020, 2021, 2022, 2023, 2024, 2025, 2026, 2027],
    "revenue": [tỷ VND theo thứ tự năm — hist+forecast],
    "npat":    [tỷ VND — hist+forecast],
    "eps":     [VND/cổ phiếu — hist+forecast],
    "equity":  [tỷ VND VCSH — hist+forecast]
  },

  "thesis": [
    "Luận điểm đầu tư 1 (~50 từ, định lượng, tiếng Việt)",
    "Luận điểm đầu tư 2 (~50 từ)",
    "Luận điểm đầu tư 3 (~50 từ)"
  ],

  "risks": [
    "Rủi ro 1 (~50 từ, nêu tác động định lượng, tiếng Việt)",
    "Rủi ro 2 (~50 từ)",
    "Rủi ro 3 (~50 từ)"
  ],

  "moats": {
    "Network Effect":    {"score": 4, "desc": "Giải thích ngắn gọn tại sao score là X/5"},
    "Cost Advantage":    {"score": 3, "desc": "..."},
    "Switching Cost":    {"score": 4, "desc": "..."},
    "Intangible Assets": {"score": 4, "desc": "..."},
    "Efficient Scale":   {"score": 3, "desc": "..."}
  },

  "pestle": [
    {"factor": "Political",    "content": "Mô tả tác động chính sách...", "impact": "Positive"},
    {"factor": "Economic",     "content": "Mô tả tác động kinh tế...",    "impact": "Neutral"},
    {"factor": "Social",       "content": "Mô tả tác động xã hội...",     "impact": "Positive"},
    {"factor": "Technological","content": "Mô tả tác động công nghệ...", "impact": "Positive"},
    {"factor": "Legal",        "content": "Mô tả tác động pháp lý...",    "impact": "Neutral"},
    {"factor": "Environmental","content": "Mô tả tác động môi trường...", "impact": "Negative"}
  ],

  "valuation": {
    "bear": 18000,
    "base": 25000,
    "bull": 32000
  },

  "comments": {
    "businessModel":        "~100 từ tiếng Việt mô tả mô hình kinh doanh cốt lõi",
    "financialPerformance": "~100 từ tiếng Việt đánh giá sức khỏe tài chính",
    "valuationText":        "~100 từ tiếng Việt giải thích định giá và khuyến nghị"
  },

  "pe_hist":        [8.2, 7.1, 6.3, 5.8, 6.5],
  "pb_hist":        [1.2, 1.1, 0.95, 0.84, 1.0],
  "pe_quarters":    [8.2, 7.9, 7.5, 7.1, 6.8, 6.3, 6.0, 5.8, 6.1, 6.3, 6.5, 6.8],
  "pb_quarters":    [1.2, 1.18, 1.15, 1.1, 1.05, 0.95, 0.90, 0.84, 0.88, 0.92, 0.96, 1.0],
  "quarter_labels": ["2021-Q1","2021-Q2","2021-Q3","2021-Q4",
                     "2022-Q1","2022-Q2","2022-Q3","2022-Q4",
                     "2023-Q1","2023-Q2","2023-Q3","2023-Q4"]
}
```

### Quy tắc bắt buộc

| Field | Quy tắc |
|-------|---------|
| `data.years` | Phải bao gồm CẢ lịch sử lẫn forecast (hist + 2-3 năm E) |
| `data.revenue/npat/eps/equity` | Cùng độ dài với `data.years` |
| `moats` | Đúng 5 keys, score từ 1-5 |
| `pestle` | Đúng 6 items, `impact` ∈ `["Positive", "Neutral", "Negative"]` |
| `valuation` | Phải có `bear`, `base`, `bull` |
| `comments` | 3 keys: `businessModel`, `financialPerformance`, `valuationText` |
| `pe_quarters` / `pb_quarters` | Lấy từ Vietcap `statistics-financial` API, TOÀN BỘ quý có sẵn |
| `quarter_labels` | Cùng độ dài với `pe_quarters`, format `"YYYY-QN"` |
| `gdriveExcelUrl` / `gdrivePdfUrl` | Khởi tạo là `null` — `run_analysis.py` patch sau upload |

### Code mẫu Python builder

```python
def save_json_summary(ticker, company_name, sector, current_price, market_cap, shares,
                      hist_years, fc_years,
                      rev_hist, rev_fc, npat_hist, npat_fc,
                      eps_hist, eps_fc, equity_hist, equity_fc,
                      pe_hist, pb_hist, pe_quarters, pb_quarters, quarter_labels,
                      thesis, risks, moats, pestle, valuation, comments):
    out_dir = os.path.join(os.path.dirname(__file__), "data")
    os.makedirs(out_dir, exist_ok=True)

    payload = {
        "ticker": ticker,
        "companyName": company_name,
        "sector": sector,
        "currentPrice": current_price,
        "marketCap": market_cap,
        "shares": shares,
        "gdriveExcelUrl": None,
        "gdrivePdfUrl": None,
        "data": {
            "years":   hist_years + fc_years,
            "revenue": [round(v, 1) for v in rev_hist + rev_fc],
            "npat":    [round(v, 1) for v in npat_hist + npat_fc],
            "eps":     [round(v, 0) for v in eps_hist + eps_fc],
            "equity":  [round(v, 1) for v in equity_hist + equity_fc],
        },
        "thesis":         thesis,     # list[str] x3
        "risks":          risks,      # list[str] x3
        "moats":          moats,      # dict{name: {score, desc}} x5
        "pestle":         pestle,     # list[{factor, content, impact}] x6
        "valuation":      valuation,  # {bear, base, bull}
        "comments":       comments,   # {businessModel, financialPerformance, valuationText}
        "pe_hist":        pe_hist,
        "pb_hist":        pb_hist,
        "pe_quarters":    pe_quarters,
        "pb_quarters":    pb_quarters,
        "quarter_labels": quarter_labels,
    }

    out_path = os.path.join(out_dir, f"{ticker}.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    print(f"[OK] JSON saved: {out_path}")
```

### Checklist JSON Export

- [ ] `data.years` = hist + forecast (ví dụ `[2020..2027]`)
- [ ] Tất cả arrays trong `data` có cùng độ dài với `data.years`
- [ ] `moats` đủ 5 keys
- [ ] `pestle` đủ 6 items, `impact` đúng giá trị
- [ ] `pe_quarters` / `pb_quarters` / `quarter_labels` cùng độ dài
- [ ] `gdriveExcelUrl` = `null` (không tự upload, để orchestrator xử lý)
- [ ] File lưu tại `data/<TICKER>.json` (KHÔNG phải `Bao cao/`)
- [ ] JSON hợp lệ (không trailing comma, không comment `//`)

