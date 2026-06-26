# Plan: Phân tích cổ phiếu VAB (Viet A Bank)

## Mục tiêu
- Ra quyết định đầu tư: Mua / Bán / Nắm giữ
- Output: Excel model (.xlsx) + PDF báo cáo phân tích chi tiết + Python builder

## Dữ liệu VAB đã có (từ Vietcap API Q1/2026)

### Tổng quan
| Chỉ tiêu | Giá trị |
|----------|---------|
| MCap | ~9,062 tỷ |
| Shares | 816,360,672 CP |
| Giá ~11,100 VND | P/B 0.86x, P/E 6.3x |
| Vốn điều lệ | 8,164 tỷ |
| Bank nhỏ nhất | trong số ~27 NH niêm yết |

### Income Statement (2023→2025)
| Khoản mục (tỷ) | 2023 | 2024 | 2025 | CAGR |
|----------------|------|------|------|------|
| NII (isb27) | 1,810 | 2,328 | 3,402 | 37% |
| TOI (isb38) | 2,513 | 2,662 | 3,787 | 23% |
| PPOP (isb40) | 1,604 | 1,641 | 2,814 | 32% |
| Provision (isb41) | -687 | -555 | -1,169 | 30% |
| PBT (isa16) | 917 | 1,086 | 1,646 | 34% |
| NPAT (isa20) | 744 | 867 | 1,320 | 33% |

### Balance Sheet (2023→2025)
| Khoản mục (tỷ) | 2023 | 2024 | 2025 | CAGR |
|----------------|------|------|------|------|
| Total Assets | 112,196 | 119,832 | 140,486 | 12% |
| Gross Loans | 68,312 | 79,157 | 87,680 | 13% |
| Customer Dep | 86,695 | 90,289 | 99,080 | 7% |
| Equity | 7,997 | 8,857 | 10,155 | 13% |
| Charter Capital | 5,400 | 5,400 | 8,164 | 23% |
| Bonds issued | 486 | 2,145 | 4,959 | 219% |

### Key Ratios Trend
| Chỉ số | Q1/2024 | Q1/2025 | Q1/2026 |
|--------|---------|---------|---------|
| NIM | 2.01% | 2.18% | 2.69% |
| NPL | 2.35% | 0.63% | 1.29% |
| Gr2 | 0.03% | 1.71% | 0.02% |
| CASA | 4.59% | 4.34% | 4.77% |
| CIR | 35.5% | 33.9% | 33.0% |
| ROE | 9.37% | 10.9% | 14.4% |
| LDR | 81.1% | 90.4% | 88.7% |
| CAR | ~9.2% | ~9.1% | ~9.4% |

---

## Execution Steps

### Step 1: Xây Python Builder (`build_vab_model.py`)
**File đích:** `E:\1. Projects\4. AIC - FA\build_vab_model.py`

**Modules cần xây:**

#### Module 1: Data Fetching (`fetch_vab_data`)
- Fetch từ Vietcap API: IS, BS, NOTE (years + quarters), statistics-financial
- Lưu cache vào `.cache\vab_*.txt`
- Mapping field: isb* (giống banking skill), bsb*, nob*

#### Module 2: Financial Model (`build_financial_model`)
- **IS Model** (2023A, 2024A, 2025A, 2026F, 2027F, 2028F):
  - NII = IEA × NIM (dự phóng: NIM giả định 2.8%→3.0%)
  - TOI = NII / grossMargin (giả định 72-74%)
  - PPOP = TOI - OPEX (CIR giả định 32-35%)
  - Provision = Dư nợ BQ × CoC (giả định 1.2-1.4%)
  - PBT, NPAT (thuế suất 20%)
- **BS Model**:
  - IEA growth (giả định 12-15% tín dụng)
  - Funding: Deposits, Bonds, Interbank
  - Equity: giữ lại LNST - cổ tức + tăng vốn

#### Module 3: Valuation (`valuation`)
- **Residual Income (RI) Model**:
  - COE = 12-14% (CAPM: rf=4.5%, β, ERP)
  - BV/share start = 12,441 VND
  - ROE forecast (14-15%)
  - RI = (ROE - COE) × BV đầu kỳ
  - Terminal value (growth 3%, ROE→COE)
- **P/B Multiple**:
  - Historic P/B range: 0.56x → 0.86x
  - Target P/B = 0.8-1.0x (50% weight)
- **Weighted fair value** = 50% RI + 50% P/B

#### Module 4: Excel Export (`export_excel`)
- Sheet 1: Tổng quan & Key Metrics Dashboard
- Sheet 2: IS Model (3Y historical + 3Y forecast)
- Sheet 3: BS Model
- Sheet 4: Ratios (NIM, NPL, CASA, CIR, ROE, ROA, LDR, CAR)
- Sheet 5: Valuation (RI Model + P/B + Weighted)
- Sheet 6: Data Dump (raw API data)

#### Module 5: PDF Report (`export_pdf`)
- Section 1: Executive Summary (1 page)
- Section 2: Asset Quality (NPL, Gr2, LLR, CoC, CAR)
- Section 3: Growth Analysis (credit, deposit, charter capital)
- Section 4: Profitability (PPOP quality, NII/TOI, non-interest income)
- Section 5: Efficiency (NIM decomposition, CIR, ROE DuPont)
- Section 6: Valuation (RI model step-by-step, P/B, sensitivity)
- Section 7: SWOT & Recommendation

### Step 2: Test & Verify
- `python build_vab_model.py --fetch-only` — verify data
- `python build_vab_model.py --build-model` — verify model
- `python build_vab_model.py --export-excel` — check Excel output
- `python build_vab_model.py --export-pdf` — check PDF output
- Verify ratios match API pre-computed values

### Step 3: Review & Decision
- Review model assumptions
- Review valuation output
- Final recommendation (Mua/Bán/Nắm giữ)
- Present to user

---

## Risk Factors & Cảnh báo cho VAB

| Rủi ro | Mức độ | Giải thích |
|--------|--------|------------|
| CASA quá thấp (4.8%) | CAO | COF cao → NIM thấp, kém cạnh tranh |
| Quy mô nhỏ | CAO | Thanh khoản thấp, khó thu hút tổ chức |
| CAR ~9.4% | CAO | Gần ngưỡng 8%, room tín dụng hẹp |
| LDR tăng nhanh | TRUNG BÌNH | 79%→89% trong 3 năm |
| CIR quá thấp | TRUNG BÌNH | Co thể do chưa đầu tư, cần verify |
| Tăng vốn gần đây | THẤP | Pha loãng ROE ngắn hạn nhưng tạo room |

## Câu hỏi cần trả lời trong báo cáo
1. VAB có thể tăng CASA từ 4.8% lên 10%+ không? (động lực NIM)
2. CAR sau tăng vốn đủ cho mở rộng 3 năm tới?
3. CoC = 1.33% — sẽ giảm khi NPL giảm?
4. Ban lãnh đạo có uy tín? Kế hoạch tăng vốn?
5. So sánh định giá với peer bank nhỏ: NAB, KLB, SGB, BAB
6. P/B <1x là value trap hay cơ hội?
