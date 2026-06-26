# Plan: Bổ sung Biểu đồ Phân tích Ngân hàng

> Cập nhật từ ACB/MBS report (2026-06): Mở rộng 9 → 12 chart types + 5 framework

## ✅ Phase 1: SKILL.md (HOÀN THÀNH)
- Đã thêm section 12: 12 chart specs (types A-L) — `ngan-hang/SKILL.md`
- Đã thêm section 13: Mẫu báo cáo chuẩn ACB/MBS format
- SKILL.md hiện có 13 sections, 358+ lines

## Phase 2: build_vab_model.py

**Thay thế 3 charts hiện tại (line 652-729) bằng 12 charts mới (theo 12 types A-L từ ACB):**

### Nhóm A — NIM & Thu nhập (4 charts)

| Chart | Type | Code | Format |
|---|---|---|---|
| chart1 | A — NIM Decomposition | `chart1_nim_decomp` | 3 lines: YOEA, COF, NIM |
| chart2 | I — NIM Delta Peer | `chart2_nim_delta_peer` | Horizontal bars peer ΔNIM |
| chart3 | J — Non-II Breakdown | `chart3_nonii_breakdown` | Stacked bar: fee/FX/CK/other |
| chart4 | K — TOI+NPAT Stacked | `chart4_toi_npat` | Bar TOI + overlay NPAT line |

### Nhóm B — Chất lượng tài sản (4 charts)

| Chart | Type | Code | Format |
|---|---|---|---|
| chart5 | B — Peer NPL | `chart5_peer_npl` | Grouped bars nhiều bank |
| chart6 | G — NPL+Gr2 Combo | `chart6_npl_gr2` | Dual axis: bars + lines |
| chart7 | H — LLR+CoC Dual | `chart7_llr_coc` | 2 lines LLR + CoC |
| chart8 | C — Peer Credit Growth | `chart8_peer_credit` | Horizontal bars peer credit% |

### Nhóm C — Cơ cấu & Thị trường (4 charts)

| Chart | Type | Code | Format |
|---|---|---|---|
| chart9 | D — Bank vs Industry | `chart9_vs_industry` | 2 lines: bank + trung bình ngành |
| chart10 | E — Loan Structure | `chart10_loan_structure` | Stacked bar % retail/corp/SME |
| chart11 | F — Earning Assets | `chart11_earning_assets` | Stacked bar % |
| chart12 | L — Peer Comparison Table | `chart12_peer_table` | Table 15 metrics × 8-12 banks |

### Peer data structure cần thêm:
```python
PEER_BANKS = ["ACB","VCB","BID","CTG","MBB","TCB","VPB","HDB","VIB","LPB","STB","SHB","EIB","MSB","OCB","NAB","BAB","KLB","SGB"]
PEER_DATA = {
    "NPL": {"ACB": 1.07, "VCB": 1.14, "MBB": 1.43, ...},
    "NIM": {"ACB": 5.49, "VCB": 3.21, ...},
    "CREDIT_GROWTH": {"ACB": 15.6, ...},
    "ROE": {...}, "CIR": {...}, "CASA": {...}, "LLR": {...}, "P_B": {...}, ...
}
```

### ROE DuPont Factor — giữ lại như một optional computed metric:
```
ROE = NPM (NPAT/TOI) × AU (TOI/IEA) × EM (IEA/Equity) × (1-Tax)
```
Có thể thêm vào phần text PDF thay vì chart riêng.

## Phase 3: Chạy & Kiểm tra
1. `python build_vab_model.py` — verify 12 PNG outputs
2. Kiểm tra PDF có 12 charts đúng vị trí
3. Kiểm tra Excel không lỗi

## Điều chỉnh so với plan cũ
| Thay đổi | Từ | Thành | Lý do |
|----------|----|-------|-------|
| Chart 3 cũ (PPOP vs NPAT) | Giữ chart riêng | → Lồng vào chart4 (TOI+NPAT) | Tránh trùng, PPOP growth thể hiện trong text |
| Chart 4 cũ (ROE DuPont) | Waterfall riêng | → Computed metric trong text | Khó waterfall trong matplotlib, đủ tính toán giải thích trong text |
| Chart 8 cũ (P/B-ROE) | Scatter riêng | → Lồng vào chart12 (Peer Table) | Dữ liệu đã có trong bảng peer, scatter dễ rối |
| Chart 9 cũ (CIR Frontier) | Scatter riêng | → Bỏ | Phức tạp, dữ liệu khó thu thập đủ |
| Mới (B, C) | — | Peer NPL + Peer Credit | 2 chart so sánh peer chuẩn ACB |
| Mới (D, L) | — | Industry avg + Peer Table | Framework industry context |
