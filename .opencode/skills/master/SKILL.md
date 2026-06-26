---
name: master
description: BẮT BUỘC dùng skill này ĐẦU TIÊN khi nhận bất kỳ ticker VN nào. Front-loaded keywords: phân tích cổ phiếu, phân tích ngành, master skill, main skill, điều phối, orchestrate, run full analysis.
---

# SKILL MASTER — ĐIỀU PHỐI PHÂN TÍCH CỔ PHIẾU TOÀN DIỆN

> **Quy tắc bắt buộc**: Skill này phải được load ĐẦU TIÊN khi nhận ticker. Nó điều phối thứ tự chạy các skill khác. KHÔNG ĐƯỢC BỎ QUA BẤT KỲ BƯỚC NÀO.

## Quy trình bắt buộc (4 Phase, tuần tự)

```
NHẬN TICKER (VD: "phân tích HPG", "định giá TCB")
    │
    ▼
┌────────────────────────────────────────────────────────────────┐
│ PHASE 0: RECEIVE TICKER                                        │
│ • Tra bảng phan-loai-nganh → xác định ngành                    │
│ • Fetch data từ Vietcap API (BCTC + statistics-financial)      │
│ • Kiểm tra đủ dữ liệu? Nếu thiếu → DỪNG LẠI báo user          │
│ GHI NHẬN: ticker, ngành, ngày phân tích, năm A/E               │
└────────────────────────────────────────────────────────────────┘
    │
    ▼ CHECKLIST: đã có đủ BCTC 3-5 năm + PE/PB/EVEBITDA lịch sử?
    │
    ▼
┌────────────────────────────────────────────────────────────────┐
│ PHASE 1: RUN SKILL FA (phân tích 6 tầng)                      │
│ • Load skill `fa` → thực hiện tuần tự Bước 1 → Bước 2 (6 tầng)│
│ • Mỗi tầng phải có artifact output (bảng/checklist)            │
│ • KHÔNG ĐƯỢC SKIP bất kỳ tầng nào                              │
│ OUTPUT: 6 artifact + bảng A/E + bảng dữ liệu thu thập          │
└────────────────────────────────────────────────────────────────┘
    │
    ▼ CHECKLIST: đủ 6 artifact? Step 1 (A/E) pass? Data đủ?
    │
    ▼
┌────────────────────────────────────────────────────────────────┐
│ PHASE 2: RUN SKILL NGÀNH                                       │
│ • Dùng ngành đã xác định ở Phase 0 → load skill tương ứng:     │
│   - Ngân hàng → skill `ngan-hang`                              │
│   - Thép → skill `thep`                                        │
│   - Bán lẻ → skill `ban-le`                                    │
│   - Ngành khác → skill `fa` (dùng phương pháp chung)          │
│ • Thực hiện TẤT CẢ các bước trong skill ngành                  │
│ • Mỗi bước có output riêng, không skip                         │
│ • Nếu skill ngành yêu cầu fetch thêm dữ liệu → phải fetch     │
│ • Có thể tạo thêm sheet/phân tích không giới hạn               │
│ OUTPUT: forecast model, valuation, charts, sheets bổ sung       │
└────────────────────────────────────────────────────────────────┘
    │
    ▼ CHECKLIST: đủ forecast? Công thức Excel đúng? sheets đủ?
    │
    ▼
┌────────────────────────────────────────────────────────────────┐
│ PHASE 3: BUILD EXCEL + PDF (xuat-bao-cao)                      │
│ • Load skill `xuat-bao-cao` → build file theo đúng spec       │
│ • TẤT CẢ ô forecast PHẢI là công thức Excel (KHÔNG hardcode)   │
│ • Recalc = 0 lỗi                                                │
│ • Output: Bao cao/<TICKER>/<TICKER>_Model_<YYYY-MM>.xlsx       │
│ • Output: Bao cao/<TICKER>/<TICKER>_Phan_Tich_<YYYY-MM>.pdf    │
└────────────────────────────────────────────────────────────────┘
    │
    ▼
HOÀN THÀNH: Báo user đường dẫn file + khuyến nghị.
```

## Master Checklist (kiểm tra sau mỗi Phase)

### SAU PHASE 0:
```
☐ Ticker hợp lệ? Có giao dịch trên HSX/HNX/UPCOM?
☐ Ngành xác định đúng? (tra bảng phan-loai-nganh)
☐ BCTC 3-5 năm đã fetch đủ? (Income Statement, Balance Sheet, Cash Flow)
☐ PE/PB/EVEBITDA lịch sử đã fetch đủ? (statistics-financial endpoint)
☐ Giá + vốn hóa + số CP đã fetch?
☐ Ngày phân tích + năm A/E đã xác định?
```

### SAU PHASE 1 (FA):
```
☐ Tầng 1 — Bảng chuỗi giá trị (≥5 bước, mỗi bước có biên)
☐ Tầng 2 — Porter 5F + PESTLE + Chu kỳ ngành + TAM
☐ Tầng 3 — Moat 5 loại + ROIC vs WACC + Quản lý
☐ Tầng 4 — Driver table + P&L forecast (3 năm +)
☐ Tầng 5 — Risk matrix + Catalyst table
☐ Tầng 6 — Định giá Bear/Base/Bull + Sensitivity
```

### SAU PHASE 2 (NGÀNH):
```
☐ Đã thực hiện tất cả các bước trong skill ngành?
☐ Bước nào bị skip? (Nếu có: DỪNG, yêu cầu xử lý)
☐ Dữ liệu PE/PB từng quý đã xử lý: LNST âm → P/E carry?
☐ Median năm đã tính? Median all-time đã tính?
☐ Forecast drivers đã có? (credit growth, NIM, CASA, ...)
```

### SAU PHASE 3 (XUẤT):
```
☐ File Excel đã tạo? (Bao cao/<TICKER>/<TICKER>_Model_*.xlsx)
☐ File PDF đã tạo? (Bao cao/<TICKER>/<TICKER>_Phan_Tich_*.pdf)
☐ Mở Excel: mọi ô forecast là formula (KHÔNG hardcode số)?
☐ Mở Excel: sheet 13_PE_PB_History có MEDIAN cuối bảng?
☐ PDF: tất cả chart hiển thị? (tối thiểu 10 chart cho ngân hàng)
☐ Recalc = 0? (không có lỗi #REF!, #DIV/0!, #VALUE!)
```

## Nguyên tắc bất di bất dịch

1. **KHÔNG SKIP**: Mỗi phase, mỗi bước trong skill đều phải thực hiện. Nếu bước nào không làm được → DỪNG LẠI báo user.
2. **KHÔNG GIỚI HẠN SHEET**: Có thể tạo thêm sheet bất kỳ lúc nào để phục vụ tính toán trung gian. Sheet tên: `<số thứ tự>_<Tên>`.
3. **TẤT CẢ DỮ LIỆU TỪ API**: Mọi số liệu tài chính phải lấy từ Vietcap API (BCTC + statistics-financial). Không hardcode số BCTC.
4. **MEDIAN, KHÔNG PHẢI AVERAGE**: Định giá dùng median toàn bộ lịch sử (PE, PB, EV/EBITDA). Năm nào cũng lấy median của 4 quý trong năm đó.
5. **LNST ÂM → P/E CARRY**: Quý nào LNST âm, P/E quý đó = P/E quý trước. Vẽ biểu đồ với marker khác màu.
6. **A/E PHÂN KỲ**: Năm A = số thực từ BCTC. Năm E = số dự báo. Mọi ô E trong Excel PHẢI là công thức.
7. **FORMULA BẮT BUỘC**: Mọi ô forecast trong Excel PHẢI là công thức tham chiếu chéo sheet. Tuyệt đối không hardcode số dự báo.
