# Changelog — Skill Thép

## 2026-07-02 (m) — Audit toàn diện D&A/CAPEX/Nợ vay/Tiền mặt/Tỷ giá/HRC 2024/EBITDA — dữ liệu Vietcap bị bỏ quên không đồng bộ

### Bối cảnh
User nghi ngờ dữ liệu Vietcap bị fetch/ghi sai ở nhiều chỗ trong `02_Assumptions`/`04_PnL`, và nghi ngờ
sản lượng HRC Q2-Q4/2024 bất thường thấp (chỉ ~1/3 so với Q1/2025). Đây là lần audit toàn diện nhất
trong phiên làm việc — phát hiện các mảng "giả định lịch sử" trong `02_Assumptions` (D&A, CAPEX, Nợ
vay, Tiền mặt) chưa BAO GIỜ được đồng bộ với dữ liệu THẬT mà chính codebase đã fetch và dùng đúng ở
CÁC SHEET KHÁC (05/06) — cùng loại bug với SL_HRC_A ở mục (l), nhưng lan rộng hơn nhiều.

### Phát hiện (đối chiếu số chết vs dữ liệu thật Vietcap `.cache/HPG_bctc.json`)
| Chỉ tiêu | Field Vietcap | 2021 | 2022 | 2023 | 2024 | 2025 |
|---|---|---|---|---|---|---|
| D&A (số chết cũ) | — | 3.000 | 3.500 | 4.000 | 4.800 | 5.500 |
| D&A (thật, cfa2) | Khấu hao TSCĐ&BĐSĐT | 6.077 | 6.759 | 6.762 | 6.916 | 8.471 |
| CAPEX (số chết cũ) | — | 12.000 | 15.000 | 18.000 | 22.000 | 25.000 |
| CAPEX (thật, cfa19) | — | 11.621 | 17.888 | 17.374 | 35.495 | 25.748 |
| Nợ vay (số chết cũ) | — | 45.000 | 55.000 | 65.000 | 72.000 | 80.000 |
| Nợ vay (thật, bsa56+71) | Vay NH+DH | 57.213 | 57.900 | 65.381 | 82.963 | 92.174 |
| Tiền mặt (số chết cũ) | — | 12.000 | 15.000 | 13.000 | 16.000 | 20.000 |
| Tiền mặt (thật, bsa2+bsa5) | Tiền&TĐT + ĐT ngắn hạn | 40.708 | 34.593 | 34.429 | 25.862 | 27.785 |

- **D&A 2024/2025 lệch ~1,4-1,5x** — khớp đúng nhận định user "khấu hao 2024 2025 lớn hơn số trong đó
  nhiều". Vì `04_PnL!row16` (EBITDA) dự phóng dùng công thức `=EBIT + Assumptions!D&A` chain từ gốc
  2025A SAI (5.500) → EBITDA 2026-2028E bị THẤP GIẢ TẠO trong khi LNST tính độc lập vẫn tăng đúng theo
  margin thật → đúng hiện tượng user mô tả ("LNST tăng mạnh nhưng EBITDA lại thấp quá khiến EV/EBITDA
  sai"). NI(LNST) không phụ thuộc D&A nên không bị ảnh hưởng, chỉ EBITDA/EV-EBITDA bị méo.
- **Nợ vay số chết cũ 2025 (80.000) SAI, thật là 92.174** — đối chiếu chéo với ghi chú có sẵn trong
  chính ô đó "Q1/2026: 90,6k giảm 2k" → 92.174 - ~2.000 ≈ 90.174, khớp RẤT SÁT với "90,6k" hơn hẳn số
  chết cũ 80.000 — tự xác nhận số thật đúng hướng trước khi sửa.
- **Tiền mặt: user hỏi nên lấy "tiền tương đương tiền" (bsa2) hay "đầu tư tài chính ngắn hạn" (bsa5)** —
  câu trả lời là CẢ HAI CỘNG LẠI, vì dòng "Tiền mặt" ở Assumptions chỉ dùng cho công thức Net Debt (Nợ
  vay - Tiền mặt) trong các phép tính EV — Đầu tư ngắn hạn của HPG chủ yếu là tiền gửi kỳ hạn, thanh
  khoản cao tương đương tiền mặt cho mục đích này. `cash_hist` (bsa2 riêng) vẫn giữ nguyên dùng cho
  sheet 05 dòng "Tiền & tương đương" (khái niệm kế toán riêng, không được gộp).
- **Tỷ giá quy đổi doanh thu HRC/XD bottom-up dùng 1 hằng số `FX_RATE=25400` cho CẢ 8 năm** (2021-2028)
  — sai vì USD/VND tăng đáng kể qua các năm (~23.100 năm 2021 → ~26.200 hiện tại). Kể cả trong Excel,
  công thức doanh thu HRC/XD (`03_Revenue_Model!row8-9`, `row24`) hardcode SỐ CHẾT "25400" trực tiếp
  trong formula string cho MỌI cột năm — cùng loại lỗi nhưng ở tầng Excel.
- **HRC Q2-Q4/2024 SAI nghiêm trọng** — verify qua báo chí (WebSearch, thitruongtaichinhtiente.vn,
  nhipsongkinhdoanh.vn): số cũ [464, 312, 399] hoàn toàn sai; số THẬT: Q3/2024 = 738 nghìn tấn (khớp
  CHÍNH XÁC bài báo), 9 tháng đầu 2024 = 2,27 triệu tấn (khớp CHÍNH XÁC 805+727+738=2270 → suy ra
  Q2=727, khớp với thông tin "Q2 giảm ~10% so với Q1" = 805×0,9≈725); Q4/2024 CHƯA có số công bố trực
  tiếp, ước tính bằng phần dư từ "sản xuất cả năm hơn 3 triệu tấn (+5% so với 2023)" → Q4≈780 (cần
  verify lại nếu tìm được báo cáo Q4/2024 cụ thể hơn — đã ghi chú rõ trong code là số ước tính phần dư).
- **EBITDA lịch sử (`04_PnL!row16`) là SỐ CHẾT Python** (`ebit + da_hist[j-2]`, không phải formula) —
  dù bản thân SỐ đã đúng (da_hist là thật), user không bấm vào ô kiểm chứng được. Đổi hẳn sang công
  thức sống `=EBIT(row7) + Assumptions!D&A(row9)` cho MỌI năm (trước đây chỉ formula cho năm dự phóng).

### Thay đổi (`build_hpg_model.py`)
- Thêm `st_invest_hist` (bsa5) và `cash_for_valuation_hist` (bsa2+bsa5) — module-level, cạnh
  `cash_hist`/`total_debt_hist` đã có.
- `02_Assumptions`: D&A/CAPEX/Nợ vay/Tiền mặt (rows 9-12) — 5 cột lịch sử (2021-2025) đổi từ số chết
  sang `da_hist`/`capex_hist`/`total_debt_hist`/`cash_for_valuation_hist` (thật, đã fetch sẵn nhưng bị
  bỏ quên không dùng ở Assumptions). Nợ vay/Tiền mặt dự phóng (2026E-2028E) GIỮ NGUYÊN tỷ lệ %YoY của
  giả định cũ nhưng neo lại đúng gốc 2025A thật (tránh bước nhảy phi lý 2025A→2026E). D&A dự phóng
  không cần sửa vì đã là công thức sống chain từ năm trước — tự động đúng khi gốc đúng.
- Thêm `FX_RATE_HIST_A`/`FX_RATE_A` (8 năm, thay `FX_RATE` hằng số) — 2021-2025 dùng bình quân năm
  tham khảo Vietcombank/SBV (nghiên cứu thủ công, Vietcap API không có trường tỷ giá); 2026E dùng
  `_usd_vnd_rate` đã fetch sống (investing.com); 2027E/2028E nối tiếp xu hướng mất giá ~2%/năm. Thêm
  dòng mới `02_Assumptions!row25` "Tỷ giá USD/VND" để `03_Revenue_Model` LINK công thức sống thay vì
  hardcode "25400" (đã sửa `row8`/`row9`/`row24`).
  `hrc_rev_all`/`xd_rev_all` (Python) đổi sang dùng `FX_RATE_A[i]` theo từng năm thay vì `FX_RATE` cố định.
- `HRC_SALES_HIST_KT` 2024Q2-Q4: `[464, 312, 399]` → `[727, 738, 780]` (2 số đầu verify chính xác qua
  báo chí, số cuối ước tính phần dư — xem chú thích trong code) — tự động lan qua `SL_HRC_A[3]` (2024,
  đã link sum-4-quý từ mục (l)) và mọi chỗ dùng `HRC_SALES_HIST_KT`.
- `04_PnL!row16` (EBITDA): bỏ nhánh Python tính sẵn cho lịch sử, dùng CHUNG 1 công thức sống
  `=EBIT+Assumptions!D&A` cho cả 8 năm.

### Kết quả thực tế (verify)
- Đối chiếu Nợ vay 2025A mới (92.174) với ghi chú có sẵn "Q1/2026: 90,6k giảm 2k" — khớp rất sát,
  TỰ XÁC NHẬN hướng sửa đúng trước khi chạy thử.
- Chạy full pipeline (`build_hpg_model.py` + `run_analysis.py HPG`) không lỗi, đọc lại Excel qua
  openpyxl xác nhận đúng: D&A/CAPEX/Nợ vay/Tiền mặt hiển thị số thật, `03_Revenue_Model!row8` link
  đúng `'02_Assumptions'!{col}25` (tỷ giá năm đó) thay vì "25400", `04_PnL!row16` là formula
  `=B7+'02_Assumptions'!B9` cho MỌI cột (kể cả lịch sử).

### Chưa làm
- Q4/2024 HRC vẫn là SỐ ƯỚC TÍNH PHẦN DƯ (780kt), chưa tìm được báo cáo/BCTC công bố trực tiếp — cần
  verify lại nếu có nguồn chính xác hơn (khác với Q2/Q3/2024 đã verify chính xác qua báo chí).
- FX_RATE_HIST_A (2021-2025) là số nghiên cứu thủ công tham khảo (Vietcombank/SBV), CHƯA fetch tự động
  — khác với thông lệ "fetch sống" đã áp dụng cho commodity prices; Vietcap API không có trường tỷ giá
  nên chưa tìm được nguồn miễn phí đủ dài/đủ tin cậy để tự động hoá hoàn toàn như World Bank Pink Sheet.
- Chưa rà soát toàn bộ các sheet khác (07-16) xem còn "số chết" nào tương tự chưa đồng bộ với dữ liệu
  Vietcap thật — audit này tập trung đúng phạm vi user chỉ ra (Assumptions row9/11/12, PnL row16, FX,
  HRC 2024), có thể còn sót ở nơi khác.

---

## 2026-07-02 (l) — Sửa SL_HRC/XD sai lệch số thật + link Spread hàng năm sang sheet 17 + BLNG dự phóng theo n quý thực tế

### Bối cảnh
User kiểm tra kỹ 02_Assumptions/03_Revenue_Model và phát hiện 2 bug nghiêm trọng làm sai lệch cả model:
1. `03_Revenue_Model!row4` (SL HRC năm): 2025 hiển thị 2.8 triệu tấn nhưng lũy kế 4 quý THẬT (đã có
   trong `HRC_SALES_HIST_KT`) là 5.0 triệu tấn — mảng Python `SL_HRC_A`/`SL_XD_A` là giả định TĨNH cũ,
   không được đồng bộ lại sau khi có dữ liệu quý thật, khiến doanh thu 2025 bị tính THIẾU và %tăng
   trưởng DT 2026E bị THỔI PHỒNG giả tạo (so với mẫu số 2025 sai thấp).
2. `02_Assumptions!row45` ("Spread hàng năm") dùng công thức ĐỘC LẬP tự trừ lại Quặng/Than thay vì LINK
   sang sheet `17_Gia_Hang_Hoa` (nơi đã có sẵn Spread HRC/All theo năm, tính đúng lag-1-quý/AD20/bình
   quân gia quyền) — 2 nguồn số liệu độc lập dễ lệch nhau âm thầm giống hệt bug SL_HRC_A. User yêu cầu:
   ≤2022 dùng MEDIAN Spread HRC ("spread thép") của 4 quý, ≥2023 dùng MEDIAN Spread All của 4 quý.
3. `02_Assumptions!row6` (BLNG dự phóng 2026E) dùng công thức đơn giản (BLNG_Q1 x tỉ lệ Spread) — user
   yêu cầu công thức TỔNG QUÁT theo n quý ĐÃ CÓ BCTC (n=1,2,3,4), dùng Spread ALL (không phải Spread
   HRC) làm tỉ lệ điều chỉnh phần doanh thu CÒN LẠI chưa biết.

### Thay đổi (`build_hpg_model.py`)
- **Sửa SL_HRC_A/SL_XD_A (module-level, các năm 2023-2025):** ghi đè bằng SUM 4 quý thật từ
  `HRC_SALES_HIST_KT`/`XD_SALES_HIST_KT` (chỉ ghi đè khi năm đó có ĐỦ 4 quý dữ liệu — 2021/2022 chưa có
  dữ liệu tách quý nên giữ giả định cũ). Kết quả xác nhận: 2023=2.8Mt (khớp thật), 2024=1.98Mt (khớp
  thật), 2025=5.0Mt (khớp thật, đúng con số user nêu) — trước đó lần lượt là 2.5/2.8/3.2Mt (sai).
- **`03_Revenue_Model!row4-5`** (SL HRC/XD) năm 2023-2025: đổi từ SỐ CHẾT (Python literal) sang CÔNG
  THỨC SUM SỐNG link `15_Quarterly_Data` (nguồn duy nhất) — đặt ở khối patch CUỐI hàm `build_excel()`
  (không phải lúc build sheet 03) vì cần `R_QV_HRC`/`R_QV_XD` của sheet 15, mà sheet 03 được build
  TRƯỚC sheet 15 trong thứ tự code hiện tại.
- **SL_HRC_A/SL_XD_A[6] (2027E)** — theo yêu cầu user, đổi từ giả định tĩnh sang công thức: SL 2 quý
  GẦN NHẤT đã biết (nửa năm) x2 (năm hóa) x1.05 (tăng trưởng giả định); **[7] (2028E)** nối tiếp x1.05
  vì chưa có dữ liệu quý mới hơn để re-anchor.
- **Thêm `_r17_annual_row_layout()`** (module-level, gọi 1 lần ra `R17_LAYOUT`): tính TRƯỚC vị trí dòng
  cố định của sheet 17 (chỉ phụ thuộc `len(Q18_LABELS)`, không đổi giữa các lần chạy) để
  `02_Assumptions` (build TRƯỚC sheet 17) có thể LINK công thức đúng ô mà không cần dựng sheet 17
  trước. Thêm `assert` đối chiếu layout thật của sheet 17 khi build với `R17_LAYOUT` — lệch nhau sẽ báo
  lỗi ngay thay vì âm thầm link sai ô (đúng bài học từ bug SL_HRC_A).
- **`02_Assumptions!row45`**: link `={S17}!{col}{R17_ANN_SPREAD}` (≤2022) hoặc
  `={S17}!{col}{R17_ANN_SPREAD_ALL}` (≥2023) — không còn tự trừ lại độc lập.
- **`02_Assumptions!row6`** (BLNG dự phóng) — viết lại hoàn toàn theo công thức user (áp dụng SONG SONG
  cho cả Excel formula VÀ biến Python `gpm_2026/2027/2028` — 2 nơi PHẢI khớp nhau vì Python còn dùng
  cho narrative PDF/JSON):
  - Tổng quát hóa "Q1/2026" (rows 40-43) thành "lũy kế N quý" dùng `cumulative_actual_quarters()` (dùng
    lại hàm đã có ở `fetch_data.py`) — tự động đúng khi rerun script vào quý sau (N=2,3,4).
  - `LNG năm = LNG lũy kế N quý + (4-N)/4 x Doanh thu ước tính năm x (LNG lũy kế N/Doanh thu lũy kế N)
    x (Spread All hiện tại/Spread All quý gần nhất đã biết)`; N=4 → dùng thẳng lũy kế 4 quý, không ước
    tính. "Spread All quý gần nhất" LUÔN lấy dòng cuối bảng 18 quý sheet 17 (`R17_Q_LAST`) — quy ước
    bảng này phải được bảo trì cập nhật mỗi khi có BCTC quý mới, nên tự động khớp đúng N mà không cần
    logic riêng theo từng giá trị N.
  - Thêm dòng mới `R_SPR_ALL_NOW` (Assumptions!row47, "Spread All hiện tại") — trước đây chỉ có Spread
    HRC hiện tại (row46, đổi tên rõ ràng "Spread HRC hiện tại" để phân biệt).
  - 2027E/2028E (chưa có quý nào của năm đó): `BLNG năm = BLNG năm TRƯỚC x (Spread All hiện tại/Spread
    All NĂM TRƯỚC)` — dùng CÙNG "Spread All hiện tại" cho cả 2 năm (không chain qua SPREAD_ALL_A như
    công thức cũ), khớp đúng yêu cầu user.

### Kết quả thực tế (verify)
- Chạy full pipeline, in debug tạm thời xác nhận: `SL_HRC_A = [2.0, 2.2, 2.8, 1.98, 5.0, 6.27, 6.3,
  6.62]` (2023-2025 khớp số thật), `revenue_growth_fc = [7.9, 17.8, 7.5]` (2026E từ mức "quá cao" trước
  đây giảm về 7.9% — hợp lý), `gp_margin_fc = [16.0, 15.6, 13.7]`.
- Chạy `python run_analysis.py HPG` full pipeline (Excel + PDF + JSON + registry) không lỗi.

### Chưa làm
- Chưa cập nhật narrative PDF (các đoạn text mô tả tăng trưởng DT/BLNG 2026E bằng số liệu cũ) — cần
  soát lại các chỗ hard-code con số cũ trong text tường thuật (không phải bảng số, các bảng số đã tự
  cập nhật qua biến Python).
- Assert đối chiếu `R17_LAYOUT` mới bảo vệ được Sheet 17's row layout — CHƯA có cơ chế tương tự bảo vệ
  sheet 15's `R_QV_HRC`/`R_QV_XD` (hiện vẫn tính cục bộ, không có shared layout function) nếu sau này
  cần link thêm từ sheet khác.

---

## 2026-07-02 (k) — 3 biểu đồ Spread HRC/Rebar/All: mở rộng 20 quý, làm mềm đường, phân biệt điểm dự báo

### Bối cảnh
Sau khi có 3 biểu đồ Spread HRC/Rebar/All (mục (j)), user yêu cầu: (1) mở rộng cửa sổ hiển thị lên 20
quý (từ 12) để dễ kiểm tra tính "đồng pha" (co-movement) giữa Spread và BLNG; (2) đường Spread vẽ tới
tận điểm "Spread hiện tại"; (3) làm MỀM cả 3 đường (Spread & BLNG) ở các đoạn gấp khúc tăng/giảm cho
trực quan hơn; (4) hỏi lại tại sao gọi là "lag 1 quý" — làm rõ chỉ chi phí (quặng/than) bị lag, giá bán
vẫn là giá quý hiện tại (khớp đúng công thức đã cài, chỉ là chỉnh lại câu chữ tiêu đề cho khỏi gây hiểu
lầm); (5) phát hiện điểm BLNG quý 2/2026 (chưa có BCTC thật) đang bị vẽ y hệt số liệu thật — yêu cầu
tách biệt màu sắc/kiểu vẽ để biết đó là DỰ BÁO, áp dụng cho cả 3 biểu đồ.

### Thay đổi
- Thêm dependency `scipy` (`requirements.txt`) — dùng `PchipInterpolator` để làm mềm đường. **Lưu ý quan
  trọng:** thử `make_interp_spline` (cubic B-spline tự do) trước, nhưng phát hiện OVERSHOOT/UNDERSHOOT
  nghiêm trọng giữa 2 điểm zigzag liên tiếp (tạo đỉnh/đáy ẢO không có trong số liệu thật — thấy rõ qua
  ảnh render). Đổi sang PCHIP (Piecewise Cubic Hermite) — đảm bảo đường cong không vượt khoảng giá trị
  của 2 điểm lân cận, mềm hơn đường thẳng nhưng KHÔNG bịa số liệu.
- `N_RECENT_Q`: 12 → 20 quý cho cả 3 biểu đồ (dùng chung hàm `_make_spread_gpm_chart`).
- Sửa tiêu đề biểu đồ: `"(lag 1 quý)"` → `"(giá bán quý này, chi phí quặng/than lag 1 quý)"` — làm rõ
  chỉ chi phí đầu vào bị lag, không phải toàn bộ Spread (công thức Python/Excel không đổi, chỉ đổi câu
  chữ hiển thị theo đúng bản chất công thức đã cài từ mục (a)/(j)).
- **Sửa bug logic "phân biệt điểm dự báo":** code cũ kiểm tra `if q_gpm[-1] is None` để nhận diện quý
  đang chạy (2026Q2) — nhưng `q_gpm[-1]` KHÔNG BAO GIỜ là `None` vì vị trí đó luôn được điền sẵn bằng
  công thức ước tính ngay từ lúc dựng mảng (`q_gpm_all_`), nên điều kiện không bao giờ đúng — điểm dự
  báo bị vẽ y hệt số liệu thật (vuông cam), đúng như user phát hiện. Sửa bằng cách nhận diện quý dự báo
  qua NHÃN (`q_lbls[-1] == _cur_lbl`) thay vì suy từ giá trị None/not-None. Verify bằng debug print thực
  tế giá trị `q_gpm[-1]`/`q_gpm_all_[-1]` trước khi sửa để xác nhận đúng nguyên nhân (không đoán mò).
- Điểm dự báo giờ vẽ: marker tam giác xanh lá (`^`, `#27AE60`) thay vì vuông cam, nối NÉT ĐỨT (không
  làm mềm) từ điểm thật gần nhất, có label riêng "BLNG ƯỚC TÍNH {quý} (dự báo — chưa có BCTC)" — áp
  dụng đồng nhất cho cả 3 biểu đồ Spread HRC/Rebar/All.
- Đường BLNG chỉ làm mềm PHẦN THỰC TẾ liên tục (loại điểm dự báo ra trước khi spline) — tránh PCHIP
  "kéo cong" đoạn cuối để khớp 1 điểm dự báo, gây hiểu lầm điểm đó cũng là số liệu thật.

### Kết quả thực tế (verify)
- Render lại cả 3 chart, xác nhận bằng mắt: không còn overshoot/undershoot giả ở các đoạn zigzag (VD
  Q1'23-Q4'23 Spread All trước khi sửa dip ảo xuống dưới 105, sau khi đổi PCHIP hết hẳn); điểm dự báo
  2026Q2 hiển thị rõ tam giác xanh + nét đứt, tách biệt hoàn toàn khỏi chuỗi số liệu thật.
- Chạy `python run_analysis.py HPG` full pipeline — gặp `OSError: [Errno 22] Invalid argument` khi ghi
  file .xlsx (transient, không liên quan tới code — retry ngay lần 2 thành công, nghi do antivirus/file
  index quét file đúng lúc ghi). Không phải lỗi cần sửa trong code.

---

## 2026-07-02 (j) — Tách Spread HRC/Rebar/All (thuế CBPG AD20 + giá thép XD quý thật + BLNG theo Spread All)

### Bối cảnh
User yêu cầu tách Spread thép thành 2 loại theo sản phẩm (HRC vs thép xây dựng/rebar) thay vì 1 Spread
chung dựa trên giá HRC như trước, vì 2 sản phẩm có cơ chế giá và chính sách thuế khác nhau:
`total_cost = 1.6×Giá quặng + 0.6×Giá than + OTHER_COST_USD`; `Spread HRC = Giá HRC - total_cost`;
`Spread Rebar = Giá thép XD - total_cost`. Đồng thời yêu cầu kiểm tra thuế chống bán phá giá (CBPG) HRC
Trung Quốc và áp hệ số 1.15 vào giá HRC từ ngày có hiệu lực để phản ánh đúng hơn kinh tế học HRC nội
địa, thêm `Spread All` (bình quân gia quyền theo sản lượng HRC/rebar), và đổi công thức dự phóng BLNG
năm sang dùng tỉ lệ Spread All (thay vì Spread HRC). **Yêu cầu cứng:** mọi công thức mới ở Excel
(assumption/PnL/sheet spread) phải là công thức sống, không số chết.

### Nghiên cứu trước khi code (theo yêu cầu user "dò đủ 18 quý VSA trước rồi mới code")
- Xác nhận vụ việc CBPG đúng là **AD20** (HRC, không phải AD19 = thép mạ như tài liệu cũ trong repo ghi
  nhầm) — thuế chính thức khổ hẹp hiệu lực **06/07/2025**, mở rộng chống lẩn tránh khổ rộng **17/04/2026**.
  User chọn mốc 06/07/2025 (rơi vào 2025Q3) làm điểm bắt đầu áp hệ số.
- Thử tìm giá thép XD theo quý thật từ VSA (bản tin tháng "Tình hình thị trường thép Việt Nam") — sau khi
  tải và đọc đúng nội dung (phát hiện lỗi debug ban đầu: in ra console bằng `.encode('ascii','replace')`
  làm mất dấu tiếng Việt, khiến tưởng nhầm regex sai — nội dung file gốc là UTF-8 hợp lệ) xác nhận **VSA
  bản tin tháng KHÔNG công bố giá bán rebar** (chỉ giá nguyên liệu đầu vào + sản lượng) → không dùng được.
- User cung cấp 2 nguồn thay thế: investing.com Steel Rebar futures (SRRc1) và SteelOnline.vn. Verify:
  SRRc1 niêm yết THẲNG USD/tấn (không cần quy đổi CNY như nhầm tưởng ban đầu), có API nội bộ
  (`api.investing.com/api/financialdata/historical/996702`) trả về lịch sử THÁNG đầy đủ 2021-10/2025,
  nhưng hợp đồng NGỪNG giao dịch thật sau 10/2025 (giá đứng yên, khối lượng = 0). SteelOnline.vn có giá
  Hòa Phát D10/CB240 THỰC nhưng chỉ có giá hiện tại (không lưu trữ lịch sử). Đối chiếu 2 nguồn cùng thời
  điểm: SRRc1 ~545 USD/tấn vs SteelOnline quy đổi ~581 USD/tấn (lệch ~6%, chấp nhận được).

### Thay đổi (`build_hpg_model.py`)
- Thêm `AD_HRC_EFFECTIVE_Q="2025Q3"`, `AD_HRC_MULTIPLIER=1.15`, `_hrc_ad_adjusted()` — áp hệ số vào MỌI
  công thức Spread HRC (quý/năm/hiện tại/dự phóng xa) kể cả 2027E-2028E (lỗi ban đầu: quên áp cho 2 năm
  này, đã tự phát hiện và sửa khi cross-check Excel formula vs Python).
- Thêm `Q18_XD`/`Q18_XD_SRC` (giá + nguồn thép XD 18 quý): 16/18 quý dùng MEDIAN THẬT SRRc1 (`"INV"`),
  2026Q1 dùng neo nội địa VSA đặc biệt quy đổi USD (`"VN"`), 2025Q4 nội suy tuyến tính (`"NC"`). Thêm
  `fetch_investing_rebar_monthly()`, `fetch_usd_vnd_rate()`, `fetch_steelonline_rebar_price()`.
- Thêm `Q18_SPREAD_HRC` (đổi tên/nội dung `Q18_SPREAD` cũ, giữ alias tương thích ngược), `Q18_SPREAD_REBAR`,
  `Q18_SPREAD_ALL` (bình quân gia quyền theo `HRC_SALES_HIST_KT`/`XD_SALES_HIST_KT` — chỉ có từ 2023Q1,
  5 quý đầu để `None`); annual `SPREAD_A`/`SPREAD_REBAR_A`/`SPREAD_ALL_A`; "hiện tại"
  `SPREAD_HRC_NOW`/`SPREAD_REBAR_NOW`/`SPREAD_ALL_NOW`.
- Đổi công thức dự phóng BLNG năm: `q1_spread_all = Q18_SPREAD_ALL[-1]` (dùng SL của CHÍNH quý gần nhất),
  tỉ lệ = `SPREAD_ALL_NOW / q1_spread_all` (trước đây dùng `SPREAD_A`/`q1_spread`, tức Spread HRC).
- Sheet `17_Gia_Hang_Hoa`: thêm cột G-L (Giá/nguồn thép XD, Spread Rebar, SL XD/HRC quý link sheet
  `15_Quarterly_Data`, Spread All) — TẤT CẢ công thức sống; cột E (Spread HRC) sửa thêm `*1.15` cho các
  quý ≥2025Q3; khối "Giá hiện tại" thêm dòng XD/Spread Rebar/Spread All hiện tại; khối "GIÁ NĂM" thêm
  dòng Giá thép XD, Spread Rebar, Spread All theo năm (Spread All năm dùng `SL_HRC_A`/`SL_XD_A` làm
  quyền số, ghi trực tiếp 2 dòng phụ SL vì đây là assumption Python, chưa có sheet Excel riêng theo năm).
- 3 biểu đồ quý riêng biệt (thay vì 1 chart HRC cũ): `spread_gp_quarterly.png` (HRC),
  `spread_rebar_gpm_quarterly.png`, `spread_all_gpm_quarterly.png` — dùng chung 1 hàm
  `_make_spread_gpm_chart()`. Cả 3 đều nhúng vào PDF (biểu đồ 7A/7A2/7A3) và JSON
  (`commodityQuarterly.spreadHrcUsd/spreadRebarUsd/spreadAllUsd`, `annualTables.spreadRebarUsd/spreadAllUsd`).

### Kết quả thực tế (verify)
- Chạy full pipeline nhiều lần, cross-check thủ công bằng tay 1 cell (`L27` = Spread All hiện tại):
  tính tay ra 156.4 → khớp đúng "156" hiển thị trên chart điểm cuối (2026Q2) — xác nhận công thức Excel
  và Python nhất quán nhau độc lập.
- KHÔNG verify được bằng win32com (mở Excel thật) như quy trình chuẩn của project — Excel COM tự động mở
  file lỗi `-2146827284` trong môi trường này (không phải lỗi code, không tìm được nguyên nhân trong thời
  gian cho phép) — đã bù bằng cross-check thủ công công thức + giá trị số như trên.
- Phát hiện investing.com có thể tạm chặn/rate-limit sau khi gọi curl nhiều lần liên tục trong 1 phiên
  làm việc dài (test connectivity riêng xác nhận `curl` tới investing.com trả về lỗi kết nối trong khi
  các domain khác — Google, World Bank, SteelOnline — vẫn OK) — script vẫn chạy đúng nhờ fallback graceful
  (không dừng), nhưng cần lưu ý khi user thấy nhiều quý XD/HRC/Iron/Coal "hiện tại" giống hệt giá quý gần
  nhất (dấu hiệu fetch lỗi tạm thời, không phải bug).

### Chưa làm
- Chưa cập nhật narrative PDF (mục "SPREAD & YẾU TỐ ĐẦU VÀO") để diễn giải bằng lời về Spread
  HRC/Rebar/All riêng biệt — hiện narrative cũ vẫn nói chung chung "Spread thép" (đúng số nhưng chưa
  phân tách rõ 3 loại trong văn bản).
- `XD_PRICE_A` (giả định giá thép XD theo năm dùng cho doanh thu bottom-up ở `03_Revenue_Model`) CHƯA
  đồng bộ với `Q18_XD` mới — vẫn là mảng giả định cũ độc lập, chỉ dùng `Q18_XD` cho các phép tính Spread.
  Nếu muốn nhất quán hoàn toàn cần thay `XD_PRICE_A[:5]` bằng median `Q18_XD` theo năm (tương tự
  `HRC_PRICE_A`).
- Chưa test lại toàn bộ sau khi investing.com hết bị chặn (để xác nhận `Q18_XD`/`hrc_now`/`iron_now`/
  `coal_now` fetch thật hoạt động đúng như lần chạy đầu tiên thành công).

---

## 2026-07-02 (i) — DTTC/CPTC/sản lượng 2026E blend với số quý ĐÃ CÓ báo cáo (không ngoại suy Q1×4)

### Bối cảnh
User phát hiện bug: `FIN_INCOME_FC = [2500, 2800, 3000]` (giả định DTTC năm 2026 cố định) THẤP HƠN
CHÍNH số DTTC Q1/2026 thực tế một mình (~5.938 tỷ, đột biến do lãi tỷ giá/cổ tức — xem mục 4C) — bỏ
qua hoàn toàn thực tế đã công bố. Nhưng user cũng lưu ý: KHÔNG được sửa bằng cách ngoại suy tuyến tính
Q1×4 (sẽ thổi phồng sai cả năm theo đúng yếu tố đột biến 1 quý). Yêu cầu công thức tổng quát, áp dụng
cả cho HPG và template ngân hàng — xem chi tiết công thức tại skill `ngan-hang` (mục "Cập nhật KQKD &
tài sản dự phóng NĂM theo diễn biến quý ĐÃ CÓ báo cáo").

### Thay đổi
- Thêm `cumulative_actual_quarters()`/`blend_annual_estimate()` vào `fetch_data.py` (dùng chung với
  `template_banking.py`) — công thức: Ước tính năm = Lũy kế thực tế n quý đã biết + Giả định ban đầu
  cả năm × (4-n)/4.
- `FIN_INCOME_FC[0]`/`FIN_COST_FC[0]` (2026E): blend với DTTC/CPTC Q1/2026 thực tế (`isa6`/`isa7`)
  thay vì giữ nguyên giả định cố định — cascades tự nhiên qua `ebt_fc`/`ni_fc` (không cần sửa gì thêm
  ở chuỗi tính phía sau, giống cách sửa `SL_HRC_A`/`SL_XD_A` production volume trước đó).
- `SL_HRC_A[5]`/`SL_XD_A[5]` (sản lượng 2026E, xem bản (e)-(g)): đổi từ cách tính "run-rate = TB(2 quý
  đã biết)×4" (ad-hoc, KHÔNG giữ giả định gốc cho phần chưa biết) sang `blend_annual_estimate()` chuẩn
  — coi Q1 (n=1, luôn có trong `HRC_SALES_HIST_KT`) + quý đang chạy (n=2, MIỄN LÀ có `CUR_Q_TOTAL_KT`
  dù chính thức hay ước tính từ tháng) là "đã biết", 2 quý còn lại vẫn theo giả định gốc 6.0/3.0 triệu
  tấn × 2/4 (không suy diễn từ H1 forward như cách cũ).

### Kết quả thực tế (2026-07-02)
- DTTC 2026E: 2.500 → 7.813 tỷ (= 5.938 Q1 thực tế + 2.500×3/4). CPTC 2026E: 3.800 → 4.719 tỷ.
- `ni_fc[0]` (LNST 2026E): tăng từ ~18.854 lên **26.618 tỷ** — khớp RẤT SÁT với ước tính riêng trong
  narrative PDF mục 4B ("Nếu HPG duy trì LNST bình quân 6.000-7.000 tỷ/quý Q2-Q4, cả năm có thể đạt
  26-28.000 tỷ") — xác nhận bug cũ (LNST dự phóng cũ 18.854 tỷ) thực sự MÂU THUẪN với chính narrative
  của báo cáo, đã hết mâu thuẫn sau khi sửa.
- `SL_HRC_A[5]`: 6.53→6.27 triệu tấn, `SL_XD_A[5]`: 5.43→4.22 triệu tấn (thận trọng hơn cách run-rate
  cũ vì giữ nguyên 1/2 giả định gốc cho 2 quý chưa biết thay vì ngoại suy từ H1).

---

## 2026-07-02 (h) — Giá quặng "hiện tại" đổi sang World Bank (đồng nhất phương pháp luận với lịch sử)

### Bối cảnh
User hỏi: giá quặng/HRC/than LỊCH SỬ lấy từ World Bank, nhưng giá HIỆN TẠI lại lấy từ investing.com —
2 nguồn có đồng nhất không? Và nghi ngờ "spread hiện tại" thấp bất thường có phải do lệch nguồn.

### Chẩn đoán (đã trace code, không đoán)
- **"Spread hiện tại" (`SPREAD_NOW`) KHÔNG bị ảnh hưởng bởi investing.com** — công thức
  `_lag_spread(_hrc_avg_now, Q18_IRON[-1], Q18_COAL[-1])` dùng `Q18_IRON[-1]` (median World Bank của
  Q1/2026, đã đúng từ bản (d)) làm chi phí quặng, KHÔNG dùng `iron_now` (investing.com) — nên nghi ngờ
  "spread thấp do quặng investing.com" là SAI, đã xác nhận bằng cách chạy thử: đổi nguồn quặng hiện tại
  chỉ làm SPREAD_NOW đổi từ 69,9 → 69,0 (chênh do giá HRC live thay đổi giữa 2 lần chạy, không phải do
  quặng).
- **Nguyên nhân thật của spread thấp**: than cốc Q1/2026 (220 USD/t, theo dữi liệu 5 quý gần nhất
  182→184→190→212→220 — TĂNG DẦN) trong khi giá HRC gần như đi ngang (460-480 USD/t) — chi phí than
  tăng đều trong khi giá bán không tăng theo là nguyên nhân chính, không phải lỗi nguồn dữ liệu quặng.
- **NHƯNG** user vẫn đúng ở điểm khác: `iron_now` (investing.com futures SGX TSI 62% Fe) và
  `Q18_IRON`/World Bank Pink Sheet (spot bình quân tháng) là 2 PHƯƠNG PHÁP LUẬN KHÁC NHAU (futures vs
  spot trung bình tháng) — trộn lẫn trong cùng 1 chuỗi "hiện tại vs lịch sử" (cụ thể: dòng "Giá hiện
  tại" C25 sheet 17, và `IRON_ORE_A[5]` dùng để dự phóng 2027E/2028E) là không nhất quán, dù không phải
  nguyên nhân gây spread thấp cụ thể lúc này.

### Thay đổi
- Sau khi fetch `iron_now` từ investing.com, GHI ĐÈ bằng giá tháng MỚI NHẤT trong `IRON_MONTHLY` (World
  Bank, đã fetch sẵn cho việc tính median quý) nếu có — cùng nguồn/phương pháp luận với 18 quý lịch sử.
  Investing.com chỉ còn là FALLBACK khi World Bank fetch lỗi. HRC & than cốc VẪN dùng investing.com
  (chưa có nguồn World Bank tương đương cho 2 mặt hàng này).
- Đánh đổi: World Bank có độ trễ ~1 tháng (không phải giá "hôm nay" thật), investing.com theo giờ
  nhưng khác cơ sở — chọn đồng nhất phương pháp luận (World Bank) hơn là độ mới (investing.com) cho
  quặng, vì quặng vốn đã 100% World Bank ở phần lịch sử.
- Cập nhật ghi chú trong sheet 17 (tiêu đề mục "GIÁ HIỆN TẠI", cột nguồn dữ liệu) để nói rõ quặng dùng
  World Bank tháng mới nhất, không phải investing.com nữa.

### Kết quả thực tế (2026-07-02)
- Quặng "hiện tại": investing.com 98,36 → World Bank tháng 5/2026 = 108,64 USD/t (chênh ~10%, đúng như
  user nghi ngờ có lệch giữa futures và spot).
- `SPREAD_NOW` gần như không đổi (69,9→69,0, do HRC live thay đổi giữa 2 lần chạy, không phải do quặng)
  — xác nhận quặng KHÔNG phải nguyên nhân spread thấp.
- `IRON_ORE_A[5]` (2026E) tăng nhẹ (101,2→106,6) do quặng hiện tại giờ cao hơn → `SPREAD_A[6]`/`[7]`
  (2027E/2028E, có lag chi phí từ 2026E) giảm nhẹ (81,3→73,0 và 81,0→72,4) — hợp lý, phản ánh chi phí
  quặng nền cao hơn khi dùng nguồn nhất quán.

---

## 2026-07-02 (g) — Thêm nguồn dautucophieu.net (báo cáo HSC Research, cadence hàng tháng đều đặn)

### Bối cảnh
User hỏi tiếp về nguồn `dautucophieu.net` (đăng lại báo cáo HSC Research) — kiểm tra xem có tích hợp
được không, vì bài có "ảnh thống kê" (lo ngại số liệu chỉ nằm trong ảnh, không đọc được bằng text).

### Điều tra
- Biểu đồ "Sản lượng tiêu thụ theo tháng" ĐÚNG LÀ ảnh (không đọc số liệu từ ảnh được, không dùng OCR/AI
  theo yêu cầu user), NHƯNG số liệu QUAN TRỌNG NHẤT vẫn được nhắc lại bằng TEXT trong bài, mẫu câu:
  "HPG bán được X tấn {thép xây dựng | HRC} trong tháng N/YYYY, tăng Y% so với cùng kỳ..." — verify khớp
  với bài tháng 12/2025 (585.000 tấn thép xây dựng).
- Trang có `/tag/hpg/` — WordPress tag page SERVER-RENDERED (không cần JS), liệt kê bài mới nhất TRƯỚC,
  kèm ngày đăng — cơ chế dò tin ĐƠN GIẢN NHẤT trong 3 nguồn đã tích hợp (chỉ 1 lần fetch, không cần
  sitemap nhiều trang).
- Cadence xuất bản RẤT ĐỀU trong lịch sử 2023-2025 (gần như tháng nào cũng có bài riêng cho HPG) —
  nhưng bài HPG gần nhất tính đến 2026-07-02 chỉ là "Hai tháng đầu năm 2026" (đăng 06/04/2026, dữ liệu
  lũy kế tháng 1+2/2026) — CHƯA có gì mới hơn cho Q2/2026, không giúp ích ngay cho quý đang chạy, nhưng
  là nguồn cross-check tốt cho các quý đã qua và sẽ tự động bắt được nếu site đăng lại.

### Thay đổi
- Thêm `fetch_dautucophieu_production_updates()`: fetch `/tag/hpg/`, lọc URL chứa `-hpg-` + tiêu đề có
  "sản lượng"/"tháng", parse theo mẫu "HPG bán được X tấn {sản phẩm} trong tháng N[/YYYY]" — CHỈ lấy số
  liệu ĐÃ XÁC NHẬN (không lấy các câu "dự kiến đạt khoảng..." vì đó là ước tính của người viết báo cáo
  tại thời điểm bài đăng, không phải số chính thức — tránh chồng thêm 1 lớp ước tính lên ước tính).
- Gộp vào cùng luồng ưu tiên với nguoiquansat.vn (tách HRC/XD riêng) khi trùng tháng/năm với
  hoaphat.com.vn (chỉ có tổng gộp).
- Sửa 1 bug nhỏ trong lúc test: regex tìm ngày đăng ban đầu tìm sai vị trí (do test bằng text đã strip
  HTML tag nhưng code thật chạy trên HTML gốc) — sửa lại đúng theo cấu trúc thật:
  `<span class="entry-meta">...Thứ Năm, DD/MM/YYYY - HH:MM</span>`.

### Kết quả thực tế (2026-07-02)
- Tìm được bản ghi tháng 12/2025 (585kt thép xây dựng) — không đổi kết quả Q2/2026 hiện tại (vẫn dựa
  vào nguoiquansat.vn tháng 5/2026 là nguồn duy nhất cho quý đang chạy) nhưng chứng minh cơ chế hoạt
  động đúng và sẽ tự bắt bài mới nếu dautucophieu.net đăng lại cho Q2/2026.

### Tổng kết 3 nguồn đang chạy (tính đến bản này)
| Nguồn | Cơ chế dò tin | Độ chi tiết | Trạng thái |
|---|---|---|---|
| hoaphat.com.vn | "Tin liên quan" (5 tin mới nhất) + sitemap 4 trang | Chỉ tổng gộp | Có Q1/2026 (chính thức) |
| nguoiquansat.vn | sitemap theo ngày, quét 70 ngày | Tách riêng HRC/XD | Có tháng 5/2026 |
| dautucophieu.net | `/tag/hpg/` (1 lần fetch) | Tách riêng HRC/XD (khi có) | Có tháng 12/2025, chưa có Q2/2026 |

---

## 2026-07-02 (f) — Thêm nguồn nguoiquansat.vn (tách riêng HRC/XD theo tháng) — verify với dữ liệu thật

### Bối cảnh
User cung cấp danh sách các nguồn hay đăng tin sản lượng HPG (hoaphat.com.vn, cafef.vn, tapchicongthuong.vn,
nguoiquansat.vn, mekongasean.vn, vietstock.vn, tinnhanhchungkhoan.vn, 24hmoney.vn, hoaphatdungquat.vn,
dautucophieu.net, và báo cáo cập nhật sản lượng của VCBS/Vietstock/Vietcap) để chủ động cập nhật sản
lượng tháng — bổ sung cho bản (e) mới build hôm trước (lúc đó CUR_Q_SOURCE="FALLBACK" vì chưa tìm ra
nguồn có dữ liệu tháng 4/5/6-2026).

### Kết quả điều tra (đã thử từng nguồn, không đoán)
- **nguoiquansat.vn — THÀNH CÔNG, giá trị cao nhất**: bài "Hòa Phát (HPG): Sản lượng bán thép xây dựng
  tháng 5/2026 giảm, HRC vẫn lập kỷ lục mới" (đăng lại báo cáo Vietcap Research, 26/06/2026) cho SỐ
  LIỆU THẬT tháng 5/2026: **HRC 622.000 tấn, thép xây dựng 429.000 tấn** — TÁCH RIÊNG 2 sản phẩm, chi
  tiết hơn hẳn tin PR gộp chung của hoaphat.com.vn. Trang có `sitemap-article-YYYY-MM-DD.xml` (sitemap
  THEO NGÀY, tĩnh, có sẵn tiêu đề trong `<image:title>` nên lọc được mà không cần fetch từng bài) —
  quét ngược 70 ngày để dò bài mới.
- **cafef.vn, tapchicongthuong.vn, mekongasean.vn, tinnhanhchungkhoan.vn, 24hmoney.vn,
  dautucophieu.net, hoaphatdungquat.vn**: là các bài viết ĐƠN LẺ (không phải trang danh sách/index có
  thể quét lặp lại) — hữu ích để đối chiếu thủ công nhưng KHÔNG dùng làm nguồn tự động hoá định kỳ vì
  không có cơ chế "tìm bài mới nhất" tương tự sitemap. Chưa tích hợp.
- **finance.vietstock.vn/HPG/tin-tuc-su-kien.htm**: trang tổng hợp tin theo mã CP — chưa kiểm tra sâu
  cấu trúc (JS-rendered hay không), để dành cho lần sau nếu cần thêm nguồn.
- **Báo cáo PDF của VCBS/Vietstock/Vietcap** (link download trực tiếp): có khả năng chứa bảng số liệu
  sản lượng có cấu trúc (đáng tin cậy hơn parse văn xuôi), nhưng cần thư viện đọc PDF (`pdfplumber`
  hoặc tương tự) và link download có token/hash đổi theo từng báo cáo (không cố định như World Bank
  Pink Sheet) — CHƯA tích hợp, để dành việc sau nếu user muốn.

### Thay đổi
- Thêm `fetch_nguoiquansat_production_updates()`: quét `sitemap-article-YYYY-MM-DD.xml` 70 ngày gần
  nhất (lọc tiêu đề qua `<image:title>` có sẵn trong sitemap, không cần fetch từng bài để lọc — chỉ
  fetch bài đã lọc khớp từ khóa "Hòa Phát"/"HPG" + "sản lượng"), regex tách riêng "thép xây dựng" và
  "thép cuộn cán nóng (HRC)" theo mẫu câu "sản lượng {sản phẩm}...trong tháng N[/YYYY] đạt X tấn".
- Gộp kết quả 2 nguồn (hoaphat.com.vn + nguoiquansat.vn) theo (loại, năm, quý/tháng) — nếu trùng, ƯU
  TIÊN bản ghi có tách HRC/XD riêng (nguoiquansat/Vietcap, chính xác hơn) thay vì bản gộp chung
  (hoaphat.com.vn PR).
- Khi CUR_Q_SOURCE="ESTIMATED" VÀ tất cả tháng đã biết đều có tách HRC/XD trực tiếp → dùng thẳng
  (`CUR_Q_HRC_KT_DIRECT`/`CUR_Q_XD_KT_DIRECT`) thay vì tỷ lệ lịch sử Q1/2026 (chính xác hơn).
- Sheet 15 mục E: thêm 2 cột riêng "SL HRC" và "SL XD" (trước đây chỉ có 1 cột tổng gộp) + cột "Tổng"
  — mỗi cột có công thức ước tính quý riêng (`AVERAGE()×3`).

### Kết quả thực tế (2026-07-02, sau khi thêm nguồn)
- `CUR_Q_SOURCE` chuyển từ `"FALLBACK"` → `"ESTIMATED"`: tìm được tháng 5/2026 (HRC 622kt, XD 429kt),
  ước tính Q2/2026 = 3.153 tấn (HRC 1.866kt, XD 1.287kt — nhân 3 từ 1 tháng duy nhất, vẫn CHỈ 1/3 tháng
  nên độ tin cậy còn hạn chế, cần chạy lại khi có thêm tháng 4 hoặc 6).
- `SL_HRC_A[5]`/`SL_XD_A[5]` (2026E, dùng cho `03_Revenue_Model!G4/G5` → LNST dự phóng) tự cập nhật
  theo run-rate mới: HRC 6.0→6.53 triệu tấn, XD 3.0→5.43 triệu tấn (tăng mạnh do tháng 5 là tháng yếu
  của XD theo báo cáo Vietcap — "giảm 19% YoY, giảm 9% MoM" — cần theo dõi thêm khi có tháng 4/6 để bức
  tranh đầy đủ hơn, hiện chỉ dựa vào 1 tháng).
- **Bug console encoding lặp lại**: lại quên quy ước ASCII-only cho `print()` khi thêm dòng log mới
  (`tách trực tiếp từ nguồn` có dấu) — làm crash trên console Windows cp1252 giống bản (d). Đã sửa.
  **Lưu ý cho lần sau: MỌI `print()` mới thêm vào file này phải kiểm tra lại bằng ASCII thuần trước khi
  coi là xong, không chỉ dựa vào 1 lần chạy thử test (build_pdf/build_excel) mà cần chạy đúng lệnh
  `python build_hpg_model.py` từ đầu tới cuối.**

### Chưa làm
- Chỉ có 1/3 tháng cho Q2/2026 — nên chạy lại script định kỳ (theo tuần) để bắt thêm tháng 4/6 khi có,
  hoặc tháng 6 khi Vietcap/nguoiquansat đăng cập nhật mới.
- Chưa tích hợp đọc PDF báo cáo cập nhật sản lượng (VCBS/Vietstock/Vietcap) — nếu làm sẽ chính xác và
  có cấu trúc hơn parse HTML/văn xuôi, nhưng cần thư viện đọc PDF + xử lý link download không cố định.
- Chưa tích hợp finance.vietstock.vn (trang tổng hợp tin theo mã CP) — có thể là nguồn tốt, chưa kiểm
  tra cấu trúc.

---

## 2026-07-02 (e) — Tự động dò sản lượng thép HPG (tháng/quý) từ hoaphat.com.vn — ước tính sớm LNST

### Bối cảnh
User muốn có sớm số liệu sản lượng thép HPG theo tháng/quý (trước khi có BCTC/tin công bố quý chính
thức) để ước tính doanh thu/LNST quý đang chạy sớm hơn — thuần Python, không dùng AI lúc chạy.

### Điều tra nguồn (đã thực hiện, không đoán)
- `hoaphat.com.vn` (IR site chính chủ) — trang danh sách "Tin tức"/"Công bố thông tin" render JS,
  không lấy được qua `curl`. NHƯNG mỗi trang BÀI VIẾT (server-rendered, lấy được qua `curl`) có khung
  "Tin liên quan" hiển thị 5 tin MỚI NHẤT toàn site (không phụ thuộc bài đang xem) — dùng làm cửa sổ
  dò tin mới. Ngoài ra site có `sitemap.xml` → `sitemap-tintuc-page-N.xml` (17 trang, tĩnh, liệt kê
  toàn bộ URL bài viết từ 2008 đến nay) — lọc theo từ khóa slug (`san-luong`, `trieu-tan`, `tan-thep`...)
  để tìm ứng viên bài công bố sản lượng.
- HPG có đăng định kỳ bài dạng "Sản lượng bán hàng thép Hòa Phát đạt X (triệu/nghìn) tấn trong {tháng
  N | quý N/YYYY}" — mẫu tiêu đề khá nhất quán qua nhiều năm (verify: quý 2/2025, tháng 10/2023, quý
  I/2026 — đều khớp cùng 1 regex).
- `vsa.com.vn` (Hiệp hội Thép) chỉ có số liệu TỔNG NGÀNH, không tách HPG — không dùng được cho mục
  đích này (vẫn dùng cho phần vĩ mô ngành thép sẵn có).

### Thay đổi
- Thêm `fetch_hpg_production_updates()`: dò tin qua (1) khung "Tin liên quan" từ 1 URL seed cố định,
  (2) quét 4 trang sitemap đầu (giới hạn số lần fetch mỗi lần chạy, KHÔNG quét toàn bộ 17 trang/lịch
  sử). Parse tiêu đề bằng regex `_HPG_TITLE_RE` (hỗ trợ quý số Ả Rập lẫn La Mã "quý I/II/III/IV", có
  hoặc không kèm năm — suy năm từ ngày đăng bài nếu thiếu). Trả về `[]` nếu lỗi bất kỳ bước nào —
  KHÔNG BAO GIỜ crash script.
- **Đã thử và BỎ việc tự parse riêng số liệu HRC từ nội dung bài** — phát hiện không tin cậy: câu văn
  thường viết "...thép cuộn cán nóng (HRC), thép xây dựng, phôi thép đạt X tấn" (X là TỔNG cả 3 sản
  phẩm) nên regex tìm "HRC...đạt X tấn" dễ bắt NHẦM tổng thành riêng HRC (đã verify bug này bằng test
  thực tế trước khi bỏ). Chỉ giữ lại TỔNG sản lượng từ TIÊU ĐỀ bài (đáng tin cậy, verify khớp nhiều bài
  mẫu) — tách HRC/XD dùng tỷ lệ lịch sử Q1/2026 thực tế, không suy từ văn bản.
- Xác định "quý đang chạy" = quý kế tiếp sau quý cuối cùng đã có số liệu (`SALES_Q_LABELS[-1]`), tự
  động tính từ dữ liệu hiện có (không hardcode). Có bài công bố QUÝ chính thức → dùng trực tiếp
  (`CUR_Q_SOURCE="OFFICIAL"`); chỉ có bài THÁNG → ước tính = TB(các tháng đã biết) × 3
  (`CUR_Q_SOURCE="ESTIMATED"`); chưa có gì → giữ giả định cũ (`CUR_Q_SOURCE="FALLBACK"`), không chặn.
- Sheet `15_Quarterly_Data` thêm 3 mục mới:
  - **D** (mở rộng bảng cũ): thêm 1 cột cho quý đang chạy nếu có dữ liệu, đánh dấu `(*)` + ghi rõ
    nguồn (CHÍNH THỨC hay ƯỚC TÍNH).
  - **E** (mới): bảng 3 tháng của quý đang chạy — tháng nào đã có tin thì điền số + ngày đăng + LINK
    bài viết (hyperlink); dòng "Ước tính quý" = `AVERAGE(...)×3` — Excel formula SỐNG, không phải số
    Python dán sẵn, tự cập nhật khi user điền thêm tháng mới thủ công nếu muốn.
  - **F** (mới): dự phóng sản lượng NĂM 2026E = run-rate (SL Q1/2026 + SL quý đang chạy)/2×4 — formula
    sống tham chiếu trực tiếp tới mục D.
- `03_Revenue_Model!G4/G5` (Sản lượng HRC/XD 2026E) đổi từ số Python dán sẵn (6.0/3.0 triệu tấn) sang
  LINK trực tiếp `='15_Quarterly_Data'!B86/B87` — vì toàn bộ chuỗi Doanh thu→GVHB→GP→EBIT→EBT→LNST
  cho các cột dự phóng ĐÃ SẴN là formula sống (không cần sửa gì thêm ở `04_PnL`), nên chỉ cần sửa
  ĐÚNG Ô GỐC sản lượng là LNST 2026E tự động cập nhật theo sản lượng mới — đúng yêu cầu "link công
  thức tính toán để dự báo LNST năm và quý này".
- Gộp `hrc_data`/`xd_data` (sheet 15) và `hrc_sales`/`xd_sales` (JSON export) — 2 mảng trùng lặp dễ
  lệch số như trước — thành `SALES_Q_LABELS`/`HRC_SALES_HIST_KT`/`XD_SALES_HIST_KT` (module-level,
  nguồn duy nhất).

### Kết quả thực tế lần chạy đầu (2026-07-02)
- Scraper chạy đúng, tìm thấy 3 bài công bố thật (tháng 10/2023, quý 2/2025, quý I/2026 — số quý
  I/2026 = 3 triệu tấn khớp hợp lý với số Q1/2026 đã có sẵn trong `HRC_SALES_HIST_KT`/`XD_SALES_HIST_KT`
  = 2,83 triệu tấn, chỉ chênh do làm tròn/nguồn khác nhau — cross-check tốt).
- KHÔNG tìm thấy dữ liệu quý 2/2026 (chưa có tin tháng 4/5/6 hay quý 2/2026) — `CUR_Q_SOURCE="FALLBACK"`,
  sheet 15 mục E hiển thị "Chưa có tin" cho cả 3 tháng, mục F giữ nguyên giả định 6.0/3.0 triệu tấn cũ.
  Hợp lý: Q2/2026 vừa kết thúc 30/06, bài công bố Q2/2025 năm ngoái ra ngày 09/07 (~9 ngày sau khi hết
  quý) — nhiều khả năng Q2/2026 cũng sẽ ra khoảng đầu-giữa tháng 7/2026, SAU thời điểm chạy script này.

### Chưa làm — cần user xác nhận nếu muốn
- Chưa test được kịch bản THẬT có dữ liệu (do chưa tới ngày HPG công bố Q2/2026) — cần chạy lại script
  sau khi HPG đăng tin tháng 6 hoặc quý 2/2026 để verify end-to-end mục E/F hoạt động đúng như thiết kế.
- Nếu HPG đổi mẫu tiêu đề bài viết trong tương lai, `_HPG_TITLE_RE` cần cập nhật lại — không có cảnh
  báo tự động khi mẫu không khớp nữa (chỉ âm thầm trả về `CUR_Q_SOURCE="FALLBACK"` như khi thật sự
  chưa có tin).

---

## 2026-07-02 (d) — Quặng sắt: MEDIAN THẬT tự động từ World Bank Pink Sheet (xác thực theo yêu cầu user)

### Bối cảnh
User hỏi xác nhận: "giá than quặng hrc ở quý là tính như nào, có lấy dữ liệu các kỳ trong quý để tính
median hay đang fix chết?" — trả lời trung thực: TRƯỚC bản này, cả 18 quý lịch sử là số Python gõ tay 1
điểm/quý cho CẢ 3 mặt hàng (không phải median nhiều điểm), vì không có nguồn giá ngày/tháng miễn phí.

### Thay đổi
- Thêm `fetch_worldbank_iron_ore_monthly()`: fetch tự động (curl) trang tĩnh World Bank để discover URL
  file `CMO-Historical-Data-Monthly.xlsx` (URL có hash đổi mỗi lần cập nhật), tải file, parse sheet
  `Monthly Prices` cột `Iron ore, cfr spot` bằng openpyxl. Trả về `{}` nếu bất kỳ bước nào lỗi — KHÔNG
  BAO GIỜ crash script.
- Với mỗi quý trong `Q18_LABELS`, nếu có đủ 3 tháng dữ liệu World Bank → ghi đè `Q18_IRON[i]` bằng
  `statistics.median()` thật của 3 tháng đó, đánh dấu `Q18_IRON_SRC[i] = "WB"`; nếu thiếu dữ liệu →
  giữ nguyên số nghiên cứu thủ công cũ, đánh dấu `"NC"`. Kết quả thực tế (2026-07-02): tất cả 18/18 quý
  fetch thành công, giá trị mới rất sát số nghiên cứu cũ (VD 2026Q1: 104.5 vs 104 cũ) — xác nhận số
  nghiên cứu thủ công trước đây khá chính xác, nhưng giờ có thể kiểm chứng bằng công thức sống thật.
- Sheet `17_Gia_Hang_Hoa` thêm cột F "Nguồn giá Quặng" hiển thị `WB — median 3 tháng` hay
  `NC — nghiên cứu thủ công` cho từng quý, để user tự kiểm chứng trực quan. Sửa lại ghi chú đầu sheet và
  mục NGUỒN DỮ LIỆU cho ĐÚNG bản chất theo từng mặt hàng (quặng = median thật; HRC/than = vẫn là số
  nghiên cứu, KHÔNG gọi nhầm là median).
- **Bug phát sinh & đã sửa ngay:** các dòng `print()` mới chứa dấu tiếng Việt (ă, ơ, ộ...) làm script
  CRASH với `UnicodeEncodeError` khi console Windows dùng codepage cp1252 (không phải lỗi riêng của môi
  trường dev — xảy ra cả khi chạy `python build_hpg_model.py` trực tiếp). Toàn bộ `print()` hiện có trong
  file (trước bản vá này) đều cố tình dùng tiếng Anh/ASCII thuần — đây là quy ước ngầm của codebase để
  tránh đúng lỗi này trên Windows. Đã sửa lại toàn bộ print() mới sang ASCII để nhất quán.

### Chưa làm
- Than cốc & HRC vẫn CHƯA có nguồn giá tháng/ngày miễn phí — nếu tìm được nguồn (kể cả trả phí) cho 2 mặt
  hàng này, áp dụng lại đúng pattern `fetch_worldbank_iron_ore_monthly()` đã làm cho quặng sắt.
- Log file tích lũy nhiều lần fetch investing.com trong 1 quý (mục "Chưa làm" từ bản 2026-07-02 (a)) vẫn
  chưa làm — vẫn cần cho "giá quý ĐANG CHẠY" (HRC, than cốc) có median thật thay vì median(2 điểm).

---

## 2026-07-02 (c) — Bỏ dự phóng DIO/DSO, thêm biểu đồ quý + tương quan Spread-BLNG, giá quý = MEDIAN

### Bối cảnh
User phản hồi sau bản (b): (1) không cần dự phóng DIO/DSO vì ít liên quan tới model định giá, (2) cần vẽ
DIO/DSO CẢ theo quý (không chỉ theo năm), (3) cần thêm biểu đồ tương quan (scatter) giữa Biên LNG và
Spread — cả năm và quý, (4) làm rõ giá HRC/quặng/than dùng trong Spread phải là MEDIAN các điểm giá quan
sát trong quý, rồi mới tính Spread theo công thức lag.

### Thay đổi
- `DIO_A`/`DSO_A` thu hẹp về 5 phần tử (2021-2025) — bỏ hẳn phần dự phóng. Cột 2026E-2028E trong sheet
  `14_Steel_Analysis` (Hàng tồn kho/Vòng quay HTK/DIO/Phải thu/DSO) để trống `—`; D/E và tỷ lệ vay NH/tổng
  vay vẫn dự phóng đủ (liên quan trực tiếp đòn bẩy).
- Thêm `DIO_Q`/`DSO_Q` (theo quý, năm hóa GVHB/DT quý ×4, khớp `Q18_LABELS`) + chart `turnover_quarterly.png`.
- Thêm 2 chart tương quan scatter+hồi quy: `spread_gpm_corr_annual.png` (5 điểm năm) và
  `spread_gpm_corr_quarterly.png` (mọi quý có đủ BCTC) — chỉ dùng dữ liệu THỰC TẾ, có annotate hệ số Pearson
  r ngay trên chart, kèm narrative PDF giải thích n nhỏ dễ bị outlier chi phối.
- Đổi mọi công thức blend "giá quý đang chạy" từ `AVERAGE()`/`(a+b)/2` sang `MEDIAN()`/`statistics.median()`
  (Python: `_hrc_avg_now`, `HRC_PRICE_A[5]`/`IRON_ORE_A[5]`/`COKE_A[5]`, `SPREAD_A[5]`; Excel: `R17_AVG` row,
  `c2026` price cells, `c2026s` spread cell) — với 2 điểm dữ liệu, median = trung bình cộng về mặt số học,
  nhưng đổi tên hàm cho đúng bản chất phương pháp (và sẵn sàng khi có thêm điểm dữ liệu tích lũy theo quý).

### Chưa làm
- Log file tích lũy nhiều lần fetch trong 1 quý (để MEDIAN có >2 điểm dữ liệu thực) — vẫn là việc "chưa làm"
  ghi trong bản 2026-07-02 (a).

---

## 2026-07-02 (b) — Spread lag-1-quý + DIO/DSO đầy đủ (vị trí, công thức, biểu đồ, đánh giá)

### Bối cảnh
User yêu cầu: (1) bổ sung vị trí, công thức tính, biểu đồ và phần đánh giá cho DIO (số ngày tồn kho BQ) và
DSO (số ngày phải thu BQ) để xem hiệu quả thu tiền/luân chuyển tồn kho có cải thiện không, (2) sửa TOÀN BỘ
công thức Spread sang LAG 1 QUÝ (giá HRC quý này − 1.6×quặng quý trước − 0.6×than quý trước − 100 USD)
vì DIO của HPG dao động quanh ~90 ngày (~1 quý), (3) rebuild + đánh giá biểu đồ Spread kết hợp Biên LNG
theo 12 quý gần nhất và theo năm, (4) kiểm tra giá HRC/quặng/than có đúng là median hay còn "giá chết".

### Các thay đổi đã áp dụng

| # | Nội dung sửa | Lý do |
|---|---|---|
| 1 | Spread mọi nơi (sheet `17_Gia_Hang_Hoa` cột E, sheet `14_Steel_Analysis`, Profit Bridge, PDF, JSON dashboard) đổi từ "cùng kỳ" sang LAG 1 QUÝ: `HRC(T) - 1.6×Quặng(T-1) - 0.6×Than(T-1) - 100` | Giá vốn kỳ này phản ánh giá NVL mua vào kỳ trước (DIO ~90 ngày ~ 1 quý), không phải giá NVL cùng kỳ với giá bán |
| 2 | Spread NĂM (2021-2025) = MEDIAN Spread quý (lag-1-quý) trong năm, thay vì trừ trực tiếp giá HRC/Quặng/Than cùng năm | Nhất quán với "Giá năm = MEDIAN 4 quý" đã dùng cho giá hàng hóa |
| 3 | Spread năm dự phóng xa (2027E/2028E, chưa có dữ liệu quý) = lag-1-NĂM (dùng giá quặng/than năm TRƯỚC) | Xấp xỉ tốt nhất khi chỉ có granularity theo năm |
| 4 | Phát hiện & sửa bug: Chart "Spread & Biên LNG theo Quý" và JSON dashboard (`sl_spread`) trước đây lấy Spread NĂM lặp lại cho MỌI quý trong năm đó (không phải Spread thực của từng quý) — giờ dùng đúng `Q18_SPREAD` (mảng 18 quý thực, lag-1-quý) | Biểu đồ/dashboard trước đó trông như bậc thang thay vì diễn biến quý thực, sai lệch với ý đồ "Spread kết hợp BLNG theo 12 quý gần nhất" |
| 5 | Thêm DIO/DSO: vị trí `14_Steel_Analysis` mục 4 (Excel formula sống link `04_PnL`+`05_Balance_Sheet`, số dư BÌNH QUÂN đầu+cuối kỳ, cả lịch sử lẫn dự phóng), biểu đồ `turnover.png` (dùng `DIO_A`/`DSO_A` tính động thay vì mảng số cứng cũ), đánh giá xu hướng CẢI THIỆN/XẤU ĐI tại PDF mục 4C | User: "xem hiệu quả thu tiền và luân chuyển lưu kho có cải thiện không" — cần công thức kiểm chứng được + đánh giá rõ ràng, không phải số Python gõ tay |
| 6 | Sửa 2 bug tiềm ẩn phát hiện trong lúc sửa Section 4 (`14_Steel_Analysis`): công thức D/E dự phóng tham chiếu nhầm biến Python `col` còn sót từ vòng lặp khác (ra công thức sai hoàn toàn cho 2026E-2028E); fallback tồn kho bình quân năm 2021 dùng nhầm `cash_hist` thay vì tồn kho | Bug âm thầm không crash nhưng cho số sai — phát hiện khi audit lại toàn bộ Section 4 để thêm DSO |
| 7 | Sửa bug đơn vị: câu văn "Hàng tồn kho ... ngày doanh thu" ở mục 4B lấy tồn kho chia DOANH THU 1 QUÝ rồi nhân 365 (thổi phồng ~4 lần) | Không quy đổi theo năm trước khi nhân 365 |
| 8 | `q1_spread` (mẫu số nội suy Biên LNG dự phóng) đổi từ xấp xỉ bằng Spread NĂM 2026 sang lấy đúng `Q18_SPREAD[-1]` (Spread Q1/2026 thực, lag-1-quý) | Trước đây tỷ lệ nội suy luôn = 1 (vì tử số = mẫu số), không phản ánh đúng phương pháp "tỷ lệ Spread dự phóng/Spread quý gần nhất" đã mô tả trong skill |
| 9 | Xác nhận giá HRC/quặng/than NĂM đã đúng median (không phải giá chết) — nguồn gốc thực của khiếu nại "giá chết" là do biểu đồ quý/JSON dùng nhầm giá năm lặp lại (mục 4), không phải giá năm bản thân nó | Tránh sửa nhầm chỗ không có bug |

### Chưa làm — cần user xác nhận nếu muốn
- DIO/DSO dự phóng (2026E-2028E) còn phụ thuộc tồn kho/phải thu dự phóng TĨNH (`INVENTORY_FC`/
  `RECEIVABLES_FC`), chưa gắn driver doanh thu/sản lượng — có thể khiến DIO dự phóng giảm nhanh hơn thực tế.
- Bảng "Leading Indicators" (Chỉ báo dẫn dắt) có dòng "Giá HRC (USD/tấn)" ngưỡng 900-1100 hoàn toàn lệch
  pha với giá thực tế model đang dùng (~460-650 USD/tấn) — có vẻ là dữ liệu cũ/đơn vị khác, chưa sửa vì
  ngoài phạm vi yêu cầu lần này (đã tạo task riêng để theo dõi).

---

## 2026-07-02 — Tự động hoá giá hàng hóa (HRC/quặng/than) không cần AI

### Bối cảnh
User yêu cầu: (1) sửa GP margin không được lấy trực tiếp từ Spread (chỉ dùng tỉ lệ để nội suy từ BLNG
thực tế quý gần nhất), (2) công thức Spread cộng thêm chi phí SX khác CỐ ĐỊNH thay vì mảng biến đổi/việc
double-count lương nhân công, (3) **quan trọng nhất**: khi chạy `build_hpg_model.py` local không có AI hỗ
trợ, script vẫn phải tự cập nhật được giá HRC/quặng/than chính xác — chỉ dùng Python thuần.

### Các thay đổi đã áp dụng vào skill

| # | Nội dung sửa | Lý do |
|---|---|---|
| 1 | Spread = Giá - 1.6×Quặng - 0.6×Than - Chi phí SX khác CỐ ĐỊNH (hằng số `OTHER_COST_USD`, không phải mảng biến đổi theo năm) | Tránh double-count lương nhân công, đơn giản hoá theo yêu cầu user; hằng số chỉnh trực tiếp theo report (200→100 USD/tấn khi số liệu ra âm bất hợp lý) |
| 2 | BLNG dự phóng: neo vào BLNG THỰC TẾ quý gần nhất, nội suy theo TỈ LỆ spread dự phóng/spread quý gần nhất — không lấy Spread trực tiếp làm BLNG | Spread chỉ phản ánh xu hướng giá benchmark quốc tế, không phải biên lợi nhuận thực tế của DN (có hợp đồng dài hạn, tự chủ nguyên liệu...) |
| 3 | Sheet Excel riêng `17_Gia_Hang_Hoa` — nguồn giá hàng hóa DUY NHẤT, các sheet khác link công thức về đây | Trước đó mỗi sheet có mảng Python độc lập, dễ lệch số khi sửa 1 chỗ quên sửa chỗ khác |
| 4 | Fetch giá hiện tại qua `subprocess` + `curl` (KHÔNG dùng `requests`) | investing.com chặn vân tay TLS của `requests`/`urllib3` (Cloudflare 403) nhưng không chặn `curl` — đã verify nhiều lần loại trừ rate-limit |
| 5 | Cột "Spread quý" (Excel formula sống) cho từng quý trong sheet 17 | User: "tôi thấy số chết nên tôi không kiểm chứng được" — mọi Spread hiển thị phải là công thức bấm vào xem được, không phải số Python tính sẵn |
| 6 | Ghi rõ nguồn URL từng mặt hàng NGAY TRONG sheet Excel (không chỉ trong code/skill) | User cần tự đối chiếu lại số liệu sau này |
| 7 | 18 quý lịch sử là số liệu nghiên cứu thủ công (không phải median 3 mốc/quý thật) — đã caveat rõ độ tin cậy theo từng mặt hàng | Không có API giá lịch sử theo ngày miễn phí; tránh gây hiểu nhầm là số liệu 100% chính xác |

### Chưa làm — cần user xác nhận nếu muốn
- Log file tích lũy giá fetch mỗi lần chạy script trong quý → "giá quý hiện tại" dần trở thành median THẬT
  (thay vì 1 điểm dữ liệu) khi user chạy script nhiều lần trong quý. Chỉ có giá trị từ quý hiện tại trở đi.
- Tự động hoá tương tự cho **sản lượng thép** (user đã đề cập muốn làm nhưng chưa gửi URL nguồn cụ thể).
- `app_steel.js`/`steel.html` (sau khi tách khỏi `app.js` dùng chung) bị thiếu 1 số section thép-riêng
  (factory table, harvest signals, peer comparison, leading indicators) từng có ở `app.js` cũ — `app.js` giờ
  không còn được load ở đâu (orphaned). Cần quyết định có khôi phục lại vào `app_steel.js` không.

---

## 2026-06-25 — Học từ báo cáo mẫu BSI (Mau phan tich HPG1.md)

### Phân tích framework từ reference

**Nguồn:** `Hoc mau phan tich/Mau phan tich HPG1.md` (BETA Securities, Lê Nguyên Anh Phương, 10/04/2026)

**Cấu trúc báo cáo phát hiện:**
1. Thông tin CP & Cơ cấu cổ đông → Khuyến nghị & Định giá → Tóm tắt TC → Phân tích ngành toàn cầu → Triển vọng DN → Rủi ro

**Industry analysis phát hiện thêm:**
- BOF Inflexibility: chu kỳ 15-25 năm, không thể dừng lò — giải thích dư cung cấu trúc
- CBAM chênh lệch chi phí carbon BOF vs EAF: 150-200 EUR/tấn
- 8 lĩnh vực phân bổ nhu cầu thép toàn cầu (xây dựng 49-50%, cơ khí 16%, v.v.)
- Dư cung 640M tấn (26%), dự báo 721M tấn (2027F)

**Business analysis phát hiện thêm:**
- Tỷ lệ nhập khẩu nguyên liệu: 90% than cốc, 52% quặng, 47% phế
- Bảng Cung-Cầu HRC nội địa: Formosa (FMS) vs HPG vs total demand
- Khái niệm "Bảo hộ toàn phần" — AD narrow + wide HRC
- Tái cơ cấu thị trường: nội địa 84% (↑38%), XK 16% (↓ từ 31%)
- Đầu tư công 2026-2030: 8.5 triệu tỷ (3x giai đoạn trước)

**Định giá:**
- Phương pháp kết hợp: DCF 70% + P/E 30% (khác với tỷ trọng EV/EBITDA+P/B trong model hiện tại)

### Các thay đổi đã áp dụng vào skill

| # | Nội dung thêm | Lý do |
|---|---|---|
| 1 | KLGD bình quân 20 phiên vào cấu trúc báo cáo | Reference có chỉ tiêu này |
| 2 | Kế hoạch vs Thực hiện định dạng chuẩn | Reference format giúp rõ ràng |
| 3 | Cơ cấu sản lượng toàn cầu (TQ 52%, Ấn 8.9%, EU 6.8%...) | Reference cung cấp số liệu cụ thể |
| 4 | Nên trình bày dữ liệu thô biểu đồ dạng bảng | Tăng tính kiểm chứng |
| 5 | Checklist so sánh KH ĐHĐCĐ vs thực tế | Reference làm rõ mục này |
| 6 | Checklist các chỉ số cần có (KLGD, biến động, v.v.) | Đảm bảo báo cáo đầy đủ |

### So sánh với model hiện tại

**Điểm mạnh của reference so với model:**
- Phân tích dư cung toàn cầu chi tiết (số liệu biểu đồ thô + phân bổ quốc gia)
- BOF inflexibility như một luận điểm cấu trúc (model chưa có)
- CBAM chi phí carbon định lượng (model chưa có)
- Bảng supply-demand HRC có đối thủ FMS (model chỉ có HPG share)
- Export restructuring chi tiết (model mới có nội địa là chính)

**Điểm mạnh của model so với reference:**
- Có quarterly data + trend analysis (reference không có)
- Có steel accounting + harvest signs (reference không có)
- Có hơn 16 sheets dữ liệu tài chính full (reference chỉ 1 bảng)
- Có A/E labeling cho từng năm
- Có P/B vs P/E trap analysis (reference vẫn dùng P/E 30%)

### Ghi chú
- File `Promt.docx` không đọc được (định dạng docx). Có thể chứa prompt gốc dùng để tạo báo cáo này.
- Lần sau nếu có file tham khảo mới, cần systematically:
  1. Đọc toàn bộ
  2. Extract: Industry framework, Business framework, Forecast method, Valuation method, Key metrics
  3. So sánh với skill hiện tại (từng mục)
  4. Update skill
  5. Ghi changelog
