---
name: thep
description: Use ONLY when analyzing Vietnamese steel stocks. Front-loaded keywords: thép, steel, HPG, HSG, NKG, TVN, tôn mạ, HRC, thép xây dựng, quặng sắt, than cốc, spread, lò cao, BOF, rebar, commodity cyclical, overcapacity, CBAM, bảo hộ thép.
---

# KĨ NĂNG PHÂN TÍCH NGÀNH THÉP

## Cấu trúc báo cáo chuẩn (BSI format)
Tham khảo báo cáo mẫu: `Hoc mau phan tich/Mau phan tich HPG1.md`

1. **Thông tin CP & Cơ cấu cổ đông**: Giá, vốn hóa, cp lưu hành, 52w high/low, KLGD bình quân 20 phiên (cp + tỷ đồng), EPS, BVPS, P/E, P/B, biến động 1m/3m/YTD vs VNINDEX, tỷ lệ sở hữu founder & gia đình
2. **Khuyến nghị & Định giá**: Giá mục tiêu, upside, phương pháp + tỷ trọng (VD: DCF 70% + PE 30%), kế hoạch vs thực hiện (định dạng: Kế hoạch: xxx | Thực hiện: xxx | Hoàn thành: xx%)
3. **Tóm tắt chỉ tiêu TC**: 3 năm quá khứ + 2 năm dự phóng (DT, GP, LNST, ROE, ROA, D/E, EPS)
4. **Phân tích ngành thép toàn cầu** → **Phân tích DN** → **Định giá** → **Rủi ro**
5. Kết thúc bằng khuyến nghị rõ ràng, ngắn gọn
6. **Luôn so sánh KQKD thực tế vs kế hoạch ĐHĐCĐ** (doanh thu, LNST) kèm % hoàn thành

---

## Top-Down 5 bước

1. **Vĩ mô & Chu kỳ**: Lãi suất, đầu tư công (kế hoạch trung hạn 2026-2030), BĐS (đầu ra chính), gói kích thích Trung Quốc
2. **Chuỗi giá trị**: Quặng sắt/Than cốc → Lò cao (BOF) → Thép lỏng → HRC / Thép xây dựng
3. **Lợi thế cạnh tranh**: Quy mô, tự chủ nguyên liệu, vị trí nhà máy, công nghệ, bảo hộ thương mại
4. **BCTC**: Hàng tồn kho (trái tim), TSCĐ & CIP (quy mô), Nợ vay, Phải thu, chất lượng LN (CFO/LNST)
5. **Định giá**: DCF + P/B là chính, EV/EBITDA phụ — **KHÔNG dùng P/E** cho cổ phiếu chu kỳ

---

## Công thức cốt lõi: Spread

Ngành thép sống dựa vào **chênh lệch giá (Spread)** giữa đầu vào và đầu ra.

**Định mức tiêu chuẩn (Lò cao BOF)**:
- 1 tấn thép cần ~1.6 tấn quặng sắt + ~0.6 tấn than cốc

**⚠️ Spread LAG 1 QUÝ (2026-07 — chốt theo yêu cầu user, thay thế công thức "cùng kỳ" cũ):**
- Vì số ngày tồn kho bình quân (DIO) của HPG dao động quanh ~90 ngày (~1 quý), giá vốn hàng bán ghi
  nhận trong quý T phần lớn phản ánh giá nguyên liệu MUA VÀO ở quý T-1, trong khi giá bán vẫn là giá
  của quý T. `total_cost(quý T) = 1.6×Giá quặng bình quân quý T-1 + 0.6×Giá than cốc bình quân quý
  T-1 + Chi phí SX khác CỐ ĐỊNH (USD/tấn, `OTHER_COST_USD`)`.

**⚠️ 3 LOẠI SPREAD RIÊNG BIỆT (2026-07 — chốt theo yêu cầu user, thay thế 1 spread HRC duy nhất cũ):**
- **Spread HRC** = Giá HRC bình quân quý T `× 1.15` (nếu quý T ≥ `AD_HRC_EFFECTIVE_Q` = 2025Q3) −
  `total_cost(quý T)`. Hệ số 1.15 phản ánh **thuế chống bán phá giá (CBPG) HRC Trung Quốc — vụ AD20**:
  Bộ Công Thương áp thuế chính thức khổ hẹp hiệu lực **06/07/2025** (rơi vào 2025Q3), mở rộng chống
  lẩn tránh khổ rộng từ 17/04/2026. Không áp dụng thuế này thì Spread HRC tính từ giá CFR quốc tế
  (investing.com) sẽ bị đánh giá THẤP hơn thực tế vì bỏ qua chi phí thuế NK thực trả.
- **Spread Rebar** (thép xây dựng) = Giá thép XD bình quân quý T − `total_cost(quý T)` — KHÔNG nhân
  hệ số CBPG (thuế chỉ áp cho HRC nhập khẩu, không áp cho rebar).
- **Spread All** = bình quân gia quyền theo sản lượng THỰC TẾ CỦA CHÍNH QUÝ ĐÓ:
  `(SL_XD(T)×Spread_Rebar(T) + SL_HRC(T)×Spread_HRC(T)) / (SL_XD(T)+SL_HRC(T))`. Chỉ tính được cho
  các quý có dữ liệu sản lượng tách riêng (2023Q1 trở đi — xem `HRC_SALES_HIST_KT`/`XD_SALES_HIST_KT`);
  5 quý đầu (2021Q4-2022Q4) để `None`, không suy diễn thiếu căn cứ.
- **Spread năm** (2021-2025, đủ dữ liệu quý) = MEDIAN các Spread quý (lag-1-quý) thuộc năm đó cho MỖI
  loại — nhất quán với cách tính "Giá năm = MEDIAN 4 quý".
- **Spread năm dự phóng xa** (2027E, 2028E — chưa có dữ liệu quý) = xấp xỉ lag-1-quý bằng lag-1-NĂM
  dùng giá quặng/than năm TRƯỚC (Spread HRC vẫn nhân 1.15 vì đã chắc chắn sau 2025Q3).
- **Spread năm hiện tại** (2026E) = TB(Spread quý gần nhất đã biết đầy đủ, Spread hiện tại quý đang chạy).
- Biến Python: `Q18_SPREAD_HRC` (= alias `Q18_SPREAD`, tương thích ngược), `Q18_SPREAD_REBAR`,
  `Q18_SPREAD_ALL` (mảng 18 quý, lag-1-quý); `SPREAD_A`/`SPREAD_REBAR_A`/`SPREAD_ALL_A` (8 năm);
  `SPREAD_HRC_NOW`/`SPREAD_REBAR_NOW`/`SPREAD_ALL_NOW` (quý đang chạy — `SPREAD_ALL_NOW` dùng SL quý
  đang chạy làm quyền số, ưu tiên `CUR_Q_HRC_KT_DIRECT`/`CUR_Q_XD_KT_DIRECT`, fallback SL quý gần
  nhất). Quý đầu tiên trong bảng lịch sử (2021Q4) không có quý trước (2021Q3) nên dùng tạm giá CÙNG
  quý — chỉ 1/18 điểm dữ liệu bị ảnh hưởng, không đại diện cho công thức chuẩn.
- **Excel**: sheet `17_Gia_Hang_Hoa` cột G-L (giá/nguồn XD, Spread Rebar, SL XD/HRC link sheet
  `15_Quarterly_Data`, Spread All) — TẤT CẢ là công thức sống, không số chết. 3 biểu đồ riêng biệt
  (Spread HRC/Rebar/All vs BLNG quý) hiển thị ở Excel/PDF/JSON — xem `_make_spread_gpm_chart()`.

**Lợi nhuận** = Spread × Sản lượng - Chi phí cố định

**⚠️ Chi phí SX khác (2026-07 — đã chốt qua nhiều lần chỉnh với user):**
- Dùng **một hằng số USD/tấn cố định** cho toàn bộ các năm (gồm nhân công, nhiên liệu, điện, khấu hao,
  CPQL), KHÔNG dùng mảng biến đổi theo năm kiểu `CONV_A = [.., .., ..]` (từng gây double-count rủi ro khi
  cộng thêm cả "lương nhân công/tấn" như một khoản riêng — user đã yêu cầu bỏ, quy về 1 số cố định duy nhất).
- Giá trị hằng số này **do user chỉnh trực tiếp theo từng report** dựa trên đối chiếu BCTC thực tế (đã đổi
  từ 200 → 100 USD/tấn khi số liệu spread tính ra âm bất hợp lý so với biên LN gộp thực tế công bố) — biến
  Python là `OTHER_COST_USD`, KHÔNG hardcode giá trị này rải rác nhiều chỗ, chỉ định nghĩa 1 lần rồi mọi
  công thức/text đều tham chiếu qua biến (kể cả câu chữ trong PDF, không viết số cứng "100 USD/tấn" trực
  tiếp trong docstring/narrative).
- **Luôn hiện công thức Spread dưới dạng Excel formula SỐNG** (không phải số Python tính sẵn) ở MỌI nơi
  hiển thị spread theo quý/năm, để user tự bấm vào ô kiểm chứng — xem mục "Dữ liệu giá hàng hóa" bên dưới.

---

## Phân Tích Ngành Thép Toàn Cầu

### Dư cung toàn cầu (Overcapacity)
- **Định lượng dư cung hiện tại:** ~640 triệu tấn (26% công suất chưa dùng)
- **Phân bổ:** Trung Quốc (~400M tấn), Ấn Độ (~50M), EU (~40M), CIS (~30M)
- **Xu hướng:** Các kế hoạch mở rộng công suất (+6.7% 2025-2027) → dư cung có thể đạt **721 triệu tấn vào 2027F**
- **BOF Inflexibility (rất quan trọng):** Lò cao (BOF) có chu kỳ vận hành liên tục **15-25 năm**, không thể dừng lò ngay lập tức khi giá giảm sâu. Đây là nguyên nhân cấu trúc khiến dư cung kéo dài.
- **Sản lượng toàn cầu 2025:** ~1,849 triệu tấn. Cơ cấu theo quốc gia: TQ 52%, Ấn Độ 8.9%, EU 6.8%, Nhật Bản 4.4%, Hoa Kỳ 4.4%, Nga 3.7%, còn lại 19.8%.
- **Nên trình bày dạng bảng dữ liệu thô cho biểu đồ** (VD: Công suất tối đa vs Nhu cầu tiêu thụ theo năm 2019-2027F) — giúp người đọc tự kiểm chứng số liệu.

### Dịch chuyển nhu cầu theo khu vực & lĩnh vực
- **Khu vực:** TQ suy yếu (BĐS đóng băng), Ấn Độ & ASEAN là động lực mới (~4% tăng trưởng)
- **Phân bổ nhu cầu toàn cầu (8 lĩnh vực):**
  - Xây dựng kết cấu, tòa nhà: **49-50%**
  - Kỹ thuật cơ khí: **16%**
  - Sản phẩm kim loại: **11%**
  - Ô tô: **8%**
  - Dầu khí: **6%**
  - Đóng tàu & đường sắt: **4%**
  - Thiết bị gia dụng: **3%**
  - Quốc phòng: **1%**
  - Khác: **2%**

### Bảo hộ thương mại & CBAM (Thuế carbon biên giới)
- Lượng thép giá rẻ dư thừa từ TQ → các nước áp thuế chống bán phá giá (AD) hàng loạt
- **CBAM của EU:** Chênh lệch chi phí carbon giữa BOF (thép truyền thống) và EAF (thép xanh) có thể lên tới **150-200 EUR/tấn** (2026). Buộc DN xuất khẩu phải chuyển đổi xanh.
- **Khái niệm "Bảo hộ toàn phần":** Khi cả HRC khổ hẹp (AD narrow) và HRC khổ rộng (AD anti-circumvention) đều bị áp thuế → thị trường đóng cửa với hàng TQ.

---

## Dữ liệu giá hàng hóa — tự động hoá không cần AI (2026-07)

**Bối cảnh:** User chạy `build_hpg_model.py` local, không có AI hỗ trợ tại thời điểm chạy — mọi việc lấy
giá HRC/quặng sắt/than cốc phải tự động 100% bằng Python thuần, không được dựa vào research thủ công mỗi lần.

**Kiến trúc:**
1. **Sheet Excel riêng `17_Gia_Hang_Hoa`** — nguồn duy nhất cho giá hàng hóa, mọi sheet khác
   (`02_Assumptions`, `03_Revenue_Model`, `14_Steel_Analysis`) **LINK công thức** về sheet này, không dùng
   mảng Python độc lập cho từng sheet (tránh lệch số).
2. **Bảng 18 quý lịch sử** (2021Q4 → hiện tại), độ tin cậy khác nhau theo mặt hàng:
   - **Quặng sắt (cột C) — MEDIAN THẬT, tự động (2026-07):** Fetch tự động (curl) file
     `CMO-Historical-Data-Monthly.xlsx` từ World Bank Commodity Markets "Pink Sheet"
     (`https://www.worldbank.org/en/research/commodity-markets` — URL file đổi hash mỗi lần cập nhật nên
     phải discover qua trang tĩnh trước, KHÔNG hardcode URL file), sheet `Monthly Prices`, cột
     `Iron ore, cfr spot`. Với mỗi quý, lấy 3 giá THÁNG thật rồi tính `statistics.median()` → **median
     thật của nhiều điểm giá trong quý**, không phải số nghiên cứu tĩnh. Cột F "Nguồn giá Quặng" trong
     sheet 17 đánh dấu `WB` (median thật) hay `NC` (fallback nghiên cứu thủ công, khi World Bank chưa có
     đủ 3 tháng hoặc fetch lỗi — script KHÔNG BAO GIỜ dừng vì lỗi fetch này).
   - **Than cốc & HRC (cột B, D) — vẫn là số nghiên cứu thủ công, KHÔNG phải median:** Chưa tìm được
     nguồn giá THÁNG/NGÀY miễn phí tương đương World Bank cho 2 mặt hàng này (đã thử: World Bank Pink
     Sheet KHÔNG có than cốc luyện kim/HRC). Vẫn là 1 số đại diện/quý tổng hợp thủ công (than cốc: FPTS/
     VCBS/Argus — trung bình; HRC: Mysteel/SteelBenchmarker/giá XK HPG — chỉ đúng xu hướng/độ lớn) — ghi
     chú rõ trong sheet 17 (không gọi nhầm là "median" để tránh hiểu lầm).
3. **Giá năm** = Excel formula `MEDIAN()` của 4 quý trong năm đó (không phải trung bình cộng — median bền
   hơn với outlier).
4. **Cột "Spread quý" riêng** (cột E trong sheet 17) — công thức SỐNG cho từng quý, không phải số Python
   tính sẵn, để user bấm vào ô kiểm chứng trực tiếp.
5. **Giá hiện tại — fetch tự động khi chạy script**, dùng `subprocess` gọi `curl` (KHÔNG dùng thư viện
   `requests`):
   ```python
   def fetch_via_curl(url, timeout=10):
       r = subprocess.run(["curl", "-sL", "-A", UA, "--max-time", str(timeout), url],
                           capture_output=True, text=True, encoding="utf-8", errors="ignore")
       return r.stdout if r.returncode == 0 else ""
   ```
   **Lý do bắt buộc dùng `curl` thay vì `requests`:** investing.com chặn vân tay TLS (JA3/ClientHello) của
   thư viện `requests`/`urllib3` bằng Cloudflare (trả 403 "Attention Required"), nhưng KHÔNG chặn `curl` dù
   User-Agent giống hệt — đã verify nhiều lần, không phải do rate-limit. `curl` có sẵn mặc định trên
   Windows 10+/macOS/Linux nên không cần cài thêm gì.
   - Parse giá: regex `data-test="instrument-price-last"[^>]*>([\d,\.]+)<` trên HTML server-rendered.
   - Parse đơn vị tiền tệ: regex `currency-in-label"[^>]*>.*?<span[^>]*>([A-Z]{3})</span>` — nếu CNY thì
     quy đổi USD qua hằng số tỷ giá riêng (VD: than cốc niêm yết trên sàn Đại Liên bằng CNY).
   - **Luôn có try/except fallback về giá quý gần nhất** nếu fetch lỗi (mất mạng, site đổi cấu trúc HTML) —
     script KHÔNG BAO GIỜ được dừng vì lỗi fetch, giống pattern `fetch_rf_vietnam()` đã có sẵn trong
     `template_banking.py`.
6. **"Spread hiện tại"** = TB(giá đầu quý gần nhất đã biết, giá hiện tại fetch live) cho cả 3 mặt hàng, rồi
   áp công thức spread chuẩn — KHÔNG dùng giá dự phóng cả năm để tính "spread hiện tại" (lỗi từng gặp).
7. **Ghi rõ nguồn URL trong chính sheet Excel** (không chỉ trong skill/code comment) để user tự đối chiếu
   khi nghi ngờ số liệu — mục "NGUỒN DỮ LIỆU" ở cuối sheet 17, liệt kê từng mặt hàng + URL cụ thể.
8. **Giá thép xây dựng (Rebar/XD) — MEDIAN THẬT + neo nội địa (2026-07):** Trước đây KHÔNG có nguồn giá
   quý thật cho rebar (chỉ có 1 giả định năm `XD_PRICE_A`, chưa đủ granularity để tính Spread Rebar/All).
   Đã research và xác nhận VSA (Hiệp hội Thép, bản tin tháng) KHÔNG công bố giá bán rebar (chỉ có giá
   nguyên liệu đầu vào + sản lượng) → không dùng được. Nguồn thay thế đã verify:
   - **investing.com Steel Rebar futures (SRRc1 — Shanghai rebar continuous)**: niêm yết THẲNG USD/tấn
     (không cần quy đổi CNY). Dùng API nội bộ `api.investing.com/api/financialdata/historical/996702`
     (header `Domain-Id: vn`, `time-frame=Monthly`) để lấy TOÀN BỘ lịch sử tháng 1 lần fetch (khác cách
     scrape từng trang quý của HRC/than cốc) → tính MEDIAN THẬT 3 tháng/quý như quặng sắt. **Hạn chế:**
     hợp đồng NGỪNG cập nhật sau 10/2025 (khối lượng giao dịch về 0, giá đứng yên) — chỉ dùng được cho
     16/18 quý (2021Q4-2025Q3), đánh dấu nguồn `"INV"`.
   - **SteelOnline.vn** (`bang-gia-thep-xay-dung-hom-nay`): giá thép XD Hòa Phát D10/CB240 NỘI ĐỊA THỰC
     (VND/kg) nhưng CHỈ có giá HIỆN TẠI (không có kho lưu trữ lịch sử) — bảng ĐẦU TIÊN trên trang là bảng
     Hòa Phát (xác nhận qua đoạn mô tả ngay sau bảng). Dùng làm "giá hiện tại" (`xd_now`), ưu tiên hơn
     SRRc1 futures đã hết thanh khoản.
   - **2026Q1** (futures không có dữ liệu): dùng điểm neo THẬT nội địa từ VSA — bài "Thị trường thép xây
     dựng Quý I/2026" (báo cáo họp Tổ Điều hành Thị trường trong nước, KHÁC bản tin tháng thường kỳ):
     giá dao động 15.100-15.700 đồng/kg → median 15.400 quy đổi USD/tấn theo tỷ giá USD/VND fetch live
     (`fetch_usd_vnd_rate()`, fallback investing.com currencies/usd-vnd). Đánh dấu nguồn `"VN"`.
   - **2025Q4** (futures chỉ còn 1/3 tháng thật — T10, hết thanh khoản T11-T12): nội suy tuyến tính giữa
     2025Q3 (thật) và 2026Q1 (thật) — đánh dấu nguồn `"NC"`.
   - Đối chiếu 2 nguồn tại cùng thời điểm (~10/2025): SRRc1 ~545 USD/tấn vs SteelOnline hiện tại quy đổi
     ~581 USD/tấn — lệch ~6%, chấp nhận được để ĐO XU HƯỚNG tương đối (không phải mốc tuyệt đối tuyệt đối
     chính xác — 2 nguồn khác thị trường: futures quốc tế vs bán lẻ nội địa Việt Nam).
   - Biến Python: `Q18_XD` (18 quý), `Q18_XD_SRC` (nhãn nguồn từng quý), `xd_now`, `XD_NOW_SRC`.

**URLs đã verify hoạt động qua `curl` (2026-07):**
- HRC: `https://www.investing.com/commodities/lme-steel-hrc-fob-china-futures`
- Quặng sắt 62% Fe CFR: `https://vn.investing.com/commodities/iron-ore-62-cfr-futures`
- Than cốc luyện kim (niêm yết CNY, sàn Đại Liên): `https://vn.investing.com/commodities/metallurgical-coke-futures`
- Thép XD/Rebar futures (SRRc1, niêm yết thẳng USD): `https://vn.investing.com/commodities/steel-rebar`
  (trang) hoặc `https://api.investing.com/api/financialdata/historical/996702` (API lịch sử tháng)
- Thép XD nội địa Hòa Phát (VND/kg, chỉ giá hiện tại): `https://www.steelonline.vn/bang-gia-thep-xay-dung-hom-nay`
- Tỷ giá USD/VND: `https://vn.investing.com/currencies/usd-vnd`
- Nguồn thay thế cho quặng sắt (miễn phí, không cần key, nhưng URL có hash đổi mỗi lần cập nhật — phải
  discover qua trang tĩnh `worldbank.org/en/research/commodity-markets` trước khi tải): World Bank
  Commodity Markets "Pink Sheet" (`CMO-Historical-Data-Monthly.xlsx`) — có "Iron ore, cfr spot" nhưng
  KHÔNG có than cốc luyện kim/HRC, nên vẫn cần investing.com cho 2 mặt hàng đó.

**Chưa làm (đề xuất cho lần sau nếu user muốn chính xác hơn):** log file lưu mỗi lần fetch kèm ngày, để
"giá quý hiện tại" dần trở thành median THẬT của nhiều lần chạy script trong quý (thay vì 1 điểm dữ liệu
duy nhất) — chỉ có giá trị cho quý đang chạy trở về sau, không backfill được lịch sử.

---

## Sản lượng thép HPG — tự động dò tin công bố sớm (2026-07)

**Mục đích:** ước tính sản lượng/doanh thu/LNST quý ĐANG CHẠY sớm hơn BCTC chính thức, bằng cách tự
động dò bài công bố sản lượng tháng/quý mới nhất mà HPG đăng trên chính site của họ.

**Cơ chế — 3 nguồn kết hợp (ưu tiên nguồn tách riêng HRC/XD khi trùng tháng/quý):**
- **`fetch_hpg_production_updates()`** (`hoaphat.com.vn`): trang danh sách tin tức render JS (không lấy
  được qua `curl`) NHƯNG mỗi trang bài viết có khung "Tin liên quan" hiện 5 tin mới nhất TOÀN SITE
  (không phụ thuộc bài đang xem) — dùng làm cửa sổ dò tin. Bổ sung quét `sitemap.xml` →
  `sitemap-tintuc-page-N.xml` (tĩnh, không cần JS), lọc URL theo từ khóa (`san-luong`, `trieu-tan`,
  `tan-thep`). Tiêu đề bài công bố có mẫu khá nhất quán: "Sản lượng bán hàng thép Hòa Phát đạt X
  (triệu/nghìn) tấn trong {tháng N | quý N/YYYY}" — quý có thể viết số Ả Rập hoặc La Mã, có/không kèm
  năm. **CHỈ lấy TỔNG sản lượng từ tiêu đề** (đáng tin cậy) — KHÔNG tự tách riêng số HRC từ nội dung
  bài (đã thử và bỏ: câu văn "...(HRC), thép xây dựng, phôi thép đạt X tấn" khiến regex dễ bắt nhầm
  TỔNG thành riêng HRC).
- **`fetch_nguoiquansat_production_updates()`** (`nguoiquansat.vn`, 2026-07 thêm — user cung cấp
  nguồn): đăng lại báo cáo cập nhật sản lượng THÁNG của Vietcap Research, **tách riêng HRC và thép
  xây dựng** (chi tiết & tin cậy hơn hoaphat.com.vn). Quét `sitemap-article-YYYY-MM-DD.xml` (sitemap
  THEO NGÀY, có sẵn tiêu đề trong `<image:title>` nên lọc mà không cần fetch từng bài) ngược 70 ngày.
  Mẫu câu: "sản lượng {thép xây dựng | thép cuộn cán nóng (HRC)} ... trong tháng N[/YYYY] đạt X tấn".
- **`fetch_dautucophieu_production_updates()`** (`dautucophieu.net`, 2026-07 thêm — user cung cấp
  nguồn): đăng lại báo cáo cập nhật hàng tháng của HSC Research, cadence RẤT ĐỀU trong lịch sử (gần
  như tháng nào cũng có bài riêng cho HPG). Discovery qua `/tag/hpg/` (WordPress tag page
  SERVER-RENDERED, liệt kê bài mới nhất TRƯỚC — chỉ 1 lần fetch, đơn giản nhất trong 3 nguồn). Biểu
  đồ sản lượng theo tháng trong bài là ẢNH (không đọc được, không dùng OCR/AI) nhưng số liệu quan
  trọng vẫn nhắc lại bằng TEXT: "HPG bán được X tấn {sản phẩm} trong tháng N[/YYYY]". CHỈ lấy số đã
  XÁC NHẬN, bỏ qua các câu "dự kiến đạt khoảng..." (ước tính của người viết báo cáo, tránh chồng thêm
  ước tính lên ước tính).
- Gộp 3 nguồn theo (loại, năm, quý/tháng) — trùng thì ưu tiên bản có tách HRC/XD riêng
  (nguoiquansat/dautucophieu) hơn bản gộp chung (hoaphat.com.vn).
- Quý đang chạy: có bài công bố QUÝ chính thức → dùng thẳng; chỉ có bài THÁNG → ước tính =
  TB(tháng đã biết)×3 (dùng số HRC/XD tách trực tiếp nếu tất cả tháng đã biết đều có, nếu không mới
  dùng TỶ LỆ lịch sử của quý gần nhất — không suy tỷ lệ từ văn bản).
- **Vị trí Excel**: sheet `15_Quarterly_Data` mục D (bảng quý, có cột quý đang chạy đánh dấu nguồn),
  E (bảng tháng, cột riêng HRC/XD/Tổng + link bài viết + công thức AVERAGE×3 sống cho từng cột), F (dự
  phóng năm run-rate, formula sống). `03_Revenue_Model!G4/G5` LINK trực tiếp về mục F — chuỗi Doanh
  thu→LNST 2026E đã sẵn formula sống nên chỉ cần sửa đúng ô gốc sản lượng.
- VSA (Hiệp hội Thép) chỉ có số TỔNG NGÀNH, không tách HPG — không dùng được cho mục đích ước tính riêng
  HPG (chỉ dùng cho phần vĩ mô ngành đã có sẵn).
- **Nguồn đã xem xét nhưng CHƯA tích hợp** (user gợi ý, xem changelog 2026-07-02(f)/(g) để biết lý do
  từng nguồn): cafef.vn, tapchicongthuong.vn, mekongasean.vn, tinnhanhchungkhoan.vn, 24hmoney.vn,
  hoaphatdungquat.vn (bài đơn lẻ, không có cơ chế dò bài mới); finance.vietstock.vn (chưa kiểm tra cấu
  trúc); báo cáo PDF VCBS/Vietstock/Vietcap (cần thư viện đọc PDF + link tải không cố định).
- **Lưu ý bắt buộc:** MỌI `print()` thêm mới trong các hàm fetch phải là ASCII thuần (không dấu tiếng
  Việt) — console Windows mặc định dùng codepage cp1252, in ký tự có dấu sẽ crash `UnicodeEncodeError`
  (đã xảy ra 2 lần trong quá trình phát triển tính năng này, ở cả bản (d) lẫn bản (e)/(f)).

---

## Dữ liệu & Nguồn (khác — sản lượng, vĩ mô)

**Đầu ra & Ngành**:
- Giá HRC/Rebar: SHFE (Sàn Thượng Hải) trên Trading Economics — dùng để đối chiếu chéo, nguồn fetch chính là investing.com (xem mục trên)
- Hiệp hội Thép VN (VSA): sản lượng, bán hàng, tồn kho nội địa
- Tổng cục Thống kê (GSO): GDP, FDI, giải ngân đầu tư công (kế hoạch trung hạn 2026-2030)
- Tổng cục Hải quan: XNK sắt thép hàng tháng

**Doanh nghiệp**:
- IR websites: hoaphat.com.vn, hoasengroup.vn, namkimgroup.vn (sản lượng hàng tháng/quý)
- CTCK Reports: SSI Research, VNDirect Research, MBS Research, BSI Research

**API Vietcap — Lịch sử P/E, P/B, EV/EBITDA cho HPG và peer thép**:
  - `GET /api/iq-insight-service/v1/company/{ticker}/statistics-financial`
  - Handshake: GET `https://trading.vietcap.com.vn/iq/company?ticker={ticker}` + User-Agent browser
  - Trả về TTM quarterly data: pe, pb, evToEbitda, marketCap, ebitda, roe, roa, roic

**API CafeF — Số CP lưu hành (bsa80 = Vốn góp)**:
  - `/api/iq-insight-service/v1/company/{ticker}/balance-sheet`
  - Số CP = Vốn góp (bsa80) / 10,000 mỗi năm

**API CafeF — Biểu đồ quý PE/PB/EV riêng lẻ**:
  - Mỗi chỉ số vẽ một biểu đồ line chart riêng (3 charts)
  - Dữ liệu quý từ HPG_RATIOS, trim trailing None
  - Labels quý: Q1-YYYY format, spaced every N quarters

**API CafeF — BCTC quý cho dữ liệu volume & sản lượng**:
  - `GET /api/iq-insight-service/v1/company/{ticker}/income-statement`

**PDF Font (multi-platform)**:
  - Windows: `C:/Windows/Fonts/arial.ttf` (Arial, có Vietnamese)
  - Linux (GitHub Actions): `/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf` (DejaVu Sans, có Vietnamese)
  - Workflow: `sudo apt-get install fonts-dejavu-core`
  - Fallback: Helvetica (không hỗ trợ Vietnamese — chỉ dùng khi không có font nào khác)

---

## Phân tích Doanh nghiệp Thép

### Chuỗi cung ứng đầu vào
- **Tỷ lệ nhập khẩu nguyên liệu:** Than cốc (~90%), Quặng sắt (~52%), Thép phế (~47%)
- **Nguồn:** Úc, Nga, Indonesia (than); Úc, Brazil (quặng); Nhật, ASEAN (phế)
- **Đánh giá rủi ro:** Mức độ tập trung nguồn cung, rủi ro địa chính trị (eo biển Hormuz, xung đột TQ-Úc)

### Thị trường đầu ra & Bảo hộ thương mại
- **Các biện pháp AD hiện tại:** Thuế HRC khổ hẹp (19.38-27.83%), HRC khổ rộng chống lẩn tránh (27.83%)
- **Bảng Cung-Cầu HRC nội địa:** Formosa (FMS) vs HPG vs Tổng nhu cầu (tấn)
- **Chiến lược thị trường:** Tỷ trọng nội địa (84%, ↑38% YoY) vs xuất khẩu (16%, ↓ từ 31%)
  - Phân bổ XK: Châu Á 8%, Châu Âu 4%, Châu Mỹ 3%, Châu Úc 1%

### Đầu tư công — Động lực dài hạn
- **Kế hoạch trung hạn 2026-2030:** ~8.5 triệu tỷ đồng (gấp ~3 lần giai đoạn 2021-2025)
- **Kế hoạch 2026:** ~995.4 nghìn tỷ (+10.4% YoY)

---

## Phân biệt Thép Xây dựng vs HRC

| Tiêu chí | Thép xây dựng (Rebar) | HRC (cuộn cán nóng) |
|---|---|---|
| Quy trình | Phôi vuông → Cán dài | Phôi dẹt → Lò nung → Cán nóng liên tục |
| Capex | Thấp hơn | Rất cao (tỷ USD) |
| Khấu hao/tấn | Thấp | Cao |
| Tiêu hao điện | Ít | Nhiều |
| Hợp kim | Cơ bản | Cao hơn, yêu cầu độ tinh khiết cao |

**Công thức giá vốn chuẩn**:
- Thép xây dựng: 1.6×Quặng + 0.6×Than + Chi phí SX khác cố định (`OTHER_COST_USD`)
- HRC: 1.6×Quặng + 0.6×Than + Chi phí SX khác cố định (`OTHER_COST_USD`) — model hiện dùng CHUNG một hằng số
  cho cả 2 loại thép (đơn giản hoá theo yêu cầu user), dù về lý thuyết CP cán nóng HRC có thể cao hơn

**Kiểm chứng từ BCTC**: Vào Thuyết minh → "Chi phí SXKD theo yếu tố" → chia cho sản lượng → ra chi phí/tấn thực tế → so với công thức chuẩn.

---

## Chỉ số quan trọng

**Biên LNG**: Biến động mạnh theo chu kỳ. Đỉnh >20%, đáy có thể âm.
**D/E**: <1.0 an toàn, >1.5 rủi ro (đặc biệt khi lãi suất tăng)
**Vòng quay hàng tồn kho**: Càng cao càng tốt. Tồn kho ứ đọng khi giá thép giảm = thảm họa
**CFO/LNST**: >1.0 = chất lượng LN tốt. <0.8 liên tục = cảnh báo LN ảo

**DIO (Số ngày tồn kho BQ) & DSO (Số ngày phải thu BQ) — 2026-07 (CHỈ LỊCH SỬ, không dự phóng):**
- Công thức: DIO = 365 × Tồn kho bình quân (đầu kỳ+cuối kỳ)/2 ÷ GVHB; DSO = 365 × Phải thu bình quân
  (đầu kỳ+cuối kỳ)/2 ÷ Doanh thu. Dùng số dư BÌNH QUÂN (không phải cuối kỳ) để nhất quán với vòng quay
  HTK; năm/quý đầu tiên trong chuỗi không có số dư đầu kỳ nên dùng số dư cuối kỳ.
- **KHÔNG dự phóng 2026E-2028E** (theo yêu cầu user: DIO/DSO ít ảnh hưởng tới model định giá) — cột dự
  phóng trong Excel để trống (`—`), chỉ D/E và tỷ lệ vay NH/tổng vay vẫn dự phóng đủ (liên quan trực tiếp
  đòn bẩy/rủi ro tài chính).
- **Vị trí công thức sống**: sheet `14_Steel_Analysis`, mục "4. HÀNG TỒN KHO, PHẢI THU & ĐÒN BẨY" — link
  trực tiếp `04_PnL` (Doanh thu/GVHB) và `05_Balance_Sheet` (Phải thu/Tồn kho, chỉ cột B-F lịch sử).
- **Biểu đồ + đánh giá — CẢ NĂM VÀ QUÝ**: `turnover.png` (DIO/DSO theo Năm, 2021-2025) và
  `turnover_quarterly.png` (DIO/DSO theo Quý, năm hóa GVHB/DT quý ×4, khớp Q18_LABELS 2021Q4-2026Q1) +
  narrative tại PDF mục "4C. CHẤT LƯỢNG LỢI NHUẬN & KẾ TOÁN" (so sánh năm đầu vs năm cuối lịch sử, kết
  luận CẢI THIỆN/XẤU ĐI cho từng chỉ số).
- **Biểu đồ tương quan Spread ↔ Biên LNG (BLNG)**: `spread_gpm_corr_annual.png` (scatter + hồi quy, 5
  điểm 2021-2025) và `spread_gpm_corr_quarterly.png` (scatter + hồi quy, tất cả quý có đủ BCTC khớp
  Q18_SPREAD) — CHỈ dùng dữ liệu THỰC TẾ (không đưa số dự phóng vào vì BLNG dự phóng được nội suy TỪ
  chính tỷ lệ Spread, sẽ tạo tương quan giả/circular). Với n nhỏ (đặc biệt chuỗi năm chỉ 5 điểm), hệ số
  Pearson r dễ bị 1 outlier chi phối (VD 2023 — xem giải thích lag-1-quý bên trên) — chỉ mang tính tham
  khảo xu hướng, không phải kết luận thống kê chắc chắn.

---

## Thủ thuật kế toán

**Công cụ**: Trích lập & Hoàn nhập dự phòng giảm giá hàng tồn kho.

- **Giấu lãi**: Trích lập dự phòng tồn kho lớn (đưa vào CP) → giảm LN, cất tiền để dành
- **Book lãi ảo**: Hoàn nhập dự phòng cũ vào DT tài chính/Giảm giá vốn → LN tăng vọt dù ế ẩm

**Check**: Thuyết minh BCTC → mục Hàng tồn kho → cột "Dự phòng". Kiểm tra Phải thu ngắn hạn tăng đột biến nhưng dòng tiền KD âm → bán hàng cho công ty sân sau.

---

## Dấu hiệu sắp hái quả ngọt

1. **Vĩ mô đảo chiều**: BĐS ấm lại, giải ngân đầu tư công tăng, TQ tung kích thích
2. **Spread nở ra**: Giá quặng/than đi ngang/giảm, giá HRC bắt đầu nhích lên
3. **Hàng tồn kho giá rẻ**: Công ty nhập được lượng lớn nguyên liệu lúc đáy giá → quý sau biên LN nổ
4. **Nhà máy mới**: CIP giảm mạnh, chuyển sang TSCĐ → nhà máy xây xong, bắt đầu chạy

---

## Định giá

### Phương pháp kết hợp (40% EV/EBITDA + 40% P/B + 20% P/E)
- **EV/EBITDA (40%, chiết khấu 10%):** Loại bỏ yếu tố cấu trúc vốn và khấu hao, dùng cho tầm nhìn M&A và so sánh quốc tế. Chiết khấu 10% vì tính thận trọng.
- **P/B (40%):** Mua khi P/B ~0.7-1.0x (ngành bi đát nhất). Bán khi P/B ~2.0-2.5x (mọi người hô hào). Công cụ chính cho chu kỳ.
- **P/E (20%, tham khảo, chiết khấu 10%):** P/E là bẫy cho cổ phiếu chu kỳ — P/E thấp (2-5x) = lúc LN đỉnh, P/E cao/âm = đáy chu kỳ. Dùng tỷ trọng nhỏ để bổ sung. Chiết khấu 10% cho an toàn.

**QUY TẮC multiple mục tiêu — dùng lịch sử của CHÍNH cổ phiếu:**
  - Fetch data từ Vietcap API `statistics-financial` (TTM quarterly, quarter != 5)
  - Tính **median** P/B, EV/EBITDA, P/E của toàn bộ lịch sử TTM
  - Dùng median làm multiple mục tiêu cho Base case
  - Kịch bản P/B: Hấp dẫn = median × 0.8, Base = median, Cao = median × 1.2
  - Vẽ biểu đồ riêng cho từng multiple (P/E, P/B, EV/EBITDA) dạng line chart quý
  - **KHÔNG lấy trung bình ngành làm primary** — chỉ để cross-check

**Công thức định giá tổng hợp:**
  - Giá EV/EBITDA = (EBITDA × Multiple - Nợ ròng) × 1e9 / Số CP × 0.90
  - Giá P/B = Multiple × BVPS
  - Giá P/E = Multiple × EPS × 0.90
  - Giá mục tiêu = EV/EBITDA × 40% + P/B × 40% + P/E × 20%

### GP Margin Forecast Methodology
  - **KHÔNG** lấy spread thay thế cho GP margin (spread chỉ đồng pha, không bằng nhau)
  - **⚠️ Dùng Spread ALL (không phải Spread HRC riêng lẻ), TỔNG QUÁT theo N quý ĐÃ CÓ BCTC — chốt
    2026-07 theo yêu cầu user (thay thế công thức đơn giản "chỉ neo Q1" trước đó):**
    - Gọi N = số quý ĐÃ CÓ báo cáo thực tế của năm dự phóng đầu tiên (dùng lại
      `cumulative_actual_quarters()` ở `fetch_data.py` để xác định N và lũy kế Doanh thu/LN gộp N quý).
    - **N < 4:**
      `LNG năm = LNG lũy kế N quý + (4-N)/4 × Doanh thu ước tính năm × (LNG lũy kế N quý / Doanh thu
      lũy kế N quý) × (Spread All hiện tại / Spread All quý gần nhất đã biết)`, rồi
      `BLNG năm = LNG năm / Doanh thu ước tính năm`.
      "Spread All quý gần nhất đã biết" PHẢI dùng sản lượng CỦA CHÍNH quý đó làm quyền số (không phải
      quyền số quý đang chạy) — luôn lấy dòng CUỐI bảng 18 quý ở sheet `17_Gia_Hang_Hoa` (quy ước: bảng
      này phải được bảo trì thêm quý mới mỗi khi có BCTC, nên tự động khớp đúng N). "Spread All hiện
      tại" dùng quyền số = sản lượng quý ĐANG CHẠY (ưu tiên số tách trực tiếp từ nguồn tin, fallback
      sản lượng quý gần nhất đã biết đầy đủ nếu chưa có).
    - **N = 4** (đủ cả năm): `BLNG năm = LNG lũy kế 4 quý / Doanh thu lũy kế 4 quý` — không cần ước
      tính phần còn lại nữa.
    - **Năm dự phóng SAU** (VD 2027E khi đang ở 2026, CHƯA có quý nào của năm đó):
      `BLNG năm = BLNG năm TRƯỚC × (Spread All hiện tại / Spread All NĂM TRƯỚC)` — dùng annual Spread
      All ước tính của CHÍNH năm liền trước (`SPREAD_ALL_A`), và dùng CÙNG "Spread All hiện tại" (không
      chain qua Spread All của từng năm dự phóng xa) cho mọi năm sau đó.
  - Spread All = bình quân gia quyền Spread HRC (đã có thuế CBPG)/Spread Rebar theo sản lượng — xem mục
    "3 LOẠI SPREAD RIÊNG BIỆT" ở trên. Biến Python: `q1_spread_all` (Spread All quý gần nhất thực tế,
    = `Q18_SPREAD_ALL[-1]`), `SPREAD_ALL_NOW` (Spread All hiện tại), `_n_q_known_mod` (N quý đã biết).
  - **Quan trọng:** Không hardcode BLNG dự phóng (VD: 17.5%), phải tính từ số thực tế
  - **Excel formula sống:** `02_Assumptions!row6` (BLNG) và `row45` (Spread hàng năm — LINK sang sheet
    `17_Gia_Hang_Hoa`, KHÔNG tự trừ lại độc lập: ≤2022 dùng Spread HRC, ≥2023 dùng Spread All) phải là
    công thức Excel tham chiếu — để user bấm vào từng ô kiểm chứng độc lập. Biến Python
    (`gpm_2026/2027/2028`) PHẢI tính theo ĐÚNG công thức y hệt Excel (dùng cho narrative PDF/JSON) —
    2 nơi tính riêng nên khi sửa 1 bên phải soát lại bên kia (bài học từ bug SL_HRC_A/SL_XD_A dưới đây).
  - **Vị trí dòng sheet 17 CỐ ĐỊNH, tính trước qua `_r17_annual_row_layout()`** (chỉ phụ thuộc
    `len(Q18_LABELS)`) — cần vì `02_Assumptions` được build TRƯỚC `17_Gia_Hang_Hoa` trong code nhưng
    phải link công thức sang đó. Có `assert` đối chiếu layout thật khi build sheet 17 — lệch sẽ báo lỗi
    ngay thay vì âm thầm link sai ô.

### ⚠️ Bug lớn đã gặp: SL_HRC_A/SL_XD_A lệch xa số liệu thật (2026-07)
  - `SL_HRC_A`/`SL_XD_A` (module-level, sản lượng HRC/XD theo năm) là GIẢ ĐỊNH TĨNH viết tay từ đầu dự
    án — sau khi có `HRC_SALES_HIST_KT`/`XD_SALES_HIST_KT` (dữ liệu quý THẬT, tổng hợp sau này) thì
    KHÔNG được đồng bộ lại, khiến lệch xa số thật (VD 2025 giả định 3.2 triệu tấn HRC, thật là 5.0 triệu
    tấn) — kéo theo doanh thu 2025 (`03_Revenue_Model`) bị tính THIẾU, %tăng trưởng DT 2026E bị THỔI
    PHỒNG giả tạo dù bản chất 2026E không đổi.
  - **Bài học:** bất kỳ mảng "giả định" nào có SỐ THẬT tương ứng ở nơi khác trong model (dù không cùng
    granularity — SL_HRC_A theo NĂM, HRC_SALES_HIST_KT theo QUÝ) đều phải kiểm tra định kỳ có bị lệch
    không, đặc biệt sau khi thêm nguồn dữ liệu thật mới (rất dễ quên đồng bộ ngược lại các chỗ dùng cũ).
  - **Fix:** năm nào có ĐỦ 4 quý dữ liệu thật thì GHI ĐÈ bằng SUM 4 quý (Python để tính SL_HRC_A/
    SL_XD_A dùng nội bộ, Excel dùng công thức SUM sống link sheet `15_Quarterly_Data`) — năm chưa có dữ
    liệu quý (2021-2022) giữ giả định cũ.

---

## Yếu tố Rủi Ro

1. **Biến động sản lượng tiêu thụ** thực tế
2. **Sự dịch chuyển khó lường của giá thép** thế giới
3. **Rủi ro địa chính trị:** Xung đột Trung Đông, phong tỏa eo biển Hormuz → tăng giá than cốc, thép phế, logistics

---

## GitHub Actions — Bảo vệ Custom Builder
  - **CUSTOM_TICKERS whitelist:** HPG, MBB, VCB, BID, ACB, CTG — không bao giờ bị xoá/dùng Gemini generate lại
  - **Size guard:** Builder >500 dòng tự động coi là custom, không bị regenerate
  - **Error handling:** Wrap `build_pdf()` trong try/except để `save_json_summary()` luôn chạy (tránh mất data JSON khi PDF lỗi font)

## Lưu ý Khi Viết Báo Cáo

- **Ngắn gọn, số liệu là chính:** Dùng bullet points, bảng biểu, không viết văn dài
- **Forecast horizon ngắn:** Chỉ 3 năm quá khứ + 2 năm dự phóng
- **Đơn vị:** Tỷ đồng (VN), USD/tấn (quốc tế)
- **Luôn so sánh kế hoạch vs thực hiện** (kèm % hoàn thành)
- **Kết thúc mỗi phần bằng bảng/bullet, không xuống dòng văn dài**
- **Embed số liệu thô của biểu đồ** dưới dạng bảng (VD: Công suất/Nhu cầu theo năm) để reader tự kiểm chứng
- **Mỗi nhận định phải có căn cứ số liệu**
- **Các chỉ số cần thiết cho mỗi cổ phiếu thép:**
  - Thị trường: KLGD 20 phiên, biến động 1m/3m/YTD, so sánh với VNINDEX
  - TC: Kế hoạch vs thực hiện DT/LNST, ROE, ROA, D/E, EPS
  - Ngành: Công suất dư thừa, giá HRC/quặng/than, spread, FMS vs HPG vs total demand**
