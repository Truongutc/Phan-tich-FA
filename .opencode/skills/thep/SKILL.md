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

**Spread = Giá thép − 1.6×Giá quặng sắt − 0.6×Giá than cốc − Chi phí SX khác CỐ ĐỊNH (USD/tấn)**

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
2. **Bảng 18 quý lịch sử** (2021Q4 → hiện tại): tổng hợp thủ công từ nhiều nguồn công khai (xem ghi chú
   nguồn ngay trong sheet, cột B "NGUỒN DỮ LIỆU"). Đây là **số liệu nghiên cứu, không phải median 3 mốc
   trong quý** (đầu/giữa/cuối quý) như lý tưởng — vì không có API lịch sử giá theo ngày miễn phí. Độ tin cậy
   khác nhau theo mặt hàng: quặng sắt (cao, đối chiếu World Bank Pink Sheet), than cốc (trung bình, FPTS/
   VCBS/Argus), HRC (thấp hơn, chỉ đúng xu hướng — Mysteel/SteelBenchmarker/giá XK HPG).
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

**URLs đã verify hoạt động qua `curl` (2026-07):**
- HRC: `https://www.investing.com/commodities/lme-steel-hrc-fob-china-futures`
- Quặng sắt 62% Fe CFR: `https://vn.investing.com/commodities/iron-ore-62-cfr-futures`
- Than cốc luyện kim (niêm yết CNY, sàn Đại Liên): `https://vn.investing.com/commodities/metallurgical-coke-futures`
- Nguồn thay thế cho quặng sắt (miễn phí, không cần key, nhưng URL có hash đổi mỗi lần cập nhật — phải
  discover qua trang tĩnh `worldbank.org/en/research/commodity-markets` trước khi tải): World Bank
  Commodity Markets "Pink Sheet" (`CMO-Historical-Data-Monthly.xlsx`) — có "Iron ore, cfr spot" nhưng
  KHÔNG có than cốc luyện kim/HRC, nên vẫn cần investing.com cho 2 mặt hàng đó.

**Chưa làm (đề xuất cho lần sau nếu user muốn chính xác hơn):** log file lưu mỗi lần fetch kèm ngày, để
"giá quý hiện tại" dần trở thành median THẬT của nhiều lần chạy script trong quý (thay vì 1 điểm dữ liệu
duy nhất) — chỉ có giá trị cho quý đang chạy trở về sau, không backfill được lịch sử.

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
  - **Cách tính (bottom-up từ quý gần nhất):**
    1. Lấy LNG & doanh thu quý gần nhất từ BCTC (VD: Q1/2026)
    2. Tính BLNG quý gần nhất = LNG / Doanh thu
    3. Tính tỉ lệ Spread = Spread hiện tại / Spread quý gần nhất
    4. BLNG các quý còn lại = BLNG quý gần nhất × tỉ lệ Spread
    5. LNG 2026 = LNG lũy kế + Doanh thu ước tính các quý còn lại × BLNG các quý còn lại
    6. BLNG 2026 = LNG 2026 / Tổng doanh thu 2026
  - Spread = Giá HRC - 1.6×Giá quặng - 0.6×Giá than cốc - Chi phí SX khác CỐ ĐỊNH (`OTHER_COST_USD`, USD/tấn)
  - **Quan trọng:** Không hardcode BLNG dự phóng (VD: 17.5%), phải tính từ số thực tế
  - **Excel formula sống:** cả Spread quý gần nhất VÀ Spread hàng năm (mẫu số + tử số của tỉ lệ ở bước 3)
    phải là công thức Excel tham chiếu tới sheet `17_Gia_Hang_Hoa` (KHÔNG phải Python tính sẵn rồi ghi số
    vào ô) — để user bấm vào từng ô kiểm chứng độc lập cách tính BLNG dự phóng.

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
