# Changelog — Skill Thép

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
