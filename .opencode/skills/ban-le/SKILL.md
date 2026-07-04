---
name: ban-le
description: Use ONLY when analyzing Vietnamese retail stocks. Front-loaded keywords: bán lẻ, retail, chuỗi cửa hàng, MWG, PNJ, FRT, DGW, bách hóa, trang sức, điện máy, dược phẩm, SSSG, cùng store sales, cửa hàng.
---

# KĨ NĂNG PHÂN TÍCH NGÀNH BÁN LẺ

## Top-Down 5 bước

1. **Vĩ mô & Sức mua**: CPI (lạm phát), thu nhập khả dụng, tín dụng tiêu dùng
2. **Phân loại phân khúc**: Hàng thiết yếu (dược, thực phẩm) vs không thiết yếu (ICT, trang sức, thời trang). Kênh: GT truyền thống, MT hiện đại, E-commerce
3. **Năng lực cạnh tranh**: Mặt bằng đắc địa, buying power với NCC, logistics & chuỗi cung ứng
4. **BCTC**: Trái tim là **Hàng tồn kho** (60-80% TS ngắn hạn) và **Chi phí bán hàng** (thuê mặt bằng + lương NV)
5. **Định giá & Điểm rơi LN**: Xác định khi chuỗi chạm điểm hòa vốn

---

## Dữ liệu & Nguồn

**Vĩ mô**: Tổng mức bán lẻ hàng hóa & DV tiêu dùng (GSO hàng tháng), CPI (GSO)
**Phân khúc**:
- Trang sức (PNJ): Kitco Gold (giá vàng TG), SJC/Doji (giá vàng trong nước)
- ICT/Điện máy (MWG, FRT, DGW): Chu kỳ ra mắt sp mới (Apple, Samsung), tín dụng tiêu dùng (FE Credit, Home Credit)
**Doanh nghiệp**:
- MWG IR: Doanh thu hàng tháng từng chuỗi (TGDD, ĐMX, BHX)
- FRT IR: Tốc độ mở chuỗi Long Châu
- PNJ IR: KQKD định kỳ hàng tháng
**Báo cáo CTCK**: SSI Research, MBS Research

---

## Quy trình phân tích & dự phóng (áp dụng MWG trước, cùng cấu trúc cho FRT/PNJ)

### 1. Phân tích cơ cấu doanh thu

Phân chia kết cấu doanh thu theo từng mảng kinh doanh (VD MWG: TGDD, ĐMX, BHX, An Khang, EraBlue campuchia/indo...).

**Nguồn dữ liệu**: Thuyết minh BCTC, báo cáo IR doanh nghiệp, báo cáo thường niên, báo chí/nguồn uy tín (xem mục "Dữ liệu & Nguồn" bên dưới).

Sau khi có kết cấu doanh thu từng mảng, đánh giá cho MỖI mảng:
- Quy mô thị trường tiêu thụ còn rộng hay đã bão hòa
- Ngành còn tăng trưởng hay đã chững lại
- Thị phần hiện tại của doanh nghiệp trong mảng đó
- Lợi thế cạnh tranh là gì, có đủ mạnh để duy trì dài hạn không

### 2. Đánh giá hiệu quả từng cửa hàng (quan trọng nhất)

Cửa hàng mới mở ~3 tháng đầu thường CHƯA ổn định (ramp-up) → phải lấy số lượng cửa hàng LÙI 3 THÁNG so với kỳ doanh thu đang tính, không dùng số cửa hàng cuối kỳ hiện tại.

**TH1 — có dữ liệu doanh thu & số cửa hàng theo THÁNG:**
```
Hiệu quả CH/tháng = Doanh thu tháng hiện tại / Số cửa hàng của mảng đó CÁCH thời điểm hiện tại 3 THÁNG
```

**TH2 — chỉ có dữ liệu theo QUÝ:**
```
Hiệu quả CH/tháng = Doanh thu quý của mảng / Số cửa hàng của mảng CÁCH thời điểm hiện tại 3 THÁNG / 3
```

Sau đó: vẽ biểu đồ hiệu quả/cửa hàng theo thời gian → đánh giá xu hướng (tăng/chững/giảm) → xác định mảng đã bão hòa hay chưa. Đây chính là cách tính SSSG-tương-đương của skill này (xem thêm mục SSSG bên dưới — 2 công thức bổ trợ nhau, công thức lag-3-tháng chính xác hơn khi chuỗi đang mở nhanh).

### 3. Dự báo doanh thu

Dự phóng doanh thu từng mảng dựa trên 3 yếu tố kết hợp:
1. Tăng trưởng cơ học của từng mảng (ngành, thị phần)
2. Số lượng cửa hàng dự kiến (kế hoạch mở mới/đóng của doanh nghiệp, IR)
3. Hiệu quả doanh thu/cửa hàng (xu hướng từ mục 2 ở trên)

`Doanh thu mảng = Số cửa hàng dự phóng × Hiệu quả DT/cửa hàng dự phóng` (điều chỉnh theo tăng trưởng cơ học/SSSG).

### 4. Dự báo lợi nhuận

- **Biên lợi nhuận gộp**: ước tính từ biên LNG của **2 quý gần nhất** (không dùng trung bình dài hạn — biên LNG bán lẻ đổi nhanh theo cơ cấu sản phẩm/rebate).
- **Chi phí bán hàng & QLDN (SG&A)**: dự phóng theo xu hướng tỷ lệ % trên doanh thu của **4 quý gần nhất**.
- Nội suy ra LNST từ Doanh thu → GP (biên LNG) → trừ SG&A → LNST.

### 5. Chỉ số quan trọng nhất

#### SSSG (Same-Store Sales Growth)

SSSG dương > lạm phát → thương hiệu mạnh, giữ chân KH cũ.
DT tổng tăng 20% nhưng SSSG âm → các CH cũ ế ẩm, đang đốt tiền mở CH mới lấy tăng trưởng ảo.

Tính: Doanh thu 1 CH/tháng = DT mảng trong quý / Số CH BQ trong kỳ / 3 (bản đơn giản — xem mục 2 ở trên cho công thức lag-3-tháng chính xác hơn).

#### Chu kỳ tiền mặt (CCC)

Bán lẻ đỉnh cao: bán thu tiền ngay, ngâm tiền NCC vài tháng mới trả.

DIO = Hàng tồn kho BQ / (Giá vốn / 365) — Số ngày tồn kho
DSO = Phải thu KH BQ / (Doanh thu / 365) — Số ngày thu tiền
DPO = Phải trả NB BQ / (Giá vốn / 365) — Số ngày trả tiền

CCC = DIO + DSO - DPO

**Tiêu chuẩn**: CCC càng nhỏ càng tốt (1-5 ngày). CCC âm = siêu đẳng (không cần vốn tự có).

---

## Dấu hiệu thủ thuật kế toán

**Book lãi ảo**:
- Ghi nhận sớm Rebates/Chiết khấu TM → giảm giá vốn, làm đẹp LN
- Tuồn hàng ra đại lý sân sau → DT tăng nhưng Phải thu phình to bất thường
- Chậm trích lập dự phòng hàng lỗi mốt/cận date

**Giấu lãi**:
- Trích trước chi phí sửa chữa, khuyến mãi quá mức hồi quý lãi đậm

**Check**: Thuyết minh BCTC → mục Hàng tồn kho (cột Dự phòng), Chi phí trả trước, Phải trả ngắn hạn khác.

---

## Dấu hiệu sắp hái quả ngọt

1. **Chuỗi mới chạm điểm hòa vốn (Inflection Point)**: DT/CH đạt ngưỡng → biên LN thuần nổ theo chữ V
2. **DIO giảm mạnh** nhưng DT tăng → logistics tối ưu, dòng tiền về
3. **SG&A/Sales giảm dần** → hiệu ứng quy mô, mở CH mới không tốn thêm CP vận hành

---

## Định giá (2026-07 — cập nhật theo quy trình chi tiết user cung cấp)

Ưu tiên **định giá theo từng mảng kinh doanh** thay vì định giá cả công ty gộp chung — đặc thù bán lẻ có nhiều mảng ở giai đoạn trưởng thành khác nhau (VD MWG: TGDD/ĐMX đã bão hòa, BHX/An Khang đang scale).

> Lưu ý trước đây skill này ghi "KHÔNG dùng P/B" (theo `Ban le.docx` — tài sản chủ yếu đi thuê, giá trị sổ sách thấp không ý nghĩa). User đã xác nhận vẫn dùng P/B với trọng số nhỏ (20%) trong bộ 5 phương pháp dưới đây, kết hợp cùng các phương pháp khác để cân bằng lại nhược điểm của P/B.

### 5 phương pháp định giá tổng hợp (trọng số đã chốt với user, 2026-07)

| # | Phương pháp | Trọng số | Ghi chú |
|---|---|---|---|
| 1 | P/E Median | 20% | Median lịch sử TTM của chính cổ phiếu (giống cách tính ở skill `thep`/`ngan-hang`) |
| 2 | P/B Median | 20% | Đã xác nhận dùng dù tài sản chủ yếu đi thuê — trọng số nhỏ để cân bằng |
| 3 | Residual Income (RI), cả công ty | 10% | Cho toàn công ty (khác với RI theo mảng — hiện KHÔNG dùng RI ở cấp mảng nữa, xem bên dưới) |
| 4 | Định giá theo từng mảng — P/S + P/E | 35% | Áp dụng cho mảng đã trưởng thành, có P/E tham chiếu đáng tin cậy (xem quy tắc bên dưới) |
| 5 | Định giá theo từng mảng — P/S + P/B | 15% | Áp dụng cho mảng đã trưởng thành nhưng KHÔNG có P/E tham chiếu phù hợp/đủ tin cậy — dùng P/B thay thế |

Tổng = 20+20+10+35+15 = **100%**.

### Nguyên tắc chọn phương pháp cho TỪNG MẢNG (bước #4, #5)

**Dùng P/S khi mảng có:**
- Biên lợi nhuận rất thấp hoặc đang lỗ
- Ngành có quy mô thị trường lớn
- Đang giảm lỗ hoặc mới chớm có lãi
- Ngành còn tốc độ tăng trưởng cao
- Doanh nghiệp chưa chiếm thị phần quá lớn, thị trường chưa bão hòa
- Mẹo: tra P/S retail cùng ngành ở Đông Nam Á (Thái Lan, Indonesia) làm mốc tham chiếu

**Với mảng đã trưởng thành** (hiệu quả DT/cửa hàng đi ngang hoặc tăng rất chậm, quy mô thị trường tăng trưởng thấp, đã dần bão hòa) — chọn 1 trong 2:

**P/S + P/E** (trọng số 35%, ưu tiên khi có đủ dữ liệu tham chiếu):
- P/E tham chiếu = P/E doanh nghiệp cùng ngành ở nước ngoài, hoặc DN tương đồng trong nước
- Hoặc P/E = 0.9 × Growth (tốc độ tăng trưởng %/năm của mảng đó)
- **BẮT BUỘC lấy giá trị NHỎ HƠN** giữa (P/E tham chiếu) và (0.9 × Growth) — tránh trường hợp mảng tăng trưởng đột biến 1 năm (VD +30%) bị máy móc suy ra P/E=27, định giá quá cao không bền vững.

**P/S + P/B** (trọng số 15%, dùng khi KHÔNG có P/E tham chiếu phù hợp/đủ tin cậy cho mảng đó):
- P/B tham chiếu theo median lịch sử của chính mảng/công ty tương đồng, tương tự cách tính P/B Median ở phương pháp #2.

### Vì sao KHÔNG chỉ dùng EV/EBITDA đơn thuần (bổ sung, không mâu thuẫn 5 phương pháp trên)
Bán lẻ thuê mặt bằng rất nhiều, chịu ảnh hưởng hạch toán thuê tài sản (IFRS 16) — EV/EBITDA giúp loại bỏ khấu hao/cấu trúc nợ khi so sánh chuỗi khác nhau, có thể dùng làm cross-check bổ sung ngoài 5 phương pháp chính.
