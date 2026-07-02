# Changelog — Skill Ngân hàng

## 2026-07-02 — Dự phóng KQKD/tài sản NĂM blend với số quý ĐÃ CÓ báo cáo thực tế (dùng chung với HPG)

### Bối cảnh
User phát hiện bug tương tự ở `build_hpg_model.py` (DTTC Q1/2026 thực tế ~5.900 tỷ nhưng giả định năm
2026 chỉ 2.500 tỷ — SAI vì bỏ qua thực tế) và yêu cầu tổng quát hóa: KHÔNG ngoại suy tuyến tính Q1×4
(sai theo hướng khác nếu Q1 đột biến 1 lần), mà blend theo công thức: "Ước tính năm = lũy kế thực tế
quý đã biết + giả định ban đầu × phần quý CHƯA biết/4" — áp dụng cho CẢ HPG lẫn template ngân hàng.

### Thay đổi
- Thêm `cumulative_actual_quarters()`, `blend_annual_estimate()` (chỉ tiêu KQKD lũy kế) và
  `latest_actual_quarter_value()`, `blend_annual_estimate_stock()` (chỉ tiêu số dư cuối kỳ — dư
  nợ/huy động/tổng tài sản, KHÔNG cộng dồn được như KQKD) vào `fetch_data.py` — dùng chung cho mọi
  ticker builder (không riêng bank).
- `template_banking.py`: blend 3 điểm trong "Build Forecast Model" (đầu năm dự phóng đầu tiên,
  `years_fc[0]`):
  1. `loans_fc[0]` (Dư nợ tín dụng, `bsb103`) — re-anchor về số dư quý gần nhất, áp phần tăng trưởng
     giả định gốc còn lại (dùng `blend_annual_estimate_stock`, vì đây là số dư không phải dòng chảy).
  2. `nii_fc[0]` (NII, `isb27`) — blend lũy kế thực tế + giả định gốc phần còn lại, cascades tự nhiên
     qua `toi_fc`→`ppop_fc`.
  3. `np_fc[0]` (LNST, `isa20`) — blend độc lập theo thực tế, rồi BACK-SOLVE `prov_fc[0]` làm biến điều
     chỉnh để giữ nhất quán PPOP→PBT→Thuế→LNST (không ghi đè LNST rồi để lệch hẳn PBT-Thuế trong sheet).
- Verify bằng cách chạy thật `python run_analysis.py TCB` (không chỉ đọc code) — kết quả: Dư nợ Q1/2026
  785.613 tỷ → blend 895.658 tỷ cả năm; NII Q1 9.522 tỷ → blend 44.028 tỷ; LNST Q1 6.950 tỷ → blend
  28.669 tỷ. Script chạy hết pipeline (Excel + win32com readback + PDF + JSON) không lỗi.

### Chưa làm
- Chỉ blend NII/LNST/Dư nợ — CHƯA blend Tổng tài sản (`bsa53`), Huy động (`bsb113`), Non-II riêng —
  nếu cần chính xác hơn có thể áp `blend_annual_estimate`/`blend_annual_estimate_stock` tương tự.
- Chưa test được kịch bản Q2/Q3/Q4 thực tế (n=2,3) vì hiện tại mới chỉ có Q1/2026 cho các NH — logic
  đã viết tổng quát cho n=0..4 nhưng cần verify lại khi có đủ dữ liệu các quý sau.

---

## 2026-07-01 — Sửa lỗi LDR, tăng trưởng tín dụng, P/B naming (từ báo cáo TCB)

### Bối cảnh
User phát hiện `06_Ratios_Quarterly`/`06_Ratios` tính LDR sai (dùng "Cho vay khách hàng" làm tử số thay
vì "Tổng tín dụng"), tỷ lệ giữ lại tiền gửi kho bạc (KBNN) sai thứ tự/giá trị, và dự phóng tăng trưởng
tín dụng tương lai bị "số chết" (hardcode 14% thay vì tính từ công thức).

### Các thay đổi đã áp dụng vào skill

| # | Nội dung sửa | Lý do |
|---|---|---|
| 1 | LDR = Tổng tín dụng / Tổng huy động (không phải Cho vay KH / Tiền gửi KH) | Đúng định nghĩa pháp lý Thông tư 22/2019/TT-NHNN |
| 2 | Tổng huy động phải cộng Tiền gửi TCTD khác (bsb270) | Điều 20, TT 22/2019 — thiếu dòng này khiến LDR bị đẩy sai lên ~98-100% thay vì ~89-90% thực tế |
| 3 | Tỷ lệ giữ lại KBNN đúng: `[0, 0, 50%, 40%, 20%]` theo TT 26/2022, khôi phục 20% cho 2026+ theo TT 08/2026 | Mảng cũ `[0,0,35%,50%,60%]` sai thứ tự/giá trị |
| 4 | "Tổng tín dụng"/"Tổng huy động" phải là dòng riêng tường minh ở Balance Sheet, sheet Ratios chỉ LINK công thức | Tránh 2 công thức tính độc lập lệch nhau âm thầm |
| 5 | Dự phóng tăng trưởng tín dụng = `AVERAGE(3 năm gần nhất)*0.9`, không hardcode | User yêu cầu rõ "đừng để số chết" — công thức sống để tự kiểm chứng |
| 6 | Đổi tên `pb_target` → `pb_over` (Excel/PDF/web) | "Target" dễ hiểu nhầm là mức kỳ vọng đạt tới; bản chất chỉ là median nửa trên phân phối lịch sử dùng cho vùng chốt lời |
| 7 | Không hiển thị target price cá nhân 1 analyst (VD: "Quan Vu" từ Vietcap API) | Chỉ nên so sánh theo CTCK/nguồn tổ chức, không gắn tên cá nhân |
| 8 | Bỏ `SEQUENCE()` trong công thức median-nửa, dùng CSE array formula `MEDIAN(SMALL/LARGE(range, ROW(INDIRECT("1:"&k))))` | `SEQUENCE()` là dynamic array Excel 365, gây `#NAME?` trên bản Excel cũ hơn của user |

### Ghi chú
- Tất cả các fix trên đã áp dụng cho `template_banking.py` (TCB) — dùng làm baseline khi phân tích bank khác.
- Verify formula bằng `win32com` (mở lại file Excel, đọc giá trị tính toán) trước khi báo hoàn thành —
  không chỉ dựa vào openpyxl đọc formula string, vì có thể có lỗi cú pháp không phát hiện được nếu không mở
  bằng chính Excel thật.
