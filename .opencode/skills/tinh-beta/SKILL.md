---
name: tinh-beta
description: Tính toán hệ số rủi ro Beta dựa trên 100 phiên giao dịch gần nhất của cổ phiếu so với VN-Index.
---

# Hướng dẫn Tính toán Hệ số Beta cho Cổ phiếu

Tài liệu này định nghĩa phương pháp tính toán hệ số Beta của cổ phiếu tương quan với chỉ số VN-Index dựa trên 100 phiên giao dịch gần nhất. Hệ số Beta này sau đó được sử dụng trong mô hình CAPM để tính chi phí vốn chủ sở hữu (COE) phục vụ định giá Residual Income (RI).

## 1. Công thức Toán học

Hệ số Beta được xác định theo công thức:

$$\beta = \frac{Covar(R_i, R_m)}{Var(R_m)}$$

Trong đó:
- $R_i$: Tỷ suất sinh lời hàng ngày của cổ phiếu.
- $R_m$: Tỷ suất sinh lời hàng ngày của thị trường (chỉ số VN-Index).
- $Covar(R_i, R_m)$: Hiệp phương sai của tỷ suất sinh lời cổ phiếu và thị trường.
- $Var(R_m)$: Phương sai của tỷ suất sinh lời thị trường.

Tỷ suất sinh lời hàng ngày ($R$) được tính bằng:

$$R_t = \frac{P_t - P_{t-1}}{P_{t-1}}$$

Với:
- $P_t$: Giá đóng cửa điều chỉnh tại ngày $t$.
- $P_{t-1}$: Giá đóng cửa điều chỉnh tại ngày $t-1$.

## 2. Quy tắc số phiên giao dịch và Nguồn dữ liệu

Hệ thống sẽ tải tối đa dữ liệu lịch sử giá của Cổ phiếu và VN-Index (lên tới 2 năm, khoảng 500 phiên).
- **Dưới 250 phiên (< 1 năm)**: 
  - Vẫn tạo trang tính `00_Beta` và điền dữ liệu kèm công thức tính toán.
  - Tuy nhiên, trong trang tính `00_COE` ô `B5` (Hệ số Beta), giá trị sẽ lấy **theo Web** (Vietstock/Vietcap API) thay vì liên kết sang `00_Beta`.
  - Bên cạnh ô Beta (ô `C5`), hệ thống chèn một liên kết tìm kiếm: `=HYPERLINK("...", "Tra cứu Beta trên Vietcap (Số phiên < 1 năm)")` để người dùng dễ dàng kiểm tra.
- **Từ 250 phiên trở lên (>= 1 năm)**:
  - Tự động lấy hệ số Beta tự tính toán từ trang tính `00_Beta`.
  - Liên kết trực tiếp `00_COE!B5` sang `='00_Beta'!C1`.
  - Cột `C5` sẽ hiển thị liên kết tra cứu tham khảo thông thường.

## 3. Tích hợp trong Excel Model

Hệ thống sẽ tự động tạo một trang tính riêng tên là `00_Beta` nằm trước trang tính `00_COE`.
- **Dữ liệu nguồn**: Tối đa 501 dòng giá đóng cửa điều chỉnh gần nhất của Cổ phiếu và VN-Index xếp theo thứ tự thời gian tăng dần.
- **Công thức tính tỷ suất sinh lời (Cột C và E)**: `=(B6-B5)/B5`
- **Công thức tính Beta (Ô C1)**: `=COVAR(C6:C{last_row}, E6:E{last_row})/VAR(E6:E{last_row})`
- **Số phiên thực tế (Ô C2)**: `=COUNT(C6:C{last_row})`
