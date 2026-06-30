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

## 2. Quy tắc số phiên giao dịch

- **Dưới 30 phiên**: Không thực hiện tính toán Beta (sử dụng hệ số fallback hoặc Vietcap details API).
- **Từ 30 đến dưới 500 phiên**: Tính Beta dựa trên toàn bộ dữ liệu lịch sử từ khi niêm yết đến phiên gần nhất.
- **Từ 500 phiên trở lên**: Chỉ tính toán dựa trên **500 phiên giao dịch gần nhất** (tương đương khoảng 2 năm giao dịch) để phản ánh chính xác rủi ro hệ thống hiện tại và ổn định trong trung hạn.

## 3. Tích hợp trong Excel Model

Hệ thống sẽ tự động tạo một trang tính riêng tên là `00_Beta` nằm trước trang tính `00_COE`.
- **Dữ liệu nguồn**: Tối đa 501 dòng giá đóng cửa điều chỉnh gần nhất của Cổ phiếu và VN-Index xếp theo thứ tự thời gian tăng dần.
- **Công thức tính tỷ suất sinh lời (Cột C và E)**: `=(B5-B4)/B4`
- **Công thức tính Beta (Ô C1)**: `=COVAR(C6:C{last_row}, E6:E{last_row})/VAR(E6:E{last_row})`
- **Liên kết sang COE**: Trong trang tính `00_COE`, ô hệ số Beta (`B5`) sẽ được liên kết trực tiếp bằng công thức: `='00_Beta'!C1`.
