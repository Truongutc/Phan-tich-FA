# Schema `data/fvtpl_holdings/<TICKER>.json`

Kho dữ liệu danh mục tự doanh FVTPL/AFS cho CTCK, trích trực tiếp từ thuyết minh BCTC hợp nhất
(mục "Tài sản tài chính ghi nhận thông qua lãi/lỗ (FVTPL)" và "...sẵn sàng để bán (AFS)").
**Vị trí thuyết minh**: luôn nằm ngay sau thuyết minh "Tiền và các khoản tương đương tiền" (mục 3.1),
thường là mục 3.2 (FVTPL) và 3.3/3.4 (AFS) trong phần "THÔNG TIN BỔ SUNG BÁO CÁO TÌNH HÌNH TÀI CHÍNH".
Khi convert PDF→MD, phần này hay bị cắt do MD cắt 1/4 trên + 1/4 dưới — **luôn ưu tiên render ảnh
trực tiếp bằng `bctc_pdf_tool.py render` thay vì đọc bản MD** cho các note này.

Đơn vị: VND (giữ nguyên, KHÔNG quy đổi tỷ VND, vì holdings từng mã có giá trị nhỏ dễ méo khi làm tròn).

5 nhóm tài sản CỐ ĐỊNH (không chi tiết sub-item, chỉ lấy tổng lớn theo yêu cầu):
- `CoPhieuNiemYet` — Cổ phiếu niêm yết + giao dịch trên UPCoM (gộp cả "Tài sản cơ sở cho hoạt động
  phòng ngừa rủi ro chứng quyền" nếu CTCK có nghiệp vụ chứng quyền — bản chất vẫn là cổ phiếu niêm yết).
- `CoPhieuChuaNiemYet` — Cổ phiếu chưa niêm yết (OTC).
- `TraiPhieu` — Trái phiếu (niêm yết + trái phiếu riêng lẻ trên HNX).
- `ChungChiQuy` — Chứng chỉ Quỹ (gộp ETF + Quỹ mở).
- `ChungChiTienGui` — Chứng chỉ tiền gửi có thể chuyển nhượng.

```json
{
  "ticker": "HCM",
  "companyName": "CTCP Chứng khoán TP.HCM",
  "unit": "vnd",
  "lastUpdated": "2026-07-11",
  "categories": {
    "CoPhieuNiemYet":     {"label": "Cổ phiếu niêm yết/UPCoM", "order": 1, "color": "#3b82f6"},
    "CoPhieuChuaNiemYet": {"label": "Cổ phiếu chưa niêm yết",  "order": 2, "color": "#8b5cf6"},
    "TraiPhieu":          {"label": "Trái phiếu",              "order": 3, "color": "#f59e0b"},
    "ChungChiQuy":        {"label": "Chứng chỉ quỹ (ETF+Quỹ mở)", "order": 4, "color": "#10b981"},
    "ChungChiTienGui":    {"label": "Chứng chỉ tiền gửi",      "order": 5, "color": "#64748b"}
  },
  "fvtpl": {
    "quarterly": {
      "2025Q2": {
        "CoPhieuNiemYet": {"cost": 3802461608772, "fairValue": 3851051710436, "gain": 48855617248, "loss": -265515584,
                            "source": "BCTC HN Q2/2025 TM 3.2", "sourceType": "bctc_pdf"}
      }
    },
    "yearly": { "2024": { "...": "..." } }
  },
  "afs": { "quarterly": {}, "yearly": {} },
  "holdings": {
    "note": "Danh mục cổ phiếu niêm yết cụ thể trong nhóm CoPhieuNiemYet (fair value, VND). 'Khac' = phần dư gộp các mã nhỏ không nêu tên.",
    "quarterly": {
      "2025Q2": {"TCB": 1219893480262, "FPT": 891074340000, "Khac": 76413469636}
    },
    "yearly": {}
  }
}
```

Công cụ: dùng chung `bctc_pdf_tool.py` (list/download/render) đã xây cho KCN — không cần công cụ
riêng vì cơ chế tải PDF BCTC hợp nhất giống hệt nhau cho mọi loại doanh nghiệp niêm yết.
