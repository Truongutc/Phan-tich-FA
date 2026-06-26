---
name: phan-loai-nganh
description: Use ONLY when needing to classify a Vietnamese stock ticker into its industry group. Front-loaded keywords: ngành, nhóm ngành, phân loại, industry, sector, classification, ticker VN thuộc ngành nào, cùng ngành với.
---

# PHÂN LOẠI NGÀNH & CỔ PHIẾU VIỆT NAM

Khi nhận ticker, tra bảng này để xác định nhóm ngành. Sau đó dùng skill `fa` để áp dụng phương pháp dự báo doanh thu & định giá tương ứng.

| # | Nhóm ngành | Cổ phiếu tiêu biểu |
|---|---|---|
| 1 | **Ngân hàng** | VCB, BID, TCB, MBB, VPB, ACB |
| 2 | **Chứng khoán** | SSI, VND, HCM, VCI, FTS, SHS, BSI |
| 3 | **Bất động sản dân cư** | VHM, NLG, KDH, DXG, PDR, NVL |
| 4 | **Bất động sản KCN** | IDC, KBC, BCM, SZC, VGC, SIP |
| 5 | **Xây dựng** | CTD, VCG, HBC, FCN, LCG |
| 6 | **Hạ tầng - BOT - Đầu tư công** | HHV, C4G, VCG, FCN, LCG |
| 7 | **Thép** | HPG, HSG, NKG, TVN |
| 8 | **Tôn mạ** | HSG, NKG, GDA |
| 9 | **Dầu khí thượng nguồn** | PVD, PVS, PVC |
| 10 | **Dầu khí trung - hạ nguồn** | GAS, BSR, OIL, PLX |
| 11 | **Điện** | POW, NT2, QTP, PPC, GEG, REE |
| 12 | **Nước sạch** | BWE, TDM, DNW |
| 13 | **Phân bón** | DPM, DCM, LAS, BFC |
| 14 | **Hóa chất** | DGC, CSV, DDV |
| 15 | **Dược phẩm** | DHG, TRA, IMP, DBD, DHT |
| 16 | **Bán lẻ** | MWG, FRT, PNJ, DGW |
| 17 | **Công nghệ** | FPT, CMG, ELC |
| 18 | **Viễn thông** | CTR, FOX, VGI |
| 19 | **Thủy sản** | VHC, ANV, FMC, MPC, CMX |
| 20 | **Dệt may** | TNG, MSH, STK, VGT, ADS |
| 21 | **Sợi** | STK, ADS |
| 22 | **Cảng biển** | GMD, VSC, PHP, SGP |
| 23 | **Vận tải biển** | HAH, VOS, PVT, MVN |
| 24 | **Logistics** | GMD, VSC, TCL, PDN |
| 25 | **Gỗ - Nội thất** | PTB, GDT, SAV |
| 26 | **Vật liệu xây dựng** | VCS, HT1, BCC, BTS |
| 27 | **Bao bì - Giấy** | DHC, HHP, GVT |
| 28 | **Nhựa** | AAA, BMP, NTP |
| 29 | **Cao su** | GVR, DPR, PHR, TRC |
| 30 | **Hàng không** | HVN, VJC, ACV |
| 31 | **Du lịch - Khách sạn** | VNG, HOT, OCH |
| 32 | **Nông nghiệp** | HAG, BAF, DBC, PAN |
| 33 | **Chăn nuôi** | DBC, BAF, HAG |
| 34 | **Thực phẩm - Đồ uống** | VNM, MSN, SAB, QNS |
| 35 | **Đá - Gạch - VLXD cao cấp** | VCS, CVT |
| 36 | **Xuất khẩu sản xuất** | PTB, VCS, DHC, HHP, ACG, TCM, MSH, TNI, TNG |

---

## CÁCH DÙNG

1. Người dùng đưa ticker (VD: `HPG`)
2. Tra bảng trên → dòng 7: **Thép**
3. Tra skill `fa` để biết phương pháp:
   - Dự báo doanh thu: Sản lượng × Giá bán (Commodity)
   - Định giá: EV/EBITDA + DCF (mid-cycle)
