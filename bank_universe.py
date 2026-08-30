#!/usr/bin/env python3
"""
bank_universe.py — Danh sách mã NGÂN HÀNG niêm yết/UPCoM DUY NHẤT dùng chung cho toàn bộ pipeline
(run_analysis.py phân loại ngành cho 1 mã lẻ, bank_system_risk.py tổng hợp rủi ro toàn hệ thống).
Trước 2026-08 danh sách này bị khai báo LẶP LẠI trực tiếp trong run_analysis.py (22 mã, THIẾU CTG —
VietinBank, 1 trong 4 ngân hàng quốc doanh lớn nhất — cùng SSB/BVB/PGB/VBB) — gộp về 1 hằng số DUY
NHẤT để 2 pipeline không bao giờ lệch danh sách nhau nữa. Nguồn: ảnh chụp dashboard tham chiếu của
user (2026-08-30, nhãn "26 mã").
"""

BANKING_TICKERS = frozenset({
    "VCB", "BID", "CTG", "TCB", "MBB", "LPB", "HDB", "STB", "ACB", "SHB",
    "SSB", "VIB", "MSB", "TPB", "EIB", "OCB", "NAB", "NVB", "ABB", "BAB",
    "VAB", "BVB", "PGB", "KLB", "SGB", "VBB",
})


def is_banking_ticker(ticker):
    return (ticker or "").upper() in BANKING_TICKERS
