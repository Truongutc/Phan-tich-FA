#!/usr/bin/env python3
"""
template_nhietdien.py — Universal calculation engine for Vietnamese thermal power
generation companies: POW, NT2, PPC, QTP.

Giai đoạn 1 (xem "Hướng dẫn Định giá Doanh nghiệp Nhóm Nhiệt điện Việt Nam" trong
Logic phan tich cac nganh/): định giá từ BCTC chuẩn (Vietcap) + giá nhiên liệu than/
khí/dầu/tỷ giá (cào qua curl, không AI). CHƯA cào sản lượng điện (MWh)/hệ số tải (%)
từ IR từng công ty/EVN — để dành giai đoạn 2 sau khi lõi định giá đã chạy ổn định
(giống HPG: nguồn sản lượng thép mất nhiều vòng lặp thực tế mới ổn định).

REE bị loại khỏi nhóm này (công ty đa ngành, không phải nhiệt điện thuần) — không xử
lý trong file này.

Định giá: DCF 50% + P/E 20% + P/B 15% + Asset-based 15% (khác KCN: 40%P/E+40%P/B+20%RI
— đây là DCF thật đầu tiên trong hệ thống, các template khác chưa có WACC/DCF nào).
Asset-based diễn giải = Book Value of Equity thuộc cổ đông mẹ (bsa78-bsa210), TƯƠNG
ĐƯƠNG "P/B sàn 1.0x" — không chép nguyên công thức ví dụ trong tài liệu hướng dẫn (ví
dụ gốc tự mâu thuẫn: trừ nợ 2 lần vì "giá trị sổ sách" trong ví dụ đã là VCSH, vốn đã
net nợ sẵn).

Output: Excel model + PDF report + data/<TICKER>.json (cho web dashboard nhietdien.html).
"""
import os
import sys

# Fix Windows console encoding (cp1252 không hỗ trợ tiếng Việt)
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
import re
import json
import datetime
import subprocess
import statistics as stats
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, white, grey, black
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import requests

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))

NHIETDIEN_COMPANY_NAMES = {
    "POW": "Tổng Công ty Điện lực Dầu khí Việt Nam - CTCP",
    "NT2": "CTCP Điện lực Dầu khí Nhơn Trạch 2",
    "PPC": "CTCP Nhiệt điện Phả Lại",
    "QTP": "CTCP Nhiệt điện Quảng Ninh",
}


# ══════════════════════════════════════════════════════════════════════════
# VIETNAMESE FONT REGISTRATION (copy nguyên từ template_kcn.py)
# ══════════════════════════════════════════════════════════════════════════
def register_vn_fonts():
    font_paths_to_try = [
        ("C:/Windows/Fonts/arial.ttf", "Arial"),
        ("C:/Windows/Fonts/arialbd.ttf", "Arial-Bold"),
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "Arial"),
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", "Arial-Bold"),
        ("/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf", "Arial"),
        ("/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf", "Arial-Bold"),
    ]
    found = {}
    for path, freg in font_paths_to_try:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont(freg, path))
                found[freg] = path
            except Exception:
                pass
    return found


_VN_FONTS = register_vn_fonts()
FONT_REG = 'Arial' if 'Arial' in _VN_FONTS else 'Helvetica'
FONT_BOLD = 'Arial-Bold' if 'Arial-Bold' in _VN_FONTS else 'Helvetica-Bold'


# ══════════════════════════════════════════════════════════════════════════
# CAPM INPUTS: Rf + Beta (COPY NGUYÊN VẸN từ template_kcn.py — không có module dùng
# chung, mỗi template tự copy khối này, xem plan)
# ══════════════════════════════════════════════════════════════════════════
UA_STR = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"


def fetch_via_curl(url, timeout=10, label=None):
    """Tải HTML/text qua curl subprocess (KHÔNG dùng thư viện requests) — investing.com
    chặn vân tay TLS của Python requests/urllib3 (Cloudflare 403) nhưng vẫn cho phép
    curl. Copy nguyên từ build_hpg_model.py — KHÔNG import trực tiếp file đó vì nó chạy
    toàn bộ pipeline HPG ngay khi import (side-effect ở module scope)."""
    tag = f" [{label}]" if label else ""
    try:
        r = subprocess.run(
            ["curl", "-sL", "-A", UA_STR, "--max-time", str(timeout), "-w", "\n__HTTP_STATUS__:%{http_code}", url],
            capture_output=True, text=True, encoding="utf-8", errors="ignore", timeout=timeout + 5,
        )
        out = r.stdout
        status = None
        marker = "\n__HTTP_STATUS__:"
        if marker in out:
            out, status_str = out.rsplit(marker, 1)
            status = status_str.strip()
        if r.returncode != 0 or not out.strip() or status not in (None, "200"):
            print(f"  [DIAG]{tag} fetch_via_curl weak/empty result: curl_exit={r.returncode} http_status={status} "
                  f"len={len(out)} url={url[:90]}" + (f" stderr={r.stderr[:150].strip()}" if r.stderr else ""))
        return out if r.returncode == 0 else ""
    except Exception as e:
        print(f"  [DIAG]{tag} fetch_via_curl exception: {url[:90]} -> {e}")
        return ""


def fetch_rf_vietnam(timeout=15):
    FALLBACK_RF = 0.045
    try:
        html = fetch_via_curl("https://vn.investing.com/rates-bonds/vietnam-10-year-bond-yield", timeout=timeout, label="investing-rf")
        if html:
            m = re.search(r'data-test="instrument-price-last"[^>]*>([\d.,]+)', html)
            if m:
                rf = float(m.group(1).replace(",", "")) / 100
                if 0.01 <= rf <= 0.15:
                    return rf, "investing.com"
    except Exception as e:
        print(f"  [WARN] investing.com Rf failed: {e}")
    try:
        r = requests.get("https://www.worldgovernmentbonds.com/bond-yield/vietnam/10-years/",
                          headers={"User-Agent": "Mozilla/5.0"}, timeout=timeout)
        if r.status_code == 200:
            matches = re.findall(r'(\d+\.\d+)%', r.text[:5000])
            if matches:
                rf = float(matches[0]) / 100
                if 0.01 <= rf <= 0.15:
                    return rf, "worldgovernmentbonds.com"
    except Exception as e:
        print(f"  [WARN] WorldGovernmentBonds Rf failed: {e}")
    return FALLBACK_RF, "Fallback (manual)"


def fetch_aligned_history(ticker, days=720, timeout=15):
    from_time = 1577836800
    to_time = 2000000000
    url_stock = f"https://dchart-api.vndirect.com.vn/dchart/history?symbol={ticker}&resolution=D&from={from_time}&to={to_time}"
    url_index = f"https://dchart-api.vndirect.com.vn/dchart/history?symbol=VNINDEX&resolution=D&from={from_time}&to={to_time}"
    headers = {
        "User-Agent": UA_STR, "Accept": "application/json, text/plain, */*",
        "Referer": "https://dchart.vndirect.com.vn/",
    }
    try:
        r_stock = requests.get(url_stock, headers=headers, timeout=timeout)
        r_index = requests.get(url_index, headers=headers, timeout=timeout)
        if r_stock.status_code == 200 and r_index.status_code == 200:
            d_stock, d_index = r_stock.json(), r_index.json()
            t_s, c_s = d_stock.get("t") or [], d_stock.get("c") or []
            t_m, c_m = d_index.get("t") or [], d_index.get("c") or []
            map_stock = {t_s[i]: c_s[i] for i in range(min(len(t_s), len(c_s))) if c_s[i] is not None and c_s[i] > 0}
            map_index = {t_m[i]: c_m[i] for i in range(min(len(t_m), len(c_m))) if c_m[i] is not None and c_m[i] > 0}
            common_t = sorted(list(set(map_stock.keys()) & set(map_index.keys())))
            aligned = []
            for t in common_t:
                date_str = datetime.datetime.fromtimestamp(t).strftime("%Y-%m-%d")
                p_s = map_stock[t]
                if p_s < 1000:
                    p_s = p_s * 1000
                aligned.append((date_str, p_s, map_index[t]))
            return aligned
    except Exception as e:
        print(f"[Beta Calc] Error fetching history: {e}")
    return []


def fetch_beta_vietstock(ticker, timeout=15):
    try:
        search_url = f"https://finance.vietstock.vn/search?query={ticker}"
        headers = {'User-Agent': UA_STR}
        r1 = requests.get(search_url, headers=headers, timeout=timeout)
        if r1.status_code == 200:
            data = json.loads(r1.text).get("data", "")
            target_url = ""
            for line in data.split('\r\n'):
                parts = line.split('|')
                if len(parts) >= 3 and parts[0].strip().upper() == ticker.upper():
                    target_url = parts[2]
                    break
            if target_url:
                r2 = requests.get(target_url, headers={'User-Agent': UA_STR, 'Referer': 'https://finance.vietstock.vn/'}, timeout=timeout)
                if r2.status_code == 200:
                    m = re.search(r'\"Beta\":\"([\d\.]+)\"', r2.text)
                    if m:
                        beta = float(m.group(1))
                        if 0.3 <= beta <= 2.5:
                            return beta
    except Exception as e:
        print(f"  [WARN] Vietstock scrape failed: {e}")
    return None


def fetch_beta_vietcap(ticker, timeout=15):
    try:
        url = f"https://trading.vietcap.com.vn/api/iq-insight-service/v1/company/details?ticker={ticker}"
        r = requests.get(url, headers={"User-Agent": UA_STR, "Referer": "https://trading.vietcap.com.vn/"}, timeout=timeout)
        if r.status_code == 200:
            d = r.json().get("data", {})
            beta = d.get("beta")
            if beta is not None and 0.3 <= float(beta) <= 2.5:
                return float(beta)
    except Exception:
        pass
    return None


def fetch_and_calc_beta(ticker, days=720, timeout=20, fallback=1.1):
    print(f"  [INFO] Đang tải lịch sử giá để tự tính Beta cho {ticker}...")
    aligned_data = fetch_aligned_history(ticker, days=days, timeout=timeout)
    latest_price = aligned_data[-1][1] if aligned_data else None
    num_sessions = len(aligned_data)
    web_beta = fetch_beta_vietstock(ticker, timeout) or fetch_beta_vietcap(ticker, timeout) or fallback
    calculated_beta, is_enough_sessions = fallback, False
    if num_sessions >= 30:
        sliced_data = aligned_data[-501:] if num_sessions > 500 else aligned_data
        s = [x[1] for x in sliced_data]
        m = [x[2] for x in sliced_data]
        rs = [(s[i] - s[i-1]) / s[i-1] for i in range(1, len(s))]
        rm = [(m[i] - m[i-1]) / m[i-1] for i in range(1, len(m))]
        n_ret = len(rs)
        mean_rs, mean_rm = sum(rs) / n_ret, sum(rm) / n_ret
        cov_sm = sum((rs[i]-mean_rs)*(rm[i]-mean_rm) for i in range(n_ret)) / (n_ret-1) if n_ret > 1 else 0
        var_m = sum((rm[i]-mean_rm)**2 for i in range(n_ret)) / (n_ret-1) if n_ret > 1 else 1.0
        calculated_beta = round(max(0.3, min(2.5, cov_sm / var_m if var_m > 0 else fallback)), 4)
        if num_sessions >= 250:
            is_enough_sessions = True
        aligned_data = aligned_data[-501:] if num_sessions > 500 else aligned_data
    beta_src = f"Tự tính toán ({num_sessions} phiên)" if is_enough_sessions else f"Web/API ({web_beta:.2f}) - lịch sử chỉ {num_sessions} phiên"
    return calculated_beta, web_beta, is_enough_sessions, beta_src, latest_price, aligned_data


# ══════════════════════════════════════════════════════════════════════════
# FIELD MAP VIETCAP (isa/bsa/cfa) — verify trực tiếp qua cache POW 2026-07
# CAPEX = cfa19 xác nhận CHÉO qua build_hpg_model.py + build_mwg_model.py (2 template
# khác đã dùng đúng field này), KHÔNG đoán mới — tránh lặp lại sai lầm dò field code
# một mình không kiểm chứng.
# ══════════════════════════════════════════════════════════════════════════
IS_GEN = {
    "revenue": "isa3",          # Doanh thu thuần
    "cogs": "isa4",             # Giá vốn hàng bán
    "gross_profit": "isa5",     # Lợi nhuận gộp
    "fin_income": "isa6",       # Doanh thu hoạt động tài chính
    "fin_expense": "isa7",      # Chi phí tài chính
    "interest_expense": "isa8", # Chi phí lãi vay
    "sga_sales": "isa9",        # Chi phí bán hàng
    "sga_admin": "isa10",       # Chi phí quản lý doanh nghiệp
    "operating_result": "isa11",
    "other_income_net": "isa14",
    "pbt": "isa16",             # Lãi/(lỗ) trước thuế
    "tax_current": "isa17",
    "tax_deferred": "isa18",
    "npat": "isa20",            # Lãi/(lỗ) thuần sau thuế
    "nci_income": "isa21",      # Lợi ích của cổ đông thiểu số (P&L)
    "npat_parent": "isa22",     # LNST của cổ đông công ty mẹ
    "eps_basic": "isa23",
}
BS_GEN = {
    "cash": "bsa2",              # Tiền và tương đương tiền
    "total_assets": "bsa53",     # TỔNG CỘNG TÀI SẢN
    "total_liab": "bsa54",       # NỢ PHẢI TRẢ
    "short_borrow": "bsa56",     # Vay ngắn hạn
    "long_borrow": "bsa71",      # Vay dài hạn
    "equity_total": "bsa78",     # VỐN CHỦ SỞ HỮU (đã bao gồm NCI)
    "charter_capital": "bsa80",  # Vốn góp
    "nci": "bsa210",             # Lợi ích cổ đông không kiểm soát (nested trong bsa78)
    "total_capital": "bsa96",    # TỔNG CỘNG NGUỒN VỐN
}
CF_GEN = {
    "depreciation": "cfa2",      # Khấu hao TSCĐ và BĐSĐT
    "capex": "cfa19",            # Tiền chi mua sắm/XD TSCĐ (âm) — verify chéo build_hpg_model.py/build_mwg_model.py
    "dividends_paid": "cfa32",   # Cổ tức, lợi nhuận đã trả cho chủ sở hữu (âm)
}


def _get_yr(records, year, field):
    for r in records:
        if r.get("yearReport") == year:
            v = r.get(field)
            if v is not None:
                return v / 1e9
    return 0.0


def _get_q(records, year, quarter, field):
    for r in records:
        if r.get("yearReport") == year and r.get("lengthReport") == quarter:
            v = r.get(field)
            if v is not None:
                return v / 1e9
    return 0.0


# ══════════════════════════════════════════════════════════════════════════
# ĐỊNH GIÁ: DCF 50% + P/E 20% + P/B 15% + Asset-based 15%
# (xem "Hướng dẫn Định giá..." mục 4-5 — đây là DCF THẬT đầu tiên trong hệ thống,
# tham khảo vòng lặp chiết khấu/terminal value của calc_valuation_kcn nhưng đổi từ
# Residual Income sang FCFF chiết khấu bằng WACC)
# ══════════════════════════════════════════════════════════════════════════
def _cagr_nhietdien(values):
    """CAGR từ dãy giá trị dương theo thời gian tăng dần, kẹp [-5%, +15%] — nhiệt điện
    tăng trưởng chậm & ổn định hơn KCN (KCN dùng [-10%,+20%]) vì sản lượng bị giới hạn
    bởi công suất lắp đặt cố định, không mở rộng quỹ đất/dự án như BĐS KCN."""
    vals = [v for v in values if v is not None and v > 0]
    if len(vals) < 2:
        return 0.02
    n = len(vals) - 1
    g = (vals[-1] / vals[0]) ** (1 / n) - 1
    return max(-0.05, min(0.15, g))


def _eps_parent(is_recs, year):
    """EPS cơ bản từ isa23 (VND/cp). Dùng isa22/shares làm fallback nếu isa23 = 0."""
    for r in is_recs:
        if r.get("yearReport") == year:
            eps = r.get("isa23")
            if eps is not None and eps != 0:
                return float(eps)
    return None


def _bvps_parent(bs_recs, year, shares):
    """BVPS thuộc về cổ đông công ty mẹ = (bsa78 - bsa210) / shares (VND/cp)."""
    for r in bs_recs:
        if r.get("yearReport") == year:
            equity_total = (r.get("bsa78") or 0)
            nci = (r.get("bsa210") or 0)
            vcsh_parent = equity_total - nci
            if vcsh_parent > 0 and shares > 0:
                return vcsh_parent / shares
    return None


def calc_wacc(rf, beta, erp, current_price, shares, debt_last, debt_prev, interest_expense, tax_rate=0.20):
    """WACC = tỷ trọng vốn CSH×COE + tỷ trọng nợ vay×COD×(1-thuế).
    COE = CAPM chuẩn (Rf + β×ERP, không cộng specific risk premium — dòng tiền PPA
    tương đối ổn định, khác CTCK biến động cao). COD = lãi vay / nợ vay bình quân,
    kẹp [2%, 12%] để tránh nhiễu khi nợ vay quá nhỏ/= 0 (vd PPC không vay nợ).
    Tất cả input tài chính đơn vị TỶ VND, current_price đơn vị VND/cp, shares đơn vị cp."""
    coe = rf + beta * erp
    avg_debt = (debt_last + debt_prev) / 2 if debt_prev else debt_last
    cod_raw = (interest_expense / avg_debt) if avg_debt > 0 else 0.0
    cod = max(0.02, min(0.12, cod_raw))
    market_cap = current_price * shares / 1e9  # tỷ VND
    total_cap = market_cap + debt_last
    if total_cap <= 0:
        return {"wacc": coe, "coe": coe, "cod": cod, "w_e": 1.0, "w_d": 0.0, "market_cap": market_cap}
    w_e = market_cap / total_cap
    w_d = debt_last / total_cap
    wacc_raw = w_e * coe + w_d * cod * (1 - tax_rate)
    # Sàn 8% (không phải 6%) — cổ phiếu nhiệt điện VN thanh khoản thấp thường cho beta đo
    # được RẤT thấp (vd QTP beta~0.36) khiến CAPM/WACC gần bằng Rf phi rủi ro, làm Terminal
    # Value nổ số phi lý (WACC-g_term quá mỏng). Sàn 8% ép ERP tối thiểu ~3.5-4% dù beta đo
    # được thấp — phát hiện thực tế 2026-07 khi test QTP ra upside +396% (đã sửa, xem
    # calc_dcf_valuation dùng thêm exit-multiple để tránh phụ thuộc 1 mình Gordon-growth).
    wacc = max(0.08, min(0.18, wacc_raw))
    return {"wacc": wacc, "coe": coe, "cod": cod, "w_e": w_e, "w_d": w_d, "market_cap": market_cap}


def project_fcf(hist_years, is_recs_y, cf_recs_y, n_fc=5, tax_rate=0.20):
    """Dự phóng FCFF n_fc năm tới. EBIT tính từ Lợi nhuận gộp - Chi phí bán hàng - Chi
    phí QLDN (isa5-isa9-isa10) — KHÔNG dùng isa11 (Lợi nhuận thuần HĐKD theo VAS) vì
    isa11 đã trừ cả doanh thu/chi phí tài chính (gồm lãi vay), trong khi FCFF cần EBIT
    THUẦN HOẠT ĐỘNG (chưa trừ lãi vay) — ảnh hưởng vốn vay đã nằm trong WACC, không
    được trừ trùng ở đây. Biên EBIT/D&A/CAPEX dự phóng = bình quân lịch sử (D&A và
    CAPEX không có driver vật lý rõ ràng ở Giai đoạn 1 nên neo theo % doanh thu)."""
    revenue_hist = [_get_yr(is_recs_y, y, IS_GEN["revenue"]) for y in hist_years]
    gp_hist = [_get_yr(is_recs_y, y, IS_GEN["gross_profit"]) for y in hist_years]
    sga_hist = [_get_yr(is_recs_y, y, IS_GEN["sga_sales"]) + _get_yr(is_recs_y, y, IS_GEN["sga_admin"]) for y in hist_years]
    ebit_hist = [gp_hist[i] - sga_hist[i] for i in range(len(hist_years))]
    da_hist = [_get_yr(cf_recs_y, y, CF_GEN["depreciation"]) for y in hist_years]
    capex_hist = [abs(_get_yr(cf_recs_y, y, CF_GEN["capex"])) for y in hist_years]

    rev_g = _cagr_nhietdien(revenue_hist)
    ebit_margins = [ebit_hist[i] / revenue_hist[i] for i in range(len(hist_years)) if revenue_hist[i] > 0]
    ebit_margin_fc = stats.mean(ebit_margins[-3:]) if len(ebit_margins) >= 2 else (ebit_margins[-1] if ebit_margins else 0.10)
    da_pct_hist = [da_hist[i] / revenue_hist[i] for i in range(len(hist_years)) if revenue_hist[i] > 0]
    da_pct_fc = stats.mean(da_pct_hist) if da_pct_hist else 0.08
    capex_pct_hist = [capex_hist[i] / revenue_hist[i] for i in range(len(hist_years)) if revenue_hist[i] > 0]
    capex_pct_fc = stats.mean(capex_pct_hist) if capex_pct_hist else 0.10

    fc_rows = []
    rev_t = revenue_hist[-1] if revenue_hist else 0.0
    for i in range(n_fc):
        rev_t = rev_t * (1 + rev_g)
        ebit_t = rev_t * ebit_margin_fc
        nopat_t = ebit_t * (1 - tax_rate)
        da_t = rev_t * da_pct_fc
        capex_t = rev_t * capex_pct_fc
        fcff_t = nopat_t + da_t - capex_t
        fc_rows.append({
            "year_offset": i + 1, "revenue": rev_t, "ebit": ebit_t, "nopat": nopat_t,
            "da": da_t, "capex": capex_t, "fcff": fcff_t,
        })
    assumptions = {
        "rev_g": rev_g, "ebit_margin_fc": ebit_margin_fc,
        "da_pct_fc": da_pct_fc, "capex_pct_fc": capex_pct_fc,
        "revenue_hist": revenue_hist, "ebit_hist": ebit_hist,
        "da_hist": da_hist, "capex_hist": capex_hist,
    }
    return fc_rows, assumptions


def calc_dcf_valuation(fc_rows, wacc, shares, net_debt, rev_g, exit_ev_ebitda=None):
    """PV(FCFF 5 năm) + Terminal Value chiết khấu về hiện tại ở WACC.

    Terminal Value = TRUNG BÌNH của 2 phương pháp độc lập, không chỉ dùng 1 mình
    Gordon-growth — phát hiện thực tế 2026-07 (test QTP): khi WACC thấp (beta đo được
    thấp do thanh khoản cổ phiếu kém) và mẫu số (WACC-g_term) quá mỏng, Gordon-growth
    một mình cho Terminal Value nổ số phi lý (QTP: TV chiếm 81% Enterprise Value, upside
    +396%). Neo thêm bằng EXIT-MULTIPLE (EBITDA năm cuối × EV/EBITDA mục tiêu — CÙNG bội
    số đã tính ở calc_ev_ebitda_valuation, neo vào thị trường thay vì thuần công thức
    toán) để 2 phương pháp tự kiểm chứng chéo lẫn nhau, giảm rủi ro lệch về 1 phía.
    g_term kẹp [0%, 3%] theo 40% tốc độ tăng trưởng doanh thu dự phóng — nhiệt điện là
    ngành trưởng thành, g dài hạn phải thấp hơn nhiều tăng trưởng ngắn hạn (cùng logic
    g_term của calc_valuation_kcn, chỉ đổi biên trên 4%->3% vì tăng trưởng nền thấp hơn).
    Mẫu số WACC-g_term kẹp tối thiểu 3% (không phải 1%) để tránh nổ số khi WACC sát sàn."""
    n = len(fc_rows)
    pv_sum = 0.0
    for row in fc_rows:
        pv = row["fcff"] / ((1 + wacc) ** row["year_offset"])
        row["fcff_pv"] = pv
        pv_sum += pv
    g_term = max(0.0, min(0.03, rev_g * 0.4))
    last_fcff = fc_rows[-1]["fcff"] if fc_rows else 0.0
    tv_gordon = last_fcff * (1 + g_term) / max(wacc - g_term, 0.03)

    tv_exit = None
    if exit_ev_ebitda and fc_rows:
        last_ebitda = fc_rows[-1]["ebit"] + fc_rows[-1]["da"]
        tv_exit = last_ebitda * exit_ev_ebitda
        terminal_value = (tv_gordon + tv_exit) / 2
    else:
        terminal_value = tv_gordon

    terminal_value_pv = terminal_value / ((1 + wacc) ** n) if n else terminal_value
    enterprise_value = pv_sum + terminal_value_pv
    equity_value = enterprise_value - net_debt
    fair_dcf = (equity_value * 1e9) / shares if shares > 0 else 0.0
    return {
        "pv_sum": pv_sum, "g_term": g_term, "tv_gordon": tv_gordon, "tv_exit": tv_exit,
        "terminal_value": terminal_value,
        "terminal_value_pv": terminal_value_pv, "enterprise_value": enterprise_value,
        "net_debt": net_debt, "equity_value": equity_value, "fair_dcf": fair_dcf,
    }


def calc_ev_ebitda_valuation(hist_years, is_recs_y, cf_recs_y, bs_recs_y, current_price, shares,
                              fc_rows, ev_low=4.5, ev_high=8.0):
    """Định giá theo EV/EBITDA lịch sử — thay cho P/E trong công thức blend (theo yêu cầu
    user 2026-07): nhóm nhiệt điện có KHẤU HAO và LÃI VAY rất lớn (nhà máy vốn đầu tư
    khủng, khấu hao nhanh; nhiều mã vay nợ lớn tài trợ dự án) nên LNST (dùng cho P/E) bị
    bóp méo mạnh bởi 2 khoản này — EBITDA (= EBIT + D&A, TRƯỚC lãi vay+khấu hao) phản ánh
    đúng hơn khả năng sinh tiền hoạt động thực. EV lịch sử = market cap HIỆN TẠI + net
    debt của TỪNG NĂM lịch sử (cùng quy ước với cách P/E-lịch sử/P/B-lịch sử trong
    template_kcn.py dùng current_price áp vào EPS/BVPS quá khứ — không có giá cổ phiếu
    lịch sử nên phải xấp xỉ vậy). Biên [4.5x, 8.0x] là giả định chung hợp lý cho ngành
    điện VN (không có số liệu ngành cụ thể trong tài liệu hướng dẫn — nếu có số liệu
    ngành thực tế sau này nên thay bằng biên đó)."""
    market_cap = current_price * shares / 1e9  # tỷ VND
    ebitda_hist_pairs = []
    for y in hist_years:
        revenue = _get_yr(is_recs_y, y, IS_GEN["revenue"])
        gp = _get_yr(is_recs_y, y, IS_GEN["gross_profit"])
        sga = _get_yr(is_recs_y, y, IS_GEN["sga_sales"]) + _get_yr(is_recs_y, y, IS_GEN["sga_admin"])
        da = _get_yr(cf_recs_y, y, CF_GEN["depreciation"])
        ebit = gp - sga
        ebitda = ebit + da
        debt_y = _get_yr(bs_recs_y, y, BS_GEN["short_borrow"]) + _get_yr(bs_recs_y, y, BS_GEN["long_borrow"])
        cash_y = _get_yr(bs_recs_y, y, BS_GEN["cash"])
        net_debt_y = debt_y - cash_y
        ev_y = market_cap + net_debt_y
        if ebitda > 0:
            ebitda_hist_pairs.append((y, ebitda, ev_y / ebitda))

    ev_ebitda_hist = [m for _, _, m in ebitda_hist_pairs]
    target_ev_ebitda = round(max(ev_low, min(ev_high, stats.median(ev_ebitda_hist))), 2) if ev_ebitda_hist else round((ev_low + ev_high) / 2, 2)

    ebitda_fc1 = (fc_rows[0]["ebit"] + fc_rows[0]["da"]) if fc_rows else (ebitda_hist_pairs[-1][1] if ebitda_hist_pairs else 0.0)
    debt_last = _get_yr(bs_recs_y, hist_years[-1], BS_GEN["short_borrow"]) + _get_yr(bs_recs_y, hist_years[-1], BS_GEN["long_borrow"])
    cash_last = _get_yr(bs_recs_y, hist_years[-1], BS_GEN["cash"])
    net_debt_last = debt_last - cash_last

    implied_ev = target_ev_ebitda * ebitda_fc1
    implied_equity = implied_ev - net_debt_last
    fair_ev_ebitda = (implied_equity * 1e9) / shares if shares > 0 else 0.0

    return {
        "ebitda_hist_pairs": ebitda_hist_pairs, "target_ev_ebitda": target_ev_ebitda,
        "ebitda_fc1": ebitda_fc1, "implied_ev": implied_ev, "net_debt_last": net_debt_last,
        "implied_equity": implied_equity, "fair_ev_ebitda": fair_ev_ebitda,
    }


def calc_valuation_nhietdien(ticker, is_recs_y, bs_recs_y, cf_recs_y, hist_years, shares,
                              current_price, rf, beta, erp=0.07, tax_rate=0.20, n_fc=5):
    """Entry point tính toán định giá: DCF 50% + EV/EBITDA 20% + P/B 15% + Asset-based 15%.
    (Đổi từ P/E sang EV/EBITDA theo yêu cầu user — xem docstring calc_ev_ebitda_valuation:
    D&A/lãi vay của nhóm nhiệt điện quá lớn khiến LNST/P/E bị méo). P/E vẫn được tính và
    trả về làm THAM CHIẾU hiển thị (không nằm trong trọng số blend).
    Asset-based diễn giải = Book Value of Equity mẹ hiện tại (bsa78-bsa210)/shares —
    KHÔNG chép công thức ví dụ trong tài liệu hướng dẫn (tự mâu thuẫn, trừ nợ 2 lần).
    Trả về dict đủ để ghi Excel/JSON."""
    # --- WACC ---
    debt_last = _get_yr(bs_recs_y, hist_years[-1], BS_GEN["short_borrow"]) + _get_yr(bs_recs_y, hist_years[-1], BS_GEN["long_borrow"])
    debt_prev = 0.0
    if len(hist_years) >= 2:
        debt_prev = _get_yr(bs_recs_y, hist_years[-2], BS_GEN["short_borrow"]) + _get_yr(bs_recs_y, hist_years[-2], BS_GEN["long_borrow"])
    interest_expense = abs(_get_yr(is_recs_y, hist_years[-1], IS_GEN["interest_expense"]))
    wacc_info = calc_wacc(rf, beta, erp, current_price, shares, debt_last, debt_prev, interest_expense, tax_rate)

    # --- Dự phóng FCFF (dùng chung cho cả DCF và EV/EBITDA forward) ---
    fc_rows, fc_assump = project_fcf(hist_years, is_recs_y, cf_recs_y, n_fc=n_fc, tax_rate=tax_rate)
    cash_last = _get_yr(bs_recs_y, hist_years[-1], BS_GEN["cash"])
    net_debt = debt_last - cash_last

    # --- EV/EBITDA (thay P/E trong blend) — tính TRƯỚC DCF vì Terminal Value của DCF cần
    # target_ev_ebitda làm neo exit-multiple (xem calc_dcf_valuation) ---
    ev_ebitda_result = calc_ev_ebitda_valuation(hist_years, is_recs_y, cf_recs_y, bs_recs_y, current_price, shares, fc_rows)

    # --- DCF (Terminal Value neo cả Gordon-growth lẫn exit-multiple EV/EBITDA) ---
    dcf_result = calc_dcf_valuation(fc_rows, wacc_info["wacc"], shares, net_debt, fc_assump["rev_g"],
                                     exit_ev_ebitda=ev_ebitda_result["target_ev_ebitda"])

    # --- P/E (chỉ để tham chiếu hiển thị) & P/B (nằm trong blend) ---
    eps_vals = [(y, _eps_parent(is_recs_y, y)) for y in hist_years]
    bvps_vals = [(y, _bvps_parent(bs_recs_y, y, shares)) for y in hist_years]
    eps_valid = [(y, v) for y, v in eps_vals if v is not None and v > 0]
    bvps_valid = [(y, v) for y, v in bvps_vals if v is not None and v > 0]
    eps_last = eps_valid[-1][1] if eps_valid else 1000.0
    bvps_last = bvps_valid[-1][1] if bvps_valid else 10000.0

    pe_hist = [current_price / eps for _, eps in eps_valid if current_price > 0 and eps > 0]
    target_pe = round(max(8.0, min(12.0, stats.median(pe_hist))), 1) if pe_hist else 10.0
    pb_hist = [current_price / bvps for _, bvps in bvps_valid if current_price > 0 and bvps > 0]
    target_pb = round(max(0.8, min(1.2, stats.median(pb_hist))), 2) if pb_hist else 1.0

    eps_cagr = _cagr_nhietdien([v for _, v in eps_valid[-3:]]) if len(eps_valid) >= 2 else 0.02
    eps_fc1 = eps_last * (1 + eps_cagr)
    fair_pe = target_pe * eps_fc1  # tham chiếu, không dùng trong blend
    fair_pb = target_pb * bvps_last

    # --- Asset-based (= Book Value of Equity mẹ hiện tại, xem docstring) ---
    fair_asset = bvps_last

    # --- Blend: DCF 50% + EV/EBITDA 20% + P/B 15% + Asset 15% ---
    fair_dcf = dcf_result["fair_dcf"]
    fair_ev_ebitda = ev_ebitda_result["fair_ev_ebitda"]
    fair_blend = 0.50 * fair_dcf + 0.20 * fair_ev_ebitda + 0.15 * fair_pb + 0.15 * fair_asset
    upside = (fair_blend - current_price) / current_price if current_price > 0 else 0.0

    return {
        "rf": rf, "beta": beta, "erp": erp,
        "wacc": wacc_info, "fc_rows": fc_rows, "fc_assumptions": fc_assump, "dcf": dcf_result,
        "ev_ebitda": ev_ebitda_result,
        "target_pe": target_pe, "target_pb": target_pb, "eps_last": eps_last, "bvps_last": bvps_last,
        "eps_cagr": eps_cagr, "eps_fc1": eps_fc1,
        "fair_dcf": fair_dcf, "fair_ev_ebitda": fair_ev_ebitda, "fair_pe": fair_pe,
        "fair_pb": fair_pb, "fair_asset": fair_asset,
        "fair_blend": fair_blend, "current_price": current_price, "upside": upside,
    }


if __name__ == "__main__":
    # Smoke-test Bước 1+2 của kế hoạch: fetch BCTC + Rf/Beta thật, in số liệu thô +
    # kết quả định giá WACC/DCF/blend để đối chiếu với ví dụ tính tay trong tài liệu
    # hướng dẫn mục 5.3 (không kỳ vọng khớp tuyệt đối vì input thực tế khác giả định,
    # nhưng thứ tự độ lớn phải hợp lý). Sẽ thay bằng CLI thật ở bước cuối cùng.
    from fetch_data import fetch_all
    t = sys.argv[1].upper() if len(sys.argv) > 1 else "POW"
    raw = fetch_all(t, use_cache=True)
    is_y = raw["sections"]["INCOME_STATEMENT"].get("years", [])
    bs_y = raw["sections"]["BALANCE_SHEET"].get("years", [])
    cf_y = raw["sections"]["CASH_FLOW"].get("years", [])
    hist_years = sorted({r["yearReport"] for r in is_y if r.get("yearReport")})[-5:]
    current_price = raw.get("currentPrice") or 0

    shares = 0
    bs_sorted_desc = sorted(bs_y, key=lambda r: r.get("yearReport", 0), reverse=True)
    for r in bs_sorted_desc:
        cap = r.get(BS_GEN["charter_capital"])
        if cap and cap > 0:
            shares = int(cap / 10_000)
            break
    if shares <= 0:
        shares_api = raw.get("numberOfSharesMktCap", 0)
        shares = int(shares_api) if shares_api > 0 else 100_000_000

    print(f"\n{t} — {raw.get('companyName')}")
    print(f"Current price: {current_price} | Shares: {shares:,}")
    print(f"Hist years: {hist_years}")
    for y in hist_years:
        rev = _get_yr(is_y, y, IS_GEN["revenue"])
        cogs = _get_yr(is_y, y, IS_GEN["cogs"])
        npat = _get_yr(is_y, y, IS_GEN["npat_parent"])
        da = _get_yr(cf_y, y, CF_GEN["depreciation"])
        capex = abs(_get_yr(cf_y, y, CF_GEN["capex"]))
        equity = _get_yr(bs_y, y, BS_GEN["equity_total"]) - _get_yr(bs_y, y, BS_GEN["nci"])
        debt = _get_yr(bs_y, y, BS_GEN["short_borrow"]) + _get_yr(bs_y, y, BS_GEN["long_borrow"])
        int_exp = _get_yr(is_y, y, IS_GEN["interest_expense"])
        print(f"  {y}: DT={rev:8.1f} GVHB={cogs:8.1f} LNST_me={npat:7.1f} D&A={da:6.1f} "
              f"CAPEX={capex:7.1f} VCSH_me={equity:8.1f} No_vay={debt:7.1f} LaiVay={int_exp:5.1f}")

    print("\n[Fetching Rf/Beta...]")
    rf, rf_src = fetch_rf_vietnam()
    calc_beta, web_beta, is_enough, beta_src, _, _ = fetch_and_calc_beta(t)
    beta = calc_beta if is_enough else web_beta
    print(f"Rf={rf*100:.2f}% ({rf_src}) | Beta={beta:.3f} ({beta_src})")

    val = calc_valuation_nhietdien(t, is_y, bs_y, cf_y, hist_years, shares, current_price, rf, beta)
    w = val["wacc"]
    print(f"\nWACC={w['wacc']*100:.2f}% (COE={w['coe']*100:.2f}%, COD={w['cod']*100:.2f}%, "
          f"w_e={w['w_e']*100:.1f}%, w_d={w['w_d']*100:.1f}%)")
    print("FCFF dự phóng 5 năm (tỷ VND):")
    for row in val["fc_rows"]:
        print(f"  Năm +{row['year_offset']}: DT={row['revenue']:8.1f} EBIT={row['ebit']:7.1f} "
              f"NOPAT={row['nopat']:7.1f} D&A={row['da']:6.1f} CAPEX={row['capex']:6.1f} "
              f"FCFF={row['fcff']:7.1f} PV={row['fcff_pv']:7.1f}")
    d = val["dcf"]
    print(f"Terminal Value: Gordon={d['tv_gordon']:.0f} | Exit-multiple={d['tv_exit'] if d['tv_exit'] is None else round(d['tv_exit'])} "
          f"-> TB={d['terminal_value']:.0f} (g_term={d['g_term']*100:.1f}%)")
    print(f"PV(FCFF)={d['pv_sum']:.0f} + TV_PV={d['terminal_value_pv']:.0f} "
          f"= EV={d['enterprise_value']:.0f} - NetDebt={d['net_debt']:.0f} = EquityValue={d['equity_value']:.0f} tỷ VND")
    ev = val["ev_ebitda"]
    print(f"\nEV/EBITDA lịch sử: {[(y, round(m,2)) for y, _, m in ev['ebitda_hist_pairs']]} "
          f"-> target={ev['target_ev_ebitda']}x | EBITDA dự phóng năm+1={ev['ebitda_fc1']:.0f} tỷ VND")
    print(f"Fair DCF       = {val['fair_dcf']:,.0f} VND/cp  (trọng số 50%)")
    print(f"Fair EV/EBITDA = {val['fair_ev_ebitda']:,.0f} VND/cp  (trọng số 20%, target {ev['target_ev_ebitda']}x)")
    print(f"Fair P/B       = {val['fair_pb']:,.0f} VND/cp  (trọng số 15%, target P/B={val['target_pb']}x)")
    print(f"Fair Asset     = {val['fair_asset']:,.0f} VND/cp  (trọng số 15%)")
    print(f"Fair P/E (chỉ tham chiếu, KHÔNG trong blend) = {val['fair_pe']:,.0f} VND/cp (target P/E={val['target_pe']}x)")
    print(f"=> FAIR BLEND = {val['fair_blend']:,.0f} VND/cp | Giá hiện tại = {current_price:,.0f} | "
          f"Upside = {val['upside']*100:+.1f}%")
