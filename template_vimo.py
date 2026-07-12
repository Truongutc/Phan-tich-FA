#!/usr/bin/env python3
"""
template_vimo.py — Phân tích Vĩ mô Kinh tế Việt Nam (không gắn với mã cổ phiếu nào).

Dựa theo "Hướng Dẫn Phân Tích Vĩ Mô Kinh Tế Việt Nam & Ứng Dụng Đầu Tư Chứng Khoán" trong
Logic phan tich cac nganh/ — khung Top-Down 7 chương: theo dõi ~20 chỉ báo (tăng trưởng, lạm
phát, tiền tệ, thương mại/vốn, tài khóa, lao động) + áp lực bên ngoài (Fed, DXY, dầu, Trung
Quốc) + định giá thị trường (P/E, P/B, ERP VN-Index), tổng hợp thành Scorecard Vĩ Mô 5 nhóm
(-1/0/+1 mỗi nhóm) và ma trận quyết định phân bổ vốn (Chương 6).

Dữ liệu THÔ nằm ở data/vimo_raw.json (được fetch_macro_data.py cập nhật) — file này CHỈ tính
toán/trình bày (trend, scorecard, ERP, ma trận quyết định, PDF, JSON dashboard), KHÔNG tự fetch
dữ liệu thô.

KHÔNG có Excel (theo yêu cầu user — chỉ Excel cho các sector template theo ticker).

Giai đoạn 1 (skeleton sơ bộ, theo chỉ đạo user "hoàn thiện khung trước, tinh chỉnh sau"): một số
chỉ báo (PMI, lãi suất liên ngân hàng, tín dụng, M2, dự trữ ngoại hối, dòng tiền khối ngoại/margin)
mới chỉ có seed dữ liệu thưa hoặc chưa có nguồn tự động — xem "auto_source" trong vimo_raw.json.
"""
import os
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

import re
import json
import datetime
import subprocess
import requests

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, white
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
VIMO_RAW_PATH = os.path.join(PROJECT_ROOT, "data", "vimo_raw.json")


# ══════════════════════════════════════════════════════════════════════════
# FONTS (copy pattern từ template_nhietdien.py — Arial nếu có, else Helvetica)
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

UA_STR = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
          "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")


# ══════════════════════════════════════════════════════════════════════════
# Rf (copy từ template_nhietdien.py — cần cho ERP = E/P - Rf)
# ══════════════════════════════════════════════════════════════════════════
def fetch_via_curl(url, timeout=10, label=None):
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
            print(f"  [DIAG]{tag} fetch_via_curl weak/empty: curl_exit={r.returncode} http_status={status} url={url[:90]}")
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


# ══════════════════════════════════════════════════════════════════════════
# LOAD DATA
# ══════════════════════════════════════════════════════════════════════════
def load_vimo_raw():
    with open(VIMO_RAW_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


GROUP_LABELS = {
    "growth": "Tăng trưởng",
    "inflation": "Lạm phát",
    "monetary": "Tiền tệ & Lãi suất",
    "trade": "Thương mại & Vốn",
    "fiscal": "Tài khóa",
    "labor": "Lao động",
    "external": "Áp lực bên ngoài",
    "market": "Thị trường chứng khoán",
}

# Ánh xạ 8 nhóm dữ liệu (vimo_raw.json) -> 5 nhóm Scorecard Vĩ Mô theo đúng Chương 4.1/6.1 tài
# liệu hướng dẫn (KHÔNG trùng 1-1 với 8 nhóm dữ liệu — vd tỷ giá thuộc "monetary" trong dữ liệu
# nhưng thuộc "Bên ngoài" trong scorecard, theo đúng cách tài liệu phân loại).
SCORECARD_GROUPS = {
    "Tăng trưởng": ["gdp_growth", "iip_growth", "pmi_manufacturing", "retail_sales_growth", "unemployment_rate"],
    "Lạm phát & Lãi suất": ["cpi_yoy", "core_inflation", "refinancing_rate", "interbank_rate_3m"],
    "Thanh khoản": ["credit_growth", "m2_growth", "forex_reserves"],
    "Bên ngoài": ["trade_balance", "fdi_disbursed", "usdvnd", "fed_funds_rate", "brent_oil", "dxy_proxy", "china_gdp_growth"],
    "Tâm lý thị trường": ["vnindex_pe", "vnindex_pb"],
}


# ══════════════════════════════════════════════════════════════════════════
# TREND — so latest vs kỳ liền trước + vs kỳ xa hơn (tối đa 4 kỳ trước) theo hướng "tốt" riêng
# của từng chỉ báo.
# ══════════════════════════════════════════════════════════════════════════
def calc_trend(series, good_direction):
    """series: list [{period, value, ...}] đã sort theo thời gian tăng dần (đúng thứ tự lưu
    trong vimo_raw.json). Trả dict: latest, prior, delta, pct_of_prior, label ('up'/'down'/
    'flat'), arrow, is_improving (theo good_direction), n_points."""
    valid = [p for p in series if p.get("value") is not None]
    if not valid:
        return {"latest": None, "n_points": 0, "label": "no_data", "arrow": "—"}
    latest = valid[-1]
    n = len(valid)
    result = {
        "latest": latest["value"], "latest_period": latest["period"],
        "latest_source": latest.get("source_url"), "n_points": n,
    }
    if n < 2:
        result.update({"label": "insufficient", "arrow": "—", "is_improving": None})
        return result

    prior = valid[-2]
    delta = latest["value"] - prior["value"]
    result["prior"] = prior["value"]
    result["prior_period"] = prior["period"]
    result["delta"] = round(delta, 4)

    # So thêm với kỳ xa hơn (tối đa 4 kỳ trước) nếu đủ dữ liệu — cho cái nhìn xu hướng dài hơn
    if n >= 5:
        far = valid[-5]
        result["delta_vs_4back"] = round(latest["value"] - far["value"], 4)
        result["far_period"] = far["period"]

    threshold = abs(prior["value"]) * 0.003 if prior["value"] else 0.01  # ngưỡng "flat" ~0.3%
    if abs(delta) < max(threshold, 1e-6):
        direction = "flat"
    elif delta > 0:
        direction = "up"
    else:
        direction = "down"
    result["direction"] = direction
    # value_arrow = CHIỀU THẬT của số liệu (↑ nghĩa đen là tăng, ↓ là giảm) — KHÔNG mang nghĩa
    # tốt/xấu, để tránh nhầm lẫn (vd Nợ công/GDP giảm là điều TỐT nhưng số liệu vẫn đi xuống).
    result["value_arrow"] = {"up": "↑", "down": "↓", "flat": "→"}[direction]

    if direction == "flat":
        is_improving = None
    elif good_direction == "higher":
        is_improving = direction == "up"
    else:  # "lower"
        is_improving = direction == "down"
    result["is_improving"] = is_improving

    # judgment_* = ĐÁNH GIÁ tốt lên/xấu đi (theo đúng yêu cầu user) — TÁCH RIÊNG khỏi value_arrow
    # ở trên để không lẫn "chiều số liệu" với "ý nghĩa tốt/xấu" của chiều đó.
    if is_improving is None:
        result["label"] = "flat"
        result["arrow"] = "→"
        result["judgment_label"] = "Ổn định"
        result["judgment_color"] = "#f59e0b"
    elif is_improving:
        result["label"] = "improving"
        result["arrow"] = "▲"
        result["judgment_label"] = "Tốt lên"
        result["judgment_color"] = "#10b981"
    else:
        result["label"] = "worsening"
        result["arrow"] = "▼"
        result["judgment_label"] = "Xấu đi"
        result["judgment_color"] = "#ef4444"
    return result


# ══════════════════════════════════════════════════════════════════════════
# SCORECARD VĨ MÔ — 5 nhóm, mỗi nhóm -1/0/+1 (Chương 4.1/6.1)
# ══════════════════════════════════════════════════════════════════════════
def calc_scorecard(raw, trends):
    """trends: {indicator_key: calc_trend(...) result}. Trả {group_name: {score, detail, n}}."""
    scorecard = {}
    for group_name, keys in SCORECARD_GROUPS.items():
        votes = []
        detail = []
        for k in keys:
            t = trends.get(k)
            if not t or t.get("is_improving") is None:
                continue
            votes.append(1 if t["is_improving"] else -1)
            detail.append({"indicator": k, "label": raw[k]["label"], "vote": votes[-1],
                            "arrow": t["arrow"], "judgment_label": t["judgment_label"]})
        if not votes:
            score = 0
        else:
            avg = sum(votes) / len(votes)
            score = 1 if avg > 0.2 else (-1 if avg < -0.2 else 0)
        scorecard[group_name] = {"score": score, "detail": detail, "n_votes": len(votes)}
    total = sum(g["score"] for g in scorecard.values())
    return scorecard, total


# ══════════════════════════════════════════════════════════════════════════
# ĐỊNH GIÁ THỊ TRƯỜNG — ERP = E/P (từ VN-Index P/E) - Rf (Chương 5.2/6.1)
# ══════════════════════════════════════════════════════════════════════════
def calc_market_valuation(raw, rf):
    pe_series = raw["vnindex_pe"]["series"]
    pb_series = raw["vnindex_pb"]["series"]
    pe = pe_series[-1]["value"] if pe_series else None
    pb = pb_series[-1]["value"] if pb_series else None
    erp = None
    valuation_label = "Không xác định"
    if pe and pe > 0:
        earnings_yield = 1 / pe
        erp = earnings_yield - rf
        if erp > 0.03:
            valuation_label = "Rẻ/Hấp dẫn"
        elif erp < 0.01:
            valuation_label = "Đắt/Kém hấp dẫn"
        else:
            valuation_label = "Hợp lý"
    return {
        "pe": pe, "pb": pb, "rf": rf, "erp": erp, "valuation_label": valuation_label,
        "pe_source": pe_series[-1]["source_url"] if pe_series else None,
    }


# ══════════════════════════════════════════════════════════════════════════
# MA TRẬN QUYẾT ĐỊNH PHÂN BỔ VỐN (Chương 6.2) — 2 trục thật (Scorecard, Định giá); trục "Dòng
# tiền & Tâm lý" (khối ngoại/margin) CHƯA có nguồn dữ liệu tự động trong Giai đoạn 1 — ghi chú
# rõ, không tự bịa số dòng tiền.
# ══════════════════════════════════════════════════════════════════════════
def calc_decision_matrix(scorecard_total, valuation_label):
    macro_good = scorecard_total > 0
    if macro_good:
        if valuation_label == "Rẻ/Hấp dẫn":
            return "Bung vốn mạnh", "Vĩ mô tốt + định giá rẻ — tăng tỷ trọng cổ phiếu, ưu tiên ngành hưởng lợi từ vĩ mô."
        elif valuation_label == "Đắt/Kém hấp dẫn":
            return "Giảm tỷ trọng, chờ điều chỉnh", "Vĩ mô tốt nhưng định giá đã đắt — chốt lời một phần, chờ cơ hội mua lại giá tốt hơn."
        else:
            return "Duy trì, chọn lọc", "Vĩ mô tốt, định giá hợp lý — giữ tỷ trọng hiện tại, chọn cổ phiếu nền tảng tốt."
    else:
        if valuation_label == "Rẻ/Hấp dẫn":
            return "Mua từ từ, phân kỳ", "Vĩ mô xấu nhưng định giá đã rẻ — mua tích lũy dần, tỷ trọng nhỏ, chiến lược dài hạn."
        elif valuation_label == "Đắt/Kém hấp dẫn":
            return "Giảm vốn mạnh", "Vĩ mô xấu + định giá đắt — cắt giảm tỷ trọng cổ phiếu, ưu tiên tiền mặt."
        else:
            return "Phòng thủ, giữ tiền mặt", "Vĩ mô xấu, định giá hợp lý — ưu tiên bảo toàn vốn, tăng tỷ trọng tiền mặt/tài sản phòng thủ."


# ══════════════════════════════════════════════════════════════════════════
# CHARTS — 1 chart/chỉ báo cho các chỉ báo có ≥4 điểm dữ liệu (đủ để vẽ đường xu hướng có ý
# nghĩa); chỉ báo thưa dữ liệu hơn hiển thị dạng bảng trong PDF/dashboard, không ép vẽ chart.
# ══════════════════════════════════════════════════════════════════════════
def build_charts_vimo(out_dir, raw, min_points=4):
    os.makedirs(out_dir, exist_ok=True)
    charts = {}
    plt.rcParams["font.family"] = "DejaVu Sans"
    for key, ind in raw.items():
        if key == "_meta":
            continue
        valid = [p for p in ind["series"] if p.get("value") is not None]
        if len(valid) < min_points:
            continue
        periods = [p["period"] for p in valid]
        values = [p["value"] for p in valid]
        improving = (values[-1] >= values[0]) if ind["good_direction"] == "higher" else (values[-1] <= values[0])
        color = "#10b981" if improving else "#ef4444"

        fig, ax = plt.subplots(figsize=(7.5, 3.4))
        ax.plot(periods, values, marker="o", color=color, linewidth=2)
        ax.fill_between(periods, values, alpha=0.08, color=color)
        ax.set_title(f"{ind['label']} ({ind['unit']})", fontsize=11, fontweight="bold")
        ax.tick_params(axis="x", rotation=45, labelsize=8)
        ax.grid(alpha=0.25)
        fig.tight_layout()
        path = os.path.join(out_dir, f"vimo_{key}.png")
        fig.savefig(path, dpi=130)
        plt.close(fig)
        charts[key] = path
    return charts


# ══════════════════════════════════════════════════════════════════════════
# PDF
# ══════════════════════════════════════════════════════════════════════════
def build_pdf_vimo(pdf_path, raw, trends, scorecard, scorecard_total, valuation, decision_label,
                    decision_text, charts):
    doc = SimpleDocTemplate(pdf_path, pagesize=A4, rightMargin=15 * mm, leftMargin=15 * mm,
                             topMargin=15 * mm, bottomMargin=15 * mm)
    styles = getSampleStyleSheet()
    title_st = ParagraphStyle("VM_Title", parent=styles["Heading1"], fontName=FONT_BOLD, fontSize=18,
                               leading=22, textColor=HexColor("#1F4E78"), spaceAfter=12)
    h1_st = ParagraphStyle("VM_H1", parent=styles["Heading2"], fontName=FONT_BOLD, fontSize=13,
                            leading=17, textColor=HexColor("#2E75B6"), spaceBefore=14, spaceAfter=7)
    h2_st = ParagraphStyle("VM_H2", parent=styles["Heading3"], fontName=FONT_BOLD, fontSize=11,
                            leading=15, textColor=HexColor("#404040"), spaceBefore=8, spaceAfter=4)
    body_st = ParagraphStyle("VM_Body", parent=styles["Normal"], fontName=FONT_REG, fontSize=10,
                              leading=14, textColor=HexColor("#2D3748"), spaceAfter=6)
    italic_st = ParagraphStyle("VM_Italic", parent=styles["Normal"], fontName=FONT_REG, fontSize=8,
                                leading=11, textColor=HexColor("#718096"), italic=True)
    small_st = ParagraphStyle("VM_Small", parent=styles["Normal"], fontName=FONT_REG, fontSize=8,
                               leading=11, textColor=HexColor("#4A5568"))

    BLUE_DARK = HexColor("#1F4E78")
    LIGHT_BLUE = HexColor("#DDEBF7")

    def tbl_style(header_col=BLUE_DARK, alt_col=LIGHT_BLUE, font_size=9):
        return TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), header_col), ("TEXTCOLOR", (0, 0), (-1, 0), white),
            ("FONTNAME", (0, 0), (-1, 0), FONT_BOLD), ("FONTSIZE", (0, 0), (-1, -1), font_size),
            ("FONTNAME", (0, 1), (-1, -1), FONT_REG), ("ROWBACKGROUNDS", (0, 1), (-1, -1), [white, alt_col]),
            ("ALIGN", (1, 0), (-1, -1), "CENTER"), ("ALIGN", (0, 0), (0, -1), "LEFT"),
            ("GRID", (0, 0), (-1, -1), 0.4, HexColor("#CBD5E1")),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4), ("TOPPADDING", (0, 0), (-1, -1), 4),
        ])

    story = []
    today_str = datetime.datetime.now().strftime("%d/%m/%Y")
    story.append(Paragraph("PHÂN TÍCH VĨ MÔ KINH TẾ VIỆT NAM", title_st))
    story.append(Paragraph(f"Khung phân tích Top-Down — Scorecard Vĩ Mô & Ma trận Quyết định Phân bổ Vốn | "
                            f"Ngày lập: {today_str}", body_st))
    story.append(Spacer(1, 8))

    scorecard_color = "#10b981" if scorecard_total > 0 else ("#ef4444" if scorecard_total < 0 else "#f59e0b")
    summary_data = [
        ["Scorecard Vĩ Mô (tổng)", "Định giá thị trường", "ERP", "Khuyến nghị"],
        [f"{scorecard_total:+d} / 5", valuation["valuation_label"],
         f"{valuation['erp']*100:.2f}%" if valuation["erp"] is not None else "N/A", decision_label],
    ]
    t_sum = Table(summary_data, colWidths=[42 * mm, 42 * mm, 32 * mm, 55 * mm])
    t_sum.setStyle(tbl_style())
    story.append(t_sum)
    story.append(Spacer(1, 6))
    story.append(Paragraph(f"<b>Khuyến nghị phân bổ vốn:</b> {decision_text}", body_st))
    story.append(Spacer(1, 10))

    # ── Scorecard chi tiết 5 nhóm ──
    story.append(Paragraph("1. Scorecard Vĩ Mô — 5 nhóm chỉ số (Chương 4.1/6.1)", h1_st))
    sc_rows = [["Nhóm", "Điểm", "Số chỉ báo có xu hướng rõ", "Chi tiết"]]
    for gname, gdata in scorecard.items():
        score_txt = {1: "+1 (Tốt)", 0: "0 (Trung tính)", -1: "-1 (Xấu)"}[gdata["score"]]
        detail_txt = ", ".join(f"{d['label']} ({d['judgment_label']})" for d in gdata["detail"]) or "Chưa đủ dữ liệu"
        sc_rows.append([gname, score_txt, str(gdata["n_votes"]), detail_txt])
    t_sc = Table(sc_rows, colWidths=[32 * mm, 26 * mm, 26 * mm, 85 * mm])
    t_sc.setStyle(tbl_style())
    story.append(t_sc)
    story.append(Spacer(1, 10))

    # ── Từng nhóm chỉ báo chi tiết + chart ──
    story.append(Paragraph("2. Chi tiết theo từng nhóm chỉ báo (Chương 3)", h1_st))
    groups_order = ["growth", "inflation", "monetary", "trade", "fiscal", "labor", "external", "market"]
    for grp in groups_order:
        keys = [k for k, v in raw.items() if k != "_meta" and v["group"] == grp]
        if not keys:
            continue
        story.append(Paragraph(f"2.{groups_order.index(grp)+1}. {GROUP_LABELS[grp]}", h2_st))
        rows = [["Chỉ báo", "Giá trị mới nhất", "Kỳ", "Số liệu", "Đánh giá", "Nguồn cập nhật"]]
        for k in keys:
            t = trends.get(k, {})
            ind = raw[k]
            src_label = {"worldbank": "World Bank API", "imf": "IMF DataMapper API", "fred": "FRED API",
                         "fx_api": "exchangerate-api.com", "pe_ratio_api": "worldperatio.com",
                         "nso_scrape": "nso.gov.vn (tự động)", "manual": "Nghiên cứu thủ công"}.get(ind["auto_source"], ind["auto_source"])
            val_txt = f"{t['latest']:.2f} {ind['unit']}" if t.get("latest") is not None else "N/A"
            # value_arrow = chiều số liệu THẬT (↑/↓/→), judgment_label = đánh giá tốt lên/xấu đi
            # riêng biệt (vd Nợ công/GDP giảm ↓ nhưng đánh giá là "Tốt lên") — TRÁNH nhầm 2 khái
            # niệm này với nhau, đúng yêu cầu user.
            rows.append([ind["label"], val_txt, t.get("latest_period", "—"),
                         t.get("value_arrow", "—"), t.get("judgment_label", "—"), src_label])
        t_grp = Table(rows, colWidths=[40 * mm, 26 * mm, 18 * mm, 14 * mm, 22 * mm, 35 * mm])
        t_grp.setStyle(tbl_style())
        story.append(t_grp)
        story.append(Spacer(1, 4))
        for k in keys:
            if k in charts:
                story.append(Image(charts[k], width=140 * mm, height=63 * mm))
        story.append(Spacer(1, 8))

    # ── Nguồn tham khảo đầy đủ ──
    story.append(Paragraph("3. Nguồn dữ liệu đầy đủ (theo từng chỉ báo)", h1_st))
    for key, ind in raw.items():
        if key == "_meta":
            continue
        urls = sorted({p.get("source_url") for p in ind["series"] if p.get("source_url")})
        if not urls:
            continue
        story.append(Paragraph(f"<b>{ind['label']}:</b>", small_st))
        for u in urls[:4]:
            story.append(Paragraph(f"• {u}", small_st))
    story.append(Spacer(1, 8))

    story.append(Paragraph(
        "⚠ Giai đoạn 1 (khung sơ bộ): một số chỉ báo (PMI, lãi suất liên ngân hàng, tăng trưởng "
        "tín dụng, cung tiền M2, dự trữ ngoại hối, lạm phát cơ bản) hiện chỉ có seed dữ liệu thưa "
        "hoặc chưa có nguồn tự động cập nhật — xem cột 'Nguồn cập nhật' ở mỗi bảng. Trục 'Dòng "
        "tiền & Tâm lý' (mua/bán ròng khối ngoại, tỷ lệ margin — Chương 5.3/6.1) CHƯA được tích "
        "hợp vào ma trận quyết định do chưa có nguồn dữ liệu tự động đáng tin cậy — ma trận hiện "
        "tại chỉ dựa trên 2 trục Scorecard Vĩ Mô và Định giá thị trường.", italic_st))

    doc.build(story)
    return pdf_path


# ══════════════════════════════════════════════════════════════════════════
# JSON EXPORT
# ══════════════════════════════════════════════════════════════════════════
def save_json_vimo(raw, trends, scorecard, scorecard_total, valuation, decision_label, decision_text):
    out = {
        "sector": "Vĩ mô",
        "lastUpdated": raw.get("_meta", {}).get("last_auto_update") or raw.get("_meta", {}).get("seeded_at"),
        "scorecard": {
            "groups": {gname: {"score": g["score"], "nVotes": g["n_votes"],
                                "detail": g["detail"]} for gname, g in scorecard.items()},
            "total": scorecard_total,
        },
        "marketValuation": valuation,
        "decision": {"label": decision_label, "text": decision_text},
        "indicators": {},
    }
    for key, ind in raw.items():
        if key == "_meta":
            continue
        t = trends.get(key, {})
        out["indicators"][key] = {
            "group": ind["group"], "label": ind["label"], "unit": ind["unit"],
            "goodDirection": ind["good_direction"], "autoSource": ind["auto_source"],
            "series": ind["series"], "trend": t, "note": ind.get("note"),
        }
    json_path = os.path.join(PROJECT_ROOT, "data", "vimo.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"  [OK] JSON: {json_path}")
    return json_path


# ══════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════
def run_vimo_analysis():
    print("\n" + "=" * 60)
    print("  PHÂN TÍCH VĨ MÔ KINH TẾ VIỆT NAM")
    print("=" * 60)

    raw = load_vimo_raw()

    print("[INFO] Fetching Rf (lãi suất TPCP 10 năm) cho ERP...")
    rf, rf_src = fetch_rf_vietnam()
    print(f"  Rf={rf*100:.2f}% ({rf_src})")

    print("[INFO] Tính xu hướng từng chỉ báo...")
    trends = {}
    for key, ind in raw.items():
        if key == "_meta":
            continue
        trends[key] = calc_trend(ind["series"], ind["good_direction"])
        t = trends[key]
        print(f"  {ind['label']}: {t.get('arrow', '—')} (latest={t.get('latest')}, n={t.get('n_points', 0)})")

    print("[INFO] Tính Scorecard Vĩ Mô...")
    scorecard, scorecard_total = calc_scorecard(raw, trends)
    for gname, g in scorecard.items():
        print(f"  {gname}: {g['score']:+d} ({g['n_votes']} chỉ báo có xu hướng)")
    print(f"  => TỔNG SCORECARD: {scorecard_total:+d} / 5")

    print("[INFO] Tính định giá thị trường (P/E, ERP)...")
    valuation = calc_market_valuation(raw, rf)
    print(f"  P/E={valuation['pe']} | ERP={valuation['erp']*100:.2f}%" if valuation['erp'] is not None else "  P/E/ERP: N/A")
    print(f"  => {valuation['valuation_label']}")

    decision_label, decision_text = calc_decision_matrix(scorecard_total, valuation["valuation_label"])
    print(f"[INFO] Ma trận quyết định: {decision_label} — {decision_text}")

    out_dir = os.path.join(PROJECT_ROOT, "Bao cao", "VIMO")
    os.makedirs(out_dir, exist_ok=True)
    print("[INFO] Building charts...")
    charts = build_charts_vimo(out_dir, raw)
    print(f"  -> {len(charts)} charts")

    date_str = datetime.datetime.now().strftime("%Y-%m-%d")
    pdf_path = os.path.join(out_dir, f"VIMO_Report_{date_str}.pdf")
    print("[INFO] Building PDF...")
    build_pdf_vimo(pdf_path, raw, trends, scorecard, scorecard_total, valuation, decision_label, decision_text, charts)
    print(f"  [OK] PDF: {pdf_path}")

    print("[INFO] Saving JSON dashboard...")
    save_json_vimo(raw, trends, scorecard, scorecard_total, valuation, decision_label, decision_text)

    for p in charts.values():
        try:
            os.remove(p)
        except Exception:
            pass

    print(f"\n{'='*60}")
    print(f"  HOÀN THÀNH — PDF: {pdf_path}")
    print(f"{'='*60}\n")
    return True


if __name__ == "__main__":
    run_vimo_analysis()
