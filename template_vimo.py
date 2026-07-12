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
    "Thanh khoản": ["credit_growth", "m2_growth", "forex_reserves", "omo_rate_7d"],
    "Bên ngoài": ["trade_balance", "export_growth", "fdi_disbursed", "usdvnd", "fed_funds_rate", "brent_oil", "dxy_proxy", "china_gdp_growth"],
    "Tâm lý thị trường": ["vnindex_pe"],
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
    # VN-Index P/B ĐÃ LOẠI KHỎI phạm vi (theo quyết định user — không phải tiêu chí chính, và
    # không tìm được nguồn API tự động cho riêng cấp độ CHỈ SỐ, chỉ có API cấp từng công ty —
    # xem lịch sử trao đổi/git log). Chỉ còn P/E làm cơ sở tính ERP.
    pe_series = raw["vnindex_pe"]["series"]
    pe = pe_series[-1]["value"] if pe_series else None
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
        "pe": pe, "rf": rf, "erp": erp, "valuation_label": valuation_label,
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
# TỔNG HỢP PHÂN TÍCH ĐA CHỈ SỐ — "bức tranh tổng thể" (user yêu cầu rõ: không chỉ liệt kê số
# liệu/Scorecard, mà phải có ĐÁNH GIÁ TÁC ĐỘNG thật sự tới kinh tế Việt Nam và TTCK, viết chuyên
# nghiệp, đủ chi tiết — không phải 1-2 câu ngắn). RULE-BASED, KHÔNG dùng AI/LLM (user: "tôi không
# cần gemini api key, vì tôi không muốn dùng ai") — mọi câu chữ được ráp từ template có điều kiện,
# mọi con số lấy trực tiếp từ raw/trends/scorecard/valuation, không suy diễn ngoài dữ liệu thật.
# ══════════════════════════════════════════════════════════════════════════
def _vn_join(items):
    items = list(items)
    if not items:
        return ""
    if len(items) == 1:
        return items[0]
    return ", ".join(items[:-1]) + " và " + items[-1]


def _indicator_txt(raw, trends, key):
    """Trả (label, value_text, judgment_label) cho 1 chỉ báo nếu có dữ liệu, else None."""
    ind = raw.get(key)
    t = trends.get(key, {})
    if not ind or t.get("latest") is None:
        return None
    val_txt = f"{t['latest']:.2f} {ind['unit']}"
    return ind["label"], val_txt, t.get("judgment_label")


def _join_indicators(raw, trends, keys):
    parts = []
    for k in keys:
        info = _indicator_txt(raw, trends, k)
        if info:
            label, val, judgment = info
            j = f", {judgment.lower()}" if judgment else ""
            parts.append(f"{label} đạt {val}{j}")
    return parts


def _build_overview(raw, trends, scorecard, scorecard_total, valuation, decision_label, decision_text):
    pos = [g for g, d in scorecard.items() if d["score"] == 1]
    neg = [g for g, d in scorecard.items() if d["score"] == -1]
    neu = [g for g, d in scorecard.items() if d["score"] == 0 and d["n_votes"] > 0]
    no_data = [g for g, d in scorecard.items() if d["n_votes"] == 0]

    posture = "tích cực" if scorecard_total > 0 else ("tiêu cực" if scorecard_total < 0 else "trung tính, chưa nghiêng hẳn về hướng nào")
    sentences = [f"Scorecard Vĩ Mô tổng hợp đạt {scorecard_total:+d}/5 điểm, cho thấy bức tranh vĩ mô Việt Nam hiện nghiêng {posture}."]
    if pos:
        sentences.append(f"Động lực tích cực đến từ nhóm {_vn_join(pos)}.")
    if neg:
        sentences.append(f"Ngược lại, nhóm {_vn_join(neg)} đang là điểm nghẽn, kéo lùi điểm tổng.")
    if neu:
        sentences.append(f"Nhóm {_vn_join(neu)} giữ trạng thái trung tính — các chỉ báo trong nhóm chưa đồng thuận rõ hướng.")
    if no_data:
        sentences.append(f"Riêng nhóm {_vn_join(no_data)} chưa có đủ chỉ báo có xu hướng rõ để tính điểm, nên tạm giữ 0 theo quy ước thận trọng (không suy diễn khi thiếu dữ liệu).")
    para1 = " ".join(sentences)

    # Liệt kê CỤ THỂ điểm sáng/điểm nghẽn từ TOÀN BỘ chỉ báo (không chỉ giới hạn trong Scorecard)
    # — mirror phong cách "ĐÁNH GIÁ CHU KỲ" của báo cáo vĩ mô DSC/VikkiBankS (T6/2026, user cung
    # cấp làm mẫu tham khảo): nêu số cụ thể, đánh số rõ ràng từng điểm thay vì chỉ nói chung chung.
    positives, risks = [], []
    for key, ind in raw.items():
        if key == "_meta":
            continue
        t = trends.get(key, {})
        if t.get("latest") is None:
            continue
        item_txt = f"{ind['label']} đạt {t['latest']:.2f} {ind['unit']}"
        if t.get("judgment_label") == "Tốt lên":
            positives.append(item_txt)
        elif t.get("judgment_label") == "Xấu đi":
            risks.append(item_txt)

    para2_parts = []
    if positives:
        para2_parts.append("Các điểm sáng cụ thể: " + "; ".join(f"({i+1}) {p}" for i, p in enumerate(positives)) + ".")
    if risks:
        para2_parts.append("Ngược lại, một số điểm nghẽn/rủi ro cần lưu ý: " + "; ".join(f"({i+1}) {r}" for i, r in enumerate(risks)) + ".")
    para2 = " ".join(para2_parts)

    if valuation.get("erp") is not None:
        para3 = (f"Về định giá, VN-Index đang giao dịch ở P/E {valuation['pe']:.2f}x, tương ứng ERP (lợi suất thu nhập trừ "
                 f"lãi suất phi rủi ro) {valuation['erp']*100:.2f}% — mức định giá được xếp loại \"{valuation['valuation_label']}\". "
                 f"Kết hợp Scorecard vĩ mô ({scorecard_total:+d}/5) và mức định giá này, khuyến nghị phân bổ vốn hiện tại là "
                 f"\"{decision_label}\": {decision_text}")
    else:
        para3 = (f"Chưa đủ dữ liệu P/E để tính ERP — phần định giá thị trường tạm thời để trống, không suy diễn. "
                 f"Khuyến nghị phân bổ vốn dựa trên Scorecard: \"{decision_label}\" — {decision_text}")

    paras = [para1]
    if para2:
        paras.append(para2)
    paras.append(para3)
    return "\n\n".join(paras)


def _build_economy_impact(raw, trends):
    paras = []

    growth_keys = ["gdp_growth", "iip_growth", "pmi_manufacturing", "retail_sales_growth"]
    inflation_keys = ["cpi_yoy", "core_inflation"]
    g_parts = _join_indicators(raw, trends, growth_keys)
    i_parts = _join_indicators(raw, trends, inflation_keys)
    if g_parts or i_parts:
        txt = "Về tăng trưởng và lạm phát: "
        if g_parts:
            txt += "; ".join(g_parts) + ". "
        # Bóc tách ĐỘNG LỰC tăng trưởng thật (theo NSO) thay vì chỉ nêu con số GDP đơn lẻ — trả
        # lời trực tiếp câu hỏi "GDP tăng do đâu, do chi tiêu công hay đầu tư tư nhân".
        industry_series = raw.get("gdp_industry_contribution", {}).get("series", [])
        invest_series = raw.get("gdp_investment_growth", {}).get("series", [])
        if industry_series:
            p = industry_series[-1]
            txt += (f"Theo NSO, động lực tăng trưởng GDP {p['period']} chủ yếu đến từ khu vực Công nghiệp-Xây dựng "
                     f"(đóng góp {p['value']:.2f}% vào mức tăng chung), trong đó ngành chế biến-chế tạo là hạt nhân. ")
        if invest_series:
            p = invest_series[-1]
            txt += (f"Xét theo phía cầu (phương pháp sử dụng GDP), tích lũy tài sản — tức ĐẦU TƯ, gồm cả đầu tư công, "
                     f"tư nhân và FDI — tăng {p['value']:.2f}% ({p['period']}), vượt xa tốc độ tăng tiêu dùng cuối cùng, "
                     "cho thấy tăng trưởng đang được dẫn dắt chủ yếu bởi đầu tư và xuất khẩu hơn là cầu tiêu dùng nội địa. ")
        if i_parts:
            txt += "Ở chiều giá cả, " + "; ".join(i_parts) + ". "
        brent_j = _indicator_txt(raw, trends, "brent_oil")
        if brent_j and brent_j[2] == "Tốt lên":
            txt += ("Giá dầu Brent hạ nhiệt là một kênh giảm áp lực CPI đáng chú ý — tác động trực tiếp nhất qua cấu "
                    "phần chi phí vận tải/nhiên liệu trong rổ CPI (nhóm Giao thông), gián tiếp qua chi phí logistics "
                    "đầu vào của doanh nghiệp sản xuất, chứ không phải do giá lương thực-thực phẩm hạ. ")
        usdvnd_j = _indicator_txt(raw, trends, "usdvnd")
        if usdvnd_j and usdvnd_j[2] == "Xấu đi":
            txt += ("Tỷ giá USD/VND đang chịu áp lực tăng, làm tăng chi phí nhập khẩu nguyên nhiên liệu và có thể cộng "
                    "thêm áp lực lên CPI trong các kỳ tới qua kênh lạm phát nhập khẩu.")
        paras.append(txt.strip())

    # Thanh khoản/lãi suất — đây là mục QUAN TRỌNG NHẤT cần trung thực: biểu lãi suất huy động
    # niêm yết CÔNG KHAI (VCB/VietinBank/NamABank, xem dep_parts bên dưới) thấp hơn NHIỀU so với
    # thực tế thị trường — user đã chỉ ra và cung cấp nguồn thật (VnExpress, NSO) xác nhận độc
    # lập: KHÔNG được kết luận "vĩ mô ổn định" chỉ vì lãi suất niêm yết thấp và đi ngang.
    omo = _indicator_txt(raw, trends, "omo_rate_7d")
    interbank_on_series = raw.get("interbank_rate_on", {}).get("series", [])
    interbank_9m_series = raw.get("interbank_rate_9m", {}).get("series", [])
    refin = _indicator_txt(raw, trends, "refinancing_rate")
    if omo or interbank_on_series or refin:
        txt2 = ("Về thanh khoản hệ thống: lãi suất OMO kỳ hạn 7 ngày là công cụ NHNN dùng để bơm/hút thanh khoản NGẮN "
                "HẠN tại thị trường liên ngân hàng (thị trường 2), KHÔNG phải thay đổi cung tiền M2 dài hạn — ")
        if omo:
            txt2 += f"hiện ở mức {omo[1]}. "
        if interbank_on_series and interbank_9m_series:
            on_v = interbank_on_series[-1]["value"]
            m9_v = interbank_9m_series[-1]["value"]
            slope = "dốc lên" if m9_v > on_v else ("gần như phẳng" if abs(m9_v - on_v) < 0.3 else "đảo ngược nhẹ")
            txt2 += (f"Đường cong lãi suất liên ngân hàng VNIBOR hiện {slope} (O/N {on_v:.2f}% so với 9 tháng {m9_v:.2f}%), "
                     "phản ánh kỳ vọng thị trường về mức độ căng thẳng thanh khoản ở các kỳ hạn dài hơn. ")
        txt2 += ("Tác động từ OMO lan tỏa GIÁN TIẾP sang lãi suất huy động/cho vay tại thị trường 1 (doanh nghiệp và dân "
                 "cư) thông qua kênh truyền dẫn lãi suất liên ngân hàng — không tức thời và không 1-1, thường có độ trễ "
                 "vài tuần đến vài tháng tùy thanh khoản từng ngân hàng.")
        dep_parts = _join_indicators(raw, trends, ["deposit_rate_12m_vcb", "deposit_rate_12m_ctg", "deposit_rate_12m_nab"])
        if dep_parts:
            txt2 += " Biểu niêm yết CHÍNH THỨC kỳ hạn 12 tháng của nhóm ngân hàng lớn: " + "; ".join(dep_parts) + "."
        paras.append(txt2.strip())

        neg_series = raw.get("deposit_rate_negotiated_max", {}).get("series", [])
        txt3 = "LƯU Ý QUAN TRỌNG — không nên chỉ nhìn biểu lãi suất niêm yết để kết luận thanh khoản ổn định. "
        if neg_series:
            p = neg_series[-1]
            txt3 += (f"Theo NSO (tính đến 26/6/2026: huy động toàn hệ thống +5,02% YTD trong khi tín dụng +7,41% YTD, "
                     f"tín dụng vượt huy động khoảng 1,48 lần) và báo chí (VnExpress, {p['period']}): một số ngân hàng "
                     f"đã phải chào lãi suất huy động THỎA THUẬN (ngoài biểu niêm yết) lên tới {p['value']:.1f}%/năm cho "
                     "khoản tiền gửi lớn (200 triệu - 1 tỷ đồng trở lên) để bù đắp khoảng cách này. Bảng lãi suất ONLINE "
                     "công khai đa ngân hàng (24hmoney.vn, 07/2026) cũng cho thấy nhiều NHTM (SHB, Saigonbank, PGBank, "
                     "Sacombank...) đã niêm yết công khai 6,8-7,8%/năm kỳ hạn 12 tháng — cao hơn hẳn mức ~5,9% của "
                     "riêng nhóm Big4 nêu trên. Đây là dấu hiệu hệ thống ngân hàng đang THỰC SỰ CĂNG THẲNG thanh "
                     "khoản để đáp ứng nhu cầu tín dụng, khác hẳn ấn tượng ổn định nếu chỉ nhìn lãi suất niêm yết Big4.")
        if neg_series:
            paras.append(txt3.strip())

    trade_keys = ["export_growth", "import_growth", "trade_balance", "fdi_disbursed"]
    fiscal_keys = ["budget_revenue_growth", "budget_expenditure_growth", "public_investment_growth"]
    t_parts = _join_indicators(raw, trends, trade_keys)
    f_parts = _join_indicators(raw, trends, fiscal_keys)
    if t_parts or f_parts:
        txt3 = ""
        if t_parts:
            txt3 += "Về thương mại và dòng vốn: " + "; ".join(t_parts) + ". "
            exp_v = trends.get("export_growth", {}).get("latest")
            imp_v = trends.get("import_growth", {}).get("latest")
            if exp_v is not None and imp_v is not None:
                if imp_v > exp_v:
                    txt3 += "Tốc độ tăng nhập khẩu đang cao hơn xuất khẩu, cho thấy áp lực thu hẹp cán cân thương mại nếu xu hướng này kéo dài. "
                else:
                    txt3 += "Xuất khẩu đang tăng nhanh hơn nhập khẩu, hỗ trợ tích cực cho cán cân thương mại và nguồn cung ngoại tệ. "
        if f_parts:
            txt3 += ("Về đầu tư công, " + "; ".join(f_parts) + " — mức độ giải ngân thực tế quyết định hiệu ứng lan tỏa "
                     "tới các ngành xây dựng, vật liệu, hạ tầng.")
        paras.append(txt3.strip())

    unemp = _indicator_txt(raw, trends, "unemployment_rate")
    ext_keys = ["fed_funds_rate", "brent_oil", "dxy_proxy", "china_gdp_growth"]
    e_parts = _join_indicators(raw, trends, ext_keys)
    if unemp or e_parts:
        txt4 = ""
        if unemp:
            txt4 += f"Về lao động, {unemp[0]} ở mức {unemp[1]}. "
        if e_parts:
            txt4 += "Áp lực từ bên ngoài: " + "; ".join(e_parts) + ". "
            txt4 += ("Lãi suất Fed và chỉ số USD ảnh hưởng trực tiếp tới dòng vốn ngoại và tỷ giá USD/VND; giá dầu Brent "
                     "tác động tới chi phí đầu vào và lạm phát nhập khẩu do Việt Nam là nước nhập khẩu ròng xăng dầu "
                     "tinh chế; tăng trưởng GDP Trung Quốc ảnh hưởng tới nhu cầu nhập khẩu hàng hóa Việt Nam do đây là "
                     "một trong các đối tác thương mại lớn nhất.")
        paras.append(txt4.strip())

    if not paras:
        return "Chưa đủ dữ liệu chỉ báo có xu hướng rõ để đánh giá tác động tới nền kinh tế."
    return "\n\n".join(paras)


def _build_market_impact(raw, trends, scorecard_total, valuation, decision_label, decision_text):
    paras = []
    posture = "thuận lợi" if scorecard_total > 0 else ("bất lợi" if scorecard_total < 0 else "trung tính")
    paras.append(f"Với Scorecard vĩ mô {scorecard_total:+d}/5 điểm ({posture} cho kênh cổ phiếu) và định giá VN-Index "
                 f"được xếp loại \"{valuation.get('valuation_label', 'chưa xác định')}\", khuyến nghị phân bổ vốn hiện "
                 f"tại là \"{decision_label}\": {decision_text}")

    if valuation.get("erp") is not None and valuation.get("rf") is not None:
        p2 = (f"Chênh lệch lợi suất thu nhập cổ phiếu so với lãi suất phi rủi ro (ERP) hiện ở mức {valuation['erp']*100:.2f}%, "
              f"trên nền lãi suất phi rủi ro tham chiếu {valuation['rf']*100:.2f}%. ")
        if valuation["erp"] < 0.02:
            p2 += ("Biên độ bù đắp rủi ro này tương đối mỏng — nếu lãi suất huy động/trái phiếu tiếp tục tăng, sức hấp "
                   "dẫn tương đối của cổ phiếu so với kênh tiền gửi/trái phiếu sẽ giảm thêm.")
        else:
            p2 += ("Biên độ này vẫn còn dư địa, cổ phiếu vẫn giữ được sức hấp dẫn tương đối so với kênh tiền gửi/trái "
                   "phiếu ở mức lãi suất hiện tại.")
        neg_series = raw.get("deposit_rate_negotiated_max", {}).get("series", [])
        if neg_series and valuation.get("rf") is not None:
            real_gap = neg_series[-1]["value"] / 100 - valuation["rf"]
            if real_gap > 0:
                p2 += (f" LƯU Ý: Rf tham chiếu ở trên lấy từ lợi suất TPCP — thấp hơn đáng kể lãi suất huy động THỰC TẾ "
                       f"cao nhất đang ghi nhận trên thị trường ({neg_series[-1]['value']:.1f}%/năm, xem chi tiết ở mục "
                       f"tác động kinh tế). Nếu dùng mức lãi suất thực tế này làm chuẩn so sánh chi phí vốn thay vì Rf "
                       f"trái phiếu, sức hấp dẫn tương đối của cổ phiếu sẽ THẤP HƠN NHIỀU so với con số ERP nêu trên — "
                       "không nên chỉ dựa vào ERP tính theo Rf trái phiếu để kết luận định giá đang hấp dẫn.")
        paras.append(p2)

    sector_notes = []
    usdvnd_j = _indicator_txt(raw, trends, "usdvnd")
    if usdvnd_j and usdvnd_j[2] == "Xấu đi":
        sector_notes.append("VND mất giá có lợi tương đối cho nhóm xuất khẩu (dệt may, thủy sản, gỗ) nhờ tăng sức cạnh "
                             "tranh giá, nhưng gây áp lực chi phí cho nhóm nhập khẩu nguyên liệu và doanh nghiệp có dư "
                             "nợ vay ngoại tệ lớn.")
    brent_j = _indicator_txt(raw, trends, "brent_oil")
    if brent_j:
        if brent_j[2] == "Tốt lên":
            sector_notes.append("Giá dầu Brent hạ nhiệt hỗ trợ biên lợi nhuận nhóm vận tải, nhựa, phân bón (chi phí "
                                 "đầu vào giảm), nhưng bất lợi cho nhóm khai thác/dầu khí thượng nguồn.")
        elif brent_j[2] == "Xấu đi":
            sector_notes.append("Giá dầu Brent tăng có lợi cho nhóm dầu khí thượng nguồn nhưng bào mòn biên lợi nhuận "
                                 "nhóm vận tải, nhựa, phân bón do chi phí đầu vào tăng.")
    interbank_j = _indicator_txt(raw, trends, "interbank_rate_3m")
    if interbank_j and interbank_j[2] == "Xấu đi":
        sector_notes.append("Lãi suất liên ngân hàng tăng phản ánh thanh khoản hệ thống thắt chặt hơn — thường bất lợi "
                             "cho nhóm cổ phiếu dùng đòn bẩy cao (bất động sản, xây dựng) do chi phí vốn tăng.")
    if sector_notes:
        paras.append("Ở cấp độ ngành: " + " ".join(sector_notes))

    return "\n\n".join(paras)


def _build_watch_points(raw, trends, scorecard):
    worsening = []
    for gdata in scorecard.values():
        for d in gdata["detail"]:
            if d["judgment_label"] == "Xấu đi":
                worsening.append(d["label"])
    lines = []
    if worsening:
        lines.append(f"- Các chỉ báo đang xấu đi cần theo dõi sát: {_vn_join(worsening)}.")
    no_data_groups = [g for g, d in scorecard.items() if d["n_votes"] == 0]
    if no_data_groups:
        lines.append(f"- Nhóm {_vn_join(no_data_groups)} hiện chưa đủ chỉ báo có xu hướng rõ để đưa vào Scorecard — cần bổ sung nguồn dữ liệu hoặc theo dõi thủ công.")
    lines.append("- Diễn biến đường cong lãi suất liên ngân hàng VNIBOR (đặc biệt chênh lệch O/N so với các kỳ hạn dài) — dấu hiệu sớm về căng thẳng thanh khoản hệ thống.")
    lines.append("- Tỷ giá USD/VND và động thái lãi suất Fed — ảnh hưởng trực tiếp tới dòng vốn ngoại và áp lực lạm phát nhập khẩu.")
    lines.append("- Xu hướng cán cân thương mại (xuất khẩu so với nhập khẩu) trong các kỳ báo cáo tiếp theo.")
    cred_v = trends.get("credit_growth", {}).get("latest")
    dep_growth_v = trends.get("deposit_growth", {}).get("latest")
    if cred_v is not None and dep_growth_v is not None and cred_v > dep_growth_v:
        lines.append(f"- Khoảng cách tăng trưởng tín dụng ({cred_v:.2f}%) so với huy động vốn ({dep_growth_v:.2f}%) — "
                      "nếu tiếp tục nới rộng, lãi suất huy động thực tế/thỏa thuận (đã ghi nhận tới 9%/năm, cao hơn "
                      "nhiều biểu niêm yết công khai) có thể còn tăng thêm, siết chi phí vốn toàn nền kinh tế.")
    return "\n".join(lines)


def build_synthesis_vimo(raw, trends, scorecard, scorecard_total, valuation, decision_label, decision_text):
    """Tổng hợp phân tích đa chỉ số RULE-BASED (không AI) — trả dict {overview, economy_impact,
    market_impact, watch_points}, mỗi mục là 1+ đoạn văn ráp từ số liệu thật trong raw/trends."""
    return {
        "overview": _build_overview(raw, trends, scorecard, scorecard_total, valuation, decision_label, decision_text),
        "economy_impact": _build_economy_impact(raw, trends),
        "market_impact": _build_market_impact(raw, trends, scorecard_total, valuation, decision_label, decision_text),
        "watch_points": _build_watch_points(raw, trends, scorecard),
    }


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


# Danh sách chỉ báo lãi suất huy động 12T theo ngân hàng, gắn nhãn NHÓM quy mô (user yêu cầu so
# sánh theo nhóm lớn/vừa/nhỏ — hiện chỉ có đại diện nhóm lớn và nhỏ, xem note trong vimo_raw.json
# về lý do thiếu đại diện nhóm vừa — MBB/TCB không có nguồn scrape ổn định).
BANK_DEPOSIT_RATE_KEYS = [
    ("deposit_rate_12m_vcb", "VCB (lớn)"),
    ("deposit_rate_12m_ctg", "VietinBank (lớn)"),
    ("deposit_rate_12m_nab", "NamABank (nhỏ)"),
]


def build_bank_comparison_chart(out_dir, raw):
    """1 bar chart so sánh lãi suất huy động 12 tháng giữa các ngân hàng đại diện — riêng biệt
    với build_charts_vimo() (ở đó mỗi chỉ báo có chart đường thời gian RIÊNG, còn ở đây cần 1
    chart CỘT so sánh NGANG giữa nhiều ngân hàng tại cùng 1 thời điểm). Trả path hoặc None nếu
    chưa có ngân hàng nào có dữ liệu."""
    labels, values = [], []
    for key, bank_label in BANK_DEPOSIT_RATE_KEYS:
        series = raw.get(key, {}).get("series", [])
        if series:
            labels.append(bank_label)
            values.append(series[-1]["value"])
    if not values:
        return None
    fig, ax = plt.subplots(figsize=(7.5, 3.4))
    colors = ["#3b82f6", "#3b82f6", "#f59e0b"][:len(labels)]
    bars = ax.bar(labels, values, color=colors)
    for b, v in zip(bars, values):
        ax.text(b.get_x() + b.get_width() / 2, v + 0.05, f"{v:.2f}%", ha="center", fontsize=10, fontweight="bold")
    ax.set_title("Lãi suất huy động 12 tháng theo ngân hàng (%)", fontsize=11, fontweight="bold")
    ax.grid(alpha=0.25, axis="y")
    fig.tight_layout()
    path = os.path.join(out_dir, "vimo_bank_deposit_comparison.png")
    fig.savefig(path, dpi=130)
    plt.close(fig)
    return path


# Kỳ hạn VNIBOR theo đúng thứ tự tăng dần — khớp key trong vimo_raw.json (fetch_sbv_interest_rates()
# trong fetch_macro_data.py). Kiểu trình bày lấy cảm hứng từ chart lãi suất liên ngân hàng đa kỳ
# hạn của vimo.cuthongthai.vn (user chỉ ra) — nhưng dùng CHÍNH dữ liệu đã cào từ sbv.gov.vn,
# không phụ thuộc gì vào nguồn của họ.
INTERBANK_TENOR_KEYS = [
    ("interbank_rate_on", "O/N"), ("interbank_rate_1w", "1 Tuần"), ("interbank_rate_2w", "2 Tuần"),
    ("interbank_rate_1m", "1 Tháng"), ("interbank_rate_3m", "3 Tháng"),
    ("interbank_rate_6m", "6 Tháng"), ("interbank_rate_9m", "9 Tháng"),
]


def build_interbank_curve_chart(out_dir, raw):
    """1 bar chart đường cong lãi suất liên ngân hàng (VNIBOR) theo 7 kỳ hạn tại cùng 1 thời
    điểm — cho thấy hình dạng đường cong lãi suất (dốc lên = thị trường kỳ vọng thanh khoản căng
    hơn ở kỳ hạn dài). Trả path hoặc None nếu chưa đủ dữ liệu."""
    labels, values = [], []
    for key, tenor_label in INTERBANK_TENOR_KEYS:
        series = raw.get(key, {}).get("series", [])
        if series:
            labels.append(tenor_label)
            values.append(series[-1]["value"])
    if len(values) < 2:
        return None
    fig, ax = plt.subplots(figsize=(7.5, 3.4))
    ax.plot(labels, values, marker="o", color="#8b5cf6", linewidth=2)
    ax.fill_between(labels, values, alpha=0.08, color="#8b5cf6")
    for i, v in enumerate(values):
        ax.text(i, v + 0.08, f"{v:.2f}%", ha="center", fontsize=9, fontweight="bold")
    ax.set_title("Đường cong lãi suất liên ngân hàng VNIBOR theo kỳ hạn (%)", fontsize=11, fontweight="bold")
    ax.grid(alpha=0.25)
    fig.tight_layout()
    path = os.path.join(out_dir, "vimo_interbank_curve.png")
    fig.savefig(path, dpi=130)
    plt.close(fig)
    return path


# ══════════════════════════════════════════════════════════════════════════
# PDF
# ══════════════════════════════════════════════════════════════════════════
def build_pdf_vimo(pdf_path, raw, trends, scorecard, scorecard_total, valuation, decision_label,
                    decision_text, charts, synthesis):
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

    # ── 1. TỔNG HỢP PHÂN TÍCH ĐA CHỈ SỐ (AI, dựa trên số liệu thật) — đặt NGAY ĐẦU báo cáo vì
    # đây là phần quan trọng nhất theo yêu cầu user, không phải phần liệt kê số liệu thô.
    story.append(Paragraph("1. Tổng hợp Phân tích Đa Chỉ số — Bức tranh Tổng thể", h1_st))
    story.append(Paragraph("1.1. Tổng quan bức tranh vĩ mô", h2_st))
    story.append(Paragraph(synthesis["overview"], body_st))
    story.append(Paragraph("1.2. Tác động tới kinh tế Việt Nam", h2_st))
    story.append(Paragraph(synthesis["economy_impact"], body_st))
    story.append(Paragraph("1.3. Tác động tới thị trường chứng khoán", h2_st))
    story.append(Paragraph(synthesis["market_impact"], body_st))
    story.append(Paragraph("1.4. Điểm cần theo dõi tiếp", h2_st))
    story.append(Paragraph(synthesis["watch_points"], body_st))
    story.append(Spacer(1, 10))

    # ── Scorecard chi tiết 5 nhóm ──
    story.append(Paragraph("2. Scorecard Vĩ Mô — 5 nhóm chỉ số (Chương 4.1/6.1)", h1_st))
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
    story.append(Paragraph("3. Chi tiết theo từng nhóm chỉ báo (Chương 3)", h1_st))
    groups_order = ["growth", "inflation", "monetary", "trade", "fiscal", "labor", "external", "market"]
    for grp in groups_order:
        keys = [k for k, v in raw.items() if k != "_meta" and v["group"] == grp]
        if not keys:
            continue
        story.append(Paragraph(f"3.{groups_order.index(grp)+1}. {GROUP_LABELS[grp]}", h2_st))
        rows = [["Chỉ báo", "Giá trị mới nhất", "Kỳ", "Số liệu", "Đánh giá", "Nguồn cập nhật"]]
        for k in keys:
            t = trends.get(k, {})
            ind = raw[k]
            src_label = {"worldbank": "World Bank API", "imf": "IMF DataMapper API", "fred": "FRED API",
                         "fx_api": "exchangerate-api.com", "pe_ratio_api": "worldperatio.com",
                         "nso_scrape": "nso.gov.vn (báo cáo quý, tự động)",
                         "nso_chart_embed": "nso.gov.vn (biểu đồ tháng, tự động)",
                         "sbv_chart": "sbv.gov.vn (biểu đồ, tự động)",
                         "sbv_table": "sbv.gov.vn (bảng lãi suất, tự động)",
                         "vietnambiz": "data.vietnambiz.vn (tự động)",
                         "bank_page": "Trang NH chính thức (tự động)",
                         "manual": "Nghiên cứu thủ công"}.get(ind["auto_source"], ind["auto_source"])
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
        if grp == "monetary" and "interbank_curve" in charts:
            story.append(Paragraph("Đường cong lãi suất liên ngân hàng VNIBOR theo kỳ hạn:", small_st))
            story.append(Image(charts["interbank_curve"], width=140 * mm, height=63 * mm))
            story.append(Spacer(1, 4))
        if grp == "monetary" and "bank_comparison" in charts:
            story.append(Paragraph("So sánh lãi suất huy động 12 tháng theo ngân hàng đại diện:", small_st))
            story.append(Image(charts["bank_comparison"], width=140 * mm, height=63 * mm))
        story.append(Spacer(1, 8))

    # ── Nguồn tham khảo đầy đủ ──
    story.append(Paragraph("4. Nguồn dữ liệu đầy đủ (theo từng chỉ báo)", h1_st))
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
def save_json_vimo(raw, trends, scorecard, scorecard_total, valuation, decision_label, decision_text,
                    synthesis, pdf_url=None):
    out = {
        "sector": "Vĩ mô",
        "gdrivePdfUrl": pdf_url,
        "lastUpdated": raw.get("_meta", {}).get("last_auto_update") or raw.get("_meta", {}).get("seeded_at"),
        "scorecard": {
            "groups": {gname: {"score": g["score"], "nVotes": g["n_votes"],
                                "detail": g["detail"]} for gname, g in scorecard.items()},
            "total": scorecard_total,
        },
        "marketValuation": valuation,
        "decision": {"label": decision_label, "text": decision_text},
        "synthesis": synthesis,
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

    print("[INFO] Tổng hợp phân tích đa chỉ số (rule-based, dựa trên số liệu thật)...")
    synthesis = build_synthesis_vimo(raw, trends, scorecard, scorecard_total, valuation, decision_label, decision_text)
    print("  -> Đã sinh phân tích tổng hợp")

    out_dir = os.path.join(PROJECT_ROOT, "Bao cao", "VIMO")
    os.makedirs(out_dir, exist_ok=True)
    print("[INFO] Building charts...")
    charts = build_charts_vimo(out_dir, raw)
    bank_chart = build_bank_comparison_chart(out_dir, raw)
    if bank_chart:
        charts["bank_comparison"] = bank_chart
    interbank_chart = build_interbank_curve_chart(out_dir, raw)
    if interbank_chart:
        charts["interbank_curve"] = interbank_chart
    print(f"  -> {len(charts)} charts")

    date_str = datetime.datetime.now().strftime("%Y-%m-%d")
    pdf_path = os.path.join(out_dir, f"VIMO_Report_{date_str}.pdf")
    print("[INFO] Building PDF...")
    build_pdf_vimo(pdf_path, raw, trends, scorecard, scorecard_total, valuation, decision_label, decision_text,
                    charts, synthesis)
    print(f"  [OK] PDF: {pdf_path}")

    print("[INFO] Saving JSON dashboard...")
    save_json_vimo(raw, trends, scorecard, scorecard_total, valuation, decision_label, decision_text, synthesis)

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
