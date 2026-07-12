#!/usr/bin/env python3
"""
fetch_macro_data.py — Cập nhật data/vimo_raw.json bằng dữ liệu vĩ mô THẬT từ các API/nguồn đã
xác nhận hoạt động (World Bank, IMF DataMapper, FRED, exchangerate-api.com, worldperatio.com,
nso.gov.vn). Script THUẦN `requests` — không phụ thuộc tool nào của Claude Code — chạy được cả
cục bộ lẫn trong GitHub Actions runner (xem .github/workflows/update_vimo.yml).

Nguồn cần API key (FRED) mà thiếu key sẽ TỰ BỎ QUA (không lỗi, không crash pipeline) — xem
FRED_API_KEY trong GitHub Secrets.

nso.gov.vn có chứng chỉ TLS không tự verify được (đã xác nhận qua khảo sát thủ công) — dùng
verify=False có chủ đích cho riêng domain này, không áp dụng cho các nguồn khác.
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
import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
VIMO_RAW_PATH = os.path.join(PROJECT_ROOT, "data", "vimo_raw.json")

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")


def load_raw():
    with open(VIMO_RAW_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def save_raw(data):
    with open(VIMO_RAW_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _current_period(dt=None):
    dt = dt or datetime.date.today()
    return dt.strftime("%Y-%m")


def _append_point(raw, key, period, value, source_url):
    """Thêm điểm mới vào cuối series — nếu period đã tồn tại thì ghi đè (tránh trùng khi chạy
    nhiều lần trong cùng tháng)."""
    series = raw[key]["series"]
    if series and series[-1]["period"] == period:
        series[-1]["value"] = value
        series[-1]["source_url"] = source_url
    else:
        series.append({"period": period, "value": value, "source_url": source_url})


# ══════════════════════════════════════════════════════════════════════════
# NGUỒN 1: API thật, KHÔNG cần key
# ══════════════════════════════════════════════════════════════════════════
def fetch_worldbank(indicator_code, country="VN", n=10):
    """Trả list [(year_str, value), ...] mới nhất, hoặc [] nếu lỗi."""
    url = f"https://api.worldbank.org/v2/country/{country}/indicator/{indicator_code}?format=json&per_page={n}"
    try:
        r = requests.get(url, headers={"User-Agent": UA}, timeout=20)
        r.raise_for_status()
        data = r.json()
        out = [(row["date"], round(row["value"], 2)) for row in data[1] if row.get("value") is not None]
        return out
    except Exception as e:
        print(f"  [WARN] World Bank {indicator_code}/{country} thất bại: {e}")
        return []


def fetch_imf_datamapper(indicator_code, country="VNM", n=10):
    # LƯU Ý: IMF DataMapper API trả 403 (Akamai edge block) khi gửi User-Agent giả lập trình
    # duyệt — ngược đời so với các WAF thông thường (thường chặn KHÔNG có UA). Đã verify: bỏ
    # hẳn header User-Agent (dùng UA mặc định của requests) thì gọi thành công bình thường.
    url = f"https://www.imf.org/external/datamapper/api/v1/{indicator_code}/{country}"
    try:
        r = requests.get(url, timeout=20)
        r.raise_for_status()
        vals = r.json().get("values", {}).get(indicator_code, {}).get(country, {})
        years = sorted(vals.keys())[-n:]
        return [(y, round(vals[y], 2)) for y in years]
    except Exception as e:
        print(f"  [WARN] IMF DataMapper {indicator_code}/{country} thất bại: {e}")
        return []


def fetch_usdvnd_current():
    try:
        r = requests.get("https://api.exchangerate-api.com/v4/latest/USD", headers={"User-Agent": UA}, timeout=15)
        r.raise_for_status()
        return round(r.json()["rates"]["VND"], 2), "https://api.exchangerate-api.com/v4/latest/USD"
    except Exception as e:
        print(f"  [WARN] USD/VND fetch thất bại: {e}")
        return None, None


def fetch_vnindex_pe_current():
    try:
        r = requests.get("https://worldperatio.com/area/vietnam/", headers={"User-Agent": UA}, timeout=20)
        r.raise_for_status()
        # Trang thực tế: "Current P/E<br>( 15.43 )" — có tag <br> chen giữa, và có 1 dòng header
        # "Current P/E Ratio (x₀)" đứng trước KHÔNG có số thật, nên phải cho phép tag tùy ý ở
        # giữa và bắt buộc match có chữ số thật trong ngoặc.
        m = re.search(r"Current P/E(?:<[^>]+>|\s)*\(\s*([\d.]+)\s*\)", r.text)
        if m:
            return float(m.group(1)), "https://worldperatio.com/area/vietnam/"
        print("  [WARN] worldperatio.com: không tìm thấy pattern 'Current P/E (...)' — trang có thể đã đổi cấu trúc.")
        return None, None
    except Exception as e:
        print(f"  [WARN] VN-Index P/E fetch thất bại: {e}")
        return None, None


# ══════════════════════════════════════════════════════════════════════════
# NGUỒN 2: FRED (cần API key — tự skip nếu thiếu)
# ══════════════════════════════════════════════════════════════════════════
def fetch_fred(series_id, n=12):
    api_key = os.environ.get("FRED_API_KEY")
    if not api_key:
        print(f"  [SKIP] FRED {series_id}: thiếu biến môi trường FRED_API_KEY, bỏ qua.")
        return []
    try:
        url = (f"https://api.stlouisfed.org/fred/series/observations?series_id={series_id}"
               f"&api_key={api_key}&file_type=json&sort_order=desc&limit={n}")
        r = requests.get(url, timeout=20)
        r.raise_for_status()
        out = [(o["date"], round(float(o["value"]), 2))
               for o in r.json().get("observations", []) if o["value"] != "."]
        return out
    except Exception as e:
        print(f"  [WARN] FRED {series_id} thất bại: {e}")
        return []


# ══════════════════════════════════════════════════════════════════════════
# NGUỒN 3: nso.gov.vn (scrape HTML thật, cần crawl trang danh sách để tìm bài mới nhất)
# ══════════════════════════════════════════════════════════════════════════
def fetch_nso_latest_report():
    """Crawl trang danh sách nso.gov.vn tìm báo cáo kinh tế-xã hội mới nhất, trích GDP/thất
    nghiệp từ câu chữ thật. Trả dict hoặc None nếu thất bại — best-effort, không lỗi pipeline."""
    try:
        r = requests.get("https://www.nso.gov.vn/en/data-and-statistics/", headers={"User-Agent": UA},
                          timeout=20, verify=False)
        r.raise_for_status()
        links = re.findall(
            r'href="(https://www\.nso\.gov\.vn/en/[a-z\-]+/\d{4}/\d{2}/[^"]*'
            r'(?:socio-economic-situation|report-on-socio-economic)[^"]*)"', r.text)
        # "infographic-on-the-socio-economic-situation..." KHÔNG phải bài text report (chủ yếu
        # ảnh, không có câu chữ để regex trích số) — loại ra, chỉ giữ bài report dạng văn bản
        # thật. Báo cáo đầy đủ của quý mới nhất thường ra SAU infographic vài tuần — nếu quý này
        # chưa có bài report văn bản (chỉ mới có infographic), coi là bình thường, bỏ qua nhẹ
        # nhàng, lần chạy Action sau sẽ nhặt được khi bài report đã lên.
        links = [u for u in links if "infographic" not in u.lower()]
        if not links:
            print("  [INFO] Chưa có bài report văn bản (dạng text) mới cho kỳ hiện tại trên NSO "
                  "(có thể mới chỉ có infographic dạng ảnh) — bỏ qua, giữ nguyên seed cũ.")
            return None
        latest_url = links[0]
        r2 = requests.get(latest_url, headers={"User-Agent": UA}, timeout=20, verify=False)
        r2.raise_for_status()
        text = re.sub(r"<[^>]+>", " ", r2.text)
        text = re.sub(r"&#\d+;", " ", text)
        text = re.sub(r"\s+", " ", text)

        out = {"source_url": latest_url}
        m = re.search(r"GDP[^.]{0,80}?increase[d]?\s+by\s+([\d.]+)\s*%\s*year-on-year", text, re.I)
        if m:
            out["gdp_growth"] = float(m.group(1))
        m = re.search(r"unemployment rate at working age was\s+([\d.]+)\s*%", text, re.I)
        if m:
            out["unemployment_rate"] = float(m.group(1))
        return out
    except Exception as e:
        print(f"  [WARN] NSO scrape thất bại: {e}")
        return None


def fetch_nso_chart_embed(chart_slug):
    """nso.gov.vn có các trang chuyên đề (vd /cpi-vi/, /iip-vi/) nhúng iframe biểu đồ Highcharts
    tại nso.gov.vn/chart/<slug>/embed/ — trang embed chứa thẳng mảng "data":[[kỳ, giá trị], ...]
    dạng JSON trong HTML, KHÔNG cần crawl trang danh sách như fetch_nso_latest_report(). Đây là
    nguồn chi tiết theo THÁNG (tốt hơn báo cáo quý dùng cho GDP/thất nghiệp/XNK/FDI).
    Trả list [(period_label_goc, value), ...] hoặc [] nếu thất bại."""
    url = f"https://www.nso.gov.vn/chart/{chart_slug}/embed/?show=chart&width=responsive&share"
    try:
        r = requests.get(url, headers={"User-Agent": UA}, timeout=20, verify=False)
        r.raise_for_status()
        m = re.search(r'"data":(\[\[.*?\]\])', r.text)
        if not m:
            print(f"  [WARN] NSO chart embed {chart_slug}: không tìm thấy mảng 'data' — trang có thể đã đổi cấu trúc.")
            return []
        pairs = json.loads(m.group(1))
        return [(str(p[0]), float(p[1])) for p in pairs]
    except Exception as e:
        print(f"  [WARN] NSO chart embed {chart_slug} thất bại: {e}")
        return []


def _nso_period_to_iso(label):
    """NSO chart embed trả nhãn kỳ kiểu '6/2025' hoặc '01/2026' (tháng/năm, không đệm số 0 nhất
    quán) — chuẩn hóa về 'YYYY-MM' để khớp định dạng period dùng chung trong vimo_raw.json."""
    m = re.match(r"(\d{1,2})/(\d{4})", label)
    if m:
        month, year = m.groups()
        return f"{year}-{int(month):02d}"
    return label


def fetch_sbv_credit_growth():
    """sbv.gov.vn nhúng thẳng mảng JS 'const tongCong = [...]' (tăng trưởng tín dụng TỔNG theo
    tháng, %) cùng 'const labels = [...]' trên trang dư nợ tín dụng — không cần API key, không
    JS rendering. Trả list [(period_iso, value), ...] hoặc [] nếu thất bại."""
    url = "https://www.sbv.gov.vn/vi/du-no-tin-dung-doi-voi-nen-kt-dttktt"
    try:
        r = requests.get(url, headers={"User-Agent": UA}, timeout=20, verify=False)
        r.raise_for_status()
        m_labels = re.search(r"const labels\s*=\s*(\[[^\]]+\]);", r.text)
        m_total = re.search(r"const tongCong\s*=\s*(\[[\d.,\s\-]+\]);", r.text)
        if not (m_labels and m_total):
            print("  [WARN] SBV credit growth: không tìm thấy 'labels'/'tongCong' — trang có thể đã đổi cấu trúc.")
            return []
        labels = json.loads(m_labels.group(1))
        values = json.loads(m_total.group(1))
        return [(_nso_period_to_iso(lb), float(v)) for lb, v in zip(labels, values)]
    except Exception as e:
        print(f"  [WARN] SBV credit growth thất bại: {e}")
        return []


def fetch_sbv_interest_rates():
    """sbv.gov.vn/vi/lãi-suất1 — LƯU Ý: URL này bị 404 khi test bằng curl KHÔNG có domain
    'www.' phía trước hoặc thiếu -L theo redirect (đã từng kết luận nhầm là link chết ở lần
    khảo sát trước — user cung cấp lại URL và test kỹ hơn xác nhận trang THẬT SỰ hoạt động qua
    'https://www.sbv.gov.vn/...' + theo redirect). Trang chứa 2 bảng HTML thật (không phải JS
    render): (1) lãi suất tái chiết khấu/tái cấp vốn hiện hành, (2) lãi suất bình quân liên ngân
    hàng theo kỳ hạn (O/N, 1W, 2W, 1M, 3M, 6M, 9M). Số dùng dấu phẩy thập phân kiểu Việt Nam
    ('4,500%') — phải đổi ',' -> '.' trước khi ép kiểu float.
    Trả dict {"refinancing_rate": value, "interbank_rate_on": value, "interbank_rate_1w": value,
    "interbank_rate_2w": value, "interbank_rate_1m": value, "interbank_rate_3m": value,
    "interbank_rate_6m": value, "interbank_rate_9m": value} (key nào không tìm thấy thì bị bỏ
    qua, không lỗi)."""
    url = "https://www.sbv.gov.vn/vi/l%C3%A3i-su%E1%BA%A5t1"
    # Nhãn kỳ hạn TRÊN TRANG SBV -> tên field trong out dict. Thứ tự khớp đúng cột "Doanh số"
    # đứng cạnh mỗi dòng trong bảng "Lãi suất BQ liên Ngân hàng".
    TENOR_MAP = [
        ("Qua đêm", "interbank_rate_on"),
        ("1 Tuần", "interbank_rate_1w"),
        ("2 Tuần", "interbank_rate_2w"),
        ("1 Tháng", "interbank_rate_1m"),
        ("3 Tháng", "interbank_rate_3m"),
        ("6 Tháng", "interbank_rate_6m"),
        ("9 Tháng", "interbank_rate_9m"),
    ]
    try:
        r = requests.get(url, headers={"User-Agent": UA}, timeout=20, verify=False)
        r.raise_for_status()
        text = re.sub(r"<[^>]+>", " | ", r.text)
        text = re.sub(r"\s+", " ", text)

        out = {}
        m = re.search(r"Lãi suất tái cấp vốn(?:\s*\|)+\s*([\d,]+)\s*%", text)
        if m:
            out["refinancing_rate"] = float(m.group(1).replace(",", "."))
        for label, key in TENOR_MAP:
            m = re.search(re.escape(label) + r"(?:\s*\|)+\s*([\d,]+)\s", text)
            if m:
                out[key] = float(m.group(1).replace(",", "."))
        return out
    except Exception as e:
        print(f"  [WARN] SBV lãi suất thất bại: {e}")
        return {}


def fetch_sbv_omo_rate():
    """sbv.gov.vn/vi/web/sbv_portal/nghiệp-vụ-thị-trường-mở — kết quả đấu thầu OMO (mua kỳ hạn)
    mới nhất, bảng HTML thật. LƯU Ý BẢN CHẤT (user giải thích, ghi lại để không hiểu nhầm khi
    dùng dữ liệu): OMO là công cụ BƠM/HÚT THANH KHOẢN NGẮN HẠN của NHNN tại thị trường LIÊN NGÂN
    HÀNG (thị trường 2) — hoàn toàn KHÔNG PHẢI cung tiền M2 (M2 là tổng phương tiện thanh toán
    trong nền kinh tế, đo lường khác hẳn). Hoạt động bơm/hút rất ngắn hạn (7/35/63 ngày), mục
    đích là khơi thông tắc nghẽn thanh khoản tức thời, không phải tăng/giảm cung tiền dài hạn.
    Tác động lan tỏa dần từ thị trường 2 (lãi suất liên ngân hàng) sang thị trường 1 (lãi suất
    huy động/cho vay với doanh nghiệp & dân cư) qua kênh truyền dẫn lãi suất liên ngân hàng —
    KHÔNG tức thời, KHÔNG trực tiếp. Trả dict {"omo_rate_7d": value} hoặc {} nếu thất bại."""
    url = "https://www.sbv.gov.vn/vi/web/sbv_portal/nghi%E1%BB%87p-v%E1%BB%A5-th%E1%BB%8B-tr%C6%B0%E1%BB%9Dng-m%E1%BB%9F"
    try:
        r = requests.get(url, headers={"User-Agent": UA}, timeout=20, verify=False)
        r.raise_for_status()
        text = re.sub(r"<[^>]+>", " | ", r.text)
        text = re.sub(r"\s+", " ", text)

        out = {}
        m = re.search(r"Kỳ hạn 7 ngày(?:\s*\|)+\s*([\d/]+)(?:\s*\|)+\s*([\d,]+)(?:\s*\|)+\s*([\d,]+)", text)
        if m:
            out["omo_rate_7d"] = float(m.group(3).replace(",", "."))
        return out
    except Exception as e:
        print(f"  [WARN] SBV OMO thất bại: {e}")
        return {}


# ══════════════════════════════════════════════════════════════════════════
# NGUỒN 4: Lãi suất huy động 12 tháng — từng ngân hàng đại diện theo nhóm quy mô (user yêu cầu).
# LƯU Ý: đã khảo sát 6 ngân hàng (VCB/CTG nhóm lớn, MBB/TCB nhóm vừa, NAB/VAB nhóm nhỏ) — CHỈ
# VCB/CTG/NAB có nguồn cào ổn định cho lãi suất HUY ĐỘNG; KHÔNG ngân hàng nào có nguồn ổn định
# cho lãi suất CHO VAY SXKD/mua nhà (JS-render hoặc chỉ nằm trong văn bản quảng cáo không đáng
# tin cậy để tự động hóa) — không cố ép lấy, tránh vi phạm nguyên tắc "không estimate".
# ══════════════════════════════════════════════════════════════════════════
def fetch_vcb_deposit_rate_12m():
    """Vietcombank có API JSON thật (không cần JS render): trả toàn bộ biểu lãi suất huy động
    theo kỳ hạn/loại tiền. Lọc tenor='12-months', currencyCode='VND', tenorType='Savings'."""
    url = "https://www.vietcombank.com.vn/vi-VN/api/interestrates?accountType=Personal"
    try:
        r = requests.get(url, headers={"User-Agent": UA}, timeout=20)
        r.raise_for_status()
        data = r.json()
        for item in data.get("Data", []):
            if item.get("tenor") == "12-months" and item.get("currencyCode") == "VND" and item.get("tenorType") == "Savings":
                return round(item["rates"] * 100, 2)
        print("  [WARN] VCB: không tìm thấy dòng 12-months/VND/Savings trong API response.")
        return None
    except Exception as e:
        print(f"  [WARN] VCB deposit rate thất bại: {e}")
        return None


def fetch_ctg_deposit_rate_12m():
    """VietinBank — bảng HTML thật tại lai-suat-khcn. Regex khớp CHÍNH XÁC nhãn '12 tháng' (loại
    trừ 'Từ 11 tháng đến dưới 12 tháng'/'Trên 12 tháng đến 13 tháng' — các nhãn khác cũng chứa
    chuỗi '12 tháng' nên phải cẩn thận không khớp nhầm)."""
    url = "https://www.vietinbank.vn/lai-suat-khcn"
    try:
        r = requests.get(url, headers={"User-Agent": UA}, timeout=20)
        r.raise_for_status()
        m = re.search(r'text-left\s*\">12\s*tháng</td><td class="p-4">([\d,]+)<!-- -->\s*%</td>', r.text)
        if m:
            return round(float(m.group(1).replace(",", ".")), 2)
        print("  [WARN] VietinBank: không khớp được dòng '12 tháng' — trang có thể đã đổi cấu trúc.")
        return None
    except Exception as e:
        print(f"  [WARN] VietinBank deposit rate thất bại: {e}")
        return None


def fetch_nab_deposit_rate_12m():
    """NamABank — bảng HTML thật tại lai-suat-tien-gui-vnd-2. Nhãn '12 tháng, 365 ngày' xuống
    dòng qua nhiều thẻ <p>/<strong> — regex phải cho phép whitespace/tag linh hoạt giữa nhãn và
    giá trị cột đầu tiên (lãi cuối kỳ)."""
    url = "https://www.namabank.com.vn/lai-suat-tien-gui-vnd-2"
    try:
        r = requests.get(url, headers={"User-Agent": UA}, timeout=20)
        r.raise_for_status()
        m = re.search(
            r"12\s*tháng,\s*</strong></p>\s*<p><strong>365\s*ngày<br\s*/>\s*</strong></p>\s*</td>\s*"
            r"<td>\s*<p>([\d.]+)<br\s*/>", r.text)
        if m:
            return round(float(m.group(1)), 2)
        print("  [WARN] NamABank: không khớp được dòng '12 tháng, 365 ngày' — trang có thể đã đổi cấu trúc.")
        return None
    except Exception as e:
        print(f"  [WARN] NamABank deposit rate thất bại: {e}")
        return None


# VietnamBiz's Vietnamese "title" field -> indicator key trong vimo_raw.json. CHỈ map các chỉ
# báo mà VietnamBiz là nguồn TỐT NHẤT tìm được (PMI, bán lẻ — trước đây "manual" chỉ 1 điểm) —
# không map đè lên GDP/CPI/thất nghiệp/IIP vì NSO (trực tiếp từ cơ quan thống kê) đáng tin cậy
# hơn nguồn tổng hợp lại của bên thứ ba, dù VietnamBiz cũng có các chỉ báo đó làm đối chiếu.
VIETNAMBIZ_TITLE_MAP = {
    "PMI": "pmi_manufacturing",
    "Bán lẻ HH&DV (YoY)": "retail_sales_growth",
    "Thu ngân sách (YoY)": "budget_revenue_growth",
    "Chi ngân sách (YoY)": "budget_expenditure_growth",
    "Vốn đầu tư NSNN (YoY)": "public_investment_growth",
    "Xuất khẩu (YoY)": "export_growth",
    "Nhập khẩu (YoY)": "import_growth",
}


def fetch_vietnambiz_macro():
    """data.vietnambiz.vn/macro-economic nhúng __NEXT_DATA__ JSON (Next.js server-rendered,
    KHÔNG phải SPA rỗng) chứa ~25 chỉ báo vĩ mô, mỗi chỉ báo có value (kỳ mới nhất) + pre_value
    (kỳ trước) + nhãn kỳ tiếng Việt ('Tháng 06/2026'/'Quý 2/2026'/'Năm 2023'). Trả dict
    {indicator_key: (period_iso, value)} cho các chỉ báo có trong VIETNAMBIZ_TITLE_MAP."""
    url = "https://data.vietnambiz.vn/macro-economic"
    try:
        r = requests.get(url, headers={"User-Agent": UA}, timeout=20)
        r.raise_for_status()
        m = re.search(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>', r.text, re.S)
        if not m:
            print("  [WARN] VietnamBiz: không tìm thấy __NEXT_DATA__ — trang có thể đã đổi cấu trúc.")
            return {}
        data = json.loads(m.group(1))
        items = data.get("props", {}).get("pageProps", {}).get("data", [])
        out = {}
        for it in items:
            key = VIETNAMBIZ_TITLE_MAP.get(it.get("title"))
            if not key:
                continue
            ngay = it.get("ngay", "")
            m_month = re.match(r"Tháng\s+(\d{1,2})/(\d{4})", ngay)
            m_quarter = re.match(r"Quý\s+(\d)/(\d{4})", ngay)
            if m_month:
                period = f"{m_month.group(2)}-{int(m_month.group(1)):02d}"
            elif m_quarter:
                period = f"{m_quarter.group(2)}-Q{m_quarter.group(1)}"
            else:
                period = ngay
            out[key] = (period, round(float(it["value"]), 2))
        return out
    except Exception as e:
        print(f"  [WARN] VietnamBiz macro thất bại: {e}")
        return {}


VIETNAMBIZ_RATE_TITLE_MAP = {
    "Tăng trưởng huy động (YoY)": "deposit_growth",
    "Tăng trưởng cung tiền M2 (YoY)": "m2_growth",
}


def fetch_vietnambiz_rates():
    """data.vietnambiz.vn/currency-interest-rate — cùng cấu trúc __NEXT_DATA__ như
    fetch_vietnambiz_macro() nhưng trang riêng cho tiền tệ/lãi suất, chứa "Tăng trưởng huy động
    (YoY)" — chỉ báo QUAN TRỌNG để đối chiếu với credit_growth (đã có, nguồn SBV riêng): khi tín
    dụng tăng nhanh hơn huy động vốn, hệ thống ngân hàng phải cạnh tranh huy động mạnh hơn (lãi
    suất huy động thực tế/thỏa thuận thường cao hơn biểu niêm yết — xem note của deposit_growth
    trong vimo_raw.json). Cũng lấy luôn M2 growth thật (trước đây chỉ có 1 điểm seed thủ công)."""
    url = "https://data.vietnambiz.vn/currency-interest-rate"
    try:
        r = requests.get(url, headers={"User-Agent": UA}, timeout=20)
        r.raise_for_status()
        m = re.search(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>', r.text, re.S)
        if not m:
            print("  [WARN] VietnamBiz rates: không tìm thấy __NEXT_DATA__ — trang có thể đã đổi cấu trúc.")
            return {}
        data = json.loads(m.group(1))
        items = data.get("props", {}).get("pageProps", {}).get("data", [])
        out = {}
        for it in items:
            key = VIETNAMBIZ_RATE_TITLE_MAP.get(it.get("title"))
            if not key:
                continue
            ngay = it.get("ngay", "")
            m_month = re.match(r"Tháng\s+(\d{1,2})/(\d{4})", ngay)
            if m_month:
                period = f"{m_month.group(2)}-{int(m_month.group(1)):02d}"
            else:
                period = ngay
            out[key] = (period, round(float(it["value"]), 2))
        return out
    except Exception as e:
        print(f"  [WARN] VietnamBiz rates thất bại: {e}")
        return {}


def fetch_market_deposit_rate_12m():
    """24hmoney.vn/lai-suat-gui-ngan-hang — trang HTML tĩnh THẬT (đã xác nhận qua curl, không
    phải SPA rỗng), có bảng lãi suất gửi ONLINE kỳ hạn 12 tháng của ~38 ngân hàng (class
    "online-table", cột cuối trong 5 cột 1/3/6/9/12 tháng). Khác biểu niêm yết Big4 (VCB/
    VietinBank/NamABank) đang theo dõi riêng — bảng này cho thấy MẶT BẰNG rộng hơn nhiều của toàn
    thị trường, xác nhận việc chỉ nhìn Big4 sẽ đánh giá thấp mức lãi suất huy động thực tế. Trả
    (max_rate, avg_rate, n_banks) hoặc (None, None, 0) nếu lỗi/không tìm thấy bảng."""
    url = "https://24hmoney.vn/lai-suat-gui-ngan-hang"
    try:
        r = requests.get(url, headers={"User-Agent": UA}, timeout=20)
        r.raise_for_status()
        idx = r.text.find("online-table")
        if idx == -1:
            print("  [WARN] 24hmoney: không tìm thấy bảng 'online-table' — trang có thể đã đổi cấu trúc.")
            return None, None, 0
        tbody_m = re.search(r"<tbody>(.*?)</tbody>", r.text[idx:], re.S)
        if not tbody_m:
            return None, None, 0
        rows = re.findall(r"<tr[^>]*>(.*?)</tr>", tbody_m.group(1), re.S)
        rates = []
        for row in rows:
            m = re.findall(r'class="bank-interest-rate[^"]*">([\d.]+)</p>', row)
            if len(m) >= 5:  # cột thứ 5 = kỳ hạn 12 tháng
                rates.append(float(m[4]))
        if not rates:
            return None, None, 0
        return max(rates), round(sum(rates) / len(rates), 2), len(rates)
    except Exception as e:
        print(f"  [WARN] 24hmoney thất bại: {e}")
        return None, None, 0


# Lãi suất huy động THỎA THUẬN (ngoài biểu niêm yết) không có API/trang công bố chính thức nào —
# chỉ xuất hiện rải rác trong tin tức khi báo chí phát hiện/phỏng vấn. RSS_NEWS_FEEDS là các
# nguồn tin thật, tần suất cao, đã xác nhận hoạt động (không phải trang search JS-rendered).
RSS_NEWS_FEEDS = [
    "https://cafef.vn/tai-chinh-ngan-hang.rss",
    "https://vietstock.vn/144/tai-chinh-ngan-hang.rss",
]
NEGOTIATED_RATE_TITLE_KEYWORDS = ["thỏa thuận", "chạm mốc", "vượt trần", "ngầm"]


def fetch_negotiated_deposit_rate_news():
    """Quét RSS tin tức tài chính-ngân hàng (CafeF, VietStock) mỗi lần Action chạy, tìm bài viết
    có tiêu đề chứa 'lãi suất' + 1 trong các từ khóa đặc trưng cho tin lãi suất THỎA THUẬN/ngầm
    (khác hẳn tin lãi suất niêm yết định kỳ). Thể loại tin này gần như luôn nêu THẲNG con số %
    ngay trong tiêu đề (vd 'Lãi suất thỏa thuận chạm mốc 9%/năm') — chỉ trích số khi tìm thấy %
    NGAY TRONG TIÊU ĐỀ của bài khớp từ khóa, để giảm rủi ro trích nhầm số từ nội dung bài (không
    đọc/hiểu văn bản tự do — chỉ regex có điều kiện chặt). Phần lớn các lần chạy sẽ KHÔNG tìm thấy
    bài nào khớp (bình thường — đây là tin hiếm, sự kiện) — hàm trả None, KHÔNG ghi đè dữ liệu cũ.
    Trả (period, value, source_url, title) hoặc None."""
    for feed_url in RSS_NEWS_FEEDS:
        try:
            r = requests.get(feed_url, headers={"User-Agent": UA}, timeout=15)
            r.raise_for_status()
            items = re.findall(r"<item>(.*?)</item>", r.text, re.S)
            for it in items:
                title_m = re.search(r"<title>\s*(?:<!\[CDATA\[)?(.*?)(?:\]\]>)?\s*</title>", it, re.S)
                title = title_m.group(1).strip() if title_m else ""
                if "lãi suất" not in title.lower():
                    continue
                if not any(kw in title.lower() for kw in NEGOTIATED_RATE_TITLE_KEYWORDS):
                    continue
                pct_m = re.search(r"(\d+(?:[.,]\d+)?)\s*%", title)
                if not pct_m:
                    continue
                value = float(pct_m.group(1).replace(",", "."))
                if not (3.0 <= value <= 15.0):  # biên hợp lý cho lãi suất huy động VND — loại số nhiễu
                    continue
                link_m = re.search(r"<link>\s*(?:<!\[CDATA\[)?(.*?)(?:\]\]>)?\s*</link>", it, re.S)
                link = link_m.group(1).strip() if link_m else feed_url
                pubdate_m = re.search(r"<pubDate>(.*?)</pubDate>", it, re.S)
                period = _current_period()
                if pubdate_m:
                    try:
                        dt = datetime.datetime.strptime(pubdate_m.group(1).strip()[:16], "%a, %d %b %y")
                        period = dt.strftime("%Y-%m")
                    except ValueError:
                        pass
                return (period, value, link, title)
        except Exception as e:
            print(f"  [WARN] RSS {feed_url} thất bại: {e}")
    return None


# ══════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════
def update_vimo_raw():
    print("=" * 60)
    print("  CẬP NHẬT data/vimo_raw.json — fetch_macro_data.py")
    print("=" * 60)
    raw = load_raw()
    period_now = _current_period()

    print("[USD/VND]")
    v, src = fetch_usdvnd_current()
    if v:
        _append_point(raw, "usdvnd", period_now, v, src)
        print(f"  -> {period_now}: {v}")

    print("[VN-Index P/E]")
    v, src = fetch_vnindex_pe_current()
    if v:
        _append_point(raw, "vnindex_pe", period_now, v, src)
        print(f"  -> {period_now}: {v}")

    print("[World Bank — China GDP growth]")
    pts = fetch_worldbank("NY.GDP.MKTP.KD.ZG", "CN", n=8)
    if pts:
        raw["china_gdp_growth"]["series"] = [
            {"period": y, "value": v, "source_url": "https://api.worldbank.org/v2/country/cn/indicator/NY.GDP.MKTP.KD.ZG"}
            for y, v in sorted(pts)
        ]
        print(f"  -> {len(pts)} điểm")

    print("[World Bank — Vietnam forex reserves]")
    pts = fetch_worldbank("FI.RES.TOTL.CD", "VN", n=8)
    if pts:
        # World Bank trả USD thô — quy đổi tỷ USD cho khớp đơn vị đã khai báo trong vimo_raw.json
        raw["forex_reserves"]["series"] = [
            {"period": y, "value": round(v / 1e9, 2),
             "source_url": "https://api.worldbank.org/v2/country/vn/indicator/FI.RES.TOTL.CD"}
            for y, v in sorted(pts)
        ]
        print(f"  -> {len(pts)} điểm")

    print("[IMF DataMapper — Vietnam public debt/GDP]")
    pts = fetch_imf_datamapper("GG_DEBT_GDP", "VNM", n=8)
    if pts:
        raw["public_debt_gdp"]["series"] = [
            {"period": y, "value": v, "source_url": "https://www.imf.org/external/datamapper/GG_DEBT_GDP@GDD/VNM"}
            for y, v in sorted(pts)
        ]
        print(f"  -> {len(pts)} điểm")

    print("[FRED — Fed funds rate / Brent oil / Dollar index]")
    for key, sid in {"fed_funds_rate": "FEDFUNDS", "brent_oil": "DCOILBRENTEU", "dxy_proxy": "DTWEXBGS"}.items():
        pts = fetch_fred(sid, n=12)
        if pts:
            raw[key]["series"] = [
                {"period": d, "value": v, "source_url": f"https://fred.stlouisfed.org/series/{sid}"}
                for d, v in sorted(pts)
            ]
            print(f"  -> {key}: {len(pts)} điểm")

    print("[NSO — báo cáo kinh tế-xã hội mới nhất]")
    nso = fetch_nso_latest_report()
    if nso:
        today_q = f"{datetime.date.today().year}-Q{(datetime.date.today().month - 1) // 3 + 1}"
        if "gdp_growth" in nso:
            _append_point(raw, "gdp_growth", today_q, nso["gdp_growth"], nso["source_url"])
            print(f"  -> gdp_growth {today_q}: {nso['gdp_growth']}")
        if "unemployment_rate" in nso:
            _append_point(raw, "unemployment_rate", today_q, nso["unemployment_rate"], nso["source_url"])
            print(f"  -> unemployment_rate {today_q}: {nso['unemployment_rate']}")
        if "gdp_growth" not in nso and "unemployment_rate" not in nso:
            print("  [WARN] Fetch được bài báo cáo nhưng không trích được số liệu nào — có thể mẫu câu đã đổi.")

    print("[NSO — biểu đồ chuyên đề CPI (nso.gov.vn/cpi-vi/, chi tiết THEO THÁNG)]")
    pts = fetch_nso_chart_embed("cpi")
    if pts:
        raw["cpi_yoy"]["series"] = [
            {"period": _nso_period_to_iso(p), "value": v,
             "source_url": "https://www.nso.gov.vn/cpi-vi/"}
            for p, v in pts
        ]
        print(f"  -> {len(pts)} điểm (thay thế chuỗi theo quý cũ bằng chuỗi theo tháng)")

    print("[NSO — biểu đồ chuyên đề IIP (nso.gov.vn/iip-vi/, chi tiết THEO THÁNG)]")
    pts = fetch_nso_chart_embed("index-of-industrial-production")
    if pts:
        raw["iip_growth"]["series"] = [
            {"period": _nso_period_to_iso(p), "value": v,
             "source_url": "https://www.nso.gov.vn/iip-vi/"}
            for p, v in pts
        ]
        print(f"  -> {len(pts)} điểm")

    print("[SBV — tăng trưởng tín dụng theo tháng]")
    pts = fetch_sbv_credit_growth()
    if pts:
        raw["credit_growth"]["series"] = [
            {"period": p, "value": v,
             "source_url": "https://www.sbv.gov.vn/vi/du-no-tin-dung-doi-voi-nen-kt-dttktt"}
            for p, v in pts
        ]
        print(f"  -> {len(pts)} điểm")

    print("[VietnamBiz — PMI & Bán lẻ (đối chiếu, tích lũy theo lần chạy)]")
    vnb = fetch_vietnambiz_macro()
    for key, (period, value) in vnb.items():
        _append_point(raw, key, period, value, "https://data.vietnambiz.vn/macro-economic")
        print(f"  -> {key} {period}: {value}")

    print("[VietnamBiz — Tăng trưởng huy động & M2 (đối chiếu credit_growth, tích lũy theo lần chạy)]")
    vnb_rates = fetch_vietnambiz_rates()
    for key, (period, value) in vnb_rates.items():
        _append_point(raw, key, period, value, "https://data.vietnambiz.vn/currency-interest-rate")
        print(f"  -> {key} {period}: {value}")

    print("[SBV — lãi suất tái cấp vốn & liên ngân hàng 3 tháng (tích lũy theo lần chạy)]")
    rates = fetch_sbv_interest_rates()
    for key, value in rates.items():
        _append_point(raw, key, period_now, value, "https://www.sbv.gov.vn/vi/l%C3%A3i-su%E1%BA%A5t1")
        print(f"  -> {key} {period_now}: {value}")

    print("[SBV — lãi suất OMO kỳ hạn 7 ngày (bơm thanh khoản thị trường 2, tích lũy theo lần chạy)]")
    omo = fetch_sbv_omo_rate()
    for key, value in omo.items():
        _append_point(raw, key, period_now,
                       value, "https://www.sbv.gov.vn/vi/web/sbv_portal/nghi%E1%BB%87p-v%E1%BB%A5-th%E1%BB%8B-tr%C6%B0%E1%BB%9Dng-m%E1%BB%9F")
        print(f"  -> {key} {period_now}: {value}")

    print("[Ngân hàng — lãi suất huy động 12 tháng: VCB / VietinBank / NamABank (tích lũy theo lần chạy)]")
    v = fetch_vcb_deposit_rate_12m()
    if v is not None:
        _append_point(raw, "deposit_rate_12m_vcb", period_now, v,
                       "https://www.vietcombank.com.vn/vi-VN/api/interestrates?accountType=Personal")
        print(f"  -> VCB {period_now}: {v}")
    v = fetch_ctg_deposit_rate_12m()
    if v is not None:
        _append_point(raw, "deposit_rate_12m_ctg", period_now, v, "https://www.vietinbank.vn/lai-suat-khcn")
        print(f"  -> VietinBank {period_now}: {v}")
    v = fetch_nab_deposit_rate_12m()
    if v is not None:
        _append_point(raw, "deposit_rate_12m_nab", period_now, v, "https://www.namabank.com.vn/lai-suat-tien-gui-vnd-2")
        print(f"  -> NamABank {period_now}: {v}")

    print("[24hmoney — lãi suất huy động online 12 tháng, mặt bằng toàn thị trường (~38 NH, tích lũy theo lần chạy)]")
    max_r, avg_r, n = fetch_market_deposit_rate_12m()
    if max_r is not None:
        src = "https://24hmoney.vn/lai-suat-gui-ngan-hang"
        _append_point(raw, "deposit_rate_12m_market_max", period_now, max_r, src)
        _append_point(raw, "deposit_rate_12m_market_avg", period_now, avg_r, src)
        print(f"  -> {period_now}: max={max_r}% avg={avg_r}% (n={n} ngân hàng)")

    print("[RSS tin tức — lãi suất huy động THỎA THUẬN (quét CafeF/VietStock, chỉ ghi khi có tin mới khớp)]")
    hit = fetch_negotiated_deposit_rate_news()
    if hit:
        period, value, link, title = hit
        _append_point(raw, "deposit_rate_negotiated_max", period, value, link)
        print(f"  -> {period}: {value}% — \"{title}\"")
    else:
        print("  -> Không có tin mới khớp từ khóa (bình thường, đây là tin hiếm)")

    raw["_meta"]["last_auto_update"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    save_raw(raw)
    print("\n[OK] Đã ghi data/vimo_raw.json")


if __name__ == "__main__":
    update_vimo_raw()
