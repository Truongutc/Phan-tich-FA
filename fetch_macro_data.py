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
import glob
import json
import datetime
import unicodedata
import subprocess
import statistics
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


def _current_period_weekly(dt=None):
    """Kỳ THEO TUẦN (ISO week, 'YYYY-Www') — dùng riêng cho nhóm lãi suất liên ngân hàng/huy động
    (user 2026-07-24: các số liệu này Action chạy hằng ngày/tuần nhưng trước đó dùng chung
    _current_period() theo THÁNG nên mỗi tháng chỉ có 1 điểm, biểu đồ lịch sử gần như không tích
    lũy được gì dù chạy nhiều lần — đổi sang tuần để mỗi lần chạy cách nhau ≥1 tuần đều tạo điểm
    mới thay vì ghi đè điểm THÁNG cũ). Điểm THÁNG cũ đã có trong vimo_raw.json vẫn giữ nguyên làm
    mốc lịch sử xa hơn — _append_point() so khớp chuỗi period nên tự động thêm điểm mới (không
    khớp định dạng cũ) thay vì ghi đè, không cần migrate dữ liệu cũ."""
    dt = dt or datetime.date.today()
    iso_year, iso_week, _ = dt.isocalendar()
    return f"{iso_year}-W{iso_week:02d}"


def _append_point(raw, key, period, value, source_url):
    """Thêm điểm mới vào cuối series — nếu period đã tồn tại thì ghi đè (tránh trùng khi chạy
    nhiều lần trong cùng tháng)."""
    series = raw[key]["series"]
    if series and series[-1]["period"] == period:
        series[-1]["value"] = value
        series[-1]["source_url"] = source_url
    else:
        series.append({"period": period, "value": value, "source_url": source_url})


def _period_sort_key(period):
    """Khoá sắp xếp thời gian cho các định dạng period KHÁC NHAU cùng tồn tại trong 1 series
    (vd fdi_registered_usd_bn/public_investment_growth trộn 'YYYY-MM' theo tháng từ VBMA với
    'YYYY-Qn'/'YYYY-Hn'/'YYYY-9M'/'YYYY-FY' lũy kế theo quý từ NSO/VietnamBiz). Trả (year, month,
    subrank) — subrank tách các kỳ cùng tháng cuối cùng của 1 khoảng lũy kế (vd FY đứng sau 9M)."""
    m = re.match(r"(\d{4})-(\d{2})$", period)
    if m:
        return (int(m.group(1)), int(m.group(2)), 0)
    m = re.match(r"(\d{4})-Q(\d)$", period)
    if m:
        return (int(m.group(1)), int(m.group(2)) * 3, 1)
    m = re.match(r"(\d{4})-H(\d)$", period)
    if m:
        return (int(m.group(1)), int(m.group(2)) * 6, 2)
    m = re.match(r"(\d{4})-9M$", period)
    if m:
        return (int(m.group(1)), 9, 3)
    m = re.match(r"(\d{4})-FY$", period)
    if m:
        return (int(m.group(1)), 12, 4)
    m = re.match(r"(\d{4})$", period)
    if m:
        return (int(m.group(1)), 12, 5)
    return (0, 0, 0)


def _merge_vbma_points(raw, key, points, source_url):
    """Trộn danh sách [(period, value), ...] từ VBMA vào raw[key]['series'] vốn có thể đã chứa
    các điểm period KHÁC ĐỊNH DẠNG (quý/nửa năm/lũy kế) từ nguồn khác (NSO/VietnamBiz) — khác
    _append_point() (chỉ so khớp điểm CUỐI), hàm này: (1) xoá điểm cũ có period trùng CHÍNH XÁC
    với điểm mới (tránh trùng lặp), (2) thêm toàn bộ điểm mới, (3) sắp xếp lại theo thời gian
    thực sự qua _period_sort_key() (tránh chuỗi bị đảo lộn thứ tự khi trộn 2 định dạng kỳ)."""
    if not points:
        return
    new_periods = {p for p, _ in points}
    series = raw[key]["series"]
    series[:] = [pt for pt in series if pt["period"] not in new_periods]
    series.extend({"period": p, "value": v, "source_url": source_url} for p, v in points)
    series.sort(key=lambda pt: _period_sort_key(pt["period"]))


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


def fetch_fii_net_flow():
    """cafef.vn — endpoint Ajax nội bộ (KHÔNG public API chính thức, nhưng public/không cần key,
    xác nhận hoạt động qua test thủ công 2026-07-13) trả khối lượng/giá trị mua-bán của KHỐI NGOẠI
    trên sàn HOSE. Trả về NGÀY GIAO DỊCH GẦN NHẤT có dữ liệu (không nhất thiết đúng ngày truyền
    vào tham số Date — vd cuối tuần/nghỉ lễ tự lùi về phiên gần nhất). Trả (net_ty_vnd, date_iso,
    source_url) hoặc (None, None, None) nếu lỗi. net > 0 = mua ròng, net < 0 = bán ròng."""
    today_str = datetime.date.today().strftime("%d/%m/%Y")
    url = f"https://cafef.vn/du-lieu/Ajax/PageNew/DataGDNN/GDNuocNgoai.ashx?TradeCenter=hose&Date={today_str}"
    try:
        r = requests.get(url, headers={"User-Agent": UA}, timeout=15)
        r.raise_for_status()
        data = r.json().get("Data", {})
        diff_value = data.get("DiffValue")
        date_str = data.get("Date")
        if diff_value is None or not date_str:
            return None, None, None
        net_ty_vnd = round(diff_value / 1e9, 2)
        d, m, y = date_str.split("/")
        return net_ty_vnd, f"{y}-{m}-{d}", url
    except Exception as e:
        print(f"  [WARN] FII (khối ngoại HOSE) thất bại: {e}")
        return None, None, None


VNINDEX_NONVIN_URL = ("https://raw.githubusercontent.com/Truongutc/AIC---chart-nganh/main/"
                       "Output/finance/VNINDEX_NONVIN.json")


def fetch_vnindex_nonvin_data():
    """GitHub PUBLIC repo Truongutc/AIC---chart-nganh (dự án khác của user, chia sẻ 2026-07-25) —
    file Output/finance/VNINDEX_NONVIN.json là dữ liệu P/E, P/B, ROE của TOÀN BỘ thị trường sau
    khi LOẠI BỎ họ VIN (VIC/VHM/VRE/VPL, xem sector_groups.json['VNINDEX_NONVIN']['exclude']) —
    user (2026-07-25) yêu cầu dùng dữ liệu này thay vì P/E/P/B HEADLINE (có VIN) để tính toán,
    vì VIN chiếm tỷ trọng lớn + định giá bất thường làm méo mó P/E/P/B chung của VN-Index (xác
    nhận bằng số liệu thật: 24/07/2026 P/E headline ~12.4x nhưng ex-VIN chỉ 10.4x, P/B 1.92x vs
    1.62x — chênh lệch đáng kể). File public, không cần token/xác thực.
    Trả dict {"pe": float, "pb": float, "roe": float (quý gần nhất), "date": "YYYY-MM-DD",
    "daily": {"dates":[...], "pe":[...], "pb":[...]}, "quarterly": {"labels":[...], "roe":[...],
    "yoy_lnst_growth":[...]}}} hoặc None nếu lỗi."""
    try:
        r = requests.get(VNINDEX_NONVIN_URL, timeout=20)
        r.raise_for_status()
        payload = r.json()
        daily = payload.get("daily", {})
        dates, pe_list, pb_list = daily.get("dates", []), daily.get("pe", []), daily.get("pb", [])
        # Lấy điểm GẦN NHẤT có đủ cả P/E lẫn P/B (vài ngày cuối có thể null nếu báo cáo tài chính
        # quý mới nhất chưa cập nhật đủ cho toàn bộ universe).
        latest_pe = latest_pb = latest_date = None
        for i in range(len(dates) - 1, -1, -1):
            if pe_list[i] is not None and pb_list[i] is not None:
                latest_pe, latest_pb, latest_date = pe_list[i], pb_list[i], dates[i]
                break
        if latest_pe is None:
            print("  [WARN] VNINDEX_NONVIN: không tìm thấy điểm P/E+P/B hợp lệ nào.")
            return None
        q = payload.get("quarterly", {})
        latest_roe = q["roe"][-1] if q.get("roe") else None
        return {
            "pe": latest_pe, "pb": latest_pb, "roe": latest_roe, "date": latest_date,
            "daily": daily, "quarterly": q,
        }
    except Exception as e:
        print(f"  [WARN] VNINDEX_NONVIN (GitHub) thất bại: {e}")
        return None


def fetch_vnindex_pe_pb_24hmoney():
    """24hmoney.vn/indices/vn-index — trang Nuxt SPA, nhưng dữ liệu P/E và P/B (cấp CHỈ SỐ, không
    phải từng mã) đã được server render sẵn dạng JS object literal trong <script> window.__NUXT__,
    y hệt cách vietnambiz nhúng __NEXT_DATA__ đang dùng ở các chỉ báo khác trong file này — không
    cần render JS. Cụ thể dạng: keyStatistic:{pb:2.08,pe:13.53,avg_volume:...}. Trả (pe, pb, url)
    hoặc (None, None, None) nếu lỗi/đổi cấu trúc."""
    url = "https://24hmoney.vn/indices/vn-index"
    try:
        r = requests.get(url, headers={"User-Agent": UA}, timeout=20)
        r.raise_for_status()
        m = re.search(r"keyStatistic:\{pb:([\d.]+),pe:([\d.]+)", r.text)
        if m:
            return float(m.group(2)), float(m.group(1)), url
        print("  [WARN] 24hmoney vn-index: không tìm thấy 'keyStatistic{pb,pe}' — trang có thể đã đổi cấu trúc.")
        return None, None, None
    except Exception as e:
        print(f"  [WARN] 24hmoney VN-Index P/E-P/B fetch thất bại: {e}")
        return None, None, None


def fetch_vnindex_pe_current():
    # Nguồn chính: 24hmoney.vn (cùng nguồn với bảng lãi suất huy động đã dùng trong file này, cũng
    # cho luôn P/B — xem fetch_vnindex_pe_pb_24hmoney()). worldperatio.com giữ làm fallback nếu
    # 24hmoney lỗi/đổi cấu trúc (theo đúng pattern fallback đang dùng cho fetch_rf_vietnam()).
    pe, pb, src = fetch_vnindex_pe_pb_24hmoney()
    if pe:
        return pe, src
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


def fetch_vnindex_pb_current():
    pe, pb, src = fetch_vnindex_pe_pb_24hmoney()
    if pb:
        return pb, src
    return None, None


# ══════════════════════════════════════════════════════════════════════════
# NGUỒN 2: FRED (cần API key — tự skip nếu thiếu)
# ══════════════════════════════════════════════════════════════════════════
def fetch_fred(series_id, n=12, units=None):
    """units=None -> giá trị gốc. units="pc1" -> %YoY (FRED tự tính "Percent Change from Year
    Ago"). units="chg" -> thay đổi tuyệt đối so kỳ liền trước ("Change"). Dùng để lấy thẳng
    CPI/PCE YoY% hoặc số việc làm tăng/giảm trong tháng mà không cần tự viết derived-diff."""
    api_key = os.environ.get("FRED_API_KEY")
    if not api_key:
        print(f"  [SKIP] FRED {series_id}: thiếu biến môi trường FRED_API_KEY, bỏ qua.")
        return []
    try:
        url = (f"https://api.stlouisfed.org/fred/series/observations?series_id={series_id}"
               f"&api_key={api_key}&file_type=json&sort_order=desc&limit={n}")
        if units:
            url += f"&units={units}"
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


def _nso_cumulative_period(phrase, fallback_year=None):
    """Chuyển cụm từ mô tả kỳ báo cáo NSO (vd 'sáu tháng đầu năm 2026', 'quý I năm 2026', 'chín
    tháng đầu năm 2026') thành nhãn kỳ chuẩn 'YYYY-H1'/'YYYY-Q1'/'YYYY-9M'/'YYYY-FY' — khớp quy
    ước period lũy kế đã dùng cho budget_revenue_growth/public_investment_growth trong dự án.
    Trả None nếu không xác định được năm."""
    year_m = re.search(r"(20\d{2})", phrase)
    year = year_m.group(1) if year_m else fallback_year
    if not year:
        return None
    p = phrase.lower()
    if "sáu tháng" in p or "6 tháng" in p:
        return f"{year}-H1"
    if "chín tháng" in p or "9 tháng" in p:
        return f"{year}-9M"
    if "quý i" in p or "quý 1" in p:
        return f"{year}-Q1"
    if "cả năm" in p:
        return f"{year}-FY"
    return f"{year}-FY"  # không khớp mẫu nào đã biết -> coi là lũy kế cả năm (an toàn hơn báo lỗi)


# Chữ số đếm tháng kiểu "Tính chung {N} tháng đầu năm..." trong báo cáo NSO — dùng để suy ra
# THÁNG báo cáo đang nói tới (xem public_investment_disbursement_rate_pct trong
# fetch_nso_gdp_structure_report()) mà không cần regex riêng từng dạng "tháng Tư"/"tháng 4".
def _vn_number(s):
    """Số kiểu VN dùng '.' làm phân cách NGHÌN và ',' làm phân cách THẬP PHÂN (vd '1.100,1' = 1100.1)
    — PHẢI xoá dấu chấm trước rồi mới đổi dấu phẩy, khác hẳn '.replace(",", ".")' đơn giản (sẽ vỡ
    với số ≥1000, vd '1.100,1'.replace(",",".") ra '1.100.1' không parse được thành float)."""
    return float(s.replace(".", "").replace(",", "."))


_VN_MONTH_COUNT_WORDS = {
    "một": 1, "hai": 2, "ba": 3, "bốn": 4, "năm": 5, "sáu": 6, "bảy": 7,
    "tám": 8, "chín": 9, "mười": 10, "mười một": 11, "mười hai": 12,
    "2": 2, "3": 3, "4": 4, "5": 5, "6": 6, "7": 7, "8": 8, "9": 9, "10": 10, "11": 11,
}


def fetch_nso_gdp_structure_report():
    """Tự động tìm bài 'Thông cáo báo chí về tình hình kinh tế-xã hội' MỚI NHẤT (tiếng Việt) trên
    nso.gov.vn/du-lieu-va-so-lieu-thong-ke/ (index Việt — KHÁC index tiếng Anh đã dùng ở
    fetch_nso_latest_report()), rồi trích 3 nhóm câu chữ THẬT (không suy diễn):
    1) Cơ cấu GDP theo khu vực kinh tế (nông-lâm-thủy sản / công nghiệp-xây dựng / dịch vụ / thuế
       sản phẩm), tính theo % — dùng vẽ biểu đồ miền cơ cấu GDP theo khu vực.
    2) Cơ cấu vốn đầu tư thực hiện toàn xã hội theo thành phần (Nhà nước / ngoài Nhà nước / FDI),
       tính theo % — dùng vẽ biểu đồ miền cơ cấu đầu tư.
    3) Tổng vốn FDI ĐĂNG KÝ (khác fdi_disbursed đã có — đó là FDI GIẢI NGÂN).
    LƯU Ý: đây là số liệu LŨY KẾ theo kỳ báo cáo (Q1/6 tháng/9 tháng/cả năm), KHÔNG phải chuỗi quý
    độc lập — xem _nso_cumulative_period(). Trả dict hoặc {} nếu thất bại/chưa có bài mới."""
    try:
        # NSO đổi URL scheme khoảng T2/2026: báo cáo mới nằm dưới /bai-top/YYYY/MM/... (liệt kê
        # tại trang danh mục /bao-cao-tinh-hinh-kinh-te-xa-hoi-hang-thang/), KHÁC path
        # /du-lieu-va-so-lieu-thong-ke/ dùng cho các báo cáo cũ hơn (T1/2026 trở về trước) — giữ
        # cả 2 pattern để không mất khả năng lùi lịch sử nếu cần.
        r = requests.get("https://www.nso.gov.vn/bao-cao-tinh-hinh-kinh-te-xa-hoi-hang-thang/",
                          headers={"User-Agent": UA}, timeout=20, verify=False)
        r.raise_for_status()
        links = re.findall(
            r'href="(https://www\.nso\.gov\.vn/(?:bai-top|du-lieu-va-so-lieu-thong-ke)/\d{4}/\d{2}/'
            r'[^"]*(?:bao-cao-tinh-hinh-kinh-te-xa-hoi|thong-cao-bao-chi)[^"]*)"', r.text)
        if not links:
            print("  [INFO] NSO (VN): chưa tìm thấy bài thông cáo báo chí mới trên trang danh sách.")
            return {}
        latest_url = links[0]

        r2 = requests.get(latest_url, headers={"User-Agent": UA}, timeout=20, verify=False)
        r2.raise_for_status()
        text = re.sub(r"<[^>]+>", " ", r2.text)
        text = re.sub(r"&#\d+;", " ", text)
        # NSO đôi khi ghi dấu thanh điệu kiểu TỔ HỢP (vd "quý" = q+u+y+dấu-sắc-rời U+0301) thay vì
        # ký tự ĐÃ GHÉP SẴN ("ý" = U+00FD) NGAY TRONG CÙNG 1 TRANG — phát hiện khi debug regex
        # "quý III/2026..." không khớp dù mắt thường thấy giống hệt chữ đã gõ. Chuẩn hoá NFC trước
        # khi regex để tránh lặp lại lỗi này ở các cụm từ khác.
        text = unicodedata.normalize("NFC", text)
        text = re.sub(r"\s+", " ", text)

        out = {"source_url": latest_url}

        m0 = re.search(
            r"Tổng sản phẩm trong nước \(GDP\) quý ([IVX]+)/(\d{4})[^.]*?"
            r"tốc độ tăng ước đạt ([\d,]+)% so với cùng kỳ năm trước", text)
        if m0:
            roman_to_int = {"I": 1, "II": 2, "III": 3, "IV": 4}
            q = roman_to_int.get(m0.group(1))
            if q:
                out["gdp_growth_period"] = f"{m0.group(2)}-Q{q}"
                out["gdp_growth"] = float(m0.group(3).replace(",", "."))

        m = re.search(
            r"Về cơ cấu nền kinh tế ([^,]+?), khu vực nông, lâm nghiệp và thủy sản chiếm tỷ trọng "
            r"([\d,]+)%; khu vực công nghiệp và xây dựng chiếm ([\d,]+)%; khu vực dịch vụ chiếm "
            r"([\d,]+)%; thuế sản phẩm trừ trợ cấp sản phẩm chiếm ([\d,]+)%", text)
        if m:
            period = _nso_cumulative_period(m.group(1))
            if period:
                out["period"] = period
                out["gdp_share_agri"] = float(m.group(2).replace(",", "."))
                out["gdp_share_industry"] = float(m.group(3).replace(",", "."))
                out["gdp_share_services"] = float(m.group(4).replace(",", "."))
                out["gdp_share_tax"] = float(m.group(5).replace(",", "."))

        m2 = re.search(
            r"Vốn khu vực Nhà nước đạt [\d.,]+ nghìn tỷ đồng, chiếm ([\d,]+)% tổng vốn.*?"
            r"khu vực ngoài Nhà nước đạt [\d.,]+ nghìn tỷ đồng, chiếm ([\d,]+)%.*?"
            r"khu vực có vốn đầu tư trực tiếp nước ngoài đạt [\d.,]+ nghìn tỷ đồng, chiếm ([\d,]+)%", text)
        if m2:
            out["investment_share_state"] = float(m2.group(1).replace(",", "."))
            out["investment_share_private"] = float(m2.group(2).replace(",", "."))
            out["investment_share_fdi"] = float(m2.group(3).replace(",", "."))

        m3 = re.search(
            r"Tổng vốn đầu tư nước ngoài đăng ký vào Việt Nam.*?đạt ([\d,]+) tỷ USD, tăng ([\d,]+)%", text)
        if m3:
            out["fdi_registered_usd_bn"] = float(m3.group(1).replace(",", "."))

        # Vốn đầu tư thực hiện TOÀN XÃ HỘI theo GIÁ TRỊ TUYỆT ĐỐI (nghìn tỷ đồng), RIÊNG TỪNG QUÝ
        # (không phải lũy kế) — khác hẳn m/m2 ở trên (đó là % TĂNG TRƯỞNG và % CƠ CẤU). User
        # (2026-07-28): "% cơ cấu không biết đầu tư toàn xã hội giai đoạn này có mạnh hơn trước
        # không — cần thêm cột giá trị thực tế". Câu có 2 dạng: "Vốn đầu tư thực hiện toàn xã hội
        # quý N/YYYY theo giá hiện hành ước đạt X nghìn tỷ đồng, tăng Y%" (báo cáo Q1 độc lập, hoặc
        # câu ĐẦU trong báo cáo Q3+9M/Q4+FY) HOẶC nằm trong ngoặc đơn giữa câu lũy kế nửa năm (báo
        # cáo Q2+H1): "...(quý II/YYYY theo giá hiện hành ước đạt X nghìn tỷ đồng, tăng Y%)..." — cả
        # 2 dạng đều khớp cùng 1 regex vì không yêu cầu cụm dẫn "Vốn đầu tư..." đứng NGAY TRƯỚC "quý".
        m5 = re.search(
            r"qu[ýy]\s*([IVX]+)\s*/\s*(\d{4})[^.]{0,40}theo gi[áa] hi[ệe]n h[àa]nh [ưu][ớo]c đạt"
            r"\s*([\d.,]+)\s*ngh[ìi]n tỷ đồng,\s*tăng\s*([\d.,]+)%", text)
        if m5:
            roman_to_int2 = {"I": 1, "II": 2, "III": 3, "IV": 4}
            q2 = roman_to_int2.get(m5.group(1))
            if q2:
                out["investment_value_total_social_period"] = f"{int(m5.group(2)):04d}-Q{q2}"
                out["investment_value_total_social"] = _vn_number(m5.group(3))

        # Tỷ lệ giải ngân vốn đầu tư công (lũy kế, % kế hoạch năm) — CHỈ có ở báo cáo THÁNG (Q/6T/
        # 9T/cả năm không có câu này, xem note của public_investment_disbursement_rate trong
        # vimo_raw.json) nên latest_url ở trên đôi khi là báo cáo quý -> m4/m4b không khớp, bỏ
        # qua nhẹ nhàng (đã đúng ý, không phải lỗi). "Tính chung N tháng đầu năm" luôn nêu số tháng
        # bằng SỐ hoặc CHỮ (vd "bốn tháng", "mười một tháng") -> tự suy ra tháng báo cáo từ N, thay
        # vì phải regex riêng cụm "tháng Tư"/"tháng 4"/"tháng Mười Một" (nhiều biến thể hơn).
        m4 = re.search(
            r"Tính chung (\S+(?:\s+một)?) tháng (?:đầu )?năm (\d{4}), vốn đầu tư thực hiện từ nguồn "
            r"ngân sách Nhà nước ước đạt ([\d.,]+) nghìn tỷ đồng, bằng ([\d.,]+)% kế hoạch năm", text)
        m4b = None if m4 else re.search(
            r"Vốn đầu tư thực hiện từ nguồn ngân sách Nhà nước tháng 01/(\d{4}) ước đạt ([\d.,]+) nghìn "
            r"tỷ đồng, bằng ([\d.,]+)% kế hoạch năm", text)
        if m4:
            month = _VN_MONTH_COUNT_WORDS.get(m4.group(1).lower())
            if month:
                out["public_investment_disbursement_period"] = f"{int(m4.group(2)):04d}-{month:02d}"
                out["public_investment_disbursement_value_ty"] = _vn_number(m4.group(3))
                out["public_investment_disbursement_rate_pct"] = _vn_number(m4.group(4))
        elif m4b:
            out["public_investment_disbursement_period"] = f"{int(m4b.group(1)):04d}-01"
            out["public_investment_disbursement_value_ty"] = _vn_number(m4b.group(2))
            out["public_investment_disbursement_rate_pct"] = _vn_number(m4b.group(3))

        # FDI GIẢI NGÂN lũy kế theo THÁNG (tỷ USD) — user (2026-08-07) yêu cầu biểu đồ tổng quan vĩ
        # mô cần đường FDI giải ngân lũy kế trong năm. Câu "Vốn đầu tư trực tiếp nước ngoài thực
        # hiện tại Việt Nam N tháng năm YYYY ước đạt X tỷ USD" xuất hiện HÀNG THÁNG trong báo cáo
        # (khác fdi_disbursed hiện có — series đó chỉ cập nhật theo QUÝ/6T/9T/cả năm từ nguồn khác,
        # xem đầu file); cùng chỉ số nên GHI VÀO CHUNG series fdi_disbursed (period 'YYYY-MM' trộn
        # với 'YYYY-Qn/Hn/9M/FY' đã có — _period_sort_key() đã hỗ trợ sẵn kiểu trộn này, xem ở trên).
        m_fdi = re.search(
            r"Vốn đầu tư trực tiếp nước ngoài thực hiện tại Việt Nam (\S+(?:\s+một)?) tháng "
            r"(?:đầu )?năm (\d{4}) ước đạt ([\d.,]+) tỷ USD", text)
        if m_fdi:
            month = _VN_MONTH_COUNT_WORDS.get(m_fdi.group(1).lower())
            if month:
                out["fdi_disbursed_period"] = f"{int(m_fdi.group(2)):04d}-{month:02d}"
                out["fdi_disbursed_usd_bn"] = _vn_number(m_fdi.group(3))

        # IIP (Chỉ số sản xuất công nghiệp) tăng trưởng YoY THEO THÁNG riêng lẻ — dự phòng/bổ sung
        # cho iip_growth (nguồn chính vẫn là fetch_nso_chart_embed("index-of-industrial-production"),
        # NHƯNG trang embed chỉ giữ cửa sổ ~13 tháng gần nhất, không lùi được xa hơn — câu này trong
        # từng báo cáo tháng giúp lấp khoảng trống lịch sử khi cần backfill nhiều tháng cùng lúc.
        # CHỈ điền nếu chưa có sẵn giá trị cho kỳ đó (xem nơi gọi) — nguồn chart-embed đáng tin hơn.
        m_iip = re.search(
            r"Chỉ số sản xuất công nghiệp \(IIP\) tháng \S+ ước (?:tính )?tăng [\d.,]+% so với tháng "
            r"trước và tăng ([\d.,]+)% so với cùng kỳ năm trước", text)
        if m_iip and out.get("public_investment_disbursement_period"):
            out["iip_growth_period"] = out["public_investment_disbursement_period"]
            out["iip_growth_pct"] = _vn_number(m_iip.group(1))
        elif m_iip and out.get("fdi_disbursed_period"):
            out["iip_growth_period"] = out["fdi_disbursed_period"]
            out["iip_growth_pct"] = _vn_number(m_iip.group(1))

        # Tổng mức bán lẻ hàng hóa và doanh thu dịch vụ tiêu dùng — user (2026-08-08) yêu cầu dữ
        # liệu THEO TỪNG THÁNG (không phải lũy kế) + tăng trưởng YoY từng tháng, nguồn chính đề
        # xuất là vnanet.vn nhưng khảo sát cho thấy các bài đó chỉ là ẢNH (infographic) hoặc bài
        # báo phái sinh — số liệu GỐC nằm sẵn TRONG CHÍNH báo cáo NSO đang fetch (câu mẫu xác nhận
        # qua báo cáo T7/2026: "Tổng mức bán lẻ hàng hóa và doanh thu dịch vụ tiêu dùng theo giá
        # hiện hành tháng Bảy ước đạt 669,1 nghìn tỷ đồng, tăng 0,9% so với tháng trước và tăng
        # 14,5% so với cùng kỳ năm trước.") nên KHÔNG cần crawl vnanet.vn riêng — dùng lại đúng hạ
        # tầng sitemap-backfill đã có cho FDI/IIP. Tự suy kỳ báo cáo (year, month) từ câu LŨY KẾ đi
        # kèm ngay sau đó ("Tính chung N tháng năm YYYY, tổng mức bán lẻ...") thay vì phải parse
        # tên tháng chữ ("tháng Bảy"/"tháng Mười Một") — tách biệt hoàn toàn khỏi period của GDP/
        # đầu tư công ở trên để không phụ thuộc các regex kia có khớp hay không.
        m_retail_cum = re.search(
            r"Tính chung (\S+(?:\s+một)?) tháng (?:đầu )?năm (\d{4}), tổng mức bán lẻ hàng hóa và "
            r"doanh thu dịch vụ tiêu dùng theo giá hiện hành ước đạt ([\d.,]+) nghìn tỷ đồng, tăng "
            r"([\d.,]+)% so với cùng kỳ năm trước", text)
        m_retail_month = re.search(
            r"[Tt]ổng mức bán lẻ hàng hóa và doanh thu dịch vụ tiêu dùng theo giá hiện hành (?:trong )?"
            r"tháng [^\d,]+? ước đạt ([\d.,]+) nghìn tỷ đồng, tăng ([\d.,]+)% so với tháng trước(?:,)? "
            r"(?:và )?tăng (?:tới )?([\d.,]+)% so với cùng kỳ năm trước", text)
        retail_period = None
        if m_retail_cum:
            rmonth = _VN_MONTH_COUNT_WORDS.get(m_retail_cum.group(1).lower())
            if rmonth:
                retail_period = f"{int(m_retail_cum.group(2)):04d}-{rmonth:02d}"
        if m_retail_month and (retail_period or out.get("public_investment_disbursement_period")):
            out["retail_sales_period"] = retail_period or out["public_investment_disbursement_period"]
            out["retail_sales_value_ty"] = _vn_number(m_retail_month.group(1))
            out["retail_sales_mom_pct"] = _vn_number(m_retail_month.group(2))
            out["retail_sales_yoy_pct"] = _vn_number(m_retail_month.group(3))

        return out
    except Exception as e:
        print(f"  [WARN] NSO (VN) cơ cấu GDP/đầu tư thất bại: {e}")
        return {}


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


def _strip_diacritics(s):
    """Bỏ dấu tiếng Việt (NFD rồi loại combining marks) — dùng để so khớp NHÃN chữ trong text OCR
    (OCR đọc dấu tiếng Việt không ổn định, dễ rớt/nhầm dấu) — KHÔNG áp dụng cho việc parse SỐ (số
    không có dấu nên không ảnh hưởng)."""
    import unicodedata as _ud
    return "".join(c for c in _ud.normalize("NFD", s) if _ud.category(c) != "Mn")


def _ocr_number_after_label(text_nodiacritic, label_nodiacritic, window=60):
    """Tìm số kiểu VN (chấm=nghìn, phẩy=thập phân, hoặc chỉ có phẩy nếu <1000) xuất hiện trong
    khoảng `window` ký tự SAU nhãn (đã bỏ dấu) — dùng cho text OCR vốn mất định dạng dòng/cột gốc
    của ảnh nên không thể regex theo cấu trúc câu như văn bản thật. Khoảng trắng GIỮA các từ trong
    nhãn được coi là LINH HOẠT (\\s+, khớp cả xuống dòng) vì OCR đọc panel dạng thẻ/hộp thường tách
    dòng khác với văn bản gốc. Trả float hoặc None."""
    label_pattern = re.escape(label_nodiacritic).replace(r"\ ", r"\s+")
    m_label = re.search(label_pattern, text_nodiacritic)
    if not m_label:
        return None
    window_text = text_nodiacritic[m_label.end():m_label.end() + window]
    m = re.search(r"(\d{1,3}(?:[.,]\d{3})*(?:[.,]\d+)?)", window_text)
    if not m:
        return None
    raw_num = m.group(1)
    # Số VN: neu co ca '.' lan ',' thi '.'=nghin, ','=thap phan; neu chi co 1 loai dau phan cach,
    # gia dinh la thap phan (khop cach OCR/anh infographic dang bieu dien, vd "744,7" hoac "1.451,3")
    if "." in raw_num and "," in raw_num:
        return float(raw_num.replace(".", "").replace(",", "."))
    if "," in raw_num:
        return float(raw_num.replace(",", "."))
    if "." in raw_num:
        # chi co dau cham: co the la phan cach nghin (vd "1.451") HOAC thap phan kieu US - uu tien
        # nghin neu >=4 chu so nguyen truoc dau cham cuoi, it gap trong bo so nay nen coi la nghin
        return float(raw_num.replace(".", ""))
    return float(raw_num)


NSO_INFOGRAPHIC_LISTING_URL = "https://www.nso.gov.vn/do-hoa-thong-tin/"


def fetch_nso_infographic_investment():
    """Đọc cơ cấu vốn đầu tư TOÀN XÃ HỘI (Nhà nước/Ngoài NN/FDI) từ ẢNH infographic quý của NSO
    (nso.gov.vn/do-hoa-thong-tin) — CHỈ CÓ Ở DẠNG ẢNH, không có text tương ứng ở bất kỳ bài báo cáo
    nào khác đã khảo sát (xem note của investment_share_state trong vimo_raw.json). User (2026-07-30)
    tự tìm 9 bài infographic + tôi đọc số bằng mắt (vision) để backfill lịch sử — hàm này TỰ ĐỘNG
    HÓA việc đó bằng OCR (tesseract, cần cài tesseract-ocr + tesseract-ocr-vie, xem
    update_vimo.yml) cho các kỳ SAU này. RỦI RO ĐÃ BIẾT (user chấp nhận 2026-07-30): OCR ảnh
    infographic nhiều màu/font trang trí có thể đọc sai số — hàm này tự KIỂM TRA CHÉO (state% +
    private% + fdi% phải ~100, tổng 3 giá trị tuyệt đối phải ~bằng Tổng số đọc được) và BỎ QUA
    (trả None) nếu không khớp, thay vì ghi số có thể sai vào dữ liệu.

    Quy trình: (1) quét trang danh mục lấy link infographic MỚI NHẤT có chữ 'quy' trong URL (chỉ
    báo cáo quý mới có cơ cấu đầu tư, báo cáo tháng thì không); (2) tải TẤT CẢ ảnh panel trong bài
    đó (KHÔNG cố định theo tên file/số thứ tự panel — đã xác nhận thủ công tên file 'DT-XNK-CPI'
    đôi khi bị gán NHẦM ảnh ở 1 số bài "final" cuối năm); (3) OCR từng ảnh, panel nào có cụm
    'Vốn đầu tư thực hiện toàn xã hội' mới là panel đúng; (4) trích Tổng số/Nhà nước/Ngoài NN/FDI
    bằng cách tìm số xuất hiện GẦN SAU mỗi nhãn (text OCR mất cấu trúc dòng/cột gốc nên không
    regex theo câu được như văn bản thật).

    Trả {"period": "YYYY-Qn", "state_pct":.., "private_pct":.., "fdi_pct":.., "source_url": ...}
    hoặc None nếu bất kỳ bước nào thất bại/không đủ tin cậy (không lỗi, không crash pipeline)."""
    try:
        import pytesseract
        from PIL import Image
        import io
    except ImportError:
        print("  [WARN] Thiếu pytesseract/Pillow (pip install pytesseract Pillow) — bỏ qua OCR infographic NSO.")
        return None

    try:
        r = requests.get(NSO_INFOGRAPHIC_LISTING_URL, headers={"User-Agent": UA}, timeout=20, verify=False)
        r.raise_for_status()
        links = re.findall(r'href="(https://www\.nso\.gov\.vn/[^"]*infographic[^"]*)"', r.text)
        # chỉ bài QUÝ (có "quy" trong slug URL) mới có cơ cấu đầu tư — báo cáo THÁNG không có mục này
        quarterly_links = [u for u in dict.fromkeys(links) if "quy" in u.lower()]
        if not quarterly_links:
            print("  [INFO] NSO infographic: không tìm thấy bài quý nào trên trang danh mục.")
            return None
        # link đầu tiên trong danh sách là MỚI NHẤT (trang liệt kê theo thứ tự đăng, mới nhất trước)
        article_url = quarterly_links[0]

        r2 = requests.get(article_url, headers={"User-Agent": UA}, timeout=20, verify=False)
        r2.raise_for_status()
        panel_urls = re.findall(r'data-orig-src="(https://www\.nso\.gov\.vn/wp-content/uploads/[^"]+\.(?:png|jpg|jpeg))"', r2.text)
        if not panel_urls:
            print(f"  [WARN] NSO infographic {article_url}: không tìm thấy ảnh panel nào.")
            return None

        # \s+ giữa các từ (KHÔNG dùng literal " ") vì OCR panel dạng thẻ/hộp thường tách dòng khác
        # hẳn văn bản gốc — cùng lý do với _ocr_number_after_label().
        anchor_pattern = re.escape(_strip_diacritics("Vốn đầu tư thực hiện toàn xã hội").lower()).replace(r"\ ", r"\s+")
        target_text = None
        for panel_url in panel_urls:
            try:
                img_r = requests.get(panel_url, headers={"User-Agent": UA}, timeout=30, verify=False)
                img_r.raise_for_status()
                img = Image.open(io.BytesIO(img_r.content))
                ocr_text = pytesseract.image_to_string(img, lang="vie+eng")
            except Exception as e:
                print(f"  [WARN] OCR panel {panel_url} thất bại: {e}")
                continue
            ocr_nodiacritic = re.sub(r"\s+", " ", _strip_diacritics(ocr_text).lower())
            if re.search(anchor_pattern, ocr_nodiacritic):
                target_text = ocr_nodiacritic
                print(f"  -> Panel đúng: {panel_url}")
                break

        if target_text is None:
            print(f"  [WARN] NSO infographic {article_url}: không panel nào OCR ra đúng cụm 'Vốn đầu tư thực hiện toàn xã hội'.")
            return None

        total = _ocr_number_after_label(target_text, "tong so")
        state = _ocr_number_after_label(target_text, "nha nuoc")
        private = _ocr_number_after_label(target_text, "ngoai nn")
        fdi = _ocr_number_after_label(target_text, "fdi")
        if None in (total, state, private, fdi) or total <= 0:
            print(f"  [WARN] NSO infographic {article_url}: OCR thiếu 1 trong 4 số (Tổng/Nhà nước/Ngoài NN/FDI).")
            return None

        # KIỂM TRA CHÉO trước khi tin OCR — sai số cho phép 3% (làm tròn ảnh + OCR)
        sum_parts = state + private + fdi
        if abs(sum_parts - total) / total > 0.03:
            print(f"  [WARN] NSO infographic {article_url}: tổng 3 phần ({sum_parts}) lệch quá 3% so với Tổng số OCR ({total}) — bỏ qua, nghi OCR sai.")
            return None

        # suy ra kỳ báo cáo (Q1/H1/9M/FY) từ chính URL bài viết (không suy từ OCR — URL đáng tin hơn)
        year_m = re.search(r"/(\d{4})/\d{2}/", article_url)
        if not year_m:
            return None
        year = year_m.group(1)
        slug = article_url.lower()
        period = None
        if "quy-i-nam" in slug or re.search(r"quy-i-\d{4}", slug):
            period = f"{year}-Q1"
        elif "sau-thang" in slug or "6-thang" in slug:
            period = f"{year}-H1"
        elif "chin-thang" in slug or "9-thang" in slug:
            period = f"{year}-9M"
        elif "va-nam" in slug:
            period = f"{year}-FY"
        if period is None:
            print(f"  [WARN] NSO infographic {article_url}: không suy được kỳ báo cáo từ URL.")
            return None

        return {
            "period": period,
            "state_pct": round(state / total * 100, 2),
            "private_pct": round(private / total * 100, 2),
            "fdi_pct": round(fdi / total * 100, 2),
            "source_url": article_url,
        }
    except Exception as e:
        print(f"  [WARN] NSO infographic OCR thất bại: {e}")
        return None


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


def _fetch_vbma_csv_text(url):
    """Tải 1 file CSV tĩnh của vbma.org.vn và decode đúng chuẩn (UTF-16LE có BOM, server không
    khai báo charset — Content-Type trả về application/octet-stream) — dùng chung cho mọi hàm
    fetch_vbma_*. Trả str (đã bỏ BOM) hoặc None nếu lỗi."""
    r = requests.get(url, headers={"User-Agent": UA}, timeout=20)
    r.raise_for_status()
    return r.content.decode("utf-16-le").lstrip("﻿")


def _parse_vbma_wide_row(text, row_label):
    """Parse 1 dòng dữ liệu trong file CSV 'wide' của VBMA (dòng 1 = header các kỳ, các dòng
    sau = 1 chỉ báo/dòng, cột đầu là tên chỉ báo). Trả list [(header_raw, value), ...] khớp
    đúng dòng có nhãn == row_label (so khớp chính xác sau khi strip VÀ bỏ dấu ngoặc kép bao
    ngoài — nhãn có dấu phẩy như '"Nhà, điện, nước"' bị bọc quote dù file là TSV), bỏ ô rỗng.
    Trả [] nếu không tìm thấy dòng hoặc file rỗng."""
    lines = text.splitlines()
    if len(lines) < 2:
        return []
    headers = [h.strip() for h in lines[0].split("\t")]
    for line in lines[1:]:
        cols = line.split("\t")
        if not cols or cols[0].strip().strip('"') != row_label:
            continue
        out = []
        for h, v in zip(headers[1:], cols[1:]):
            v = v.strip()
            if not v:
                continue
            out.append((h, v))
        return out
    return []


def _vbma_num(raw):
    """'"19,818,534"' / '5.86%' / '-3.54' -> float. Bỏ dấu ngoặc kép bao ngoài (một số bảng
    dạng số tuyệt đối lớn có dấu phẩy ngăn cách nghìn được bọc trong "..." dù file là TSV),
    dấu phẩy ngăn cách nghìn, và ký hiệu %."""
    return float(raw.strip().strip('"').replace(",", "").replace("%", "").strip())


def fetch_vbma_money_supply():
    """vbma.org.vn/vi/market-data/money-supply — Hiệp hội Thị trường Trái phiếu VN nhúng bảng
    CUNG TIỀN M2 THEO THÁNG dưới dạng file tĩnh (không cần đăng nhập/API key/JS render):
    https://vbma.org.vn/csv/markets/tables/vi/tong_cung_tien_theo_thang.csv — trả về TOÀN BỘ
    lịch sử (T12/2018 → hiện tại, mới nhất T4/2026 tính đến 2026-07-23) thay vì chỉ 1 điểm/lần
    chạy như VietnamBiz cũ (xem note cũ trong vimo_raw.json['m2_growth']). Cột: kỳ (Txx yyyy),
    M2 (tỷ VND), % MoM, % YoY, % YTD, Tiền gửi TCKT, Tiền gửi dân cư — file là TSV mã hoá
    UTF-16LE có BOM (Content-Type trả về là application/octet-stream, không tự declare charset
    nên PHẢI decode thủ công, không dùng r.text). Trả list [(period_iso, yoy_value), ...] mới
    nhất đứng cuối, hoặc [] nếu thất bại."""
    url = "https://vbma.org.vn/csv/markets/tables/vi/tong_cung_tien_theo_thang.csv"
    try:
        r = requests.get(url, headers={"User-Agent": UA}, timeout=20)
        r.raise_for_status()
        text = r.content.decode("utf-16-le").lstrip("﻿")
        lines = text.splitlines()
        if len(lines) < 2:
            print("  [WARN] VBMA cung tiền M2: file rỗng hoặc đổi cấu trúc.")
            return []
        out = []
        for line in lines[1:]:
            cols = line.split("\t")
            if len(cols) < 4:
                continue
            m = re.match(r"T(\d{1,2})\s+(\d{4})", cols[0].strip())
            if not m:
                continue
            period = f"{m.group(2)}-{int(m.group(1)):02d}"
            yoy = cols[3].strip().rstrip("%")
            try:
                out.append((period, round(float(yoy), 2)))
            except ValueError:
                continue
        out.sort(key=lambda t: t[0])
        return out
    except Exception as e:
        print(f"  [WARN] VBMA cung tiền M2 thất bại: {e}")
        return []


def fetch_vbma_deposit_balance():
    """CÙNG file CSV với fetch_vbma_money_supply() (tong_cung_tien_theo_thang.csv) — user
    (2026-08-03) chỉ ra data.vietnambiz.vn/currency-interest-rate có "tăng trưởng huy động" nhưng
    chỉ là snapshot 1 điểm/lần (đã có sẵn ở deposit_growth, nguồn vietnambiz, tích lũy chậm). Phát
    hiện: file CSV VBMA đang dùng cho M2 CÓ SẴN 2 cột 'Tiền gửi TCKT' + 'Tiền gửi dân cư' (tỷ VND,
    theo tháng, từ T12/2018) — CỘNG LẠI ra TỔNG HUY ĐỘNG tuyệt đối, cùng chất lượng lịch sử như
    credit_balance_total, KHÔNG cần fetch thêm nguồn nào khác. Trả list [(period_iso, tong_huy_dong
    tỷ_vnd), ...] hoặc [] nếu thất bại."""
    url = "https://vbma.org.vn/csv/markets/tables/vi/tong_cung_tien_theo_thang.csv"
    try:
        r = requests.get(url, headers={"User-Agent": UA}, timeout=20)
        r.raise_for_status()
        text = r.content.decode("utf-16-le").lstrip("﻿")
        lines = text.splitlines()
        if len(lines) < 2:
            print("  [WARN] VBMA huy động: file rỗng hoặc đổi cấu trúc.")
            return []
        out = []
        for line in lines[1:]:
            cols = line.split("\t")
            if len(cols) < 7:
                continue
            m = re.match(r"T(\d{1,2})\s+(\d{4})", cols[0].strip())
            if not m:
                continue
            period = f"{m.group(2)}-{int(m.group(1)):02d}"
            try:
                total_deposit = round(_vbma_num(cols[5]) + _vbma_num(cols[6]), 0)
                out.append((period, total_deposit))
            except ValueError:
                continue
        out.sort(key=lambda t: t[0])
        return out
    except Exception as e:
        print(f"  [WARN] VBMA huy động thất bại: {e}")
        return []


def fetch_vbma_money_supply_level():
    """CÙNG file CSV với fetch_vbma_money_supply()/fetch_vbma_deposit_balance() — lấy cột M2 tuyệt
    đối (cols[1], tỷ VND) thay vì % YoY. Bổ sung MỨC TUYỆT ĐỐI cho m2_growth (vốn chỉ có %), cùng
    vai trò như credit_balance_total/deposit_balance_total — dùng để suy ra tăng trưởng SO VỚI
    CUỐI NĂM TRƯỚC (YTD) của M2 tại template_vimo.py (user 2026-08-08: muốn xem diễn biến tín
    dụng/M2/huy động TRONG NĂM thay vì so cùng kỳ, vì cùng kỳ năm trước tăng mạnh làm YoY hiện tại
    trông thấp đi không rõ ràng). Trả list [(period_iso, value_ty_vnd), ...] hoặc [] nếu thất bại."""
    url = "https://vbma.org.vn/csv/markets/tables/vi/tong_cung_tien_theo_thang.csv"
    try:
        r = requests.get(url, headers={"User-Agent": UA}, timeout=20)
        r.raise_for_status()
        text = r.content.decode("utf-16-le").lstrip("﻿")
        lines = text.splitlines()
        if len(lines) < 2:
            print("  [WARN] VBMA cung tiền M2 (mức tuyệt đối): file rỗng hoặc đổi cấu trúc.")
            return []
        out = []
        for line in lines[1:]:
            cols = line.split("\t")
            if len(cols) < 2:
                continue
            m = re.match(r"T(\d{1,2})\s+(\d{4})", cols[0].strip())
            if not m:
                continue
            period = f"{m.group(2)}-{int(m.group(1)):02d}"
            try:
                out.append((period, round(_vbma_num(cols[1]), 0)))
            except ValueError:
                continue
        out.sort(key=lambda t: t[0])
        return out
    except Exception as e:
        print(f"  [WARN] VBMA cung tiền M2 (mức tuyệt đối) thất bại: {e}")
        return []


def fetch_vbma_cpi_yoy():
    """vbma.org.vn/vi/market-data/cpi — file 'wide' (1 dòng/chỉ báo, cột = kỳ) chứa CPI YoY THEO
    THÁNG từ T1/2020 (dài hơn nhiều so với cửa sổ ~13 điểm của biểu đồ nhúng NSO hiện dùng), mới
    nhất T6/2026 tính đến 2026-07-23, khớp giá trị với nso.gov.vn (4.69%) — xác nhận đáng tin cậy.
    Trả list [(period_iso, value), ...] hoặc [] nếu thất bại."""
    url = "https://vbma.org.vn/csv/markets/charts/vi/lam_phat_so_voi_cung_ky_nam_truoc.csv"
    try:
        text = _fetch_vbma_csv_text(url)
        pairs = _parse_vbma_wide_row(text, "Lạm phát danh nghĩa (so với cùng kì)")
        out = []
        for header, val in pairs:
            m = re.match(r"T(\d{1,2})\s+(\d{4})", header)
            if not m:
                continue
            period = f"{m.group(2)}-{int(m.group(1)):02d}"
            out.append((period, round(_vbma_num(val), 2)))
        out.sort(key=lambda t: t[0])
        return out
    except Exception as e:
        print(f"  [WARN] VBMA CPI YoY thất bại: {e}")
        return []


def fetch_vbma_core_inflation():
    """vbma.org.vn/vi/market-data/cpi — CÙNG FILE với fetch_vbma_cpi_yoy() (lam_phat_so_voi_
    cung_ky_nam_truoc.csv), khác dòng: 'Lạm phát cơ bản' — chỉ báo MỚI, lấp khoảng trống
    core_inflation (trước nay để trống hoàn toàn vì không tìm được nguồn scrape được, xem note
    cũ trong vimo_raw.json — đã khảo sát nso.gov.vn/cong-nghiep/, cpi-vi/, VietnamBiz không ra).
    Theo tháng từ T1/2020. Trả list [(period_iso, value), ...] hoặc [] nếu thất bại."""
    url = "https://vbma.org.vn/csv/markets/charts/vi/lam_phat_so_voi_cung_ky_nam_truoc.csv"
    try:
        text = _fetch_vbma_csv_text(url)
        pairs = _parse_vbma_wide_row(text, "Lạm phát cơ bản")
        out = []
        for header, val in pairs:
            m = re.match(r"T(\d{1,2})\s+(\d{4})", header)
            if not m:
                continue
            period = f"{m.group(2)}-{int(m.group(1)):02d}"
            out.append((period, round(_vbma_num(val), 2)))
        out.sort(key=lambda t: t[0])
        return out
    except Exception as e:
        print(f"  [WARN] VBMA lạm phát cơ bản thất bại: {e}")
        return []


# Tên nhóm hàng trong dong_gop_vao_lam_phat.csv (VBMA) -> hậu tố key trong vimo_raw.json. File
# này cho ĐIỂM PHẦN TRĂM mỗi nhóm hàng ĐÓNG GÓP vào mức tăng CPI chung (so với cùng kỳ), KHÁC
# cpi_yoy (chỉ số tổng) — đây là phần "kết cấu" (decomposition) mà cpi_yoy không thể hiện được.
VBMA_CPI_CONTRIB_GROUPS = {
    "Thực phẩm": "food",
    "Nhà, điện, nước": "housing_utilities",
    "Y tế": "healthcare",
    "Vận tải": "transport",
    "Khác": "other",
}


def fetch_vbma_cpi_contribution():
    """vbma.org.vn/vi/market-data/cpi — dong_gop_vao_lam_phat.csv: ĐÓNG GÓP (điểm %) của 5 nhóm
    hàng chính vào mức tăng CPI chung theo tháng, từ T1/2020 — đây là 'KẾT CẤU CPI' (decomposition)
    mà cpi_yoy (chỉ số tổng hợp) không cho thấy được: vd CPI tăng chủ yếu do nhóm nào kéo. Trả
    dict {suffix: [(period_iso, value), ...]} theo VBMA_CPI_CONTRIB_GROUPS, hoặc {} nếu thất bại."""
    url = "https://vbma.org.vn/csv/markets/charts/vi/dong_gop_vao_lam_phat.csv"
    try:
        text = _fetch_vbma_csv_text(url)
        out = {}
        for label, suffix in VBMA_CPI_CONTRIB_GROUPS.items():
            pairs = _parse_vbma_wide_row(text, label)
            pts = []
            for header, val in pairs:
                m = re.match(r"T(\d{1,2})\s+(\d{4})", header)
                if not m:
                    continue
                period = f"{m.group(2)}-{int(m.group(1)):02d}"
                pts.append((period, round(_vbma_num(val), 2)))
            pts.sort(key=lambda t: t[0])
            if pts:
                out[suffix] = pts
        return out
    except Exception as e:
        print(f"  [WARN] VBMA đóng góp vào lạm phát thất bại: {e}")
        return {}


def fetch_vbma_gdp_growth():
    """vbma.org.vn/vi/market-data/gdp-growth — file 'wide' chứa TỐC ĐỘ TĂNG TRƯỞNG GDP THỰC TẾ
    THEO QUÝ từ Q1/2015 (dài hơn nhiều so với nguồn tin tức lẻ tẻ hiện dùng), mới nhất Q2/2026
    (8.4%, khớp với điểm 2026-Q2=8.39 đang có trong vimo_raw.json). Trả list
    [(period_iso 'YYYY-Qn', value), ...] hoặc [] nếu thất bại."""
    url = "https://vbma.org.vn/csv/markets/charts/vi/toc_do_tang_truong_gdp_thuc_te_(quy).csv"
    try:
        text = _fetch_vbma_csv_text(url)
        pairs = _parse_vbma_wide_row(text, "Tốc độ tăng trưởng GDP thực tế (quý)")
        out = []
        for header, val in pairs:
            m = re.match(r"Q(\d)\s+(\d{4})", header)
            if not m:
                continue
            period = f"{m.group(2)}-Q{m.group(1)}"
            out.append((period, round(_vbma_num(val), 2)))
        out.sort(key=lambda t: t[0])
        return out
    except Exception as e:
        print(f"  [WARN] VBMA GDP growth thất bại: {e}")
        return []


def fetch_vbma_pmi():
    """vbma.org.vn/vi/market-data/gdp-growth (cùng trang GDP, biểu đồ PMI riêng) — file 'wide'
    chứa PMI SẢN XUẤT THEO THÁNG từ 1/2016 (dài hơn nhiều so với VietnamBiz hiện dùng, mới tích
    lũy được 5 điểm), mới nhất T6/2026 = 51.8 (khớp VietnamBiz). Header dạng 'D/M/YYYY' (D luôn
    =1, ví dụ '1/6/2026' = tháng 6/2026). Trả list [(period_iso, value), ...] hoặc [] nếu thất
    bại."""
    url = "https://vbma.org.vn/csv/markets/charts/vi/pmi.csv"
    try:
        text = _fetch_vbma_csv_text(url)
        pairs = _parse_vbma_wide_row(text, "PMI")
        out = []
        for header, val in pairs:
            m = re.match(r"\d{1,2}/(\d{1,2})/(\d{4})", header)
            if not m:
                continue
            period = f"{m.group(2)}-{int(m.group(1)):02d}"
            out.append((period, round(_vbma_num(val), 2)))
        out.sort(key=lambda t: t[0])
        return out
    except Exception as e:
        print(f"  [WARN] VBMA PMI thất bại: {e}")
        return []


def fetch_vbma_credit_balance():
    """vbma.org.vn/vi/market-data/credit — bảng chi tiết DƯ NỢ TÍN DỤNG TOÀN NỀN KINH TẾ theo
    tháng (cột 'Tổng dư nợ', tỷ VND) — chỉ báo MỚI, bổ sung cho credit_growth (%, đã có từ SBV)
    một góc nhìn về QUY MÔ tuyệt đối. Trả list [(period_iso, value_ty_vnd), ...] hoặc [] nếu
    thất bại."""
    url = "https://vbma.org.vn/csv/markets/tables/vi/du_no_tin_dung_theo_nganh_nghe.csv"
    try:
        text = _fetch_vbma_csv_text(url)
        lines = text.splitlines()
        if len(lines) < 2:
            print("  [WARN] VBMA dư nợ tín dụng: file rỗng hoặc đổi cấu trúc.")
            return []
        out = []
        for line in lines[1:]:
            cols = line.split("\t")
            if len(cols) < 2:
                continue
            m = re.match(r"T(\d{1,2})\s+(\d{4})", cols[0].strip())
            if not m:
                continue
            period = f"{m.group(2)}-{int(m.group(1)):02d}"
            try:
                out.append((period, round(_vbma_num(cols[1]), 0)))
            except ValueError:
                continue
        out.sort(key=lambda t: t[0])
        return out
    except Exception as e:
        print(f"  [WARN] VBMA dư nợ tín dụng thất bại: {e}")
        return []


def _fetch_vbma_rolling_yearly_chart(url, value_row_regex, unit_scale=1.0):
    """Nhiều biểu đồ VBMA (FDI đăng ký, giải ngân đầu tư công...) dùng CHUNG 1 layout: header
    T1..T12, các dòng '2025_'/'2026_' là giá trị LŨY KẾ TỪ ĐẦU NĂM theo tháng (chỉ 2 năm gần
    nhất — cửa sổ trượt, KHÔNG có lịch sử xa hơn), dòng cuối '% <năm sau>/<năm trước>' là YoY —
    hàm này lấy các dòng năm (khớp regex '^(\\d{4})_?$') và trả
    {period_iso 'YYYY-MM': value_luy_ke}. Dùng value_row_regex để chọn đúng dòng (vd r'^\\d{4}_?$'
    cho giá trị tuyệt đối, hoặc r'^%\\s' cho dòng YoY)."""
    text = _fetch_vbma_csv_text(url)
    lines = text.splitlines()
    if len(lines) < 2:
        return {}
    headers = [h.strip() for h in lines[0].split("\t")]
    out = {}
    for line in lines[1:]:
        cols = line.split("\t")
        if not cols:
            continue
        label = cols[0].strip()
        if not re.match(value_row_regex, label):
            continue
        year_m = re.match(r"(\d{4})", label)
        if not year_m:
            continue
        year = year_m.group(1)
        for h, v in zip(headers[1:], cols[1:]):
            v = v.strip()
            if not v:
                continue
            hm = re.match(r"T(\d{1,2})", h)
            if not hm:
                continue
            period = f"{year}-{int(hm.group(1)):02d}"
            try:
                out[period] = round(_vbma_num(v) * unit_scale, 4)
            except ValueError:
                continue
    return out


def fetch_vbma_fdi_registered():
    """vbma.org.vn/vi/market-data/fdi — FDI ĐĂNG KÝ lũy kế theo tháng (tỷ USD), chỉ báo MỚI
    (chưa có trong vimo_raw.json — trước nay chỉ theo dõi FDI GIẢI NGÂN). Cửa sổ trượt 2 năm gần
    nhất (không có lịch sử xa hơn qua nguồn này). Trả list [(period_iso 'YYYY-MM', value), ...]
    hoặc [] nếu thất bại."""
    url = "https://vbma.org.vn/csv/markets/charts/vi/fdi_dang_ky.csv"
    try:
        data = _fetch_vbma_rolling_yearly_chart(url, r"^\d{4}_?$")
        return sorted(data.items())
    except Exception as e:
        print(f"  [WARN] VBMA FDI đăng ký thất bại: {e}")
        return []


def fetch_vbma_public_investment_growth():
    """vbma.org.vn/vi/market-data/states-budget — GIẢI NGÂN ĐẦU TƯ CÔNG, dòng '% yoy' cho tăng
    trưởng lũy kế so với cùng kỳ theo tháng (chỉ có năm hiện tại so với năm trước trong cửa sổ
    trượt 2 năm) — dùng để BỔ SUNG cho public_investment_growth (hiện chỉ có 1 điểm/lần chạy từ
    VietnamBiz), KHÔNG thay thế lịch sử cũ vì cửa sổ này không lùi xa được. Trả list
    [(period_iso 'YYYY-MM', value_pct), ...] hoặc [] nếu thất bại."""
    url = "https://vbma.org.vn/csv/markets/charts/vi/chi_dau_tu_cong.csv"
    try:
        text = _fetch_vbma_csv_text(url)
        lines = text.splitlines()
        if len(lines) < 2:
            return []
        headers = [h.strip() for h in lines[0].split("\t")]
        out = []
        # Dòng '% yoy' so sánh năm SAU (mới nhất) với năm trước đó -> gán period theo năm mới nhất
        year_rows = [l.split("\t")[0].strip() for l in lines[1:] if re.match(r"^\d{4}_?$", l.split("\t")[0].strip())]
        latest_year = max(int(y.rstrip("_")) for y in year_rows) if year_rows else None
        for line in lines[1:]:
            cols = line.split("\t")
            if not cols or cols[0].strip() != "% yoy" or latest_year is None:
                continue
            for h, v in zip(headers[1:], cols[1:]):
                v = v.strip()
                if not v:
                    continue
                hm = re.match(r"T(\d{1,2})", h)
                if not hm:
                    continue
                period = f"{latest_year}-{int(hm.group(1)):02d}"
                try:
                    out.append((period, round(_vbma_num(v), 2)))
                except ValueError:
                    continue
        out.sort(key=lambda t: t[0])
        return out
    except Exception as e:
        print(f"  [WARN] VBMA tăng trưởng đầu tư công thất bại: {e}")
        return []


def fetch_vbma_budget_deficit_pct_gdp():
    """vbma.org.vn/vi/market-data/states-budget — bảng thu/chi ngân sách THEO NĂM từ 2015, dòng
    '% GDP' = thặng dư(+)/thâm hụt(-) ngân sách tính theo %GDP mỗi năm — chỉ báo MỚI (chưa có
    trong vimo_raw.json). Trả list [(year_str, value_pct), ...] hoặc [] nếu thất bại."""
    url = "https://vbma.org.vn/csv/markets/charts/vi/thu_chi_ns_theo_nam.csv"
    try:
        text = _fetch_vbma_csv_text(url)
        pairs = _parse_vbma_wide_row(text, "% GDP")
        out = []
        for header, val in pairs:
            m = re.match(r"12T\s+(\d{4})", header)
            if not m:
                continue
            out.append((m.group(1), round(_vbma_num(val), 2)))
        out.sort(key=lambda t: t[0])
        return out
    except Exception as e:
        print(f"  [WARN] VBMA thâm hụt ngân sách/GDP thất bại: {e}")
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


VIRA_BULLETIN_URL_TMPL = ("https://vira.org.vn/tin/Ban-tin-Kinh-te-Tai-chinh-ngay/"
                           "Ban-tin-Kinh-te-Tai-chinh-ngay-{d:02d}-{m:02d}-{y}-.html")


def fetch_vira_bulletin(lookback_days=10):
    """vira.org.vn (Hội Nghiên cứu thị trường liên ngân hàng Việt Nam) — bản tin Kinh tế - Tài
    chính NGÀY, HTML tĩnh (không cần đăng nhập, xác nhận qua khảo sát thủ công 2026-07-28). VIRA
    KHÔNG ra bản tin mỗi ngày (nghỉ cuối tuần + thỉnh thoảng bỏ ngày) và mỗi bản tin ghi số liệu
    của PHIÊN TRƯỚC (vd bản tin đăng 28/07 ghi "Ngày 27/07") — nên quét lùi lookback_days ngày lịch
    theo URL, còn NGÀY THẬT của số liệu lấy từ chính cụm "Ngày DD/MM" trong text, không suy từ
    ngày URL. Mỗi bản tin cho: (1) lãi suất liên ngân hàng VND kỳ hạn ON/1W/2W/1M — đối chiếu
    fetch_sbv_interest_rates() vốn chỉ có snapshot theo TUẦN/THÁNG (user 2026-07-28: "cần dữ liệu
    cập nhật để nhìn xu hướng" — VIRA cho lịch sử NGÀY thật, dùng THAY THẾ 4 kỳ hạn này, KHÔNG cộng
    dồn chung SBV vì khác phương pháp gộp, trộn sẽ ra biểu đồ răng cưa); (2) lợi suất TPCP thứ cấp
    3Y/5Y/7Y/10Y/15Y (chỉ báo MỚI, chưa có nguồn nào khác đang theo dõi); (3) NHNN bơm ròng/hút
    ròng qua OMO kênh cầm cố (tỷ đồng, chỉ báo MỚI — dấu ÂM = hút ròng, DƯƠNG = bơm ròng); (4) lãi
    suất OMO kỳ hạn 7 ngày (thay thế snapshot tuần từ SBV, cùng lý do (1)); (5) số dư OMO đang LƯU
    HÀNH trên kênh cầm cố (tỷ đồng, chỉ báo MỚI — user 2026-07-30 chỉ ra bản tin có luôn số này,
    khác omo_net_operation là DÒNG CHẢY ròng/ngày, đây là TỒN KHO lũy kế tại thời điểm đó).
    Trả list[dict] mỗi phần tử {"date": "YYYY-MM-DD", "source_url": ..., rồi các field nào tìm
    được trong: interbank_on/1w/2w/1m, bond_3y/5y/7y/10y/15y, omo_net, omo_rate, omo_outstanding}
    — field nào không tìm thấy (bản tin đổi cấu trúc/thiếu đoạn) thì bị bỏ qua, không lỗi. Bản tin
    nào không tồn tại (404/302, ngày nghỉ) cũng bỏ qua lặng lẽ."""
    import html as htmlmod
    results = []
    today = datetime.date.today()
    for i in range(lookback_days):
        d = today - datetime.timedelta(days=i)
        url = VIRA_BULLETIN_URL_TMPL.format(d=d.day, m=d.month, y=d.year)
        try:
            r = requests.get(url, headers={"User-Agent": UA}, timeout=15)
            if r.status_code != 200:
                continue
            text = re.sub(r"<script.*?</script>", " ", r.text, flags=re.S)
            text = re.sub(r"<style.*?</style>", " ", text, flags=re.S)
            text = re.sub(r"<[^>]+>", " ", text)
            text = htmlmod.unescape(text)
            text = re.sub(r"\s+", " ", text)

            m_date = re.search(r"Ng[àa]y\s*(\d{1,2})/(\d{1,2}),\s*l[ãa]i suất b[ìi]nh qu[âa]n LNH VND", text)
            if not m_date:
                continue
            day, month = int(m_date.group(1)), int(m_date.group(2))
            year = d.year - 1 if (month == 12 and d.month == 1) else d.year
            entry = {"date": f"{year:04d}-{month:02d}-{day:02d}", "source_url": url}

            m_rates = re.search(
                r"ON\s*([\d,]+)%;\s*1W\s*([\d,]+)%;\s*2W\s*([\d,]+)%\s*v[àa]\s*1M\s*([\d,]+)%", text)
            if m_rates:
                entry["interbank_on"] = float(m_rates.group(1).replace(",", "."))
                entry["interbank_1w"] = float(m_rates.group(2).replace(",", "."))
                entry["interbank_2w"] = float(m_rates.group(3).replace(",", "."))
                entry["interbank_1m"] = float(m_rates.group(4).replace(",", "."))

            m_bond = re.search(
                r"3Y\s*([\d,]+)%;\s*5Y\s*([\d,]+)%;\s*7Y\s*([\d,]+)%;\s*10Y\s*([\d,]+)%;\s*15Y\s*([\d,]+)%", text)
            if m_bond:
                entry["bond_3y"] = float(m_bond.group(1).replace(",", "."))
                entry["bond_5y"] = float(m_bond.group(2).replace(",", "."))
                entry["bond_7y"] = float(m_bond.group(3).replace(",", "."))
                entry["bond_10y"] = float(m_bond.group(4).replace(",", "."))
                entry["bond_15y"] = float(m_bond.group(5).replace(",", "."))

            m_omo = re.search(r"NHNN\s*(h[úu]t r[òo]ng|bơm r[òo]ng)\s*([\d.,]+)\s*tỷ đồng", text)
            if m_omo:
                amt = float(m_omo.group(2).replace(".", "").replace(",", "."))
                entry["omo_net"] = -amt if m_omo.group(1).startswith(("h", "H")) else amt

            m_omo_rate = re.search(r"lãi suất đều ở mức ([\d,]+)%", text)
            if m_omo_rate:
                entry["omo_rate"] = float(m_omo_rate.group(1).replace(",", "."))

            m_outstanding = re.search(r"C[óo]\s*([\d.,]+)\s*tỷ đồng lưu hành trên kênh cầm cố", text)
            if m_outstanding:
                entry["omo_outstanding"] = _vn_number(m_outstanding.group(1))

            if len(entry) > 2:
                results.append(entry)
        except Exception as e:
            print(f"  [WARN] VIRA {d.isoformat()} thất bại: {e}")
    return results


def fetch_sbv_tin_phieu_days_since():
    """sbv.gov.vn/vi/web/sbv_portal/thong-tin-chao-ban-tin-phieu-nhnn — trang THÔNG BÁO BÁN TÍN
    PHIẾU NHNN (kênh HÚT thanh khoản, ĐỐI LẬP với OMO ở fetch_sbv_omo_rate() vốn là kênh BƠM) —
    user (2026-07-13) muốn hệ thống tự nhận biết khi NHNN chuyển sang chế độ hút bớt thanh khoản
    qua tín phiếu, để đánh giá 2 CHIỀU (bơm/hút) thay vì chỉ nhìn 1 chiều OMO. Số liệu thật (đã
    kiểm tra thủ công 2026-07-13): lần thông báo bán tín phiếu gần nhất là 30/10/2025 — nghĩa là
    SUỐT một thời gian dài KHÔNG có hoạt động hút thanh khoản, chỉ có bơm qua OMO. Trả
    (days_since_last, last_date_iso, source_url) hoặc (None, None, None) nếu lỗi/không tìm thấy
    ngày nào. KHÔNG trích được khối lượng/lãi suất tín phiếu đáng tin cậy từ trang này (thông báo
    chào bán khác thông báo kết quả trúng thầu) — chỉ theo dõi NGÀY để biết đang ở chế độ nào."""
    url = "https://www.sbv.gov.vn/vi/web/sbv_portal/thong-tin-chao-ban-tin-phieu-nhnn"
    try:
        r = requests.get(url, headers={"User-Agent": UA}, timeout=20, verify=False)
        r.raise_for_status()
        text = re.sub(r"<[^>]+>", " ", r.text)
        text = re.sub(r"\s+", " ", text)
        dates = re.findall(r"(\d{2}/\d{2}/20\d{2}) \d{2}:\d{2}:\d{2}", text)
        if not dates:
            return None, None, None
        parsed = [datetime.datetime.strptime(d, "%d/%m/%Y") for d in dates]
        latest = max(parsed)
        days_since = (datetime.datetime.now() - latest).days
        return days_since, latest.strftime("%Y-%m-%d"), url
    except Exception as e:
        print(f"  [WARN] SBV tín phiếu (kênh hút thanh khoản) thất bại: {e}")
        return None, None, None


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
# báo mà VietnamBiz là nguồn TỐT NHẤT tìm được (bán lẻ — trước đây "manual" chỉ 1 điểm) — không
# map đè lên GDP/CPI/thất nghiệp/IIP vì NSO (trực tiếp từ cơ quan thống kê) đáng tin cậy hơn
# nguồn tổng hợp lại của bên thứ ba, dù VietnamBiz cũng có các chỉ báo đó làm đối chiếu. PMI
# cũng đã chuyển sang fetch_vbma_pmi() (full lịch sử từ 2016) nên bỏ khỏi map này.
VIETNAMBIZ_TITLE_MAP = {
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
    # m2_growth: chuyển sang fetch_vbma_money_supply() — VBMA có cả lịch sử theo tháng từ
    # T12/2018, không cần tích lũy từng điểm/lần chạy như VietnamBiz nữa.
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


VNINDEX_VALUATION_HISTORY_PATH = os.path.join(PROJECT_ROOT, "data", "vnindex_valuation_history.json")


def fetch_vietcap_index_valuation(val_type):
    """trading.vietcap.com.vn/iq — API NỘI BỘ (không phải trang HTML, xem
    .agents/skills/giaodichvietcap/SKILL.md) cấp lịch sử P/E hoặc P/B THEO NGÀY của VN-Index từ
    22/04/2009 (~4300 điểm), kèm SẴN dải thống kê trung bình/±1SD/±2SD do chính Vietcap tính trên
    toàn bộ lịch sử — user (2026-07-25) yêu cầu dựng biểu đồ so sánh định giá VN-Index theo thời
    gian, và sau khi kiểm chứng bằng chính dữ liệu này phát hiện mô hình CAPM lý thuyết (P/B hợp
    lý theo ROE/COE) đặt ngưỡng THẤP HƠN CẢ ĐÁY của 3 đợt khủng hoảng gần nhất (COVID 3/2020, bear
    2022, sốc thuế quan 4/2025) — nên chuyển sang dùng dải thống kê THỰC TẾ này làm chuẩn tham
    chiếu chính, đáng tin cậy hơn vì đã qua kiểm chứng lịch sử thay vì mô hình chưa backtest.
    val_type: 'PE' hoặc 'PB'. Trả dict {"values": [{"date","value"}...], "average", "plusOneSD",
    "plusTwoSD", "minusOneSD", "minusTwoSD"} hoặc None nếu lỗi."""
    try:
        s = requests.Session()
        s.headers.update({
            "User-Agent": UA,
            "Accept": "application/json, text/plain, */*",
            "Referer": "https://trading.vietcap.com.vn/",
        })
        s.get("https://trading.vietcap.com.vn/iq/market?tab=information", timeout=15)
        url = "https://trading.vietcap.com.vn/api/iq-insight-service/v1/market-watch/index-valuation"
        r = s.get(url, params={"type": val_type, "comGroupCode": "VNINDEX", "timeFrame": "ALL"}, timeout=20)
        r.raise_for_status()
        payload = r.json()
        if not payload.get("successful"):
            print(f"  [WARN] Vietcap IQ {val_type}: API trả successful=false — {payload.get('msg')}")
            return None
        return payload["data"]
    except Exception as e:
        print(f"  [WARN] Vietcap IQ {val_type} thất bại: {e}")
        return None


def update_vnindex_valuation_history(nonvin_data=None):
    """Fetch lại TOÀN BỘ lịch sử P/E + P/B VN-Index HEADLINE (có VIN) từ Vietcap IQ, GHÉP THÊM
    lịch sử ex-VIN (nonvin_data, từ fetch_vnindex_nonvin_data() — GitHub Truongutc/AIC---chart-
    nganh, đã fetch sẵn ở bước trước để tránh gọi API 2 lần) nếu có, rồi ghi đè
    data/vnindex_valuation_history.json — file RIÊNG (không gộp vào vimo_raw.json) vì đây là
    time-series DÀY ĐẶC theo ngày (~4300-6300 điểm/chỉ báo), khác hẳn cấu trúc chỉ báo thưa của
    vimo_raw.json; ghi đè toàn bộ mỗi lần chạy (không tích lũy dần) vì cả 2 nguồn đều trả về TRỌN
    VẸN lịch sử mỗi lần gọi, không cần merge thủ công. Làm tròn 4 chữ số thập phân + KHÔNG indent
    để giảm dung lượng file/git diff mỗi lần chạy. Trả True nếu lấy được ít nhất 1 chuỗi bất kỳ."""
    def _round_data(data):
        if not data:
            return None
        data = dict(data)
        data["values"] = [{"date": p["date"], "value": round(p["value"], 4)} for p in data.get("values", [])]
        for k in ("average", "plusOneSD", "plusTwoSD", "minusOneSD", "minusTwoSD"):
            if data.get(k) is not None:
                data[k] = round(data[k], 4)
        return data

    def _build_exvin_data(dates, raw_values):
        # Vietcap trả sẵn dải average/±1SD/±2SD cho headline — ex-VIN không có nguồn nào tính sẵn
        # nên tự tính bằng statistics.mean/stdev trên TOÀN BỘ lịch sử ex-VIN, CÙNG công thức/quy
        # ước với Vietcap (mean ± n*sample_stdev) để 4 biểu đồ so sánh được với nhau (user
        # 2026-07-25: yêu cầu 4 biểu đồ riêng, mỗi cái tự so với dải/ngưỡng của chính nó).
        values = [{"date": d, "value": round(v, 4)} for d, v in zip(dates, raw_values) if v is not None]
        if len(values) < 2:
            return None
        nums = [p["value"] for p in values]
        avg = statistics.mean(nums)
        sd = statistics.stdev(nums)
        return {
            "values": values,
            "average": round(avg, 4), "plusOneSD": round(avg + sd, 4), "minusOneSD": round(avg - sd, 4),
            "plusTwoSD": round(avg + 2 * sd, 4), "minusTwoSD": round(avg - 2 * sd, 4),
        }

    pe_data = _round_data(fetch_vietcap_index_valuation("PE"))
    pb_data = _round_data(fetch_vietcap_index_valuation("PB"))

    pe_exvin = pb_exvin = None
    if nonvin_data and nonvin_data.get("daily"):
        daily = nonvin_data["daily"]
        pe_exvin = _build_exvin_data(daily["dates"], daily["pe"])
        pb_exvin = _build_exvin_data(daily["dates"], daily["pb"])

    if not pe_data and not pb_data and not pe_exvin and not pb_exvin:
        return False
    out = {
        "_meta": {"source": "trading.vietcap.com.vn (Vietcap IQ) + GitHub Truongutc/AIC---chart-nganh (ex-VIN)",
                  "updated_at": _current_period_weekly()},
        "pe": pe_data,
        "pb": pb_data,
        "pe_exvin": pe_exvin,
        "pb_exvin": pb_exvin,
    }
    with open(VNINDEX_VALUATION_HISTORY_PATH, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))
    return True


def fetch_tcbs_ipower_max_rate():
    """tcbs.com.vn/ca-nhan/san-pham/ipower/ — trang HTML tĩnh THẬT (server-render, không cần JS)
    của sản phẩm 'Tài khoản Nhân lãi' iPower (TCBS/Techcombank). Lãi suất linh hoạt theo số dư +
    giá trị giao dịch lũy kế tháng, MỨC CAO NHẤT quảng cáo ('lên đến X%/năm') — user (2026-07-24)
    yêu cầu lấy mức CAO NHẤT vì đây là kênh 'gửi tiền' thay thế ngoài ngân hàng truyền thống, khi
    hệ thống ngân hàng căng thanh khoản/huy động khó thì TCBS phải tăng lãi suất iPower để cạnh
    tranh hút tiền — mức này HẠ xuống nghĩa là áp lực huy động bên ngoài đã dịu bớt. Trang có ghi
    rõ ngày hiệu lực biểu lãi suất ('Hiệu lực từ ngày DD/MM/YYYY') — dùng làm period thay vì ngày
    fetch (biểu lãi suất có thể đã áp dụng từ trước ngày Action chạy). Trả (rate, period_iso,
    source_url) hoặc (None, None, None) nếu lỗi/không tìm thấy.

    NGOẠI LỆ so với các hàm fetch_* khác trong file này: tcbs.com.vn chặn `requests` bằng
    TLS-fingerprint-based bot detection (đã xác nhận: cùng User-Agent/header y hệt trình duyệt
    nhưng `requests` luôn nhận 403, trong khi `curl` hệt vậy lại trả 200 — khác chữ ký bắt tay
    TLS/JA3 giữa 2 thư viện, KHÔNG phải do thiếu header) — gọi `curl` qua subprocess thay vì
    `requests.get()`. GitHub Actions runner (ubuntu-latest) có sẵn curl, không cần cài thêm gì."""
    url = "https://www.tcbs.com.vn/ca-nhan/san-pham/ipower/"
    try:
        proc = subprocess.run(
            ["curl", "-sL", "-A", UA, "--max-time", "20", url],
            capture_output=True, timeout=25)
        if proc.returncode != 0 or not proc.stdout:
            print(f"  [WARN] TCBS iPower: curl thất bại (rc={proc.returncode}).")
            return None, None, None
        html = proc.stdout.decode("utf-8", errors="replace")
        text = re.sub(r"<[^>]+>", " ", html)
        text = re.sub(r"&nbsp;", " ", text)
        text = re.sub(r"\s+", " ", text)
        m = re.search(r"lên đến\s*([\d,\.]+)\s*%\s*/\s*năm", text)
        if not m:
            print("  [WARN] TCBS iPower: không tìm thấy 'lên đến X%/năm' — trang có thể đã đổi cấu trúc.")
            return None, None, None
        rate = float(m.group(1).replace(",", "."))
        period = _current_period_weekly()
        m2 = re.search(r"Hiệu lực từ ngày (\d{2})/(\d{2})/(\d{4})", text)
        if m2:
            eff_date = datetime.date(int(m2.group(3)), int(m2.group(2)), int(m2.group(1)))
            period = _current_period_weekly(eff_date)
        return rate, period, url
    except Exception as e:
        print(f"  [WARN] TCBS iPower thất bại: {e}")
        return None, None, None


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


# Thư mục CỤC BỘ (ngoài repo, do user tự tải file Excel Tổng cục Hải quan về đặt vào) — CHỈ tồn
# tại trên máy chạy thủ công, GitHub Action KHÔNG có (Action chỉ checkout đúng repo) nên hàm dưới
# đây LUÔN bỏ qua nhẹ nhàng khi chạy trên Action, không phải lỗi. User cần tự cập nhật file mới vào
# thư mục này định kỳ (Action không tự tải được vì đây không phải nguồn web công khai có URL cố
# định — là file Excel người dùng tự tải từ trang Hải quan/GSO về).
CUSTOMS_XNK_FOLDER = r"E:\1. Projects\Du lieu xnk"


def load_customs_xnk_local():
    """Đọc file Excel THỦ CÔNG (Tổng cục Hải quan — 'Trị giá và mặt hàng xuất/nhập khẩu sơ bộ các
    tháng năm YYYY') do user tự tải về CUSTOMS_XNK_FOLDER — file 'V01-*.xls' = xuất khẩu, 'V02-*.xls'
    = nhập khẩu (V03 = theo nước/khối nước, CHƯA dùng ở đây). Cấu trúc bảng cố định: dòng 0 = tiêu
    đề có 'năm YYYY', dòng 2 = nhãn kỳ ('Tháng 01', 'Tháng 02', ..., rồi 1 cột lũy kế cuối bảng kiểu
    '12 tháng'/'6 tháng' — BỎ QUA cột này vì thứ tự chữ 'N tháng' ngược với 'Tháng N' nên không khớp
    regex, đúng ý), mỗi kỳ chiếm 2 cột (Lượng rồi Trị giá — cột Trị giá luôn NGAY SAU cột Lượng cùng
    kỳ). Dòng 'Tổng số' = TỔNG GIÁ TRỊ theo TỪNG THÁNG RỜI RẠC (đã kiểm tra thủ công: tháng sau có
    thể THẤP hơn tháng trước, vd 2025-02 < 2025-01 — KHÔNG PHẢI lũy kế, khác hẳn fdi_disbursed/
    fdi_registered_usd_bn). Đơn vị gốc 1000 USD -> đổi sang tỷ USD (chia 1e6) cho khớp đơn vị
    trade_balance đã có. Trả {"export": [(period, gia_tri_ty_usd, ten_file)...], "import": [...]}
    — rỗng nếu không tìm thấy thư mục/file (không lỗi, không crash pipeline)."""
    if not os.path.isdir(CUSTOMS_XNK_FOLDER):
        return {"export": [], "import": []}
    try:
        import xlrd
    except ImportError:
        print("  [WARN] Thiếu thư viện xlrd (pip install xlrd) — bỏ qua đọc file Hải quan cục bộ.")
        return {"export": [], "import": []}

    out = {"export": [], "import": []}
    for prefix, key in [("V01", "export"), ("V02", "import")]:
        for fpath in sorted(glob.glob(os.path.join(CUSTOMS_XNK_FOLDER, f"{prefix}-*.xls"))):
            fname = os.path.basename(fpath)
            try:
                sh = xlrd.open_workbook(fpath).sheet_by_index(0)
                year_m = re.search(r"năm\s*(\d{4})", str(sh.cell_value(0, 0)))
                if not year_m:
                    print(f"  [WARN] {fname}: không tìm thấy năm ở dòng tiêu đề — bỏ qua.")
                    continue
                year = int(year_m.group(1))

                total_row = None
                for r in range(sh.nrows):
                    if str(sh.cell_value(r, 0)).strip().lower().startswith("tổng số"):
                        total_row = r
                        break
                if total_row is None:
                    print(f"  [WARN] {fname}: không tìm thấy dòng 'Tổng số' — bỏ qua.")
                    continue

                header_row = [str(sh.cell_value(2, c)) for c in range(sh.ncols)]
                n_points = 0
                for c in range(1, sh.ncols):
                    m = re.match(r"Th[áa]ng\s*0?(\d{1,2})\s*$", header_row[c].strip(), re.I)
                    if not m:
                        continue
                    month = int(m.group(1))
                    value = sh.cell_value(total_row, c + 1)  # cột Trị giá ngay sau cột Lượng
                    if isinstance(value, (int, float)) and value:
                        period = f"{year:04d}-{month:02d}"
                        out[key].append((period, round(value / 1_000_000, 4), fname))
                        n_points += 1
                print(f"  -> {fname}: {n_points} tháng")
            except Exception as e:
                print(f"  [WARN] Đọc file Hải quan {fname} thất bại: {e}")
    return out


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

    print("[FII — khối ngoại mua/bán ròng HOSE]")
    net_ty_vnd, date_iso, src = fetch_fii_net_flow()
    if net_ty_vnd is not None:
        _append_point(raw, "fii_net_flow_hose", date_iso, net_ty_vnd, src)
        print(f"  -> {date_iso}: {net_ty_vnd:+.2f} tỷ VND ({'mua ròng' if net_ty_vnd > 0 else 'bán ròng'})")

    print("[VN-Index P/E & P/B — ƯU TIÊN ex-VIN (loại VIC/VHM/VRE/VPL, ít méo mó hơn headline)]")
    nonvin = fetch_vnindex_nonvin_data()
    if nonvin:
        pe, pb, src = nonvin["pe"], nonvin["pb"], VNINDEX_NONVIN_URL
        print(f"  -> ex-VIN {nonvin['date']}: P/E={pe:.2f} P/B={pb:.2f} ROE(quý gần nhất)={nonvin.get('roe')}")
    else:
        # Fallback headline (CÓ VIN) nếu repo GitHub kia lỗi/không truy cập được — vẫn hơn không
        # có gì, nhưng đã méo mó hơn (xem note trong vimo_raw.json).
        pe, pb, src = fetch_vnindex_pe_pb_24hmoney()
        if not pe:
            pe, src = fetch_vnindex_pe_current()
            pb = None
        if pe:
            print(f"  -> [FALLBACK headline, có VIN] P/E={pe}")
    if pe:
        _append_point(raw, "vnindex_pe", period_now, pe, src)
    if pb:
        _append_point(raw, "vnindex_pb", period_now, pb, src)

    # LUÔN fetch headline (CÓ VIN) riêng, dù ex-VIN đã lấy được — để có cặp P/E, P/B SONG SONG
    # (headline vs ex-VIN) cho 2 quyết định phân bổ vốn độc lập (user 2026-07-25: "chia ra 2
    # quyết định: nếu nhìn vào VN-Index thì quyết định là gì, nếu nhìn theo VN-Index no VIN thì
    # quyết định là gì"). Khác khối trên (chỉ fetch headline làm FALLBACK khi ex-VIN lỗi).
    print("[VN-Index P/E & P/B — headline (CÓ VIN, để so sánh song song với ex-VIN)]")
    pe_head, pb_head, src_head = fetch_vnindex_pe_pb_24hmoney()
    if not pe_head:
        pe_head, src_head = fetch_vnindex_pe_current()
    if pe_head:
        _append_point(raw, "vnindex_pe_headline", period_now, pe_head, src_head)
        print(f"  -> headline {period_now}: P/E={pe_head}")
    if pb_head:
        _append_point(raw, "vnindex_pb_headline", period_now, pb_head, src_head)

    print("[Vietcap IQ — lịch sử P/E & P/B VN-Index headline theo ngày (~17 năm, kèm dải ±1SD/±2SD)]")
    if update_vnindex_valuation_history(nonvin):
        print(f"  -> đã ghi {VNINDEX_VALUATION_HISTORY_PATH}")
    else:
        print("  [WARN] Không lấy được lịch sử P/E/P/B từ Vietcap IQ — giữ nguyên file cũ (nếu có).")

    print("[World Bank — China GDP growth]")
    pts = fetch_worldbank("NY.GDP.MKTP.KD.ZG", "CN", n=8)
    if pts:
        raw["china_gdp_growth"]["series"] = [
            {"period": y, "value": v, "source_url": "https://api.worldbank.org/v2/country/cn/indicator/NY.GDP.MKTP.KD.ZG"}
            for y, v in sorted(pts)
        ]
        print(f"  -> {len(pts)} điểm")

    print("[World Bank — Vietnam nominal GDP (theo năm, dùng tính tỷ lệ Tín dụng/GDP)]")
    pts = fetch_worldbank("NY.GDP.MKTP.CN", "VN", n=10)
    if pts:
        # World Bank trả VND thô — quy đổi tỷ VND cho khớp đơn vị credit_balance_total (cùng tỷ VND)
        raw["nominal_gdp_annual"]["series"] = [
            {"period": y, "value": round(v / 1e9, 2),
             "source_url": "https://api.worldbank.org/v2/country/vn/indicator/NY.GDP.MKTP.CN"}
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

    # "Dữ liệu Quốc tế" — Mỹ/Eurozone/Trung Quốc (user 2026-07-31: đối chiếu vĩ mô VN với 3 thị
    # trường lớn, xem toàn cầu đang hành động như nào; ban đầu chỉ có 7 chỉ báo lãi suất/lợi suất
    # NHTW lớn — nay mở rộng đủ bộ: việc làm, thất nghiệp, CPI/lõi, PCE/lõi (chỉ Mỹ — Eurozone/TQ
    # không công bố PCE), GDP, sản xuất công nghiệp, đường cong lợi suất). units=pc1/chg (xem
    # fetch_fred()) lấy thẳng %YoY / thay đổi kỳ mà không cần tự tính derived-diff.
    print("[FRED — Quốc tế: Mỹ (việc làm, CPI/PCE, GDP, sản xuất CN, đường cong lợi suất)]")
    for key, sid, units in [
        ("us_unemployment_rate", "UNRATE", None),
        ("us_nonfarm_payrolls_change", "PAYEMS", "chg"),
        ("us_cpi_yoy", "CPIAUCSL", "pc1"),
        ("us_core_cpi_yoy", "CPILFESL", "pc1"),
        ("us_pce_yoy", "PCEPI", "pc1"),
        ("us_core_pce_yoy", "PCEPILFE", "pc1"),
        ("us_gdp_growth", "A191RL1Q225SBEA", None),
        ("us_industrial_production_yoy", "INDPRO", "pc1"),
        ("us_10y_yield", "DGS10", None),
        ("us_yield_3m", "DGS3MO", None), ("us_yield_1y", "DGS1", None),
        ("us_yield_2y", "DGS2", None), ("us_yield_5y", "DGS5", None), ("us_yield_30y", "DGS30", None),
    ]:
        pts = fetch_fred(sid, n=24, units=units)
        if pts:
            raw[key]["series"] = [
                {"period": d, "value": v, "source_url": f"https://fred.stlouisfed.org/series/{sid}"}
                for d, v in sorted(pts)
            ]
            print(f"  -> {key}: {len(pts)} điểm")

    print("[FRED — Quốc tế: Eurozone (thất nghiệp, CPI, GDP, sản xuất CN, ECB, lợi suất Đức)]")
    # eu_core_cpi_yoy ĐÃ BỎ (2026-08-01, theo yêu cầu user "dữ liệu cũ không còn hoạt động thì bỏ,
    # vì nó đâu có tác dụng gì đâu") — CPGRLE01EZM659N và mọi biến thể OECD MEI thử qua đều bị
    # discontinued từ 2023-01 trên FRED, không có nguồn thay thế nào còn sống.
    for key, sid, units in [
        ("eu_unemployment_rate", "LRHUTTTTEZM156S", None),
        ("eu_cpi_yoy", "CP0000EZ19M086NEST", "pc1"),
        ("eu_gdp_growth", "NAEXKP01EZQ657S", None),
        ("eu_industrial_production_yoy", "EA19PRINTO01GYSAM", None),
        ("ecb_deposit_rate", "ECBDFR", None),
        ("germany_10y_yield", "IRLTLT01DEM156N", None),
        ("boe_sonia_rate", "IUDSOIA", None),
        ("uk_10y_yield", "IRLTLT01GBM156N", None),
        ("boj_interbank_rate", "IRSTCI01JPM156N", None),
        ("japan_10y_yield", "IRLTLT01JPM156N", None),
    ]:
        pts = fetch_fred(sid, n=24, units=units)
        if pts:
            raw[key]["series"] = [
                {"period": d, "value": v, "source_url": f"https://fred.stlouisfed.org/series/{sid}"}
                for d, v in sorted(pts)
            ]
            print(f"  -> {key}: {len(pts)} điểm")

    # cn_industrial_production_yoy/cn_gdp_growth_quarterly/cn_10y_yield ĐÃ BỎ (2026-08-01, theo
    # yêu cầu user "dữ liệu cũ không còn hoạt động thì bỏ") — CHNPRINTO01GYSAM/NAEXKP01CNQ657S/
    # IRLTLT01CNM156N và mọi biến thể thử qua đều 404/không tồn tại trên FRED, không có nguồn thay
    # thế nào còn sống (khác Đức/Nhật/Anh vẫn dùng tốt cùng họ ID IRLTLT01*M156N).
    print("[FRED — Quốc tế: Trung Quốc (CPI, lãi suất NHTW proxy)]")
    for key, sid, units in [
        ("cn_cpi_yoy", "CPALTT01CNM659N", None),
        ("cn_central_bank_rate", "INTDSRCNM193N", None),
    ]:
        pts = fetch_fred(sid, n=24, units=units)
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

    print("[NSO (VN) — cơ cấu GDP theo khu vực & cơ cấu vốn đầu tư theo thành phần (lũy kế)]")
    gdp_struct = fetch_nso_gdp_structure_report()
    if gdp_struct.get("gdp_growth_period"):
        # Trích trực tiếp câu "tốc độ tăng ước đạt X% so với cùng kỳ" từ CHÍNH bài báo cáo NSO
        # (nguồn GỐC, số liệu công bố tại thời điểm ra báo cáo) — chính xác hơn today_q (đoán
        # theo ngày hệ thống) của fetch_nso_latest_report() phía trên, và không có vấn đề "số đã
        # revise" như bảng sống của VBMA (xem note gdp_growth trong vimo_raw.json).
        _append_point(raw, "gdp_growth", gdp_struct["gdp_growth_period"], gdp_struct["gdp_growth"],
                       gdp_struct["source_url"])
        print(f"  -> gdp_growth {gdp_struct['gdp_growth_period']}: {gdp_struct['gdp_growth']} (nguồn NSO VN, thay cho fallback tin tức)")
    if gdp_struct.get("period"):
        period = gdp_struct["period"]
        src = gdp_struct["source_url"]
        for key in ["gdp_share_agri", "gdp_share_industry", "gdp_share_services", "gdp_share_tax",
                    "investment_share_state", "investment_share_private", "investment_share_fdi"]:
            if key in gdp_struct:
                _append_point(raw, key, period, gdp_struct[key], src)
        if "fdi_registered_usd_bn" in gdp_struct:
            _append_point(raw, "fdi_registered_usd_bn", period, gdp_struct["fdi_registered_usd_bn"], src)
        print(f"  -> {period}: cơ cấu GDP {gdp_struct.get('gdp_share_agri')}/{gdp_struct.get('gdp_share_industry')}/"
              f"{gdp_struct.get('gdp_share_services')}/{gdp_struct.get('gdp_share_tax')}%, "
              f"cơ cấu đầu tư {gdp_struct.get('investment_share_state')}/{gdp_struct.get('investment_share_private')}/"
              f"{gdp_struct.get('investment_share_fdi')}%, FDI đăng ký {gdp_struct.get('fdi_registered_usd_bn')} tỷ USD")
    else:
        print("  [INFO] Chưa trích được cơ cấu GDP/đầu tư kỳ này — giữ nguyên seed cũ.")
    if gdp_struct.get("public_investment_disbursement_period"):
        p = gdp_struct["public_investment_disbursement_period"]
        _append_point(raw, "public_investment_disbursement_rate", p,
                       gdp_struct["public_investment_disbursement_rate_pct"], gdp_struct["source_url"])
        _append_point(raw, "public_investment_disbursement_value", p,
                       gdp_struct["public_investment_disbursement_value_ty"], gdp_struct["source_url"])
        print(f"  -> Giải ngân đầu tư công {p}: {gdp_struct['public_investment_disbursement_value_ty']} nghìn tỷ đồng "
              f"({gdp_struct['public_investment_disbursement_rate_pct']}% kế hoạch năm)")
    # FDI giải ngân lũy kế theo THÁNG (bổ sung 2026-08-07 cho biểu đồ tổng quan vĩ mô — xem note tại
    # chỗ trích trong fetch_nso_gdp_structure_report()) — ghi vào CHUNG series fdi_disbursed đã có
    # sẵn (period 'YYYY-MM' trộn với 'YYYY-Qn/Hn/9M/FY', _period_sort_key() đã hỗ trợ).
    if gdp_struct.get("fdi_disbursed_period"):
        _append_point(raw, "fdi_disbursed", gdp_struct["fdi_disbursed_period"],
                       gdp_struct["fdi_disbursed_usd_bn"], gdp_struct["source_url"])
        print(f"  -> FDI giải ngân lũy kế {gdp_struct['fdi_disbursed_period']}: {gdp_struct['fdi_disbursed_usd_bn']} tỷ USD")
    # IIP dự phòng/lấp khoảng trống (nguồn CHÍNH vẫn là fetch_nso_chart_embed bên dưới, gọi SAU nên
    # sẽ ghi đè lại đúng giá trị đáng tin hơn cho các kỳ nó phủ được — xem note tại chỗ trích).
    if gdp_struct.get("iip_growth_period"):
        _append_point(raw, "iip_growth", gdp_struct["iip_growth_period"],
                       gdp_struct["iip_growth_pct"], gdp_struct["source_url"])
        print(f"  -> IIP (dự phòng, sẽ bị ghi đè nếu chart-embed phủ được kỳ này) {gdp_struct['iip_growth_period']}: {gdp_struct['iip_growth_pct']}%")
    # Tổng mức bán lẻ hàng hóa và doanh thu dịch vụ tiêu dùng THEO THÁNG (user 2026-08-08) — giá
    # trị tuyệt đối (nghìn tỷ đồng) ghi vào retail_sales_value (chỉ báo MỚI), tăng trưởng YoY ghi
    # vào retail_sales_growth (chỉ báo ĐÃ CÓ, trước đây chỉ có vài điểm rời rạc nguồn vietnambiz —
    # từ nay được backfill dày hơn nhiều nhờ series NSO theo tháng này).
    if gdp_struct.get("retail_sales_period"):
        if "retail_sales_value" not in raw:
            raw["retail_sales_value"] = {
                "group": "growth", "label": "Tổng mức bán lẻ hàng hóa và doanh thu dịch vụ tiêu dùng",
                "unit": "nghìn tỷ đồng", "good_direction": "higher", "auto_source": "nso_scrape",
                "note": ("Trích từ báo cáo tháng NSO (nso.gov.vn/bao-cao-tinh-hinh-kinh-te-xa-hoi-hang-thang/), "
                         "theo giá hiện hành, GIÁ TRỊ CỦA RIÊNG THÁNG ĐÓ (không phải lũy kế từ đầu năm)."),
                "impact": "Đo trực tiếp sức mua/tiêu dùng nội địa theo tháng — tăng trưởng chậm lại là tín hiệu sớm về sức cầu nội địa yếu đi, ảnh hưởng nhóm bán lẻ/hàng tiêu dùng/F&B.",
                "series": [],
            }
        _append_point(raw, "retail_sales_value", gdp_struct["retail_sales_period"],
                       gdp_struct["retail_sales_value_ty"], gdp_struct["source_url"])
        _append_point(raw, "retail_sales_growth", gdp_struct["retail_sales_period"],
                       gdp_struct["retail_sales_yoy_pct"], gdp_struct["source_url"])
        print(f"  -> Tổng mức bán lẻ {gdp_struct['retail_sales_period']}: {gdp_struct['retail_sales_value_ty']} nghìn tỷ đồng (YoY {gdp_struct['retail_sales_yoy_pct']}%)")

    print("[NSO — cơ cấu vốn đầu tư qua OCR ảnh infographic (dữ liệu CHỈ có ở dạng ảnh, xem fetch_nso_infographic_investment)]")
    ocr_inv = fetch_nso_infographic_investment()
    if ocr_inv:
        _append_point(raw, "investment_share_state", ocr_inv["period"], ocr_inv["state_pct"], ocr_inv["source_url"])
        _append_point(raw, "investment_share_private", ocr_inv["period"], ocr_inv["private_pct"], ocr_inv["source_url"])
        _append_point(raw, "investment_share_fdi", ocr_inv["period"], ocr_inv["fdi_pct"], ocr_inv["source_url"])
        print(f"  -> {ocr_inv['period']}: Nhà nước {ocr_inv['state_pct']}% / Ngoài NN {ocr_inv['private_pct']}% / FDI {ocr_inv['fdi_pct']}%")
    else:
        print("  [INFO] Không lấy được cơ cấu đầu tư qua OCR kỳ này (bình thường nếu chưa có bài quý mới, hoặc kiểm tra chéo không khớp).")

    print("[VBMA — CPI YoY theo tháng (toàn bộ lịch sử từ T1/2020)]")
    pts = fetch_vbma_cpi_yoy()
    if pts:
        raw["cpi_yoy"]["series"] = [
            {"period": p, "value": v,
             "source_url": "https://vbma.org.vn/vi/market-data/cpi"}
            for p, v in pts
        ]
        print(f"  -> {len(pts)} điểm")
    else:
        print("[NSO — biểu đồ chuyên đề CPI (nso.gov.vn/cpi-vi/, chi tiết THEO THÁNG) — fallback]")
        pts = fetch_nso_chart_embed("cpi")
        if pts:
            raw["cpi_yoy"]["series"] = [
                {"period": _nso_period_to_iso(p), "value": v,
                 "source_url": "https://www.nso.gov.vn/cpi-vi/"}
                for p, v in pts
            ]
            print(f"  -> {len(pts)} điểm (thay thế chuỗi theo quý cũ bằng chuỗi theo tháng)")

    print("[VBMA — Lạm phát cơ bản theo tháng (toàn bộ lịch sử từ T1/2020)]")
    pts = fetch_vbma_core_inflation()
    if pts:
        raw["core_inflation"]["series"] = [
            {"period": p, "value": v,
             "source_url": "https://vbma.org.vn/vi/market-data/cpi"}
            for p, v in pts
        ]
        print(f"  -> {len(pts)} điểm")

    print("[VBMA — Kết cấu CPI: đóng góp từng nhóm hàng vào lạm phát chung theo tháng]")
    contrib = fetch_vbma_cpi_contribution()
    for suffix, pts in contrib.items():
        key = f"cpi_contrib_{suffix}"
        raw[key]["series"] = [
            {"period": p, "value": v,
             "source_url": "https://vbma.org.vn/vi/market-data/cpi"}
            for p, v in pts
        ]
        print(f"  -> {key}: {len(pts)} điểm")

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

    print("[VBMA — Tổng dư nợ tín dụng toàn nền kinh tế theo tháng (toàn bộ lịch sử từ T1/2018)]")
    pts = fetch_vbma_credit_balance()
    if pts:
        raw["credit_balance_total"]["series"] = [
            {"period": p, "value": v,
             "source_url": "https://vbma.org.vn/vi/market-data/credit"}
            for p, v in pts
        ]
        print(f"  -> {len(pts)} điểm")

    print("[VBMA — Tổng huy động vốn (Tiền gửi TCKT + dân cư) theo tháng (toàn bộ lịch sử từ T12/2018)]")
    pts = fetch_vbma_deposit_balance()
    if pts:
        raw["deposit_balance_total"]["series"] = [
            {"period": p, "value": v,
             "source_url": "https://vbma.org.vn/vi/market-data/money-supply"}
            for p, v in pts
        ]
        print(f"  -> {len(pts)} điểm")

    print("[VietnamBiz — Bán lẻ (đối chiếu, tích lũy theo lần chạy)]")
    vnb = fetch_vietnambiz_macro()
    for key, (period, value) in vnb.items():
        _append_point(raw, key, period, value, "https://data.vietnambiz.vn/macro-economic")
        print(f"  -> {key} {period}: {value}")

    print("[VBMA — PMI sản xuất theo tháng (toàn bộ lịch sử từ T1/2016)]")
    pts = fetch_vbma_pmi()
    if pts:
        raw["pmi_manufacturing"]["series"] = [
            {"period": p, "value": v,
             "source_url": "https://vbma.org.vn/vi/market-data/gdp-growth"}
            for p, v in pts
        ]
        print(f"  -> {len(pts)} điểm")

    print("[VBMA — FDI đăng ký lũy kế theo tháng (bổ sung fdi_registered_usd_bn, cửa sổ trượt 2 năm)]")
    pts = fetch_vbma_fdi_registered()
    _merge_vbma_points(raw, "fdi_registered_usd_bn", pts, "https://vbma.org.vn/vi/market-data/fdi")
    if pts:
        print(f"  -> {len(pts)} điểm ({pts[0][0]}..{pts[-1][0]})")

    print("[VBMA — Giải ngân đầu tư công %YoY theo tháng (bổ sung public_investment_growth, cửa sổ trượt 2 năm)]")
    pts = fetch_vbma_public_investment_growth()
    _merge_vbma_points(raw, "public_investment_growth", pts, "https://vbma.org.vn/vi/market-data/states-budget")
    if pts:
        print(f"  -> {len(pts)} điểm ({pts[0][0]}..{pts[-1][0]})")

    print("[VBMA — Thâm hụt/thặng dư ngân sách %GDP theo năm (toàn bộ lịch sử từ 2015)]")
    pts = fetch_vbma_budget_deficit_pct_gdp()
    if pts:
        raw["budget_deficit_pct_gdp"]["series"] = [
            {"period": p, "value": v,
             "source_url": "https://vbma.org.vn/vi/market-data/states-budget"}
            for p, v in pts
        ]
        print(f"  -> {len(pts)} điểm")

    print("[VietnamBiz — Tăng trưởng huy động (đối chiếu credit_growth, tích lũy theo lần chạy)]")
    vnb_rates = fetch_vietnambiz_rates()
    for key, (period, value) in vnb_rates.items():
        _append_point(raw, key, period, value, "https://data.vietnambiz.vn/currency-interest-rate")
        print(f"  -> {key} {period}: {value}")

    print("[VBMA — Cung tiền M2 tăng trưởng YoY theo tháng (toàn bộ lịch sử từ T12/2018)]")
    pts = fetch_vbma_money_supply()
    if pts:
        raw["m2_growth"]["series"] = [
            {"period": p, "value": v,
             "source_url": "https://vbma.org.vn/vi/market-data/money-supply"}
            for p, v in pts
        ]
        print(f"  -> {len(pts)} điểm")

    print("[VBMA — Cung tiền M2 mức tuyệt đối theo tháng (toàn bộ lịch sử từ T12/2018)]")
    pts = fetch_vbma_money_supply_level()
    if pts:
        if "m2_balance_total" not in raw:
            raw["m2_balance_total"] = {
                "group": "monetary", "label": "Tổng cung tiền M2", "unit": "tỷ VND",
                "good_direction": "higher", "auto_source": "vbma",
                "note": ("Cùng file CSV VBMA đang dùng cho m2_growth (%) — bổ sung góc nhìn QUY MÔ "
                         "tuyệt đối, dùng để suy ra tăng trưởng M2 SO VỚI CUỐI NĂM TRƯỚC (YTD) tại "
                         "template_vimo.py, cùng cách credit_balance_total/deposit_balance_total "
                         "phục vụ credit_growth_ytd_monthly/deposit_growth_ytd_monthly."),
                "impact": "Quy mô cung tiền tuyệt đối đối chiếu với dư nợ tín dụng/huy động tuyệt đối cho biết thanh khoản hệ thống đang nới lỏng hay thắt chặt.",
                "series": [],
            }
        raw["m2_balance_total"]["series"] = [
            {"period": p, "value": v,
             "source_url": "https://vbma.org.vn/vi/market-data/money-supply"}
            for p, v in pts
        ]
        print(f"  -> {len(pts)} điểm")

    # Nhóm lãi suất (SBV/huy động/OMO/tín phiếu) dùng kỳ THEO TUẦN (không phải period_now theo
    # tháng ở trên) — user (2026-07-24) chỉ ra chạy nhiều lần/tháng vẫn chỉ ra 1 điểm, biểu đồ lịch
    # sử liên ngân hàng gần như không tích lũy được gì. Xem _current_period_weekly().
    period_now_weekly = _current_period_weekly()

    print("[SBV — lãi suất tái cấp vốn & liên ngân hàng 3/6/9 tháng (tích lũy theo TUẦN)]")
    rates = fetch_sbv_interest_rates()
    # on/1w/2w/1m ĐÃ CHUYỂN sang nguồn VIRA (dưới đây, tích lũy theo NGÀY thật thay vì snapshot
    # theo tuần) — bỏ qua ở đây để tránh trộn 2 phương pháp gộp khác nhau vào cùng 1 series.
    _SKIP_TENORS_NOW_FROM_VIRA = {"interbank_rate_on", "interbank_rate_1w",
                                    "interbank_rate_2w", "interbank_rate_1m"}
    for key, value in rates.items():
        if key in _SKIP_TENORS_NOW_FROM_VIRA:
            continue
        _append_point(raw, key, period_now_weekly, value, "https://www.sbv.gov.vn/vi/l%C3%A3i-su%E1%BA%A5t1")
        print(f"  -> {key} {period_now_weekly}: {value}")

    # omo_rate_7d ĐÃ CHUYỂN sang nguồn VIRA (dưới đây, tích lũy theo NGÀY thật) — cùng lý do với
    # on/1w/2w/1m ở trên, tránh trộn 2 phương pháp gộp khác nhau (tuần vs ngày) vào cùng 1 series.
    # fetch_sbv_omo_rate() không còn được gọi ở đây nữa (giữ nguyên hàm để tham khảo/dự phòng).

    print("[VIRA — lãi suất liên ngân hàng ON/1W/2W/1M + lợi suất TPCP thứ cấp + OMO bơm/hút ròng (tích lũy theo NGÀY thật)]")
    # sort tăng dần theo ngày TRƯỚC khi append — fetch_vira_bulletin() quét lùi (mới nhất trước),
    # trong khi _append_point() chỉ nối vào CUỐI series (không tự sort như _merge_vbma_points).
    vira_entries = sorted(fetch_vira_bulletin(), key=lambda e: e["date"])
    VIRA_KEY_MAP = {
        "interbank_on": "interbank_rate_on", "interbank_1w": "interbank_rate_1w",
        "interbank_2w": "interbank_rate_2w", "interbank_1m": "interbank_rate_1m",
        "bond_3y": "govt_bond_yield_3y", "bond_5y": "govt_bond_yield_5y",
        "bond_7y": "govt_bond_yield_7y", "bond_10y": "govt_bond_yield_10y",
        "bond_15y": "govt_bond_yield_15y", "omo_net": "omo_net_operation",
        "omo_rate": "omo_rate_7d", "omo_outstanding": "omo_outstanding_balance",
    }
    # _append_point() chỉ so khớp điểm CUỐI series — không đủ ở đây vì lookback_days quét lùi ~10
    # ngày MỖI LẦN chạy nên phần lớn ngày đã có sẵn từ lần chạy trước (không nằm ở cuối series do
    # thứ tự append). Tự kiểm tra period đã tồn tại (ở BẤT KỲ đâu trong series) trước khi thêm, để
    # chạy lại nhiều lần/tuần không bị nhân đôi điểm.
    _vira_touched_keys = set()
    for entry in vira_entries:
        for field, raw_key in VIRA_KEY_MAP.items():
            if field not in entry:
                continue
            existing_periods = {pt["period"] for pt in raw[raw_key]["series"]}
            if entry["date"] in existing_periods:
                continue
            _append_point(raw, raw_key, entry["date"], entry[field], entry["source_url"])
            _vira_touched_keys.add(raw_key)
        print(f"  -> {entry['date']}: {', '.join(f'{k}={v}' for k, v in entry.items() if k not in ('date', 'source_url'))}")
    # lookback_days quét lùi ~10 ngày NÊN CÓ THỂ lấp được 1 ngày CŨ HƠN điểm mới nhất đã lưu (vd
    # bản tin ra trễ, hoặc lần chạy trước bị lỗi) — _append_point() chỉ nối vào CUỐI nên điểm lấp
    # trễ đó sẽ nằm SAI VỊ TRÍ (sau các ngày mới hơn) nếu không sort lại. Sort lại period dạng
    # "YYYY-MM-DD" (so sánh chuỗi = so sánh thời gian, an toàn) cho mọi key vừa được VIRA ghi thêm.
    for raw_key in _vira_touched_keys:
        raw[raw_key]["series"].sort(key=lambda pt: pt["period"])

    print("[SBV — tín phiếu NHNN (hút thanh khoản thị trường 2, đối lập OMO — 2 chiều bơm/hút, tích lũy theo TUẦN)]")
    days_since, last_date, src = fetch_sbv_tin_phieu_days_since()
    if days_since is not None:
        _append_point(raw, "tin_phieu_days_since_issuance", period_now_weekly, days_since, src)
        print(f"  -> {period_now_weekly}: {days_since} ngày kể từ lần chào bán tín phiếu gần nhất ({last_date})")

    print("[Ngân hàng — lãi suất huy động 12 tháng: VCB / VietinBank / NamABank (tích lũy theo TUẦN)]")
    v = fetch_vcb_deposit_rate_12m()
    if v is not None:
        _append_point(raw, "deposit_rate_12m_vcb", period_now_weekly, v,
                       "https://www.vietcombank.com.vn/vi-VN/api/interestrates?accountType=Personal")
        print(f"  -> VCB {period_now_weekly}: {v}")
    v = fetch_ctg_deposit_rate_12m()
    if v is not None:
        _append_point(raw, "deposit_rate_12m_ctg", period_now_weekly, v, "https://www.vietinbank.vn/lai-suat-khcn")
        print(f"  -> VietinBank {period_now_weekly}: {v}")
    v = fetch_nab_deposit_rate_12m()
    if v is not None:
        _append_point(raw, "deposit_rate_12m_nab", period_now_weekly, v, "https://www.namabank.com.vn/lai-suat-tien-gui-vnd-2")
        print(f"  -> NamABank {period_now_weekly}: {v}")

    print("[24hmoney — lãi suất huy động online 12 tháng, mặt bằng toàn thị trường (~38 NH, tích lũy theo TUẦN)]")
    max_r, avg_r, n = fetch_market_deposit_rate_12m()
    if max_r is not None:
        src = "https://24hmoney.vn/lai-suat-gui-ngan-hang"
        _append_point(raw, "deposit_rate_12m_market_max", period_now_weekly, max_r, src)
        _append_point(raw, "deposit_rate_12m_market_avg", period_now_weekly, avg_r, src)
        print(f"  -> {period_now_weekly}: max={max_r}% avg={avg_r}% (n={n} ngân hàng)")

    print("[TCBS — lãi suất iPower cao nhất (kênh gửi tiền thay thế, tích lũy theo TUẦN)]")
    tcbs_rate, tcbs_period, tcbs_src = fetch_tcbs_ipower_max_rate()
    if tcbs_rate is not None:
        _append_point(raw, "deposit_rate_tcbs_ipower_max", tcbs_period, tcbs_rate, tcbs_src)
        print(f"  -> {tcbs_period}: {tcbs_rate}%")

    print("[RSS tin tức — lãi suất huy động THỎA THUẬN (quét CafeF/VietStock, chỉ ghi khi có tin mới khớp)]")
    hit = fetch_negotiated_deposit_rate_news()
    if hit:
        period, value, link, title = hit
        _append_point(raw, "deposit_rate_negotiated_max", period, value, link)
        print(f"  -> {period}: {value}% — \"{title}\"")
    else:
        print("  -> Không có tin mới khớp từ khóa (bình thường, đây là tin hiếm)")

    print("[Hải quan — Xuất/nhập khẩu theo tháng (file Excel cục bộ, CHỈ có khi chạy thủ công trên máy có sẵn thư mục)]")
    xnk = load_customs_xnk_local()
    if xnk["export"] or xnk["import"]:
        for period, value, fname in xnk["export"]:
            _append_point(raw, "export_value_monthly", period, value, fname)
        for period, value, fname in xnk["import"]:
            _append_point(raw, "import_value_monthly", period, value, fname)
        export_by_period = {p: v for p, v, _ in xnk["export"]}
        import_by_period = {p: v for p, v, _ in xnk["import"]}
        for period in sorted(set(export_by_period) & set(import_by_period)):
            balance = round(export_by_period[period] - import_by_period[period], 4)
            _append_point(raw, "trade_balance_monthly", period, balance, "Hải quan (V01+V02, tính từ xuất trừ nhập)")
        print(f"  -> {len(xnk['export'])} tháng xuất khẩu, {len(xnk['import'])} tháng nhập khẩu")
    else:
        print("  [INFO] Không tìm thấy thư mục/file Hải quan cục bộ — bỏ qua (bình thường khi chạy trên GitHub Action).")

    raw["_meta"]["last_auto_update"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    save_raw(raw)
    print("\n[OK] Đã ghi data/vimo_raw.json")


if __name__ == "__main__":
    update_vimo_raw()
