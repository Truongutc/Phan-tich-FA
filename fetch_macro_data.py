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

    print("[FII — khối ngoại mua/bán ròng HOSE]")
    net_ty_vnd, date_iso, src = fetch_fii_net_flow()
    if net_ty_vnd is not None:
        _append_point(raw, "fii_net_flow_hose", date_iso, net_ty_vnd, src)
        print(f"  -> {date_iso}: {net_ty_vnd:+.2f} tỷ VND ({'mua ròng' if net_ty_vnd > 0 else 'bán ròng'})")

    print("[VN-Index P/E & P/B]")
    pe, pb, src = fetch_vnindex_pe_pb_24hmoney()
    if not pe:  # 24hmoney lỗi/đổi cấu trúc -> fallback P/E riêng qua worldperatio.com (không có P/B)
        pe, src = fetch_vnindex_pe_current()
    if pe:
        _append_point(raw, "vnindex_pe", period_now, pe, src)
        print(f"  -> P/E {period_now}: {pe}")
    if pb:
        _append_point(raw, "vnindex_pb", period_now, pb, src)
        print(f"  -> P/B {period_now}: {pb}")

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

    print("[SBV — tín phiếu NHNN (hút thanh khoản thị trường 2, đối lập OMO — 2 chiều bơm/hút)]")
    days_since, last_date, src = fetch_sbv_tin_phieu_days_since()
    if days_since is not None:
        _append_point(raw, "tin_phieu_days_since_issuance", period_now, days_since, src)
        print(f"  -> {period_now}: {days_since} ngày kể từ lần chào bán tín phiếu gần nhất ({last_date})")

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
