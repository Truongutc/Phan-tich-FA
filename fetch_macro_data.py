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
import csv
import io
import glob
import json
import time
import datetime
import unicodedata
import subprocess
import statistics
import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
VIMO_RAW_PATH = os.path.join(PROJECT_ROOT, "data", "vimo_raw.json")
# File RIÊNG (không phải vimo_raw.json — tránh phá vỡ quy ước "mọi key cấp 1 của raw đều là 1
# indicator dạng {group,label,unit,series,...}" mà nhiều hàm khác trong template_vimo.py giả định
# khi lặp raw.items(), xem fetch_dulieukinhte_gdp_sector/fetch_dulieukinhte_cpi_group) — lưu LEVEL
# (giá trị tuyệt đối, chưa tính YoY) theo từng ngành/nhóm hàng, dùng để dựng bảng heatmap GDP theo
# ngành + CPI theo nhóm hàng (user 2026-10-01) trong template_vimo._build_level_yoy_heatmap().
SECTOR_DETAIL_PATH = os.path.join(PROJECT_ROOT, "data", "vimo_sector_detail.json")

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


def _merge_point_anywhere(raw, key, period, value, source_url):
    """Giống _append_point() nhưng tìm period trong TOÀN BỘ series (không chỉ phần tử cuối) rồi
    sắp lại theo thời gian sau khi thêm — cần cho các nguồn có thể trả 1 điểm KHÔNG PHẢI mới nhất
    trong cùng 1 lần gọi (vd VietnamBiz trả cả value kỳ mới nhất LẪN pre_value kỳ liền trước, xem
    fetch_vietnambiz_macro() — nếu series hiện có khoảng trống ở giữa, pre_value có thể trùng 1
    period đã tồn tại nhưng KHÔNG PHẢI phần tử cuối, dùng _append_point() thường sẽ tạo bản trùng
    lệch thứ tự). An toàn dùng string-sort trực tiếp vì các chỉ báo nguồn VietnamBiz đều 1 định
    dạng period thuần nhất (toàn 'YYYY-MM' hoặc toàn 'YYYY-Qn'), không trộn lẫn như fdi_disbursed."""
    series = raw[key]["series"]
    for p in series:
        if p["period"] == period:
            p["value"] = value
            p["source_url"] = source_url
            return
    series.append({"period": period, "value": value, "source_url": source_url})
    series.sort(key=lambda p: p["period"])


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


def fetch_vcb_usd_sell_rate(date_str):
    """Vietcombank có API JSON thật theo NGÀY (tương tự fetch_vcb_deposit_rate_12m) — trả tỷ giá
    bán ra NIÊM YẾT (cash/transfer/sell) cho mọi ngoại tệ. User (2026-10-01) muốn "chênh gap tỷ
    giá chợ đen so với tỷ giá VCB bán ra" (đầu cơ/nhu cầu USD tăng nếu gap nới rộng) — lấy field
    "sell" (bán ra cho KHÁCH HÀNG, KHÁC "transfer" là tỷ giá bán chuyển khoản liên ngân hàng, và
    KHÁC "cash" là bán mặt tại quầy — "sell" mới đúng nghĩa "giá bán" so được với chợ đen, 2 bên
    đều là giá người dân mua được thực tế). date_str dạng "YYYY-MM-DD". Trả (sell_rate, source_url)
    hoặc (None, None) nếu lỗi/không có USD trong ngày đó (vd ngày nghỉ, VCB trả ngày gần nhất)."""
    url = f"https://www.vietcombank.com.vn/api/exchangerates?date={date_str}"
    try:
        r = requests.get(url, headers={"User-Agent": UA}, timeout=20)
        r.raise_for_status()
        data = r.json()
        usd = next((x for x in data.get("Data", []) if x.get("currencyCode") == "USD"), None)
        if not usd or not usd.get("sell"):
            return None, None
        actual_date = (data.get("Date") or date_str)[:10]
        return actual_date, round(float(usd["sell"]), 2)
    except Exception as e:
        print(f"  [WARN] VCB USD sell rate ({date_str}) thất bại: {e}")
        return None, None


def fetch_chogia_usd_cho_den():
    """chogia.vn/ngoai-te/usd-cho-den/ — endpoint AJAX thật (WordPress admin-ajax, KHÔNG cần JS
    render) cấp dữ liệu cho biểu đồ Highcharts trên trang, trả ~28 ngày gần nhất (gia_mua/gia_ban,
    VND). User (2026-10-01) muốn đối chiếu "giá bán" (gia_ban) với tỷ giá VCB bán ra để tính gap
    (đầu cơ/áp lực cầu USD tăng nếu gap nới rộng). Trả list[(date_iso, gia_ban)] (rỗng nếu lỗi) —
    chỉ lấy gia_ban theo đúng yêu cầu user ("chỉ lấy theo giá bán ra thôi cho dễ so sánh"), gọi
    MỖI LẦN chạy để tự tích lũy lịch sử qua _merge_point_anywhere() (nguồn chỉ trả cửa sổ ~28 ngày
    gần nhất, không có tham số lấy sâu hơn)."""
    url = "https://chogia.vn/wp-admin/admin-ajax.php"
    try:
        r = requests.post(url, data={"action": "load_gia_ngoai_te_cho_do_thi", "ma": "usd"},
                           headers={"User-Agent": UA}, timeout=20)
        r.raise_for_status()
        j = r.json()
        if not j.get("success"):
            return []
        out = []
        for row in j.get("data", []):
            ngay = row.get("ngay")
            gia_ban = row.get("gia_ban")
            if ngay and gia_ban:
                out.append((ngay, round(float(gia_ban), 2)))
        return out
    except Exception as e:
        print(f"  [WARN] chogia.vn USD chợ đen thất bại: {e}")
        return []


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
US_MACRO_SERIES = [
    ("usm_cpi", "CPIAUCSL", "CPI toàn phần (SA, 1982-84=100)", "index"),
    ("usm_core_cpi", "CPILFESL", "CPI lõi (ex food & energy, SA)", "index"),
    ("usm_cpi_food", "CPIUFDSL", "CPI nhóm Thực phẩm (SA)", "index"),
    ("usm_cpi_energy", "CPIENGSL", "CPI nhóm Năng lượng (SA)", "index"),
    ("usm_cpi_energy_services", "CUSR0000SEHF", "CPI Năng lượng — CHỈ dịch vụ (điện+gas, KHÔNG xăng dầu, SA)", "index"),
    ("usm_cpi_shelter", "CUSR0000SAH1", "CPI nhóm Nhà ở/Shelter (SA)", "index"),
    ("usm_cpi_transport", "CPITRNSL", "CPI nhóm Giao thông (SA)", "index"),
    ("usm_cpi_medical", "CPIMEDSL", "CPI nhóm Y tế (SA)", "index"),
    ("usm_cpi_apparel", "CPIAPPSL", "CPI nhóm May mặc (SA)", "index"),
    ("usm_cpi_recreation", "CPIRECSL", "CPI nhóm Giải trí (SA)", "index"),
    ("usm_cpi_education_comm", "CPIEDUSL", "CPI nhóm Giáo dục & Truyền thông (SA)", "index"),
    ("usm_cpi_other", "CPIOGSSL", "CPI nhóm Hàng hóa & dịch vụ khác (SA)", "index"),
    ("usm_cpi_goods", "CUSR0000SAC", "CPI Hàng hóa (commodities, SA)", "index"),
    ("usm_cpi_services", "CUSR0000SAS", "CPI Dịch vụ (services, SA)", "index"),
    ("usm_ppi_final_demand", "PPIFIS", "PPI cầu cuối cùng (Final demand, SA)", "index"),
    ("usm_ppi_all_commodities", "PPIACO", "PPI tất cả hàng hóa (NSA)", "index"),
    ("usm_ppi_industrial_commodities", "PPIIDC", "PPI hàng hóa công nghiệp (NSA, proxy chi phí đầu vào)", "index"),
    ("usm_import_price", "IR", "Chỉ số giá nhập khẩu (all)", "index"),
    ("usm_export_price", "IQ", "Chỉ số giá xuất khẩu (all)", "index"),
    ("usm_retail_sales", "RSAFS", "Doanh số bán lẻ & dịch vụ ăn uống (triệu USD, SA)", "million_usd"),
    ("usm_pce_nominal", "PCE", "Chi tiêu tiêu dùng cá nhân danh nghĩa (tỷ USD)", "billion_usd"),
    ("usm_pce_real", "PCEC96", "Chi tiêu tiêu dùng thực (tỷ USD 2017, SAAR)", "billion_usd_2017"),
    ("usm_pce_price", "PCEPI", "Chỉ số giá PCE", "index"),
    # THEM 2026-10-08 (user gửi tài liệu: "Fed không điều hành theo CPI 2%, Fed nhắm mục tiêu lạm
    # phát 2% theo PCE" — Core PCE là thước đo Fed ưu tiên, trước đây CHỈ fetch usm_pce_price (PCE
    # TOÀN PHẦN) mà KHÔNG có bản lõi, cũng chưa tính YoY/hiển thị ở đâu cả) — PCEPILFE = PCE lõi.
    ("usm_pce_core", "PCEPILFE", "Chỉ số giá PCE lõi (ex food & energy) — thước đo Fed ưu tiên", "index"),
    ("usm_oil_wti", "DCOILWTICO", "Giá dầu WTI (USD/thùng, daily)", "usd_per_barrel"),
    ("usm_gasoline", "GASREGW", "Giá xăng thường (USD/gallon, weekly)", "usd_per_gallon"),
    ("usm_diesel_gulf", "DDFUELUSGULF", "Giá ULSD diesel Vịnh Mexico (USD/gallon, daily, EIA)", "usd_per_gallon"),
    ("usm_gasoline_gulf", "DGASUSGULF", "Giá xăng Vịnh Mexico (USD/gallon, daily, EIA)", "usd_per_gallon"),
    ("usm_brent", "DCOILBRENTEU", "Giá dầu Brent (USD/thùng, daily)", "usd_per_barrel"),
    # THEM 2026-10-08 (user: "lấy dữ liệu bảng cân đối Fed — QE/QT, thu hẹp/mở rộng") — H.4.1, FRED
    # mirror trực tiếp từ Fed, hàng tuần (As of/Ending Wednesday). Đơn vị GỐC triệu USD.
    ("usm_fed_assets", "WALCL", "Fed — Tổng tài sản (triệu USD)", "million_usd"),
    ("usm_fed_treasury", "TREAST", "Fed — Trái phiếu Chính phủ Mỹ nắm giữ (triệu USD)", "million_usd"),
    ("usm_fed_mbs", "WSHOMCB", "Fed — MBS nắm giữ (triệu USD)", "million_usd"),
    ("usm_fed_reserves", "WRESBAL", "Fed — Dự trữ ngân hàng tại Fed (triệu USD)", "million_usd"),
    ("usm_fed_rrp", "WLRRAL", "Fed — Reverse Repo (RRP, triệu USD)", "million_usd"),
    ("usm_fed_tga", "WTREGEN", "Fed — Tài khoản Treasury (TGA, triệu USD)", "million_usd"),
    ("usm_ecb_assets", "ECBASSETSW", "ECB — Tổng tài sản (triệu EUR)", "million_eur"),
    ("usm_unemployment", "UNRATE", "Tỷ lệ thất nghiệp (%)", "pct"),
    ("usm_fed_funds", "FEDFUNDS", "Lãi suất quỹ liên bang (%)", "pct"),
    ("usm_yield_10y", "DGS10", "Lợi suất trái phiếu 10 năm (%)", "pct"),
    ("usm_spread_10y_2y", "T10Y2Y", "Chênh lệch 10Y-2Y (%)", "pct"),
    # THEM 2026-10-08 (user: "tiếp tục triển khai theo ma trận đã bàn" — Growth/Labor/
    # Treasury&Credit/USD trong "US Macro Liquidity Matrix" người dùng gửi; KHÔNG chấm điểm gộp,
    # chỉ lấy thêm số liệu thô qua FRED, giữ nguyên nguyên tắc không gộp tín hiệu thành 1 điểm số).
    ("usm_gdp_real", "GDPC1", "GDP thực (tỷ USD 2017, SAAR, theo quý)", "billion_usd_2017"),
    ("usm_indpro", "INDPRO", "Chỉ số sản xuất công nghiệp (SA, 2017=100)", "index"),
    # BỎ "usm_leading_index"/USSLIND (Leading Index, FRB Philadelphia) — kiểm tra FRED metadata
    # 2026-10-08 thấy observation_end=2020-02-01, last_updated=2020-04-14 — CHUỖI ĐÃ BỊ NGỪNG CÔNG
    # BỐ hơn 6 năm, không phải lỗi tải. Không dùng được cho tín hiệu hiện tại.
    ("usm_payrolls", "PAYEMS", "Tổng việc làm phi nông nghiệp (nghìn người, SA)", "thousand_persons"),
    ("usm_claims", "ICSA", "Trợ cấp thất nghiệp lần đầu (người/tuần, SA)", "persons"),
    ("usm_job_openings", "JTSJOL", "Số vị trí tuyển dụng — JOLTS (nghìn, SA)", "thousand"),
    ("usm_participation", "CIVPART", "Tỷ lệ tham gia lực lượng lao động (%)", "pct"),
    ("usm_avg_earnings", "CES0500000003", "Thu nhập bình quân giờ — toàn bộ nhân viên tư nhân (USD/giờ, SA)", "usd_per_hour"),
    ("usm_yield_2y", "DGS2", "Lợi suất trái phiếu 2 năm (%)", "pct"),
    ("usm_real_yield_10y", "DFII10", "Lợi suất thực TIPS 10 năm (%)", "pct"),
    ("usm_breakeven_10y", "T10YIE", "Lạm phát kỳ vọng hòa vốn 10 năm (breakeven, %)", "pct"),
    ("usm_hy_oas", "BAMLH0A0HYM2", "Chênh lệch tín dụng High Yield (OAS, điểm %)", "pct"),
    ("usm_ig_oas", "BAMLC0A0CM", "Chênh lệch tín dụng Investment Grade (OAS, điểm %)", "pct"),
    ("usm_dxy_broad", "DTWEXBGS", "Chỉ số USD trọng số thương mại (Broad, 2006=100)", "index"),
]


def fetch_fred_csv(series_id, start="2015-01-01"):
    """Tải chuỗi FRED qua fredgraph.csv (công khai, không cần API key). Trả [(YYYY-MM-DD, float)]
    đã bỏ các ô rỗng/'.' ; chỉ giữ từ ngày start. Lỗi mạng trả []."""
    api_key = os.environ.get("FRED_API_KEY")
    try:
        if api_key:
            r = requests.get("https://api.stlouisfed.org/fred/series/observations",
                             params={"series_id": series_id, "api_key": api_key, "file_type": "json",
                                     "observation_start": start, "sort_order": "asc"},
                             headers={"User-Agent": UA}, timeout=60)
            r.raise_for_status()
            return [(o["date"], float(o["value"])) for o in r.json().get("observations", [])
                    if o.get("value") not in ("", ".", None)]
        r = requests.get(f"https://fred.stlouisfed.org/graph/fredgraph.csv?id={series_id}",
                         headers={"User-Agent": UA}, timeout=60)
        r.raise_for_status()
    except Exception as e:
        print(f"  [WARN] FRED {series_id}: {e}")
        return []
    out = []
    for line in r.text.strip().split(chr(10))[1:]:
        parts = line.split(",")
        if len(parts) < 2 or parts[1] in ("", ".") or parts[0] < start:
            continue
        try:
            out.append((parts[0], float(parts[1])))
        except ValueError:
            continue
    return out


# THEM 2026-10-08 (user: "Capital Flows (TIC) chưa làm vì không có trên FRED" — tìm nguồn riêng).
# ticdata.treasury.gov CÔNG KHAI, không cần đăng ký/API key (khác bls.gov/eia.gov — đã kiểm tra
# curl trực tiếp trả 200 OK). File "mfhhis01.csv" là báo cáo "Major Foreign Holders of Treasury
# Securities" — dữ liệu THÔ dạng bảng rộng, mỗi "khối" 1 năm (12 cột Dec→Jan), nhiều khối xếp dọc
# từ năm mới nhất lùi về ~2000. Lấy 2 dòng: "Grand Total" (tổng toàn bộ nước ngoài nắm giữ) và
# "For. Official" (chỉ khối NHTW/chính phủ nước ngoài — tín hiệu "official capital flows" rõ hơn
# vì khối tư nhân mua/bán mang tính đầu cơ nhiều, khối official phản ánh dự trữ ngoại hối quốc gia).
def fetch_tic_major_foreign_holders():
    """Trả (total_dict, official_dict): {period "YYYY-MM": value tỷ USD}. Lỗi mạng trả ({}, {})."""
    MONTHS = {"Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
              "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12}
    url = "https://ticdata.treasury.gov/resource-center/data-chart-center/tic/Documents/mfhhis01.csv"
    try:
        r = requests.get(url, headers={"User-Agent": UA}, timeout=60)
        r.raise_for_status()
    except Exception as e:
        print(f"  [WARN] TIC mfhhis01.csv: {e}")
        return {}, {}
    rows = list(csv.reader(io.StringIO(r.text)))
    total, official = {}, {}
    i, n = 0, len(rows)
    while i < n:
        row = rows[i]
        if len(row) > 1 and row[0] == "" and row[1] in MONTHS and i + 1 < n and rows[i + 1] and rows[i + 1][0] == "Country":
            month_cols, year_cols = row[1:13], rows[i + 1][1:13]
            periods = [f"{y.strip()}-{MONTHS[m]:02d}" if m in MONTHS and y.strip().isdigit() else None
                       for m, y in zip(month_cols, year_cols)]
            j = i + 3  # bỏ qua dòng gạch "------"
            while j < n and not (len(rows[j]) > 1 and rows[j][0] == "" and rows[j][1] in MONTHS):
                r2 = rows[j]
                # Nhãn "For. Official" đổi thụt đầu dòng qua các năm (vd "For. Official" ở khối
                # 2023-2025 nhưng "  For. Official" thụt 2 khoảng trắng ở khối cũ hơn) — strip().
                label = r2[0].strip() if r2 else ""
                if label == "Grand Total":
                    for p, v in zip(periods, r2[1:13]):
                        if p:
                            try:
                                total[p] = float(v)
                            except ValueError:
                                pass
                elif label == "For. Official":
                    for p, v in zip(periods, r2[1:13]):
                        if p:
                            try:
                                official[p] = float(v)
                            except ValueError:
                                pass
                j += 1
            i = j
            continue
        i += 1
    return total, official


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


def fetch_aric_gdp_by_use():
    """aric.adb.org (Asian Regional Integration Center, ADB) — bảng "Economic & Financial
    Indicators" cho Việt Nam, sector "Real Sector and Prices", TẦN SUẤT NĂM (frq=5, predefined=15)
    — user (2026-08-08) muốn tách GDP theo CẤU PHẦN SỬ DỤNG (tiêu dùng tư nhân/chính phủ/đầu tư)
    thay vì chỉ theo khu vực kinh tế (đã có sẵn gdp_share_agri/industry/services qua NSO). Trang
    KHÔNG có API/CSV, nhưng bảng kết quả server-render HTML đơn giản, 1 dòng/năm, N cột theo ĐÚNG
    THỨ TỰ tên chỉ số ở dòng "Indicator:" đầu bảng (không phải thứ tự cố định — PHẢI đọc header để
    map đúng cột, tránh gán nhầm nếu ADB đổi thứ tự sau này). Chỉ tần suất NĂM có bộ chỉ số này
    (frq=3/4 chỉ trả CPI/IIP/bán lẻ hoặc GDP Growth đơn lẻ — đã khảo sát thủ công 2026-08-08), nên
    dữ liệu trễ ~1 năm so với hiện tại (thường mới nhất là năm trước). Trả {indicator_name: [(year,
    value), ...]} hoặc {} nếu thất bại — indicator_name lấy nguyên văn cột "Indicator:" của ADB."""
    url = "https://aric.adb.org/database/economic-financial-indicators-result?mode=preview&cty=160&frq=5&predefined=15&prec=3"
    try:
        r = requests.get(url, headers={"User-Agent": UA}, timeout=20)
        r.raise_for_status()
        m_table = re.search(r"<table.*?</table>", r.text, re.S)
        if not m_table:
            print("  [WARN] ARIC GDP theo cấu phần: không tìm thấy bảng — trang có thể đã đổi cấu trúc.")
            return {}
        table = m_table.group(0)

        m_header = re.search(r"<tr><th>Indicator:</th>(.*?)</tr>", table)
        if not m_header:
            print("  [WARN] ARIC GDP theo cấu phần: không tìm thấy dòng 'Indicator:' — bỏ qua.")
            return {}
        columns = re.findall(r"<th>([^<]*)</th>", m_header.group(1))
        if not columns:
            return {}

        out = {name: [] for name in columns}
        for m_row in re.finditer(r"<tr><td>(\d{4})-\d{2}</td>(.*?)</tr>", table):
            year = int(m_row.group(1))
            cells = re.findall(r'<td[^>]*>([^<]*)</td>', m_row.group(2))
            for name, cell in zip(columns, cells):
                cell = cell.strip()
                if cell:
                    try:
                        out[name].append((year, float(cell)))
                    except ValueError:
                        continue
        return out
    except Exception as e:
        print(f"  [WARN] ARIC GDP theo cấu phần thất bại: {e}")
        return {}


def fetch_40yo_vnm_data():
    """40yo.vn/vi-mo/vn (Việt Nam) — trang là SPA (JS render) nhưng gọi thẳng endpoint JSON công
    khai (không auth, không rate-limit thấy được khi khảo sát) chứa ~80 chỉ báo Việt Nam tổng hợp
    từ IMF WEO/World Bank/ADB/UNCTAD, cấu trúc {y: năm thập phân, v: giá trị, p: nhãn kỳ dạng
    'MM/YYYY' (tháng)/'Qn/YYYY' (quý)/'YYYY' (năm)} mỗi điểm — do chính trang này dùng để vẽ chart
    cho người dùng, phát hiện qua bắt network request khi mở trang bằng Playwright (2026-08-09).
    Trả dict JSON đầy đủ (khóa 'charts' chứa các nhóm biểu đồ) hoặc {} nếu thất bại."""
    url = "https://40yo.vn/data/countries/VNM.json"
    try:
        r = requests.get(url, headers={"User-Agent": UA}, timeout=30)
        r.raise_for_status()
        return r.json()
    except Exception as e:
        print(f"  [WARN] 40yo.vn VNM data thất bại: {e}")
        return {}


def _40yo_period_to_iso(p):
    """'MM/YYYY' (tháng) -> 'YYYY-MM'; 'Qn/YYYY' (quý) -> 'YYYY-Qn'; 'YYYY' (năm) -> 'YYYY' — khớp
    định dạng period chuẩn đang dùng trong toàn hệ thống."""
    m = re.match(r"^(\d{2})/(\d{4})$", p)
    if m:
        return f"{m.group(2)}-{m.group(1)}"
    m = re.match(r"^Q(\d)/(\d{4})$", p)
    if m:
        return f"{m.group(2)}-Q{m.group(1)}"
    if re.match(r"^\d{4}$", p):
        return p
    return None


def _40yo_is_fresh(iso_periods, freshness):
    """Chỉ giữ chuỗi CÒN ĐANG CẬP NHẬT — user (2026-08-09): "dữ liệu tháng nhưng lại đang dừng ở
    năm 2025 thì bỏ qua, dữ liệu quý không có Q1 2026 thì bỏ qua, dữ liệu năm không có 2025 thì bỏ
    qua vì các data đó không có tính cập nhật" (khảo sát 2026-08-09 cho thấy nhiều chuỗi trong bộ
    dữ liệu 40yo.vn đã ngừng cập nhật từ lâu, vd finacc_break "Gián tiếp cổ phiếu" dừng ở Q4/2013)."""
    if not iso_periods:
        return False
    last = max(iso_periods)
    if freshness == "monthly":
        return last[:4] >= "2026"
    if freshness == "quarterly":
        return last >= "2026-Q1"
    return last >= "2025"  # annual


# Danh sách chỉ báo GIỮ LẠI từ 40yo.vn (đã khảo sát thủ công 2026-08-09, loại các chuỗi đã ngừng
# cập nhật qua _40yo_is_fresh — xem quyết định chi tiết trong lịch sử trao đổi). Mỗi phần tử:
# (chart_key, series_name, raw_key, label, unit, good_direction, group, freshness).
_40YO_SPEC = [
    ("yields", "1Y", "govt_bond_yield_1y_40yo", "Lợi suất TPCP thứ cấp 1 năm (40yo)", "%", "lower", "monetary", "monthly"),
    ("yields", "2Y", "govt_bond_yield_2y_40yo", "Lợi suất TPCP thứ cấp 2 năm (40yo)", "%", "lower", "monetary", "monthly"),
    ("spread", "10Y − 2Y", "yield_spread_10y_2y_vn", "Chênh lệch lợi suất TPCP 10Y-2Y (VN)", "điểm %", "higher", "monetary", "monthly"),
    ("spread", "2Y − 1Y", "yield_spread_2y_1y_vn", "Chênh lệch lợi suất TPCP 2Y-1Y (VN)", "điểm %", "higher", "monetary", "monthly"),
    ("policy_vs_us", "Lợi suất 10Y − Mỹ", "yield_spread_vn_us_10y", "Chênh lệch lợi suất TPCP 10Y VN so với Mỹ", "điểm %", "higher", "monetary", "monthly"),
    ("policy_vs_us", "Lợi suất 2Y − Mỹ", "yield_spread_vn_us_2y", "Chênh lệch lợi suất TPCP 2Y VN so với Mỹ", "điểm %", "higher", "monetary", "monthly"),
    ("reserves", "Ngoại tệ (USD)", "forex_reserves_monthly", "Dự trữ ngoại hối — ngoại tệ (theo tháng)", "tỷ USD", "higher", "external", "monthly"),
    ("reserves", "SDR (USD)", "forex_reserves_sdr", "Dự trữ ngoại hối — SDR", "tỷ USD", "higher", "external", "monthly"),
    ("exch_broad", "USD/VND (phải)", "usdvnd_monthly_avg", "Tỷ giá USD/VND (bình quân tháng)", "VND", "lower", "external", "monthly"),
    ("pe", "P/E thị trường", "vnindex_pe_market_monthly", "P/E thị trường (theo tháng, từ 2016)", "lần", "lower", "market", "monthly"),
    ("index_pe", "Chỉ số giá cổ phiếu", "vnindex_level_monthly", "Chỉ số giá cổ phiếu (theo tháng)", "điểm", "higher", "market", "monthly"),
    ("bop", "Hàng hóa", "bop_goods", "Cán cân vãng lai — Hàng hóa", "tỷ USD", "higher", "trade", "quarterly"),
    ("bop", "Dịch vụ", "bop_services", "Cán cân vãng lai — Dịch vụ", "tỷ USD", "higher", "trade", "quarterly"),
    ("bop", "Thu nhập sơ cấp", "bop_primary_income", "Cán cân vãng lai — Thu nhập sơ cấp", "tỷ USD", "higher", "trade", "quarterly"),
    ("bop", "Thu nhập thứ cấp", "bop_secondary_income", "Cán cân vãng lai — Thu nhập thứ cấp", "tỷ USD", "higher", "trade", "quarterly"),
    ("finacc", "FDI · tài sản", "finacc_fdi_assets", "Cán cân tài chính — FDI (tài sản)", "tỷ USD", "higher", "trade", "quarterly"),
    ("finacc", "FDI · nợ", "finacc_fdi_liabilities", "Cán cân tài chính — FDI (nợ)", "tỷ USD", "higher", "trade", "quarterly"),
    ("finacc", "Gián tiếp · tài sản", "finacc_portfolio_assets", "Cán cân tài chính — Đầu tư gián tiếp (tài sản)", "tỷ USD", "higher", "trade", "quarterly"),
    ("finacc", "Gián tiếp · nợ", "finacc_portfolio_liabilities", "Cán cân tài chính — Đầu tư gián tiếp (nợ)", "tỷ USD", "higher", "trade", "quarterly"),
    ("finacc", "Khác · tài sản", "finacc_other_assets", "Cán cân tài chính — Khác (tài sản)", "tỷ USD", "higher", "trade", "quarterly"),
    ("finacc", "Khác · nợ", "finacc_other_liabilities", "Cán cân tài chính — Khác (nợ)", "tỷ USD", "higher", "trade", "quarterly"),
    ("finacc_break", "ĐTTT · vốn CP", "finacc_fdi_equity", "FDI — Vốn cổ phần", "tỷ USD", "higher", "trade", "quarterly"),
    ("finacc_break", "ĐTTT · công cụ nợ", "finacc_fdi_debt_instr", "FDI — Công cụ nợ", "tỷ USD", "higher", "trade", "quarterly"),
    ("finacc_break", "Khác · tiền & tiền gửi", "finacc_other_deposits", "Đầu tư khác — Tiền & tiền gửi", "tỷ USD", "higher", "trade", "quarterly"),
    ("finacc_break", "Khác · cho vay", "finacc_other_loans", "Đầu tư khác — Cho vay", "tỷ USD", "higher", "trade", "quarterly"),
    ("finacc_dep_sec", "Ngân hàng", "finacc_deposits_bank", "Tiền gửi (cán cân tài chính) — Khu vực ngân hàng", "tỷ USD", "higher", "trade", "quarterly"),
    ("finacc_dep_sec", "Khác", "finacc_deposits_other", "Tiền gửi (cán cân tài chính) — Khu vực khác", "tỷ USD", "higher", "trade", "quarterly"),
    ("finacc_loan_sec", "Chính phủ", "finacc_loans_gov", "Vay nợ (cán cân tài chính) — Chính phủ", "tỷ USD", "higher", "trade", "quarterly"),
    ("finacc_loan_sec", "Khác", "finacc_loans_other", "Vay nợ (cán cân tài chính) — Khu vực khác", "tỷ USD", "higher", "trade", "quarterly"),
    ("adb_gdp", "Chính phủ", "govt_debt_pct_gdp_adb", "Dư nợ TPCP / GDP (ADB)", "% GDP", "lower", "fiscal", "quarterly"),
    ("adb_gdp", "Doanh nghiệp", "corp_debt_pct_gdp_adb", "Dư nợ TPDN / GDP (ADB)", "% GDP", "lower", "fiscal", "quarterly"),
    ("adb_gdp", "Ngân hàng trung ương", "cb_debt_pct_gdp_adb", "Dư nợ tín phiếu NHTW / GDP (ADB)", "% GDP", "lower", "fiscal", "quarterly"),
    ("adb_fcy", "Chính phủ", "govt_fcy_debt_usd", "Dư nợ TPCP bằng ngoại tệ", "tỷ USD", "lower", "fiscal", "quarterly"),
    ("adb_fcy", "Doanh nghiệp", "corp_fcy_debt_usd", "Dư nợ TPDN bằng ngoại tệ", "tỷ USD", "lower", "fiscal", "quarterly"),
    ("adb_mat_gov", "1–3 năm", "govt_bond_maturity_1_3y", "Cơ cấu kỳ hạn TPCP — 1-3 năm", "% dư nợ", "higher", "fiscal", "quarterly"),
    ("adb_mat_gov", "3–5 năm", "govt_bond_maturity_3_5y", "Cơ cấu kỳ hạn TPCP — 3-5 năm", "% dư nợ", "higher", "fiscal", "quarterly"),
    ("adb_mat_gov", "5–10 năm", "govt_bond_maturity_5_10y", "Cơ cấu kỳ hạn TPCP — 5-10 năm", "% dư nợ", "higher", "fiscal", "quarterly"),
    ("adb_mat_gov", "Trên 10 năm", "govt_bond_maturity_10y_plus", "Cơ cấu kỳ hạn TPCP — Trên 10 năm", "% dư nợ", "higher", "fiscal", "quarterly"),
    ("adb_mat_corp", "1–3 năm", "corp_bond_maturity_1_3y", "Cơ cấu kỳ hạn TPDN — 1-3 năm", "% dư nợ", "higher", "fiscal", "quarterly"),
    ("adb_mat_corp", "3–5 năm", "corp_bond_maturity_3_5y", "Cơ cấu kỳ hạn TPDN — 3-5 năm", "% dư nợ", "higher", "fiscal", "quarterly"),
    ("adb_mat_corp", "5–10 năm", "corp_bond_maturity_5_10y", "Cơ cấu kỳ hạn TPDN — 5-10 năm", "% dư nợ", "higher", "fiscal", "quarterly"),
    ("adb_mat_corp", "Trên 10 năm", "corp_bond_maturity_10y_plus", "Cơ cấu kỳ hạn TPDN — Trên 10 năm", "% dư nợ", "higher", "fiscal", "quarterly"),
    ("adb_holder", "Ngân hàng", "govt_bond_holder_bank", "Cơ cấu nhà đầu tư nắm giữ TPCP — Ngân hàng", "%", "higher", "fiscal", "quarterly"),
    ("adb_holder", "Bảo hiểm", "govt_bond_holder_insurance", "Cơ cấu nhà đầu tư nắm giữ TPCP — Bảo hiểm", "%", "higher", "fiscal", "quarterly"),
    ("adb_holder", "Công ty chứng khoán", "govt_bond_holder_securities", "Cơ cấu nhà đầu tư nắm giữ TPCP — CTCK", "%", "higher", "fiscal", "quarterly"),
    ("adb_holder", "Quỹ đầu tư", "govt_bond_holder_fund", "Cơ cấu nhà đầu tư nắm giữ TPCP — Quỹ đầu tư", "%", "higher", "fiscal", "quarterly"),
    ("adb_holder", "Nước ngoài", "govt_bond_holder_foreign", "Cơ cấu nhà đầu tư nắm giữ TPCP — Nước ngoài", "%", "higher", "fiscal", "quarterly"),
    ("adb_holder", "Khác", "govt_bond_holder_other", "Cơ cấu nhà đầu tư nắm giữ TPCP — Khác", "%", "higher", "fiscal", "quarterly"),
    ("ids", "Chính phủ", "intl_debt_securities_gov", "Nợ chứng khoán quốc tế — Chính phủ", "tỷ USD", "lower", "fiscal", "quarterly"),
    ("ids", "DN phi tài chính", "intl_debt_securities_corp", "Nợ chứng khoán quốc tế — DN phi tài chính", "tỷ USD", "lower", "fiscal", "quarterly"),
    ("markets_size", "Vốn hóa TTCK", "stock_market_cap_pct_gdp", "Vốn hóa thị trường chứng khoán / GDP", "% GDP", "higher", "market", "annual"),
    ("fiscal_bal", "Thu ngân sách", "fiscal_revenue_pct_gdp", "Thu ngân sách / GDP", "% GDP", "higher", "fiscal", "annual"),
    ("fiscal_bal", "Chi ngân sách", "fiscal_expenditure_pct_gdp", "Chi ngân sách / GDP", "% GDP", "lower", "fiscal", "annual"),
    ("fiscal_bal", "Cân đối tổng thể", "fiscal_balance_overall_pct_gdp", "Cân đối ngân sách tổng thể / GDP", "% GDP", "higher", "fiscal", "annual"),
    ("fiscal_bal", "Nợ công / GDP", "public_debt_pct_gdp_forecast", "Nợ công / GDP (có dự báo tới 2031)", "% GDP", "lower", "fiscal", "annual"),
    ("lt_fiscal", "Cân đối sơ cấp", "fiscal_primary_balance_pct_gdp", "Cân đối ngân sách sơ cấp / GDP", "% GDP", "higher", "fiscal", "annual"),
    ("demo_urb", "Tỷ lệ đô thị hóa", "urbanization_rate", "Tỷ lệ đô thị hóa", "% dân số", "higher", "demographics", "annual"),
    ("demo_age", "0–14", "population_age_0_14", "Dân số 0-14 tuổi", "triệu người", "higher", "demographics", "annual"),
    ("demo_age", "15–64", "population_age_15_64", "Dân số 15-64 tuổi (lao động)", "triệu người", "higher", "demographics", "annual"),
    ("demo_age", "65+", "population_age_65_plus", "Dân số 65+ tuổi", "triệu người", "higher", "demographics", "annual"),
    ("demo_dep", "Chung", "dependency_ratio_total", "Tỷ lệ phụ thuộc chung", "%", "lower", "demographics", "annual"),
    ("demo_dep", "Trẻ em", "dependency_ratio_youth", "Tỷ lệ phụ thuộc — Trẻ em", "%", "lower", "demographics", "annual"),
    ("demo_dep", "Người già", "dependency_ratio_elderly", "Tỷ lệ phụ thuộc — Người già", "%", "lower", "demographics", "annual"),
    ("gdp_exp", "Tiêu dùng hộ GĐ", "gdp_share_use_consumption_household", "GDP theo sử dụng — Tiêu dùng hộ gia đình", "% GDP", "higher", "growth", "annual"),
    ("gdp_exp", "Tiêu dùng Chính phủ", "gdp_share_use_consumption_government", "GDP theo sử dụng — Tiêu dùng Chính phủ", "% GDP", "higher", "growth", "annual"),
    ("gdp_exp", "Đầu tư (GFCF)", "gdp_share_use_investment", "GDP theo sử dụng — Đầu tư (GFCF)", "% GDP", "higher", "growth", "annual"),
    ("gdp_exp", "Xuất khẩu", "gdp_share_use_export", "GDP theo sử dụng — Xuất khẩu", "% GDP", "higher", "growth", "annual"),
    ("gdp_exp", "Nhập khẩu", "gdp_share_use_import", "GDP theo sử dụng — Nhập khẩu", "% GDP", "higher", "growth", "annual"),
    ("gdp_sector", "Nông nghiệp", "gdp_share_agri_annual_40yo", "Cơ cấu GDP theo ngành — Nông-Lâm-Thủy sản (theo năm, từ 1986)", "% GDP", "higher", "growth", "annual"),
    ("gdp_sector", "Công nghiệp", "gdp_share_industry_annual_40yo", "Cơ cấu GDP theo ngành — Công nghiệp-Xây dựng (theo năm, từ 1986)", "% GDP", "higher", "growth", "annual"),
    ("gdp_sector", "Dịch vụ", "gdp_share_services_annual_40yo", "Cơ cấu GDP theo ngành — Dịch vụ (theo năm, từ 1986)", "% GDP", "higher", "growth", "annual"),
    ("inc_saving", "Tiết kiệm quốc gia", "national_savings_pct_gdp", "Tiết kiệm quốc gia / GDP", "% GDP", "higher", "growth", "annual"),
    ("inc_saving", "Tổng đầu tư", "gross_investment_pct_gdp", "Tổng đầu tư / GDP", "% GDP", "higher", "growth", "annual"),
    ("inc_pc", "Danh nghĩa (USD)", "gdp_per_capita_nominal_usd", "GDP bình quân đầu người (danh nghĩa)", "USD", "higher", "growth", "annual"),
    ("inc_pc", "PPP (Int$)", "gdp_per_capita_ppp_intl", "GDP bình quân đầu người (PPP)", "Int$", "higher", "growth", "annual"),
    ("inc_pc", "GNI/người (Atlas USD)", "gni_per_capita_atlas_usd", "GNI bình quân đầu người (Atlas)", "USD", "higher", "growth", "annual"),
]


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

        # XNK tách khu vực TRONG NƯỚC vs FDI, lũy kế theo THÁNG (tỷ USD) — user (2026-10-01) muốn
        # "FDI Trade Balance" riêng (KHÔNG dùng để suy ra USD supply/demand thực — chỉ là "trade
        # structure" theo đúng khuyến nghị, xem _add_fdi_domestic_trade_derived nếu có). Câu xuất
        # hiện HÀNG THÁNG trong báo cáo, dạng "Tính chung {N} tháng năm {YYYY}, kim ngạch xuất/
        # nhập khẩu hàng hóa đạt {total} tỷ USD, tăng {total%}%..., Trong đó/trong đó, khu vực
        # kinh tế trong nước đạt {domestic} tỷ USD, tăng {dom%}%...; khu vực có vốn đầu tư nước
        # ngoài [(kể cả dầu thô)] đạt {fdi} tỷ USD, tăng {fdi%}%". Verify khớp CHÍNH XÁC báo cáo
        # tháng 8/2026 (export: trong nước 74,47 / FDI 300,37 / tổng 374,84 tỷ USD; import: trong
        # nước 105,07 / FDI 290,23 / tổng 395,3 tỷ USD).
        m_exp_fdi = re.search(
            r"Tính chung (\S+(?:\s+một)?) tháng năm (\d{4}), kim ngạch xuất khẩu hàng hóa đạt "
            r"([\d.,]+) tỷ USD, tăng ([\d.,\-]+)% so với cùng kỳ năm trước\. [Tt]rong đó,? khu vực "
            r"kinh tế trong nước đạt ([\d.,]+) tỷ USD, tăng ([\d.,\-]+)%[^;]*; khu vực có vốn đầu "
            r"tư (?:trực tiếp )?nước ngoài[^đ]*đạt ([\d.,]+) tỷ USD, tăng ([\d.,\-]+)%", text)
        if m_exp_fdi:
            month = _VN_MONTH_COUNT_WORDS.get(m_exp_fdi.group(1).lower())
            if month:
                out["trade_fdi_split_period"] = f"{int(m_exp_fdi.group(2)):04d}-{month:02d}"
                out["export_domestic_usd_bn"] = _vn_number(m_exp_fdi.group(5))
                out["export_fdi_usd_bn"] = _vn_number(m_exp_fdi.group(7))

        m_imp_fdi = re.search(
            r"Tính chung (\S+(?:\s+một)?) tháng năm (\d{4}), kim ngạch nhập khẩu hàng hóa đạt "
            r"([\d.,]+) tỷ USD, tăng ([\d.,\-]+)% so với cùng kỳ năm trước,? trong đó khu vực "
            r"kinh tế trong nước đạt ([\d.,]+) tỷ USD, tăng ([\d.,\-]+)%[^;]*; khu vực có vốn đầu "
            r"tư (?:trực tiếp )?nước ngoài[^đ]*đạt ([\d.,]+) tỷ USD, tăng ([\d.,\-]+)%", text)
        if m_imp_fdi:
            month = _VN_MONTH_COUNT_WORDS.get(m_imp_fdi.group(1).lower())
            if month:
                period = f"{int(m_imp_fdi.group(2)):04d}-{month:02d}"
                if not out.get("trade_fdi_split_period"):
                    out["trade_fdi_split_period"] = period
                out["import_domestic_usd_bn"] = _vn_number(m_imp_fdi.group(5))
                out["import_fdi_usd_bn"] = _vn_number(m_imp_fdi.group(7))

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


# Nhãn chính xác trên bảng BOP (user 2026-10-01, cung cấp link trực tiếp) -> field_key. GIỮ NGUYÊN
# dấu câu/khoảng trắng CHÍNH XÁC như trên trang (đã verify qua khảo sát thủ công, xem
# fetch_sbv_bop_quarterly()) — regex nối thêm "\|(-?[\d]+)\|" để lấy giá trị (Triệu USD) ngay sau
# nhãn trên bảng đã làm sạch HTML. CHỈ lấy các dòng TRỌNG YẾU nhất theo đúng khung "Áp lực Ngoại
# tệ" (Cầu/Cung/Đối chiếu BOP) — KHÔNG lấy hết ~40 dòng chi tiết nhất (vd "Tổ chức tín dụng"/"Dân
# cư" lồng trong "Tiền và tiền gửi") vì quá chi tiết/dễ nhầm nhãn trùng giữa các mục cha-con.
_SBV_BOP_FIELD_MAP = [
    ("current_account", r"A\. Cán cân vãng lai"),
    ("goods_export", r"Hàng hóa: Xuất khẩu f\.o\.b"),
    ("goods_import", r"Hàng hóa: Nhập khẩu f\.o\.b"),
    ("services_export", r"Dịch vụ: Xuất khẩu"),
    ("services_import", r"Dịch vụ: Nhập khẩu"),
    ("investment_income_received", r"Thu nhập đầu tư \(Thu nhập sơ cấp\): Thu"),
    ("investment_income_paid", r"Thu nhập đầu tư \(Thu nhập sơ cấp\): Chi"),
    ("secondary_income_received", r"Chuyển giao vãng lai \(Thu nhập thứ cấp\): Thu"),
    ("secondary_income_paid", r"Chuyển giao vãng lai \(Thu nhập thứ cấp\): Chi"),
    ("financial_account", r"C\. Cán cân tài chính"),
    ("fdi_assets_bop", r"Đầu tư trực tiếp ra nước ngoài: Tài sản có"),
    ("fdi_liabilities_bop", r"Đầu tư trực tiếp vào Việt Nam: Tài sản nợ"),
    ("portfolio_assets_bop", r"Đầu tư gián tiếp ra nước ngoài: Tài sản có"),
    ("portfolio_liabilities_bop", r"Đầu tư gián tiếp vào Việt Nam: Tài sản nợ"),
    ("external_debt_net", r"Vay, trả nợ nước ngoài"),
    ("errors_omissions", r"D\. Lỗi và Sai sót"),
    ("overall_balance", r"E\. Cán cân tổng thể"),
    ("reserve_assets_change", r"Tài sản dự trữ"),
]
_ROMAN_TO_QUARTER = {"I": 1, "II": 2, "III": 3, "IV": 4}

# Nhãn CHÍNH XÁC giống _SBV_BOP_FIELD_MAP (không có escape regex vì ở đây so KHỚP CHUỖI, không
# phải regex) — dulieukinhte.com mirror lại số liệu NHNN, dùng Y HỆT tên chỉ tiêu gốc.
_DULIEUKINHTE_BOP_LABEL_MAP = [
    ("current_account", "A. Cán cân vãng lai"),
    ("goods_export", "Hàng hóa: Xuất khẩu f.o.b"),
    ("goods_import", "Hàng hóa: Nhập khẩu f.o.b"),
    ("services_export", "Dịch vụ: Xuất khẩu"),
    ("services_import", "Dịch vụ: Nhập khẩu"),
    ("investment_income_received", "Thu nhập đầu tư (Thu nhập sơ cấp): Thu"),
    ("investment_income_paid", "Thu nhập đầu tư (Thu nhập sơ cấp): Chi"),
    ("secondary_income_received", "Chuyển giao vãng lai (Thu nhập thứ cấp): Thu"),
    ("secondary_income_paid", "Chuyển giao vãng lai (Thu nhập thứ cấp): Chi"),
    ("financial_account", "C. Cán cân tài chính"),
    ("fdi_assets_bop", "Đầu tư trực tiếp ra nước ngoài: Tài sản có"),
    ("fdi_liabilities_bop", "Đầu tư trực tiếp vào Việt Nam: Tài sản nợ"),
    ("portfolio_assets_bop", "Đầu tư gián tiếp ra nước ngoài: Tài sản có"),
    ("portfolio_liabilities_bop", "Đầu tư gián tiếp vào Việt Nam: Tài sản nợ"),
    ("external_debt_net", "Vay, trả nợ nước ngoài"),
    ("errors_omissions", "D. Lỗi và Sai sót"),
    ("overall_balance", "E. Cán cân tổng thể"),
    ("reserve_assets_change", "Tài sản dự trữ"),
]


_IMF_BOP_FIELD_MAP = [
    # (field_key, BOP_ACCOUNTING_ENTRY, INDICATOR, sign) — sign=-1 nghia la IMF bao theo quy uoc
    # NGUOC voi NHNN (vd "Tai san" ben IMF duong = tang, nhung NHNN trinh bay am = tang - xem note
    # reserve_assets_change o cho goi). DA VERIFY tung field bang cach so khop GIA TRI THAT voi
    # bop_sbv_* da co (NHNN/dulieukinhte) qua 6 quy chong lan (2025-Q1..2026-Q1) - khop CHINH XAC
    # tung dong, tung quy (user 2026-10-03, phat hien IMF co BOP Viet Nam tu 1996-Q1 khi tim cach
    # "ve bieu do tu 2021 toi nay" cho overall_balance/current_account).
    ("current_account", "NETCD_T", "CAB", 1),
    ("goods_export", "CD_T", "G", 1),
    ("goods_import", "DB_T", "G", 1),
    ("services_export", "CD_T", "S", 1),
    ("services_import", "DB_T", "S", 1),
    ("investment_income_received", "CD_T", "IN1", 1),
    ("investment_income_paid", "DB_T", "IN1", 1),
    ("secondary_income_received", "CD_T", "IN2", 1),
    ("secondary_income_paid", "DB_T", "IN2", 1),
    ("financial_account", "NNAFANIL_T", "FABXRRI", -1),
    ("fdi_assets_bop", "A_NFA_T", "D_F", -1),
    ("fdi_liabilities_bop", "L_NIL_T", "D_F", 1),
    ("portfolio_assets_bop", "A_NFA_T", "P_F", -1),
    ("portfolio_liabilities_bop", "L_NIL_T", "P_F", 1),
    ("external_debt_net", "L_NIL_T", "O_F4", 1),
    ("errors_omissions", "NETCD_T", "EO", 1),
    ("overall_balance", "A_T", "R_F", 1),
]


def fetch_imf_bop_vietnam_history():
    """api.imf.org (SDMX 2.1, dataflow IMF.STA:BOP(21.0.0), datastructure DSD_BOP) — BOP Việt Nam
    ĐẦY ĐỦ từ 1996-Q1 tới hiện tại, KHÔNG CẦN đăng nhập (khác portal.api.imf.org cần sign-in —
    endpoint data trực tiếp KHÔNG yêu cầu, đã verify qua gọi thật 2026-10-03). Dùng để BACKFILL
    lịch sử xa hơn ~5 quý mà dulieukinhte.com/SBV trực tiếp có (user: "vẽ biểu đồ từ 2021 tới nay
    thì hay quá"). CHỈ 1 lệnh gọi duy nhất (wildcard BOP_ACCOUNTING_ENTRY+INDICATOR) lấy về ~150
    series, lọc đúng 17/18 field qua _IMF_BOP_FIELD_MAP (reserve_assets_change KHÔNG cần field
    riêng — LUÔN = −overall_balance theo đúng definition BPM6, xem chỗ gọi). Đơn vị USD thô (SCALE
    "6" là metadata của IMF, KHÔNG phải hệ số cần tự chia/nhân — giá trị OBS_VALUE đã là USD đầy
    đủ), quy đổi /1e6 ra triệu USD khớp đơn vị đang dùng xuyên suốt (bop_sbv_*).
    Trả {field_key: {period ('YYYY-Qn'): value_trieu_usd}} — rỗng nếu lỗi/đổi cấu trúc API."""
    try:
        url = "https://api.imf.org/external/sdmx/2.1/data/IMF.STA,BOP,21.0.0/VNM...USD.Q"
        r = requests.get(url, headers={"User-Agent": UA}, timeout=30)
        r.raise_for_status()
        text = r.text

        out = {key: {} for key, _, _, _ in _IMF_BOP_FIELD_MAP}
        for key, entry_code, indicator_code, sign in _IMF_BOP_FIELD_MAP:
            m = re.search(
                r'<Series [^>]*BOP_ACCOUNTING_ENTRY="' + re.escape(entry_code) + r'"[^>]*'
                r'INDICATOR="' + re.escape(indicator_code) + r'"[^>]*>(.*?)</Series>', text, re.S)
            if not m:
                # Thu tim theo thu tu attribute nguoc lai (INDICATOR truoc BOP_ACCOUNTING_ENTRY) -
                # XML SDMX khong dam bao thu tu attribute co dinh.
                m = re.search(
                    r'<Series [^>]*INDICATOR="' + re.escape(indicator_code) + r'"[^>]*'
                    r'BOP_ACCOUNTING_ENTRY="' + re.escape(entry_code) + r'"[^>]*>(.*?)</Series>', text, re.S)
            if not m:
                continue
            body = m.group(1)
            for period, val in re.findall(r'TIME_PERIOD="(\d{4}-Q[1-4])" OBS_VALUE="(-?[\d.]+)"', body):
                out[key][period] = round(sign * float(val) / 1e6, 2)
        return {k: v for k, v in out.items() if v}
    except Exception as e:
        print(f"  [WARN] IMF BOP Việt Nam thất bại: {e}")
        return {}


def fetch_dulieukinhte_bop():
    """dulieukinhte.com/du-lieu/can-can-thanh-toan-361 — MIRROR lại số liệu BOP quý của NHNN
    (ghi rõ "Nguồn: Ngân hàng Nhà nước Việt Nam"), nhưng trình bày HTML TĨNH thường (KHÔNG WAF
    chặn như trang NHNN gốc — user 2026-10-01: "để đỡ bị chặn" khi cần lấy gấp). Trang chỉ hiện
    5 quý gần nhất (tại thời điểm viết: Q1/2025 → Q1/2026) — KHÔNG đủ xa để thay thế hoàn toàn
    fetch_sbv_bop_quarterly() cho lịch sử sâu hơn (vd 2024), nhưng là NGUỒN ƯU TIÊN cho các quý
    GẦN NHẤT vì ít rủi ro bị chặn hơn nguồn SBV trực tiếp.

    Parse THEO TỪNG <tr> (không flatten hết trang thành 1 chuỗi text) để tránh nhầm với badge
    "N mục con" (vd `<span class="kid-count">12</span>`) chỉ xuất hiện ở các dòng NHÓM (A/C/F) —
    nếu flatten hết thành text, số badge này dễ bị hiểu nhầm là 1 giá trị kỳ thứ 6.

    Trả {period ('YYYY-Qn'): {field_key: value_trieu_usd}} — rỗng nếu lỗi/đổi cấu trúc trang."""
    url = "https://dulieukinhte.com/du-lieu/can-can-thanh-toan-361"
    try:
        r = requests.get(url, headers={"User-Agent": UA}, timeout=20)
        r.raise_for_status()
        text = r.text

        m_table = re.search(r'<table class="table-macro-data.*?</table>', text, re.S)
        if not m_table:
            print("  [WARN] dulieukinhte.com BOP: không tìm thấy bảng — có thể đổi cấu trúc trang.")
            return {}
        table_html = m_table.group(0)

        periods_raw = re.findall(r'<th scope="col">(Q[1-4]-\d{4})</th>', table_html)
        if not periods_raw:
            print("  [WARN] dulieukinhte.com BOP: không tìm thấy header kỳ (Qn-YYYY).")
            return {}
        periods = []
        for p in periods_raw:
            pm = re.match(r"Q([1-4])-(\d{4})", p)
            periods.append(f"{pm.group(2)}-Q{pm.group(1)}")

        label_to_key = dict((label, key) for key, label in _DULIEUKINHTE_BOP_LABEL_MAP)
        out = {p: {} for p in periods}
        for row_html in re.findall(r"<tr[^>]*>(.*?)</tr>", table_html, re.S):
            # class attribute CÓ THỂ là "row-link" (dòng có trang chi tiết riêng, thẻ <a>) HOẶC
            # "row-link soft-link" (dòng không có trang chi tiết, thẻ <span>) — PHẢI match class
            # bắt đầu bằng "row-link" (không yêu cầu đúng y nguyên), nếu không sẽ bỏ sót phần lớn
            # dòng (đã gặp thực tế: chỉ bắt được 6/40+ dòng khi match chính xác "row-link">).
            m_label = re.search(r'class="row-link[^"]*">([^<]+)<', row_html)
            if not m_label:
                continue
            key = label_to_key.get(m_label.group(1).strip())
            if not key:
                continue
            values = re.findall(r'<td class="">([^<]*)</td>', row_html)
            if len(values) < len(periods):
                continue
            for period, val_str in zip(periods, values):
                val_str = val_str.strip()
                if val_str and val_str != "-":
                    out[period][key] = _vn_number(val_str)
        return {p: v for p, v in out.items() if v}
    except Exception as e:
        print(f"  [WARN] dulieukinhte.com BOP thất bại: {e}")
        return {}


def fetch_dulieukinhte_kieu_hoi_hcm():
    """dulieukinhte.com/du-lieu/kieu-hoi-tp-ho-chi-minh-426 — Kiều hối ĐÚNG NGHĨA (personal
    remittances) nhưng CHỈ PHẠM VI TP.HCM (nguồn: NHNN Chi nhánh Khu vực 2, trước 01/07/2025 là
    NHNN chi nhánh TP.HCM). LƯU Ý: đây KHÔNG PHẢI kiều hối toàn quốc — TP.HCM lịch sử chiếm tỷ
    trọng lớn nhưng không phải 100%, KHÔNG được tự suy ra số toàn quốc từ số này. Khác với
    bop_sbv_secondary_income_received (toàn quốc nhưng RỘNG HƠN, gồm mọi chuyển giao vãng lai chứ
    không riêng kiều hối cá nhân) — 2 chỉ báo bổ sung cho nhau, không thay thế nhau. HTML tĩnh,
    không WAF. Trả {period ('YYYY-Qn'): value_trieu_usd} — rỗng nếu lỗi/đổi cấu trúc trang."""
    url = "https://dulieukinhte.com/du-lieu/kieu-hoi-tp-ho-chi-minh-426"
    try:
        r = requests.get(url, headers={"User-Agent": UA}, timeout=20)
        r.raise_for_status()
        text = r.text

        m_table = re.search(r'<table class="table-macro-data.*?</table>', text, re.S)
        if not m_table:
            print("  [WARN] dulieukinhte.com kiều hối TP.HCM: không tìm thấy bảng — có thể đổi cấu trúc trang.")
            return {}
        table_html = m_table.group(0)

        periods_raw = re.findall(r'<th scope="col">(Q[1-4]-\d{4})</th>', table_html)
        if not periods_raw:
            print("  [WARN] dulieukinhte.com kiều hối TP.HCM: không tìm thấy header kỳ (Qn-YYYY).")
            return {}
        periods = []
        for p in periods_raw:
            pm = re.match(r"Q([1-4])-(\d{4})", p)
            periods.append(f"{pm.group(2)}-Q{pm.group(1)}")

        out = {}
        for row_html in re.findall(r"<tr[^>]*>(.*?)</tr>", table_html, re.S):
            m_label = re.search(r'class="row-link[^"]*">([^<]+)<', row_html)
            if not m_label or "Tổng kiều hối" not in m_label.group(1):
                continue
            values = re.findall(r'<td class="">([^<]*)</td>', row_html)
            if len(values) < len(periods):
                continue
            for period, val_str in zip(periods, values):
                val_str = val_str.strip()
                if val_str and val_str != "-":
                    out[period] = _vn_number(val_str)
            break
        return out
    except Exception as e:
        print(f"  [WARN] dulieukinhte.com kiều hối TP.HCM thất bại: {e}")
        return {}


def _fetch_dulieukinhte_table(url, quarterly, label_filter=None):
    """Scraper DÙNG CHUNG cho các trang dạng bảng "table-macro-data" của dulieukinhte.com có
    NHIỀU HÀNG đều muốn lấy hết (khác fetch_dulieukinhte_bop/fetch_dulieukinhte_kieu_hoi_hcm chỉ
    lọc đúng field cần, 2 hàm đó giữ nguyên không refactor vì đang chạy ổn). quarterly=True đọc
    header dạng "Qn-YYYY" (BOP/kiều hối/GDP...), False đọc "MM-YYYY" (CPI/bán lẻ...). label_filter
    (set[str] hoặc None=lấy hết) để lọc bớt nếu chỉ cần vài hàng. Trả {label: {period: value}} —
    rỗng nếu lỗi/đổi cấu trúc trang."""
    try:
        r = requests.get(url, headers={"User-Agent": UA}, timeout=20)
        r.raise_for_status()
        text = r.text

        m_table = re.search(r'<table class="table-macro-data.*?</table>', text, re.S)
        if not m_table:
            print(f"  [WARN] dulieukinhte.com ({url}): không tìm thấy bảng — có thể đổi cấu trúc trang.")
            return {}
        table_html = m_table.group(0)

        if quarterly:
            periods_raw = re.findall(r'<th scope="col">(Q[1-4]-\d{4})</th>', table_html)
            periods = []
            for p in periods_raw:
                pm = re.match(r"Q([1-4])-(\d{4})", p)
                periods.append(f"{pm.group(2)}-Q{pm.group(1)}")
        else:
            periods_raw = re.findall(r'<th scope="col">(\d{2})-(\d{4})</th>', table_html)
            periods = [f"{y}-{m}" for m, y in periods_raw]
        if not periods:
            print(f"  [WARN] dulieukinhte.com ({url}): không tìm thấy header kỳ.")
            return {}

        out = {}
        for row_html in re.findall(r"<tr[^>]*>(.*?)</tr>", table_html, re.S):
            m_label = re.search(r'class="row-link[^"]*">([^<]+)<', row_html)
            if not m_label:
                continue
            label = m_label.group(1).strip()
            if label_filter is not None and label not in label_filter:
                continue
            values = re.findall(r'<td class="">([^<]*)</td>', row_html)
            if len(values) < len(periods):
                continue
            row_data = {}
            for period, val_str in zip(periods, values):
                val_str = val_str.strip()
                if val_str and val_str != "-":
                    row_data[period] = _vn_number(val_str)
            if row_data:
                out[label] = row_data
        return out
    except Exception as e:
        print(f"  [WARN] dulieukinhte.com ({url}) thất bại: {e}")
        return {}


def fetch_dulieukinhte_gdp_sector():
    """dulieukinhte.com/du-lieu/gdp-so-sanh-theo-quy-269 — GDP theo giá so sánh (quý), tách theo
    NGÀNH kinh tế (nông-lâm-thủy sản/công nghiệp-xây dựng/dịch vụ + các ngành con), nguồn Cục
    Thống kê (NSO). Đây là MỨC (tỷ đồng, giá so sánh — KHÔNG phải %YoY), dùng để tự tính YoY theo
    ngành ở template_vimo._build_level_yoy_heatmap(). Trả {ngành: {period: value_ty_dong}}."""
    return _fetch_dulieukinhte_table("https://dulieukinhte.com/du-lieu/gdp-so-sanh-theo-quy-269", quarterly=True)


def fetch_dulieukinhte_cpi_group():
    """dulieukinhte.com/du-lieu/chi-so-gia-tieu-dung-cpi-272 — Chỉ số giá tiêu dùng (CPI) theo
    NHÓM HÀNG (tháng), nguồn NSO. Đây là CHỈ SỐ (điểm, gốc so sánh nội bộ NSO — KHÔNG phải %YoY),
    dùng để tự tính YoY theo nhóm hàng ở template_vimo._build_level_yoy_heatmap(). Trả {nhóm hàng:
    {period: index_points}}."""
    return _fetch_dulieukinhte_table("https://dulieukinhte.com/du-lieu/chi-so-gia-tieu-dung-cpi-272", quarterly=False)


def fetch_dulieukinhte_export_commodity():
    """dulieukinhte.com/du-lieu/xuat-khau-theo-mat-hang-313 — Kim ngạch XUẤT KHẨU theo MẶT HÀNG
    (tháng, ~56 mặt hàng), nguồn Tổng cục Hải quan. MỨC (triệu USD — KHÔNG phải %YoY), dùng để tự
    tính YoY theo mặt hàng ở template_vimo._build_level_yoy_heatmap(). Trả {mặt hàng: {period:
    value_trieu_usd}}."""
    return _fetch_dulieukinhte_table("https://dulieukinhte.com/du-lieu/xuat-khau-theo-mat-hang-313", quarterly=False)


def fetch_dulieukinhte_import_commodity():
    """dulieukinhte.com/du-lieu/nhap-khau-theo-mat-hang-314 — Kim ngạch NHẬP KHẨU theo MẶT HÀNG
    (tháng, ~59 mặt hàng), nguồn Tổng cục Hải quan. MỨC (triệu USD). Trả {mặt hàng: {period:
    value_trieu_usd}}."""
    return _fetch_dulieukinhte_table("https://dulieukinhte.com/du-lieu/nhap-khau-theo-mat-hang-314", quarterly=False)


def fetch_dulieukinhte_export_price():
    """dulieukinhte.com/du-lieu/gia-xuat-khau-286 — Giá XUẤT KHẨU bình quân theo mặt hàng (tháng,
    ~19 mặt hàng, USD/tấn), nguồn Tổng cục Hải quan. MỨC (USD/tấn — KHÔNG phải %YoY). Trả {mặt
    hàng: {period: value_usd_per_ton}}."""
    return _fetch_dulieukinhte_table("https://dulieukinhte.com/du-lieu/gia-xuat-khau-286", quarterly=False)


def fetch_dulieukinhte_import_price():
    """dulieukinhte.com/du-lieu/gia-nhap-khau-287 — Giá NHẬP KHẨU bình quân theo mặt hàng (tháng,
    ~19 mặt hàng, USD/tấn), nguồn Tổng cục Hải quan. MỨC (USD/tấn). Trả {mặt hàng: {period:
    value_usd_per_ton}}."""
    return _fetch_dulieukinhte_table("https://dulieukinhte.com/du-lieu/gia-nhap-khau-287", quarterly=False)


def fetch_dulieukinhte_retail_sales():
    """dulieukinhte.com/du-lieu/ban-le-hang-hoa-va-dich-vu-273 — Tổng mức bán lẻ hàng hóa & dịch
    vụ tiêu dùng theo THÁNG (ĐƠN LẺ, giá hiện hành — DANH NGHĨA, chưa trừ lạm phát), tách 4 cấu
    phần (Bán lẻ hàng hoá / Dịch vụ lưu trú-ăn uống / Du lịch lữ hành / Dịch vụ khác), nguồn Cục
    Thống kê (NSO). User (2026-10-01) muốn dùng để đánh giá sức tiêu thụ/nhu cầu chi tiêu — xem
    _add_retail_sales_derived() trong template_vimo.py (tính YoY danh nghĩa + THỰC, trừ CPI).
    Backfill DÀY HƠN NHIỀU so với nguồn cũ (quét câu văn báo cáo NSO, dễ vỡ, hiện chỉ có 1-2 điểm
    rải rác/năm — xem gdp_struct["retail_sales_yoy_pct"]). Trả {dòng: {period: value_ty_dong}}."""
    return _fetch_dulieukinhte_table("https://dulieukinhte.com/du-lieu/ban-le-hang-hoa-va-dich-vu-273", quarterly=False)


def fetch_dulieukinhte_export_fdi_split():
    """dulieukinhte.com/du-lieu/tong-xuat-khau-292 — Kim ngạch XUẤT KHẨU theo THÁNG (ĐƠN LẺ, không
    phải lũy kế như export_domestic_usd_bn/export_fdi_usd_bn hiện có — xem fetch_nso_gdp_structure_
    report()), tách khu vực trong nước/FDI, nguồn Tổng cục Hải quan. Trả {"Tổng"/"Khu vực trong
    nước"/"Khu vực trong FDI": {period: value_trieu_usd}}."""
    return _fetch_dulieukinhte_table("https://dulieukinhte.com/du-lieu/tong-xuat-khau-292", quarterly=False)


def fetch_dulieukinhte_import_fdi_split():
    """dulieukinhte.com/du-lieu/tong-nhap-khau-293 — Kim ngạch NHẬP KHẨU theo THÁNG (ĐƠN LẺ, không
    phải lũy kế), tách khu vực trong nước/FDI, nguồn Tổng cục Hải quan. Trả {"Tổng"/"Khu vực trong
    nước"/"Khu vực trong FDI": {period: value_trieu_usd}}."""
    return _fetch_dulieukinhte_table("https://dulieukinhte.com/du-lieu/tong-nhap-khau-293", quarterly=False)


def fetch_darvas_reer_neer_vietnam():
    """bruegel.org — bộ dữ liệu NEER/REER của Zsolt Darvas (cập nhật định kỳ, KHÔNG phải nguồn
    "chính thức" IMF/BIS nhưng user xác nhận IMF TỰ DÙNG Darvas làm nguồn NEER/REER cho Việt Nam
    trong phân tích Article IV — vì BIS (WS_EER) và IMF.STA:EER (dataset mới) đều KHÔNG có Việt
    Nam (đã verify qua API thật, xem lịch sử trao đổi 2026-10-01). Darvas CÓ Việt Nam, theo
    THÁNG, từ 1993 (NEER)/1995 (REER) tới gần hiện tại — basket "120" (Broad, 120 đối tác
    thương mại) đúng loại user muốn ("Broad", không cần "Narrow").

    File .zip có TÊN CHỨA NGÀY cập nhật (vd 'REER_database_ver22Sep2026.zip') — KHÔNG hard-code
    URL, phải crawl link download MỚI NHẤT từ trang chính mỗi lần chạy (trang này tự cập nhật
    link khi Bruegel ra bản mới). File ~5MB, chấp nhận được để tải lại mỗi lần chạy (không cache
    riêng — đơn giản hơn, khớp quy ước các nguồn khác trong file này không cache riêng).

    Trả {"neer": [(period, value)...], "reer": [(period, value)...]} (period dạng 'YYYY-MM') —
    rỗng nếu lỗi/thiếu xlrd/không tìm thấy Việt Nam trong bảng."""
    try:
        import xlrd
    except ImportError:
        print("  [WARN] Thiếu thư viện xlrd (pip install xlrd) — bỏ qua Darvas NEER/REER.")
        return {"neer": [], "reer": []}
    try:
        r = requests.get(
            "https://www.bruegel.org/publications/datasets/real-effective-exchange-rates-for-178-countries-a-new-database",
            headers={"User-Agent": UA}, timeout=20)
        r.raise_for_status()
        m = re.search(r'href="([^"]*REER_database[^"]*\.zip)"', r.text)
        if not m:
            print("  [WARN] Darvas NEER/REER: không tìm thấy link download .zip trên trang — có thể đổi cấu trúc.")
            return {"neer": [], "reer": []}
        zip_url = m.group(1)
        if zip_url.startswith("/"):
            zip_url = "https://www.bruegel.org" + zip_url

        import zipfile, io
        rz = requests.get(zip_url, headers={"User-Agent": UA}, timeout=60)
        rz.raise_for_status()
        z = zipfile.ZipFile(io.BytesIO(rz.content))
        xls_name = next((n for n in z.namelist() if n.lower().endswith(".xls")), None)
        if not xls_name:
            print("  [WARN] Darvas NEER/REER: không tìm thấy file .xls trong .zip.")
            return {"neer": [], "reer": []}
        wb = xlrd.open_workbook(file_contents=z.read(xls_name))

        out = {"neer": [], "reer": []}
        for sheet_name, out_key in [("NEER_MONTHLY_120", "neer"), ("REER_MONTHLY_120", "reer")]:
            if sheet_name not in wb.sheet_names():
                continue
            sh = wb.sheet_by_name(sheet_name)
            col_prefix = f"{out_key.upper()}_120_VN"
            col = next((c for c in range(sh.ncols) if sh.cell_value(0, c) == col_prefix), None)
            if col is None:
                print(f"  [WARN] Darvas {sheet_name}: không tìm thấy cột Việt Nam ('{col_prefix}').")
                continue
            for row in range(1, sh.nrows):
                period_raw = sh.cell_value(row, 0)
                pm = re.match(r"(\d{4})M(\d{2})", str(period_raw))
                if not pm:
                    continue
                value = sh.cell_value(row, col)
                if value == "":
                    continue
                out[out_key].append((f"{pm.group(1)}-{pm.group(2)}", float(value)))
        return out
    except Exception as e:
        print(f"  [WARN] Darvas NEER/REER thất bại: {e}")
        return {"neer": [], "reer": []}


def fetch_sbv_bop_quarterly():
    """sbv.gov.vn/vi/can-can-thanh-toan-quoc-te — bảng Cán cân thanh toán quốc tế (BOP) CHÍNH
    THỨC của NHNN, chuẩn BPM6 đầy đủ (Cán cân vãng lai/vốn/tài chính, Lỗi và Sai sót, Cán cân
    tổng thể, Dự trữ) — HTML thật, KHÔNG cần JS render, nhưng có WAF (F5 BIG-IP) chặn request
    thiếu header "giống browser" (chỉ User-Agent KHÔNG ĐỦ — thiếu Accept/Accept-Language/Referer
    sẽ bị trả trang "Request Rejected" 244 bytes thay vì nội dung thật, đã verify qua khảo sát
    thủ công 2026-10-01). Trang CHÍNH LUÔN hiện QUÝ MỚI NHẤT đã công bố — KHÔNG dùng URL phụ
    "/-/asset_publisher/.../content/quý-x-năm-y" để lấy quý cũ: đã test, URL đó bị WAF chặn NGẮT
    QUÃNG (có lúc 200 kèm nội dung đầy, có lúc 200 kèm trang lỗi F5 generic "incident ID: N/A")
    — KHÔNG đủ ổn định để backfill lịch sử. Lịch sử sẽ tự tích lũy dần qua các lần Action chạy
    (mỗi quý NHNN công bố, 1 lần chạy mới sẽ bắt được).

    NGAY CẢ trang CHÍNH cũng bị chặn NGẮT QUÃNG (đã gặp thực tế: "403 ... No server for url" dù
    cùng URL/header vừa chạy thành công vài phút trước — khả năng cao là rate-limit tạm thời của
    WAF, không phải lỗi cấu trúc trang) — thử lại TỐI ĐA 3 LẦN, cách nhau vài giây, trước khi bỏ
    cuộc. KHÔNG lặp quá nhiều (tránh làm WAF nghi ngờ thêm, cũng tránh Action chạy quá lâu).
    Trả (period, {field_key: value_trieu_usd}) hoặc (None, {}) nếu lỗi/không tìm thấy nhãn quý."""
    url = "https://www.sbv.gov.vn/vi/can-can-thanh-toan-quoc-te"
    headers = {
        "User-Agent": UA,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
        "Accept-Language": "vi-VN,vi;q=0.9,en-US;q=0.8,en;q=0.7",
        "Referer": "https://www.sbv.gov.vn/",
    }
    last_err = None
    for attempt in range(3):
        try:
            r = requests.get(url, headers=headers, timeout=20, verify=False)
            r.raise_for_status()
            break
        except Exception as e:
            last_err = e
            if attempt < 2:
                time.sleep(5)
    else:
        print(f"  [WARN] SBV BOP thất bại sau 3 lần thử: {last_err}")
        return None, {}
    try:
        text = re.sub(r"<[^>]+>", "|", r.text)
        text = re.sub(r"(\|\s*)+", "|", text)

        m = re.search(r"Quý\s+([IVX]+)\s+NĂM\s+(\d{4})", text, re.IGNORECASE)
        if not m:
            print("  [WARN] SBV BOP: không tìm thấy nhãn quý trên trang")
            return None, {}
        q = _ROMAN_TO_QUARTER.get(m.group(1).upper())
        if not q:
            return None, {}
        period = f"{m.group(2)}-Q{q}"

        out = {}
        for key, label_pattern in _SBV_BOP_FIELD_MAP:
            vm = re.search(label_pattern + r"\|(-?[\d]+)\|", text)
            if vm:
                out[key] = float(vm.group(1))
        return period, out
    except Exception as e:
        print(f"  [WARN] SBV BOP thất bại: {e}")
        return None, {}


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
    (6) THEM 2026-10-09 (user hỏi về "Swap Interest Rate Curve" làm bằng chứng tại sao lãi suất
    VN khó hạ — tìm nguồn thật là VBMA FX Swap Curve nhưng bị khóa sau login hội viên, KHÔNG có
    API công khai) — bản tin VIRA CŨNG có lãi suất liên ngân hàng USD (ON/1W/2W/1M, song song VND)
    mà trước giờ chưa bắt — dùng làm PROXY cho chênh lệch lãi suất ngụ ý từ swap (không phải đúng
    đường cong swap thật, nhưng CÙNG bản chất kinh tế + có dữ liệu NGÀY thật, khác chỉ báo cũ
    vnd_usd_rate_spread_on dùng Fed Funds theo THÁNG).
    Trả list[dict] mỗi phần tử {"date": "YYYY-MM-DD", "source_url": ..., rồi các field nào tìm
    được trong: interbank_on/1w/2w/1m, interbank_usd_on/1w/2w/1m, bond_3y/5y/7y/10y/15y, omo_net,
    omo_rate, omo_outstanding} — field nào không tìm thấy (bản tin đổi cấu trúc/thiếu đoạn) thì bị
    bỏ qua, không lỗi. Bản tin nào không tồn tại (404/302, ngày nghỉ) cũng bỏ qua lặng lẽ."""
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

            # "chào " là optional — bản tin 2025 dùng "lãi suất CHÀO bình quân LNH VND", bản tin
            # 2026 bỏ chữ "chào" (phát hiện khi backfill lịch sử 2025 cho tin_phieu_outstanding_
            # balance, user 2026-08-08 — thiếu optional này khiến TOÀN BỘ bản tin 2025 bị bỏ qua
            # do không khớp được ngày, dù trang tải về bình thường).
            m_date = re.search(
                r"Ng[àa]y\s*(\d{1,2})/(\d{1,2}),\s*l[ãa]i suất (?:ch[àa]o )?b[ìi]nh qu[âa]n LNH VND", text)
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

            # THEM 2026-10-09 (user hỏi về "Swap Interest Rate Curve" — tìm được bản tin VIRA CŨNG
            # có lãi suất liên ngân hàng USD (song song VND ở trên) nhưng trước giờ CHƯA bắt — đây
            # là nguyên liệu để tính chênh lệch VND-USD CÙNG TẦN SUẤT NGÀY, CÙNG NGUỒN VIRA — khớp
            # đúng ý nghĩa "lãi suất ngụ ý từ swap" hơn chỉ báo cũ vnd_usd_rate_spread_on (dùng Fed
            # Funds — lãi suất CHÍNH SÁCH Mỹ theo THÁNG, không phải lãi suất THỊ TRƯỜNG USD thực tế
            # theo ngày). Dấu phẩy/chấm phẩy trước "1M" KHÔNG đồng nhất giữa các bản tin (vd mẫu
            # 08/10 dùng ",", bản VND ở trên dùng "và") — chấp nhận cả 3 dạng cho chắc.
            m_usd_rates = re.search(
                r"LNH USD.*?giao d[ịi]ch t[ạa]i:\s*ON\s*([\d,]+)%;\s*1W\s*([\d,]+)%;\s*2W\s*([\d,]+)%(?:,|;|\s*v[àa])\s*1M\s*([\d,]+)", text)
            if m_usd_rates:
                entry["interbank_usd_on"] = float(m_usd_rates.group(1).replace(",", "."))
                entry["interbank_usd_1w"] = float(m_usd_rates.group(2).replace(",", "."))
                entry["interbank_usd_2w"] = float(m_usd_rates.group(3).replace(",", "."))
                entry["interbank_usd_1m"] = float(m_usd_rates.group(4).replace(",", "."))

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

            # "trên"/"ở" đều gặp trong thực tế (vd bản tin 15/10/2025 dùng "lưu hành Ở kênh cầm
            # cố" thay vì "TRÊN" — quy tắc trước chỉ khớp "trên" nên bỏ sót các ngày dùng "ở").
            m_outstanding = re.search(r"C[óo]\s*([\d.,]+)\s*tỷ đồng lưu hành (?:trên|ở) kênh cầm cố", text)
            if m_outstanding:
                entry["omo_outstanding"] = _vn_number(m_outstanding.group(1))

            # Số dư TÍN PHIẾU NHNN đang lưu hành (kênh HÚT thanh khoản, đối lập omo_outstanding ở
            # trên) — user (2026-08-08) muốn biểu đồ tổng quan Tbill giống kiểu OMO. Câu CHỈ xuất
            # hiện khi có số dư tín phiếu > 0 (đã khảo sát: các ngày NHNN "không chào thầu tín
            # phiếu" — tức hiện tại và suốt từ 30/10/2025 — câu này KHÔNG xuất hiện, ngụ ý bằng 0;
            # mẫu THẬT có số liệu: bản tin 02/01/2025 "...có 50.780 tỷ đồng tín phiếu lưu hành trên
            # thị trường"). Field vắng mặt trong entry -> nơi gọi coi là 0 (không phải None/thiếu
            # dữ liệu), xem build_tin_phieu_outstanding_series()/_backfill_tin_phieu.py.
            m_tp_outstanding = re.search(
                r"c[óo]\s*([\d.,]+)\s*tỷ đồng tín phiếu lưu hành (?:trên|ở) thị trường", text)
            entry["tin_phieu_outstanding"] = _vn_number(m_tp_outstanding.group(1)) if m_tp_outstanding else 0.0

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
# nguồn tổng hợp lại của bên thứ ba, dù VietnamBiz cũng có các chỉ báo đó làm đối chiếu.
# PMI: đã chuyển hẳn sang fetch_vbma_pmi() (full lịch sử từ 2016) ngày 2026-07-23 nên từng bỏ
# khỏi map này — nhưng VBMA cập nhật TRỄ hơn VietnamBiz (user 2026-08-08 chỉ ra thiếu T7/2026;
# VBMA lúc đó mới có tới T6, trong khi VietnamBiz đã có T7=52,9 — cùng số liệu gốc S&P Global,
# chỉ khác tốc độ cập nhật). Thêm LẠI vào map để bổ sung điểm MỚI NHẤT khi VBMA chưa kịp — an
# toàn ghép chung 1 series vì fetch_vietnambiz_macro() giờ dùng _merge_point_anywhere() (tìm+ghi
# đè đúng vị trí, không nhân đôi) thay vì _append_point() cũ.
VIETNAMBIZ_TITLE_MAP = {
    "Bán lẻ HH&DV (YoY)": "retail_sales_growth",
    "Thu ngân sách (YoY)": "budget_revenue_growth",
    "Chi ngân sách (YoY)": "budget_expenditure_growth",
    "Vốn đầu tư NSNN (YoY)": "public_investment_growth",
    "Xuất khẩu (YoY)": "export_growth",
    "Nhập khẩu (YoY)": "import_growth",
    "PMI": "pmi_manufacturing",
    "Tăng trưởng CPI (YoY)": "cpi_yoy",
}


def _vietnambiz_prior_period(period):
    """Lùi period 'YYYY-MM' hoặc 'YYYY-Qn' về đúng 1 kỳ trước đó — dùng để ghép pre_value của
    VietnamBiz thành 1 điểm lịch sử THỨ HAI (xem fetch_vietnambiz_macro(): mỗi lần gọi trang trả
    cả value kỳ mới nhất LẪN pre_value kỳ liền trước, trước đây chỉ dùng value nên mỗi lần Action
    chạy chỉ tích lũy được 1 điểm — user 2026-08-08, khi nso.gov.vn (nguồn chính cho retail_sales_
    growth) tạm sập, tận dụng pre_value có sẵn để có ngay điểm thứ 2 miễn phí, không cần thêm
    request nào)."""
    m = re.match(r"(\d{4})-(\d{2})$", period)
    if m:
        y, mo = int(m.group(1)), int(m.group(2))
        y, mo = (y - 1, 12) if mo == 1 else (y, mo - 1)
        return f"{y:04d}-{mo:02d}"
    m = re.match(r"(\d{4})-Q(\d)$", period)
    if m:
        y, q = int(m.group(1)), int(m.group(2))
        y, q = (y - 1, 4) if q == 1 else (y, q - 1)
        return f"{y:04d}-Q{q}"
    return None


def fetch_vietnambiz_macro():
    """data.vietnambiz.vn/macro-economic nhúng __NEXT_DATA__ JSON (Next.js server-rendered,
    KHÔNG phải SPA rỗng) chứa ~25 chỉ báo vĩ mô, mỗi chỉ báo có value (kỳ mới nhất) + pre_value
    (kỳ trước) + nhãn kỳ tiếng Việt ('Tháng 06/2026'/'Quý 2/2026'/'Năm 2023'). Trả list
    [(indicator_key, period_iso, value), ...] (2 điểm/chỉ báo nếu pre_value hợp lệ — value kỳ mới
    nhất + pre_value kỳ liền trước, xem _vietnambiz_prior_period) cho các chỉ báo có trong
    VIETNAMBIZ_TITLE_MAP."""
    url = "https://data.vietnambiz.vn/macro-economic"
    try:
        r = requests.get(url, headers={"User-Agent": UA}, timeout=20)
        r.raise_for_status()
        m = re.search(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>', r.text, re.S)
        if not m:
            print("  [WARN] VietnamBiz: không tìm thấy __NEXT_DATA__ — trang có thể đã đổi cấu trúc.")
            return []
        data = json.loads(m.group(1))
        items = data.get("props", {}).get("pageProps", {}).get("data", [])
        out = []
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
            # pre_value luôn là kỳ LIỀN TRƯỚC value, không đảm bảo là điểm mới nhất trong series
            # hiện có (nếu series đang có khoảng trống) — nơi gọi PHẢI dùng _merge_point_anywhere()
            # (không phải _append_point() thường) để tránh tạo bản trùng lệch thứ tự thời gian.
            if it.get("pre_value") is not None:
                prior_period = _vietnambiz_prior_period(period)
                if prior_period:
                    out.append((key, prior_period, round(float(it["pre_value"]), 2)))
            out.append((key, period, round(float(it["value"]), 2)))
        return out
    except Exception as e:
        print(f"  [WARN] VietnamBiz macro thất bại: {e}")
        return []


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


def fetch_cake_max_rate():
    """cake.vn/tien-gui/tien-gui-tieu-chuan — lãi suất tiết kiệm kỳ hạn 12 tháng (Cuối kỳ hạn) CỘNG
    THÊM % ưu đãi khuyến mãi cao nhất đang chạy, của Cake (ngân hàng số thuộc VCCB) — user (2026-09-
    28) đánh giá đây là kênh tham chiếu TỐT HƠN iPower TCBS cho áp lực huy động thị trường: "họ công
    bố có 7.4% nhưng thêm ưu đãi 2.2%" — cùng logic 'kênh gửi tiền thay thế' như iPower: khi hệ
    thống ngân hàng căng huy động, Cake phải tăng lãi suất/khuyến mãi để cạnh tranh hút tiền, mức
    này HẠ xuống nghĩa là áp lực huy động đã dịu bớt (xem fetch_tcbs_ipower_max_rate() ở trên).

    Trang là Next.js SSR THẬT (không cần JS) nhưng bảng so sánh CHÍNH chỉ render sẵn 4 dòng đầu (kỳ
    hạn 1-4 tháng, các kỳ hạn dài hơn ẩn sau nút "Xem tất cả" — không có trong HTML tĩnh); riêng
    WIDGET "Chọn kỳ hạn" của công cụ tính lãi lại có SẴN toàn bộ 15 kỳ hạn (kể cả 12 tháng) ngay
    trong HTML tĩnh — dùng nguồn này thay vì bảng chính. Khuyến mãi ("Thêm ưu đãi") áp dụng cho kỳ
    hạn 6-13 tháng nên CÓ bao gồm kỳ hạn 12 tháng đang lấy — lấy mức ưu đãi CAO NHẤT đang hiển thị
    (có thể có nhiều mức cho khách gửi lần đầu/lần tiếp theo). Trả (rate, period_iso, source_url)
    hoặc (None, None, None) nếu lỗi/không tìm thấy — KHÔNG đoán mù khi cấu trúc trang đã đổi."""
    url = "https://cake.vn/tien-gui/tien-gui-tieu-chuan"
    try:
        resp = requests.get(url, timeout=20, headers={"User-Agent": UA})
        if resp.status_code != 200:
            print(f"  [WARN] Cake: HTTP {resp.status_code}.")
            return None, None, None
        text = re.sub(r"<[^>]+>", " ", resp.text)
        text = re.sub(r"&nbsp;", " ", text)
        text = re.sub(r"\s+", " ", text)

        m_base = re.search(r"Chọn kỳ hạn.*?\b12\s*Tháng\s*([\d,\.]+)\s*%", text)
        if not m_base:
            print("  [WARN] Cake: không tìm thấy lãi suất kỳ hạn 12 tháng trong widget tính lãi — trang có thể đã đổi cấu trúc.")
            return None, None, None
        base_rate = float(m_base.group(1).replace(",", "."))

        promo_max = 0.0
        promo_m = re.search(r"Thêm ưu đãi(.*?)Ước tính số tiền", text)
        if promo_m:
            for pm in re.finditer(r"Tặng(?: đến)?\s*([\d,\.]+)\s*%", promo_m.group(1)):
                promo_max = max(promo_max, float(pm.group(1).replace(",", ".")))

        rate = base_rate + promo_max
        period = _current_period_weekly()
        m2 = re.search(r"Áp dụng chính thức từ \d{1,2}h\d{2} ngày (\d{2})/(\d{2})/(\d{4})", text)
        if m2:
            eff_date = datetime.date(int(m2.group(3)), int(m2.group(2)), int(m2.group(1)))
            period = _current_period_weekly(eff_date)
        return rate, period, url
    except Exception as e:
        print(f"  [WARN] Cake thất bại: {e}")
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

    # THEM (user 2026-10-01): "gap tỷ giá chợ đen và tỷ giá VCB" (chỉ giá bán, dễ so sánh) — đầu
    # cơ/nhu cầu USD tăng nếu gap nới rộng. VCB: backfill NHIỀU NGÀY gần đây (API hỗ trợ tham số
    # date=YYYY-MM-DD, test ngược được tới cả 2023) — lần đầu chạy sẽ lấp đầy lịch sử, các lần
    # sau chỉ thêm 1-2 ngày mới (merge theo period nên không trùng). chợ đen (chogia.vn) chỉ trả
    # ~28 ngày gần nhất mỗi lần gọi — gọi 1 lần, merge toàn bộ, tự tích luỹ dần qua các lần chạy.
    print("[Vietcombank — tỷ giá USD bán ra (backfill ~45 ngày gần nhất)]")
    if "usdvnd_vcb_sell_daily" not in raw:
        raw["usdvnd_vcb_sell_daily"] = {
            "group": "external", "label": "Tỷ giá USD/VND — Vietcombank (bán ra)", "unit": "VND",
            "good_direction": "lower", "auto_source": "vcb",
            "note": "vietcombank.com.vn/api/exchangerates (API JSON thật, theo ngày) — field \"sell\" (bán ra cho khách hàng, KHÁC \"transfer\"/\"cash\").",
            "impact": "Tỷ giá bán ra chính thức của NHTM lớn nhất — dùng làm mốc đối chiếu với tỷ giá chợ đen để tính gap.",
            "series": [],
        }
    # Chỉ GỌI API cho ngày CHƯA CÓ sẵn (tránh 45 request thừa mỗi lần chạy sau khi đã backfill đủ —
    # lần đầu chạy sẽ gọi đủ 45 lần để lấp lịch sử, các lần sau chỉ gọi cho (các) ngày mới).
    existing_vcb_dates = {p["period"] for p in raw["usdvnd_vcb_sell_daily"]["series"]}
    n_vcb = 0
    for i in range(45):
        d = (datetime.date.today() - datetime.timedelta(days=i)).strftime("%Y-%m-%d")
        if d in existing_vcb_dates:
            continue
        actual_date, sell = fetch_vcb_usd_sell_rate(d)
        if actual_date and sell:
            _merge_point_anywhere(raw, "usdvnd_vcb_sell_daily", actual_date, sell,
                                   "https://www.vietcombank.com.vn/vi-VN/Vietnam-Dong-Exchange-Rate")
            n_vcb += 1
    print(f"  -> {n_vcb} ngày mới (tổng {len(raw['usdvnd_vcb_sell_daily']['series'])} điểm)")

    print("[chogia.vn — tỷ giá USD chợ đen (bán ra)]")
    if "usd_cho_den_sell_daily" not in raw:
        raw["usd_cho_den_sell_daily"] = {
            "group": "external", "label": "Tỷ giá USD chợ đen (bán ra)", "unit": "VND",
            "good_direction": "lower", "auto_source": "derived",
            "note": "chogia.vn/ngoai-te/usd-cho-den/ (endpoint AJAX nội bộ, ~28 ngày gần nhất mỗi lần gọi — tích lũy dần qua các lần chạy).",
            "impact": "Tỷ giá USD ngoài hệ thống ngân hàng — chênh lệch lớn với tỷ giá NHTM là dấu hiệu đầu cơ/nhu cầu USD ngoài luồng chính thức tăng.",
            "series": [],
        }
    cho_den_pts = fetch_chogia_usd_cho_den()
    for ngay, gia_ban in cho_den_pts:
        _merge_point_anywhere(raw, "usd_cho_den_sell_daily", ngay, gia_ban, "https://chogia.vn/ngoai-te/usd-cho-den/")
    print(f"  -> {len(cho_den_pts)} điểm từ nguồn (tổng {len(raw['usd_cho_den_sell_daily']['series'])} điểm)")

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

    # THEM 2026-10-07 (user: tab "Kinh tế Mỹ" theo mô-tuýp phân tích lạm phát: CPI cắt lớp, PPI →
    # CPI, giá nhập khẩu, bán lẻ/tiêu dùng thực). Dùng FRED CSV công khai (fredgraph.csv, KHÔNG cần
    # API key nên chạy được cả local lẫn CI), lưu từ 2015 tới nay theo tháng, giữ cả giá trị gốc.
    print("[FRED CSV — Kinh tế Mỹ chuyên sâu (CPI cắt lớp, PPI, giá nhập khẩu, bán lẻ, PCE, lãi suất)]")
    for key, sid, label, unit in US_MACRO_SERIES:
        pts = fetch_fred_csv(sid, start="2015-01-01")
        if not pts:
            print(f"  [WARN] FRED CSV {sid}: không tải được.")
            continue
        monthly = {}
        for d, v in pts:
            monthly.setdefault(d[:7], []).append(v)
        raw[key] = {
            "group": "us_macro", "label": label, "unit": unit, "auto_source": "fred_api",
            "good_direction": "none",
            "series": [{"period": m, "value": round(sum(vs) / len(vs), 4),
                        "source_url": f"https://fred.stlouisfed.org/series/{sid}"}
                       for m, vs in sorted(monthly.items())],
        }
        print(f"  -> {key} ({sid}): {len(pts)} điểm")

    print("[Treasury TIC — Nước ngoài nắm giữ Treasury Mỹ (Major Foreign Holders, KHÔNG có trên FRED)]")
    tic_total, tic_official = fetch_tic_major_foreign_holders()
    tic_url = "https://ticdata.treasury.gov/resource-center/data-chart-center/tic/Documents/mfhhis01.csv"
    if tic_total:
        raw["usm_tic_total"] = {
            "group": "us_macro", "label": "Nước ngoài nắm giữ Treasury Mỹ — Tổng (tỷ USD)", "unit": "billion_usd",
            "auto_source": "treasury_tic", "good_direction": "none",
            "series": [{"period": p, "value": v, "source_url": tic_url} for p, v in sorted(tic_total.items())],
        }
        print(f"  -> usm_tic_total: {len(tic_total)} điểm")
    if tic_official:
        raw["usm_tic_official"] = {
            "group": "us_macro", "label": "Nước ngoài nắm giữ Treasury Mỹ — Khối chính thức/NHTW (tỷ USD)", "unit": "billion_usd",
            "auto_source": "treasury_tic", "good_direction": "none",
            "series": [{"period": p, "value": v, "source_url": tic_url} for p, v in sorted(tic_official.items())],
        }
        print(f"  -> usm_tic_official: {len(tic_official)} điểm")

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
    # XNK tách khu vực trong nước/FDI, lũy kế theo THÁNG (user 2026-10-01, muốn "FDI Trade
    # Balance" làm chỉ báo TRADE STRUCTURE — KHÔNG dùng suy diễn USD supply/demand thực, xem
    # ghi chú tại chỗ trích trong fetch_nso_gdp_structure_report()). 4 series + 2 gap phái sinh
    # (FDI/Domestic Trade Balance) tính NGAY ở đây (đơn giản, không cần hàm _add_* riêng).
    if gdp_struct.get("trade_fdi_split_period"):
        p = gdp_struct["trade_fdi_split_period"]
        src = gdp_struct["source_url"]
        _TRADE_SPLIT_META = {
            "export_domestic_usd_bn": "Xuất khẩu — khu vực kinh tế trong nước (lũy kế, NSO)",
            "export_fdi_usd_bn": "Xuất khẩu — khu vực FDI (lũy kế, NSO)",
            "import_domestic_usd_bn": "Nhập khẩu — khu vực kinh tế trong nước (lũy kế, NSO)",
            "import_fdi_usd_bn": "Nhập khẩu — khu vực FDI (lũy kế, NSO)",
        }
        for key, label in _TRADE_SPLIT_META.items():
            if key not in gdp_struct:
                continue
            if key not in raw:
                raw[key] = {
                    "group": "trade", "label": label, "unit": "tỷ USD", "good_direction": "higher",
                    "auto_source": "nso", "series": [],
                    "note": ("Trích từ báo cáo tháng NSO (nso.gov.vn/bao-cao-tinh-hinh-kinh-te-xa-hoi-hang-thang/), "
                             "GIÁ TRỊ LŨY KẾ từ đầu năm (reset mỗi tháng 1), KHÔNG PHẢI số riêng 1 tháng."),
                    "impact": "Thuộc nhóm 'trade structure' (cơ cấu XNK theo khu vực DN) — KHÔNG dùng để suy ra cung/cầu USD thực tế qua hệ thống ngân hàng (DN FDI có thể dùng vốn/giữ doanh thu ở nước ngoài, xem FDI_Trade_Balance/Domestic_Trade_Balance).",
                }
            _append_point(raw, key, p, gdp_struct[key], src)
        if "export_domestic_usd_bn" in gdp_struct and "export_fdi_usd_bn" in gdp_struct:
            if "fdi_trade_balance" not in raw:
                raw["fdi_trade_balance"] = {
                    "group": "trade", "label": "FDI Trade Balance (Xuất khẩu FDI − Nhập khẩu FDI, lũy kế)",
                    "unit": "tỷ USD", "good_direction": "higher", "auto_source": "derived", "series": [],
                    "note": "= export_fdi_usd_bn − import_fdi_usd_bn (NSO, lũy kế từ đầu năm). Phái sinh tính toán, KHÔNG lưu vào vimo_raw.json.",
                    "impact": "Khu vực FDI xuất siêu lớn KHÔNG đồng nghĩa khu vực này cung USD thực tế vào hệ thống ngân hàng VN — DN FDI có thể tài trợ nhập khẩu bằng vốn công ty mẹ hoặc giữ doanh thu xuất khẩu ở tài khoản offshore. Chỉ nên đọc đây là cấu trúc thương mại, không phải dòng ngoại tệ thực.",
                }
            if "import_domestic_usd_bn" in gdp_struct and "import_fdi_usd_bn" in gdp_struct:
                _append_point(raw, "fdi_trade_balance", p,
                               round(gdp_struct["export_fdi_usd_bn"] - gdp_struct["import_fdi_usd_bn"], 2), src)
                if "domestic_trade_balance" not in raw:
                    raw["domestic_trade_balance"] = {
                        "group": "trade", "label": "Domestic Trade Balance (Xuất khẩu − Nhập khẩu khu vực trong nước, lũy kế)",
                        "unit": "tỷ USD", "good_direction": "higher", "auto_source": "derived", "series": [],
                        "note": "= export_domestic_usd_bn − import_domestic_usd_bn (NSO, lũy kế từ đầu năm). Phái sinh tính toán, KHÔNG lưu vào vimo_raw.json.",
                        "impact": "Khu vực kinh tế trong nước thường nhập siêu — phản ánh đúng hơn áp lực cầu USD THỰC của doanh nghiệp Việt Nam (ít khả năng có nguồn vốn/doanh thu offshore như khu vực FDI).",
                    }
                _append_point(raw, "domestic_trade_balance", p,
                               round(gdp_struct["export_domestic_usd_bn"] - gdp_struct["import_domestic_usd_bn"], 2), src)
        print(f"  -> XNK theo khu vực {p}: Xuất khẩu trong nước={gdp_struct.get('export_domestic_usd_bn')}/FDI={gdp_struct.get('export_fdi_usd_bn')} tỷ USD, "
              f"Nhập khẩu trong nước={gdp_struct.get('import_domestic_usd_bn')}/FDI={gdp_struct.get('import_fdi_usd_bn')} tỷ USD")
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

    # GDP theo CẤU PHẦN SỬ DỤNG (tiêu dùng tư nhân/chính phủ, đầu tư) — user (2026-08-08): "GDP
    # tăng nhờ cái gì? Tăng trưởng đến từ tiêu dùng nội địa, đầu tư, xuất khẩu hay chính phủ?" —
    # khác gdp_share_agri/industry/services (cơ cấu theo KHU VỰC KINH TẾ, NSO, theo quý) — đây là
    # cơ cấu theo LOẠI CHI TIÊU, chỉ có ở tần suất NĂM qua ARIC/ADB (đã khảo sát thủ công, xem
    # fetch_aric_gdp_by_use()). Export/import growth ĐÃ có sẵn từ nguồn khác (Hải quan/VietnamBiz),
    # không lấy lại ở đây.
    print("[ARIC/ADB — GDP theo cấu phần sử dụng: đầu tư, tiêu dùng tư nhân/chính phủ (theo NĂM)]")
    aric = fetch_aric_gdp_by_use()
    _ARIC_KEY_MAP = {
        "Gross Domestic Investment Growth (y-o-y, %)": (
            "gdp_investment_growth", "Tăng trưởng đầu tư toàn xã hội (GDP theo cấu phần sử dụng)"),
        "Private Consumption Expenditure Growth (y-o-y, %)": (
            "private_consumption_growth", "Tăng trưởng tiêu dùng tư nhân (GDP theo cấu phần sử dụng)"),
        "Public Consumption Expenditure Growth (y-o-y, %)": (
            "public_consumption_growth", "Tăng trưởng tiêu dùng chính phủ (GDP theo cấu phần sử dụng)"),
        # ĐÓNG GÓP (điểm %) vào tăng trưởng GDP chung của TỪNG cấu phần sử dụng — user (2026-08-09)
        # muốn "biểu đồ tỷ trọng GDP theo cấu phần sử dụng" nhưng ARIC (và cả báo cáo NSO gốc)
        # KHÔNG có "% GDP tuyệt đối" theo cấu phần sử dụng (chỉ NSO có %GDP theo KHU VỰC KINH TẾ —
        # gdp_share_agri/industry/services — khác hẳn) — đây là dữ liệu THẬT gần nhất trả lời đúng
        # câu hỏi "GDP tăng nhờ cái gì": 3 cột này CỘNG LẠI ≈ GDP Growth (y-o-y) cùng năm (đã kiểm
        # tra: 2025 đầu tư 2,894 + tiêu dùng tư 3,992 + tiêu dùng công 1,056 = 7,942 ≈ GDP 8,019,
        # phần chênh nhỏ là xuất khẩu ròng/thay đổi tồn kho ARIC không tách riêng) — vẽ stacked
        # area đúng bản chất "phần trăm ĐIỂM đóng góp vào tăng trưởng", KHÔNG PHẢI %GDP level.
        "Contribution to Growth: Gross Domestic Investment": (
            "gdp_use_contrib_investment", "Đóng góp của Đầu tư vào tăng trưởng GDP (điểm %)"),
        "Contribution to Growth: Private Consumption Expenditure": (
            "gdp_use_contrib_private_consumption", "Đóng góp của Tiêu dùng tư nhân vào tăng trưởng GDP (điểm %)"),
        "Contribution to Growth: Public Consumption Expenditure": (
            "gdp_use_contrib_public_consumption", "Đóng góp của Tiêu dùng chính phủ vào tăng trưởng GDP (điểm %)"),
    }
    for aric_name, (key, label) in _ARIC_KEY_MAP.items():
        points = aric.get(aric_name, [])
        if not points:
            continue
        is_contrib = key.startswith("gdp_use_contrib_")
        if key not in raw:
            raw[key] = {
                "group": "growth", "label": label, "unit": "điểm %" if is_contrib else "%",
                "good_direction": "higher", "auto_source": "manual",
                "note": ("Nguồn aric.adb.org (Asian Regional Integration Center, ADB — CEIC "
                         "database), sector 'Real Sector and Prices', tần suất NĂM (trang không "
                         "có API/CSV, chỉ có bảng HTML tần suất năm cho bộ chỉ số này). Trễ ~1 "
                         "năm so với hiện tại (số liệu năm N thường công bố đầy đủ vào năm N+1)."
                         + (" ĐÂY LÀ SỐ ĐIỂM % ĐÓNG GÓP VÀO TĂNG TRƯỞNG GDP, KHÔNG PHẢI %GDP "
                            "tuyệt đối (Việt Nam không công bố %GDP theo cấu phần sử dụng) — 3 "
                            "cấu phần (đầu tư/tiêu dùng tư/tiêu dùng công) cộng lại xấp xỉ bằng "
                            "GDP Growth cùng năm, phần chênh nhỏ là xuất khẩu ròng/tồn kho." if is_contrib else "")),
                "impact": ("Cấu phần nào đóng góp nhiều điểm % nhất vào tăng trưởng GDP năm đó — đầu tư/chi tiêu chính phủ đóng góp áp đảo so với tiêu dùng tư nhân là dấu hiệu tăng trưởng dựa vào vốn/đòn bẩy hơn là cầu tiêu dùng nội địa bền vững."
                           if is_contrib else
                           "Tách GDP theo LOẠI CHI TIÊU (khác cơ cấu theo khu vực kinh tế gdp_share_*) — tăng trưởng đến từ tiêu dùng nội địa bền vững hơn tăng trưởng dựa vào đầu tư/chi tiêu chính phủ đơn thuần."),
                "series": [],
            }
        # MERGE (không ghi đè toàn bộ series) — "gdp_investment_growth" đã tồn tại SẴN TRƯỚC đây
        # với 1 điểm THỦ CÔNG "2026-H1" trích từ báo cáo NSO (mới hơn, đáng tin hơn ARIC vốn chỉ
        # có tới năm 2025) — ghi đè thẳng `series = [...]` sẽ XÓA MẤT điểm đó (đã xảy ra thật 1
        # lần, phát hiện qua git diff khi kiểm tra UI, xem lịch sử commit 2026-08-08). Period ARIC
        # luôn dạng "YYYY" thuần (khác "YYYY-H1"/"YYYY-Q1" của các điểm thủ công) nên không bao
        # giờ trùng — an toàn dùng merge-theo-period, chỉ thêm năm CHƯA có, giữ nguyên mọi điểm
        # thủ công/nguồn khác đã có sẵn.
        existing_periods = {p["period"] for p in raw[key]["series"]}
        added = 0
        for year, v in sorted(points):
            period = f"{year:04d}"
            if period not in existing_periods:
                raw[key]["series"].append({
                    "period": period, "value": v,
                    "source_url": "https://aric.adb.org/database/economic-financial-indicators"})
                existing_periods.add(period)
                added += 1
        raw[key]["series"].sort(key=lambda p: p["period"])
        print(f"  -> {label}: +{added} điểm mới từ ARIC (tổng {len(raw[key]['series'])} điểm)")

    # 40yo.vn — kho dữ liệu tổng hợp IMF/World Bank/ADB/UNCTAD rất lớn (~80 chỉ báo VN), user
    # (2026-08-09) yêu cầu "bổ sung tất cả" nhưng CHỈ giữ chuỗi CÒN CẬP NHẬT (_40yo_is_fresh) và
    # CHỈ vẽ biểu đồ từ 2016 trở lại đây (lọc ngay tại bước ghi vào raw, không lưu phần lịch sử xa
    # hơn — trừ các chuỗi có DỰ BÁO tương lai vẫn giữ nguyên vì đó là thông tin hữu ích, không phải
    # "dữ liệu cũ"). Xem _40YO_SPEC ở trên cho danh sách đầy đủ + lý do loại từng chuỗi đã cũ.
    print("[40yo.vn — kho dữ liệu tổng hợp IMF/World Bank/ADB/UNCTAD (~70 chỉ báo, đã lọc chuỗi còn cập nhật)]")
    data_40yo = fetch_40yo_vnm_data()
    charts_40yo = data_40yo.get("charts", {})
    n_kept, n_skipped_stale = 0, 0
    for chart_key, series_name, raw_key, label, unit, good_dir, group, freshness in _40YO_SPEC:
        chart = charts_40yo.get(chart_key)
        if not chart:
            continue
        series_list = chart.get("series") or ((chart.get("left") or []) + (chart.get("right") or []))
        series = next((s for s in series_list if s.get("name") == series_name), None)
        if not series or not series.get("data"):
            continue
        points = []
        for pt in series["data"]:
            iso = _40yo_period_to_iso(pt.get("p", ""))
            if iso is None or pt.get("v") is None:
                continue
            points.append((iso, pt["v"]))
        iso_periods = [p for p, _ in points]
        if not _40yo_is_fresh(iso_periods, freshness):
            n_skipped_stale += 1
            continue
        # Chỉ giữ TỪ 2016 trở đi (user: "những dữ liệu lâu quá rồi thì chỉ vẽ biểu đồ cho từ 2016
        # trở lại đây") — áp dụng ngay khi ghi vào raw (không lưu phần xa hơn) vì các chỉ báo này
        # CHỈ dùng để hiển thị tham khảo, không cần backtest xa hơn 2016.
        points = [(p, v) for p, v in points if p[:4] >= "2016"]
        if not points:
            continue
        if raw_key not in raw:
            raw[raw_key] = {
                "group": group, "label": label, "unit": unit, "good_direction": good_dir,
                "auto_source": "40yo",
                "note": ("Nguồn 40yo.vn (tổng hợp IMF WEO/World Bank/ADB/UNCTAD), endpoint JSON "
                         "công khai 40yo.vn/data/countries/VNM.json — chỉ hiển thị từ 2016 (dữ "
                         "liệu gốc có thể sâu hơn nhiều, xem 40yo.vn/vi-mo/vn nếu cần lịch sử xa "
                         "hơn). Các năm sau hiện tại (nếu có) là DỰ BÁO của IMF/ADB, không phải "
                         "số liệu thực tế đã công bố."),
                "impact": "Chỉ báo tham khảo bổ sung, chưa dùng cho Scorecard — mở rộng phạm vi theo dõi sang các mảng hệ thống chưa có nguồn tự động trước đây (đường cong lợi suất, cơ cấu nợ TPCP/TPDN, cán cân thanh toán chi tiết, dân số, ngân sách dài hạn).",
                "series": [],
            }
        existing_periods = {p["period"] for p in raw[raw_key]["series"]}
        for period, value in points:
            if period not in existing_periods:
                raw[raw_key]["series"].append({"period": period, "value": value,
                                                 "source_url": "https://40yo.vn/vi-mo/vn"})
                existing_periods.add(period)
        raw[raw_key]["series"].sort(key=lambda p: p["period"])
        n_kept += 1
    print(f"  -> Giữ {n_kept} chuỗi (từ 2016), bỏ {n_skipped_stale} chuỗi đã ngừng cập nhật")

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

    # VBMA PMI PHẢI chạy TRƯỚC VietnamBiz — VBMA ghi ĐÈ TOÀN BỘ series (full lịch sử từ 2016,
    # không phải merge từng điểm), nên nếu chạy SAU sẽ XÓA MẤT điểm PMI tháng mới nhất mà
    # VietnamBiz vừa thêm (VietnamBiz cập nhật NHANH hơn VBMA — user 2026-08-08 chỉ ra VBMA trễ
    # 1-2 tháng — xem PMI trong VIETNAMBIZ_TITLE_MAP). Đã xác nhận lỗi này thật qua 1 lần chạy sai
    # thứ tự: log báo "pmi_manufacturing 2026-07: 52.9" nhưng vimo_raw.json cuối cùng KHÔNG có
    # điểm đó vì bước VBMA chạy sau đã ghi đè mất.
    print("[VBMA — PMI sản xuất theo tháng (toàn bộ lịch sử từ T1/2016)]")
    pts = fetch_vbma_pmi()
    if pts:
        raw["pmi_manufacturing"]["series"] = [
            {"period": p, "value": v,
             "source_url": "https://vbma.org.vn/vi/market-data/gdp-growth"}
            for p, v in pts
        ]
        print(f"  -> {len(pts)} điểm")

    print("[VietnamBiz — Bán lẻ (đối chiếu, tích lũy theo lần chạy, nay lấy cả value + pre_value)]")
    vnb = fetch_vietnambiz_macro()
    for key, period, value in vnb:
        _merge_point_anywhere(raw, key, period, value, "https://data.vietnambiz.vn/macro-economic")
        print(f"  -> {key} {period}: {value}")

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
    # THEM 2026-10-09: lãi suất liên ngân hàng USD (ON/1W/2W/1M, cùng bản tin VIRA, CÙNG ngày với
    # VND) — pre-init như tin_phieu_outstanding_balance ở trên vì key MỚI, chưa từng có trong
    # vimo_raw.json (vòng lặp bên dưới giả định raw[raw_key]["series"] đã tồn tại).
    for _k, _lbl in [("interbank_usd_on", "ON"), ("interbank_usd_1w", "1 tuần"),
                      ("interbank_usd_2w", "2 tuần"), ("interbank_usd_1m", "1 tháng")]:
        if _k not in raw:
            raw[_k] = {
                "group": "external", "label": f"Lãi suất liên ngân hàng USD kỳ hạn {_lbl} (VIRA)",
                "unit": "%/năm", "good_direction": "none", "auto_source": "vira",
                "note": ("vira.org.vn, bản tin ngày — lãi suất bình quân liên ngân hàng USD tại "
                         "Việt Nam (KHÁC Fed Funds — đây là lãi suất THỊ TRƯỜNG thực tế, không "
                         "phải lãi suất CHÍNH SÁCH). Dùng cùng interbank_rate_on/1w/2w/1m (VND) để "
                         "tính chênh lệch VND-USD theo NGÀY — proxy cho áp lực lãi suất ngụ ý từ "
                         "swap USD/VND (đường cong swap thật bị khóa sau login hội viên VBMA, "
                         "KHÔNG có nguồn công khai — xem user 2026-10-09)."),
                "series": [],
            }
    # sort tăng dần theo ngày TRƯỚC khi append — fetch_vira_bulletin() quét lùi (mới nhất trước),
    # trong khi _append_point() chỉ nối vào CUỐI series (không tự sort như _merge_vbma_points).
    vira_entries = sorted(fetch_vira_bulletin(), key=lambda e: e["date"])
    VIRA_KEY_MAP = {
        "interbank_on": "interbank_rate_on", "interbank_1w": "interbank_rate_1w",
        "interbank_2w": "interbank_rate_2w", "interbank_1m": "interbank_rate_1m",
        "interbank_usd_on": "interbank_usd_on", "interbank_usd_1w": "interbank_usd_1w",
        "interbank_usd_2w": "interbank_usd_2w", "interbank_usd_1m": "interbank_usd_1m",
        "bond_3y": "govt_bond_yield_3y", "bond_5y": "govt_bond_yield_5y",
        "bond_7y": "govt_bond_yield_7y", "bond_10y": "govt_bond_yield_10y",
        "bond_15y": "govt_bond_yield_15y", "omo_net": "omo_net_operation",
        "omo_rate": "omo_rate_7d", "omo_outstanding": "omo_outstanding_balance",
        "tin_phieu_outstanding": "tin_phieu_outstanding_balance",
    }
    if "tin_phieu_outstanding_balance" not in raw:
        raw["tin_phieu_outstanding_balance"] = {
            "group": "monetary", "label": "Số dư tín phiếu NHNN đang lưu hành", "unit": "tỷ đồng",
            "good_direction": "lower", "auto_source": "vira",
            "note": ("vira.org.vn (bản tin Kinh tế - Tài chính ngày) — câu 'có X tỷ đồng tín "
                     "phiếu lưu hành trên thị trường' CHỈ xuất hiện khi số dư > 0, nên ngày nào "
                     "không thấy câu này được ghi 0 (không phải thiếu dữ liệu). NHNN không chào "
                     "thầu tín phiếu nào kể từ 30/10/2025 (xem tin_phieu_days_since_issuance) nên "
                     "chuỗi này bằng 0 xuyên suốt từ mốc đó tới nay — user (2026-08-08) muốn thấy "
                     "trực quan kênh hút thanh khoản qua tín phiếu đang 'tắt', đối chiếu với "
                     "omo_outstanding_balance (kênh bơm, vẫn hoạt động) trong cùng 1 biểu đồ."),
            "impact": "Tín phiếu là kênh HÚT thanh khoản đối lập OMO (kênh BƠM) — NHNN quay lại phát hành tín phiếu là dấu hiệu SỚM cho thấy đang chủ động rút bớt thanh khoản dư thừa, thường đi kèm giai đoạn tỷ giá/lạm phát chịu áp lực.",
            "series": [],
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

    print("[Cake — lãi suất 12 tháng + ưu đãi cao nhất (kênh gửi tiền thay thế, tích lũy theo TUẦN)]")
    cake_rate, cake_period, cake_src = fetch_cake_max_rate()
    if cake_rate is not None:
        # Chỉ báo MỚI (2026-09-28, user đề xuất) — tự khởi tạo container nếu vimo_raw.json chưa có
        # key này (giống cách _ARIC_KEY_MAP ở trên tự tạo raw[key] lần đầu), không cần seed riêng.
        if "deposit_rate_cake_max" not in raw:
            raw["deposit_rate_cake_max"] = {
                "group": "monetary", "label": "Lãi suất Cake 12 tháng + ưu đãi (cao nhất)",
                "unit": "%", "good_direction": "lower", "auto_source": "cake_scrape",
                "note": ("Chỉ báo mới 2026-09-28 (user đề xuất, đánh giá TỐT HƠN iPower TCBS) — "
                         "cake.vn/tien-gui/tien-gui-tieu-chuan, ngân hàng số Cake (VCCB): lãi suất "
                         "niêm yết kỳ hạn 12 tháng (Cuối kỳ hạn) CỘNG mức ưu đãi khuyến mãi cao nhất "
                         "đang chạy (mục 'Thêm ưu đãi', áp dụng kỳ hạn 6-13 tháng nên bao gồm cả kỳ "
                         "hạn 12 tháng đang lấy) — vd niêm yết 7,4%/năm + ưu đãi 2,2% = 9,6%/năm. "
                         "Cùng bản chất 'kênh gửi tiền thay thế' như iPower TCBS: khi hệ thống ngân "
                         "hàng căng huy động, Cake phải tăng lãi suất/khuyến mãi để cạnh tranh hút "
                         "tiền — mức này HẠ xuống nghĩa là áp lực huy động đã dịu bớt. Period dùng "
                         "NGÀY ÁP DỤNG CHÍNH THỨC ghi trên trang (không phải ngày fetch) — xem "
                         "fetch_cake_max_rate()."),
                "impact": ("Cake hạ lãi suất 12 tháng + ưu đãi đồng nghĩa áp lực huy động/thanh khoản "
                           "bên ngoài hệ thống ngân hàng truyền thống đã dịu bớt. Ngược lại, mức này "
                           "tăng lên (niêm yết tăng hoặc khuyến mãi tăng) là tín hiệu SỚM cho thấy "
                           "căng thẳng huy động đang lan rộng."),
                "series": [],
            }
        # THEM 2026-10-05 (user: "số liệu Cake sẽ cập nhật hàng tuần") — ghi theo NGÀY CHẠY (không theo
        # ngày áp dụng trên trang, vì trang không đổi nên sẽ không có điểm mới).
        cake_day = datetime.date.today().isoformat()
        _append_point(raw, "deposit_rate_cake_max", cake_day, cake_rate, cake_src)
        print(f"  -> {cake_day}: {cake_rate}% (áp dụng từ {cake_period})")

    print("[RSS tin tức — lãi suất huy động THỎA THUẬN (quét CafeF/VietStock, chỉ ghi khi có tin mới khớp)]")
    hit = fetch_negotiated_deposit_rate_news()
    if hit:
        period, value, link, title = hit
        _append_point(raw, "deposit_rate_negotiated_max", period, value, link)
        print(f"  -> {period}: {value}% — \"{title}\"")
    else:
        print("  -> Không có tin mới khớp từ khóa (bình thường, đây là tin hiếm)")

    # THEM (user 2026-10-01, phát hiện NHNN công bố BOP quý đầy đủ CHUẨN BPM6 — "không nhất thiết
    # phải chờ scraper IMF cho E&O/Overall Balance"): lấp 3 ô ⏳ quan trọng nhất của card "Áp lực
    # Ngoại tệ" (Errors & Omissions, Overall Balance, Δ Dự trữ), cùng lúc lấy thêm ~13 dòng khác
    # (Dịch vụ/Thu nhập đầu tư/Chuyển giao vãng lai/FDI/Đầu tư gián tiếp/Vay-trả nợ nước ngoài)
    # phục vụ lớp Cầu/Cung.
    # SỬA (user 2026-10-01, "các kì mới sẽ chủ động check xem nếu thiếu thì lấy ở sbv nhé, để đỡ
    # bị chặn"): dulieukinhte.com/du-lieu/can-can-thanh-toan-361 MIRROR lại đúng số liệu NHNN, HTML
    # tĩnh (không WAF) — dùng làm nguồn CHÍNH, backfill luôn được ~5 quý gần nhất mỗi lần chạy. Chỉ
    # gọi trực tiếp sbv.gov.vn (rủi ro bị chặn) làm FALLBACK khi dulieukinhte CHƯA có quý mới hơn
    # quý đã lưu trước đó (nghĩa là NHNN có thể vừa công bố quý mới mà dulieukinhte chưa kịp cập
    # nhật) — xem fetch_dulieukinhte_bop() / fetch_sbv_bop_quarterly().
    print("[Cán cân thanh toán quốc tế (BOP) theo quý — nguồn chính dulieukinhte.com, fallback SBV]")

    # (label, good_direction) — "higher" là mặc định trung tính (KHÔNG hàm ý "tăng luôn tốt") cho
    # hầu hết dòng — chi tiết tốt/xấu thực tế đã giải thích riêng ở "impact" từng chỉ báo.
    # reserve_assets_change dùng "lower" vì NHNN trình bày theo quy ước "Tài sản có" (tăng ghi ÂM,
    # giống dòng "Đầu tư trực tiếp ra nước ngoài: Tài sản có" cùng bảng) — xem note riêng.
    _SBV_BOP_META = {
        "current_account": ("Cán cân vãng lai (BOP, NHNN)", "higher"),
        "goods_export": ("Hàng hóa: Xuất khẩu (BOP, NHNN)", "higher"),
        "goods_import": ("Hàng hóa: Nhập khẩu (BOP, NHNN)", "higher"),
        "services_export": ("Dịch vụ: Xuất khẩu (BOP, NHNN)", "higher"),
        "services_import": ("Dịch vụ: Nhập khẩu (BOP, NHNN)", "higher"),
        "investment_income_received": ("Thu nhập đầu tư: Thu (BOP, NHNN)", "higher"),
        "investment_income_paid": ("Thu nhập đầu tư: Chi (BOP, NHNN)", "higher"),
        "secondary_income_received": ("Chuyển giao vãng lai: Thu (BOP, NHNN)", "higher"),
        "secondary_income_paid": ("Chuyển giao vãng lai: Chi (BOP, NHNN)", "higher"),
        "financial_account": ("Cán cân tài chính (BOP, NHNN)", "higher"),
        "fdi_assets_bop": ("Đầu tư trực tiếp ra nước ngoài — Tài sản có (BOP, NHNN)", "higher"),
        "fdi_liabilities_bop": ("Đầu tư trực tiếp vào Việt Nam — Tài sản nợ (BOP, NHNN)", "higher"),
        "portfolio_assets_bop": ("Đầu tư gián tiếp ra nước ngoài — Tài sản có (BOP, NHNN)", "higher"),
        "portfolio_liabilities_bop": ("Đầu tư gián tiếp vào Việt Nam — Tài sản nợ (BOP, NHNN)", "higher"),
        "external_debt_net": ("Vay, trả nợ nước ngoài — ròng (BOP, NHNN)", "higher"),
        "errors_omissions": ("Lỗi và Sai sót (BOP, NHNN)", "higher"),
        "overall_balance": ("Cán cân tổng thể (BOP, NHNN)", "higher"),
        "reserve_assets_change": ("Δ Tài sản dự trữ trong kỳ (BOP, NHNN)", "lower"),
    }

    def _ensure_bop_raw_entry(key):
        raw_key = f"bop_sbv_{key}"
        if raw_key in raw:
            return raw_key
        label, good_dir = _SBV_BOP_META.get(key, (key, "higher"))
        note = ("Cán cân thanh toán (BOP) quý, chuẩn BPM6, đơn vị Triệu USD. Nguồn chính: "
                 "dulieukinhte.com (mirror lại số liệu NHNN, HTML tĩnh — ít rủi ro bị chặn hơn "
                 "gọi trực tiếp sbv.gov.vn). Khi trang này CHƯA cập nhật quý mới nhất, hệ thống tự "
                 "động kiểm tra trực tiếp sbv.gov.vn/vi/can-can-thanh-toan-quoc-te để lấp quý đó "
                 "(fallback, chỉ gọi khi cần để tránh bị WAF chặn) — nguồn cụ thể của từng điểm dữ "
                 "liệu xem ở link gắn kèm.")
        if key == "reserve_assets_change":
            note += (" LƯU Ý QUY ƯỚC DẤU: NHNN trình bày dòng này theo quy ước \"Tài sản có\" — ÂM "
                      "(-) nghĩa là Tài sản dự trữ TĂNG (tích lũy thêm dự trữ, giống dòng \"Đầu tư "
                      "trực tiếp ra nước ngoài: Tài sản có\" cùng bảng — tăng tài sản ghi âm); DƯƠNG "
                      "(+) nghĩa là dự trữ GIẢM.")
        raw[raw_key] = {
            "group": "external", "label": label, "unit": "triệu USD",
            "good_direction": good_dir, "auto_source": "sbv", "note": note,
            "impact": "Một dòng trong bảng BOP quý — xem các dòng liên quan khác (current_account/financial_account/errors_omissions/overall_balance) để hiểu bối cảnh đầy đủ, không nên đọc 1 dòng riêng lẻ.",
            "series": [],
        }
        return raw_key

    existing_bop_periods = [p["period"] for p in raw.get("bop_sbv_current_account", {}).get("series", [])]
    existing_max_period = max(existing_bop_periods) if existing_bop_periods else None

    dlkt_data = fetch_dulieukinhte_bop()
    for period, period_data in dlkt_data.items():
        for key, value in period_data.items():
            raw_key = _ensure_bop_raw_entry(key)
            _merge_point_anywhere(raw, raw_key, period, value,
                                   "https://dulieukinhte.com/du-lieu/can-can-thanh-toan-361")
    dlkt_max_period = max(dlkt_data) if dlkt_data else None
    if dlkt_data:
        n_points = sum(len(v) for v in dlkt_data.values())
        print(f"  -> dulieukinhte.com: {len(dlkt_data)} quý ({min(dlkt_data)}..{dlkt_max_period}), {n_points} điểm dữ liệu")
    else:
        print("  -> dulieukinhte.com: không lấy được (đổi cấu trúc trang?)")

    # Quý BOP luôn công bố trễ ~2 quý theo lịch (ví dụ: Q2-2026 được NHNN công bố cuối Q3-2026,
    # xem dateCreated "2026-09-29" quan sát được khi test API Liferay trước đó) — dùng mốc này để
    # suy ra "quý mới nhất LẼ RA đã có" tính tới hôm nay, rồi mới quyết có cần hỏi SBV không. Nếu
    # so sánh thẳng dlkt_max_period với existing_max_period (đã lưu), một khi dulieukinhte ổn định
    # ở 1 quý thì mọi lần chạy sau đều so "bằng" -> lúc nào cũng gọi SBV, mất hết tác dụng giảm rủi
    # ro bị chặn mà user yêu cầu.
    _today = datetime.date.today()
    _q = (_today.month - 1) // 3 + 1 - 2
    _y = _today.year
    while _q < 1:
        _q += 4
        _y -= 1
    expected_latest_bop_quarter = f"{_y}-Q{_q}"

    need_sbv_fallback = (
        (dlkt_max_period is None or dlkt_max_period < expected_latest_bop_quarter)
        and (existing_max_period is None or existing_max_period < expected_latest_bop_quarter)
    )
    if need_sbv_fallback:
        print(f"  -> Có thể thiếu quý {expected_latest_bop_quarter} (dulieukinhte mới nhất: {dlkt_max_period}, đã lưu: {existing_max_period}) — kiểm tra trực tiếp SBV (fallback)...")
        bop_period, bop_data = fetch_sbv_bop_quarterly()
        if bop_period and bop_data and bop_period not in dlkt_data:
            for key, value in bop_data.items():
                raw_key = _ensure_bop_raw_entry(key)
                _merge_point_anywhere(raw, raw_key, bop_period, value,
                                       "https://www.sbv.gov.vn/vi/can-can-thanh-toan-quoc-te")
            print(f"  -> SBV fallback: {bop_period} ({len(bop_data)}/{len(_SBV_BOP_FIELD_MAP)} dòng)")
        elif bop_period and bop_period in dlkt_data:
            print(f"  -> SBV fallback: {bop_period} đã có từ dulieukinhte.com, bỏ qua.")
        else:
            print("  -> SBV fallback: không lấy được (WAF chặn hoặc đổi cấu trúc trang) — sẽ thử lại lần sau.")
    else:
        print(f"  -> Quý kỳ vọng mới nhất ({expected_latest_bop_quarter}) đã có sẵn từ dulieukinhte "
              f"({dlkt_max_period}) hoặc đã lưu từ lần chạy trước ({existing_max_period}) — không cần "
              f"gọi SBV trực tiếp.")

    # THEM (user 2026-10-03, "vẽ biểu đồ [áp lực tỷ giá] từ 2021 tới nay thì hay quá"): dulieukinhte/
    # SBV chỉ có ~5-6 quý gần nhất — BACKFILL lịch sử xa hơn từ IMF SDMX (api.imf.org, dataflow
    # IMF.STA:BOP, KHÔNG cần đăng nhập — khác hẳn IMF.STA:EER/BOP cho remittances đã thử TRƯỚC ĐÓ
    # trong phiên này và THẤT BẠI, lần này dùng đúng indicator code BPM6 chuẩn, đã verify khớp
    # CHÍNH XÁC 17/18 field qua 6 quý chồng lấp với NHNN/dulieukinhte — xem fetch_imf_bop_vietnam_
    # history()). CHỈ fill các kỳ CHƯA CÓ (merge, không ghi đè dữ liệu mới hơn đã có từ NHNN/
    # dulieukinhte — dù IMF cũng khớp chính xác các kỳ chồng lấp, ưu tiên nguồn trực tiếp cho kỳ
    # mới nhất). reserve_assets_change KHÔNG có field riêng ở IMF — tự suy ra = −overall_balance
    # (đúng định nghĩa BPM6, đã verify khớp 100% với dữ liệu NHNN đang có: mọi kỳ hiện có đều thấy
    # reserve_assets_change = −overall_balance chính xác).
    print("[IMF SDMX (BOP, DSD_BOP) — Backfill lịch sử BOP Việt Nam từ 1996-Q1]")
    imf_bop = fetch_imf_bop_vietnam_history()
    if imf_bop:
        imf_bop["reserve_assets_change"] = {p: round(-v, 2) for p, v in imf_bop.get("overall_balance", {}).items()}
        n_new = 0
        for key, series in imf_bop.items():
            raw_key = f"bop_sbv_{key}"
            if raw_key not in raw:
                continue  # field nay chua tung duoc _ensure_bop_raw_entry tao - bo qua, khong tu dung nhan/note
            existing_periods = {p["period"] for p in raw[raw_key]["series"]}
            for period, value in series.items():
                if period in existing_periods:
                    continue
                _merge_point_anywhere(raw, raw_key, period, value,
                                       "https://api.imf.org/external/sdmx/2.1 (IMF.STA:BOP)")
                n_new += 1
        print(f"  -> Backfill thêm {n_new} điểm (giữ nguyên các kỳ đã có từ NHNN/dulieukinhte)")
    else:
        print("  -> Không lấy được (API lỗi/đổi cấu trúc)")

    # THEM (user 2026-10-01, "kiều hối đây nhé, lấy cái về HCM là được rồi, lấy để tính vào cung
    # ngoại tệ lúc nãy bị thiếu ấy"): lấp gap "Kiều hối ĐÚNG NGHĨA" ở lớp ② Cung ngoại tệ — xem
    # fetch_dulieukinhte_kieu_hoi_hcm(). CHỈ PHẠM VI TP.HCM (không phải toàn quốc) — ghi rõ trong
    # note, KHÔNG dùng để thay thế bop_sbv_secondary_income_received (toàn quốc nhưng rộng hơn).
    print("[dulieukinhte.com — Kiều hối về TP.HCM theo quý (NHNN Chi nhánh Khu vực 2)]")
    kieu_hoi_data = fetch_dulieukinhte_kieu_hoi_hcm()
    if kieu_hoi_data:
        if "kieu_hoi_hcm" not in raw:
            raw["kieu_hoi_hcm"] = {
                "group": "external", "label": "Kiều hối về TP.HCM (NHNN Chi nhánh Khu vực 2)",
                "unit": "triệu USD", "good_direction": "higher", "auto_source": "sbv",
                "note": ("dulieukinhte.com (mirror NHNN Chi nhánh Khu vực 2, trước 01/07/2025 là "
                         "NHNN chi nhánh TP.HCM), HTML tĩnh — không WAF. LƯU Ý: đây là Kiều hối "
                         "ĐÚNG NGHĨA (personal remittances) nhưng CHỈ PHẠM VI TP.HCM — KHÔNG PHẢI "
                         "số toàn quốc (TP.HCM lịch sử chiếm tỷ trọng lớn nhưng không phải 100%, "
                         "không được tự suy ra số toàn quốc từ số này). Khác với "
                         "bop_sbv_secondary_income_received (toàn quốc nhưng RỘNG HƠN — gồm mọi "
                         "chuyển giao vãng lai, không riêng kiều hối cá nhân) — 2 chỉ báo bổ sung "
                         "cho nhau, không thay thế nhau."),
                "impact": ("Kiều hối về TP.HCM tăng là nguồn cung USD thực vào hệ thống (qua kênh "
                           "công ty kiều hối/ngân hàng) — tín hiệu TÍCH CỰC cho cung ngoại tệ, dù "
                           "chỉ phản ánh 1 địa bàn, không đại diện toàn quốc."),
                "series": [],
            }
        for period, value in kieu_hoi_data.items():
            _merge_point_anywhere(raw, "kieu_hoi_hcm", period, value,
                                   "https://dulieukinhte.com/du-lieu/kieu-hoi-tp-ho-chi-minh-426")
        print(f"  -> {len(kieu_hoi_data)} quý ({min(kieu_hoi_data)}..{max(kieu_hoi_data)})")
    else:
        print("  -> Không lấy được (đổi cấu trúc trang?)")

    # THEM (user 2026-10-01, "tổng mức bán lẻ sẽ đánh giá về sức tiêu thụ và nhu cầu chi tiêu của
    # người dân ... tư duy về chỉ số này đánh giá như nào về sức khỏe nền kinh tế"): backfill DÀY
    # HƠN NHIỀU cho retail_sales_growth (trước chỉ 1-2 điểm/năm, quét câu văn báo cáo NSO rất dễ
    # vỡ — user tự thấy thiếu khi xem biểu đồ đối chiếu năm 2025), cộng 4 cấu phần MỚI để so sánh
    # cơ cấu tiêu dùng (hàng hoá vs dịch vụ lưu trú-ăn uống vs du lịch — xem fetch_dulieukinhte_
    # retail_sales()). Tính YoY (danh nghĩa + THỰC, trừ CPI) ở _add_retail_sales_derived()
    # (template_vimo.py), KHÔNG tính ở đây (cần cpi_yoy đã tính xong, chạy sau update_vimo_raw()).
    print("[dulieukinhte.com — Tổng mức bán lẻ hàng hóa & dịch vụ tiêu dùng theo tháng (NSO)]")
    _RETAIL_SALES_META = {
        "Tổng mức bán lẻ HH và DV": ("retail_sales_total_monthly", "Tổng mức bán lẻ HH & DV tiêu dùng (theo tháng, danh nghĩa)"),
        "Bán lẻ hàng hoá": ("retail_sales_goods_monthly", "Bán lẻ hàng hoá (theo tháng, danh nghĩa)"),
        "Dịch vụ lưu trữ, ăn uống": ("retail_sales_hospitality_monthly", "Dịch vụ lưu trú, ăn uống (theo tháng, danh nghĩa)"),
        "Du lịch lữ hành": ("retail_sales_travel_monthly", "Du lịch lữ hành (theo tháng, danh nghĩa)"),
        "Dịch vụ khác": ("retail_sales_other_monthly", "Dịch vụ khác (theo tháng, danh nghĩa)"),
    }
    retail_data = fetch_dulieukinhte_retail_sales()
    n_points = 0
    for row_label, (raw_key, label) in _RETAIL_SALES_META.items():
        series = retail_data.get(row_label)
        if not series:
            continue
        if raw_key not in raw:
            raw[raw_key] = {
                "group": "growth", "label": label, "unit": "tỷ đồng", "good_direction": "higher",
                "auto_source": "nso",
                "note": ("dulieukinhte.com (mirror Tổng cục Thống kê), HTML tĩnh. Theo GIÁ HIỆN "
                         "HÀNH (danh nghĩa, CHƯA trừ lạm phát) — xem retail_sales_growth (đã có "
                         "sẵn)/retail_sales_growth_real (phái sinh) để có %YoY danh nghĩa/thực."),
                "impact": "Đo trực tiếp sức mua/tiêu dùng nội địa theo tháng — so sánh tốc độ tăng giữa hàng hoá (thiết yếu) và dịch vụ lưu trú-ăn uống/du lịch (trải nghiệm, nhạy cảm với niềm tin tiêu dùng) để biết tiêu dùng đang 'khỏe' (dịch chuyển sang dịch vụ) hay chỉ tăng do giá.",
                "series": [],
            }
        for period, value in series.items():
            _merge_point_anywhere(raw, raw_key, period, value,
                                   "https://dulieukinhte.com/du-lieu/ban-le-hang-hoa-va-dich-vu-273")
        n_points += len(series)
    print(f"  -> {n_points} điểm (5 dòng: tổng + 4 cấu phần)")

    # THEM (user 2026-10-01, "dữ liệu tổng xuất khẩu nhập khẩu này có thể tính được xuất khẩu FDI
    # và nhập khẩu FDI ảnh hưởng ntn tới xnk của VN"): XNK theo THÁNG ĐƠN LẺ (khác export_domestic_
    # usd_bn/export_fdi_usd_bn hiện có — những dòng đó LŨY KẾ từ đầu năm, quét regex câu văn báo cáo
    # NSO rất dễ vỡ nếu đổi cách viết câu) tách khu vực trong nước/FDI, nguồn Hải quan qua
    # dulieukinhte.com (bảng tĩnh, ổn định hơn). Tính thêm % tỷ trọng FDI trong tổng KNXK/KNNK mỗi
    # tháng — trả lời trực tiếp câu hỏi "FDI ảnh hưởng thế nào" (xem _add_fdi_trade_share trong
    # template_vimo.py, KHÔNG lưu % này vào vimo_raw.json — tính lại mỗi lần chạy từ 3 dòng MỨC).
    _TRADE_MONTHLY_META = {
        "Tổng": ("total", "Tổng (Hải quan, theo tháng)"),
        "Khu vực trong nước": ("domestic", "Khu vực kinh tế trong nước (Hải quan, theo tháng)"),
        "Khu vực trong FDI": ("fdi", "Khu vực FDI (Hải quan, theo tháng)"),
    }
    for direction, fetch_fn, url in [
        ("export", fetch_dulieukinhte_export_fdi_split, "https://dulieukinhte.com/du-lieu/tong-xuat-khau-292"),
        ("import", fetch_dulieukinhte_import_fdi_split, "https://dulieukinhte.com/du-lieu/tong-nhap-khau-293"),
    ]:
        label_vn = "Xuất khẩu" if direction == "export" else "Nhập khẩu"
        print(f"[dulieukinhte.com — {label_vn} theo tháng, tách khu vực trong nước/FDI (Hải quan)]")
        split_data = fetch_fn()
        n_points = 0
        for row_label, (suffix, desc) in _TRADE_MONTHLY_META.items():
            series = split_data.get(row_label)
            if not series:
                continue
            raw_key = f"{direction}_monthly_{suffix}"
            if raw_key not in raw:
                raw[raw_key] = {
                    "group": "trade", "label": f"{label_vn} — {desc}", "unit": "triệu USD",
                    "good_direction": "higher", "auto_source": "customs",
                    "note": ("dulieukinhte.com (mirror Tổng cục Hải quan), HTML tĩnh. GIÁ TRỊ RIÊNG "
                             "1 THÁNG (KHÔNG phải lũy kế — khác export_domestic_usd_bn/export_fdi_"
                             "usd_bn hiện có, 2 nhóm chỉ báo bổ sung cho nhau, không trộn lẫn)."),
                    "impact": "Thuộc nhóm 'trade structure' — xem export_fdi_share_pct/import_fdi_share_pct (phái sinh) để biết tỷ trọng FDI trong tổng KNXK/KNNK.",
                    "series": [],
                }
            for period, value in series.items():
                _merge_point_anywhere(raw, raw_key, period, value, url)
            n_points += len(series)
        print(f"  -> {n_points} điểm (Tổng/trong nước/FDI)")

    # THEM (user 2026-10-01): NEER/REER Việt Nam — BIS (WS_EER) và IMF.STA:EER đều KHÔNG có Việt
    # Nam (đã verify qua API thật), nhưng Darvas (Bruegel) CÓ, theo tháng, và chính IMF cũng dùng
    # Darvas làm nguồn NEER/REER cho Việt Nam trong phân tích Article IV — xem fetch_darvas_reer_
    # neer_vietnam(). KHÔNG gọi là "IMF NEER/REER" — ghi rõ nguồn Darvas, không phải BIS/IMF chính
    # thức (khác methodology, dù tương quan cao).
    print("[Bruegel (Darvas) — NEER/REER Việt Nam theo tháng]")
    darvas = fetch_darvas_reer_neer_vietnam()
    for key, label, unit_label in [
        ("neer", "Tỷ giá hiệu lực danh nghĩa (NEER) Việt Nam — Darvas (Bruegel)", "Index (tháng cơ sở nội bộ Darvas)"),
        ("reer", "Tỷ giá hiệu lực thực (REER) Việt Nam — Darvas (Bruegel)", "Index (tháng cơ sở nội bộ Darvas)"),
    ]:
        pts = darvas.get(key, [])
        if not pts:
            continue
        raw_key = f"darvas_{key}_vn"
        if raw_key not in raw:
            raw[raw_key] = {
                "group": "external", "label": label, "unit": unit_label, "good_direction": "lower",
                "auto_source": "derived", "series": [],
                "note": ("bruegel.org (Zsolt Darvas, 'Timely measurement of REER', cập nhật định kỳ) — KHÔNG phải "
                         "nguồn chính thức BIS/IMF (methodology riêng: tỷ giá song phương + CPI + trọng số thương "
                         "mại), nhưng Việt Nam KHÔNG có trong bộ BIS (WS_EER)/IMF (IMF.STA:EER) — đã verify qua gọi "
                         "API thật, cả 2 đều 0 dữ liệu cho Việt Nam. IMF tự dùng Darvas làm nguồn NEER/REER Việt Nam "
                         "trong phân tích Article IV (tương quan cao với EER chính thức của IMF). Basket 'Broad' "
                         "(120 đối tác thương mại, KHÔNG dùng basket 'Narrow' 51 đối tác)."),
                "impact": ("NEER tăng = VND lên giá danh nghĩa so với rổ tiền tệ đối tác thương mại (KHÔNG chỉ riêng "
                            "USD như USD/VND); REER tăng = VND lên giá THỰC (đã điều chỉnh lạm phát tương đối), nghĩa "
                            "là hàng hóa VN kém cạnh tranh hơn về giá. USD/VND tăng nhưng NEER ổn định = áp lực chủ "
                            "yếu từ riêng USD (DXY mạnh lên toàn cầu); USD/VND tăng VÀ NEER giảm = áp lực tỷ giá rộng "
                            "hơn, không chỉ do USD."),
            }
        darvas_src = "https://www.bruegel.org/publications/datasets/real-effective-exchange-rates-for-178-countries-a-new-database"
        for period, value in pts:
            _merge_point_anywhere(raw, raw_key, period, round(value, 4), darvas_src)
        print(f"  -> {key.upper()}: {len(pts)} điểm (mới nhất {pts[-1][0]}={pts[-1][1]:.2f})")

    print("[Hải quan — Xuất/nhập khẩu theo tháng (file Excel cục bộ, CHỈ có khi chạy thủ công trên máy có sẵn thư mục)]")
    xnk = load_customs_xnk_local()
    if xnk["export"] or xnk["import"]:
        # PHẢI dùng _merge_point_anywhere() (không phải _append_point()) — CUSTOMS_XNK_FOLDER
        # thường có NHIỀU file V01-*.xls chồng lấn kỳ (vd file "năm 2025" phủ Tháng 01-12, file
        # "năm 2026" phủ Tháng 01-06 của năm SAU nhưng cấu trúc đọc theo TÊN FILE không theo
        # NĂM DỮ LIỆU nên các file MỚI của CÙNG năm vẫn có thể phủ lại các tháng ĐÃ CÓ) — dùng
        # _append_point() (chỉ so khớp phần tử CUỐI) khiến mỗi file quét lại từ đầu sẽ nối THÊM
        # bản trùng thay vì ghi đè, phát hiện thực tế: 144 điểm nhưng chỉ 18 tháng riêng biệt (mỗi
        # tháng lặp đúng 8 lần, khớp số file trong thư mục) — user 2026-08-08 phát hiện qua yêu
        # cầu vẽ biểu đồ cán cân thương mại theo tháng.
        for period, value, fname in xnk["export"]:
            _merge_point_anywhere(raw, "export_value_monthly", period, value, fname)
        for period, value, fname in xnk["import"]:
            _merge_point_anywhere(raw, "import_value_monthly", period, value, fname)
        export_by_period = {p: v for p, v, _ in xnk["export"]}
        import_by_period = {p: v for p, v, _ in xnk["import"]}
        for period in sorted(set(export_by_period) & set(import_by_period)):
            balance = round(export_by_period[period] - import_by_period[period], 4)
            _merge_point_anywhere(raw, "trade_balance_monthly", period, balance, "Hải quan (V01+V02, tính từ xuất trừ nhập)")
        print(f"  -> {len(set(export_by_period))} tháng xuất khẩu, {len(set(import_by_period))} tháng nhập khẩu")
    else:
        print("  [INFO] Không tìm thấy thư mục/file Hải quan cục bộ — bỏ qua (bình thường khi chạy trên GitHub Action).")

    raw["_meta"]["last_auto_update"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    save_raw(raw)
    print("\n[OK] Đã ghi data/vimo_raw.json")

    # THEM (user 2026-10-01, "làm cho tôi bảng GDP theo ngành... có màu để biểu thị nhóm nào đang
    # tốt lên xấu đi" / "cơ cấu CPI... làm biểu đồ tổng hợp... có màu để biểu thị cái nào đang nhức
    # nhối"): lưu RIÊNG vào data/vimo_sector_detail.json (KHÔNG phải vimo_raw.json ở trên — đây là
    # MỨC theo từng ngành/nhóm hàng, không phải 1 indicator dạng {group,label,series} nên không cho
    # vào raw, tránh vỡ các vòng lặp raw.items() ở template_vimo.py). Merge dồn theo kỳ (dulieukinhte
    # chỉ hiện ~13 quý/37 tháng gần nhất mỗi lần — giữ lại kỳ cũ đã lưu, không mất lịch sử khi dulieu-
    # kinhte cắt bớt cửa sổ hiển thị). template_vimo._build_level_yoy_heatmap() đọc file này để tự
    # tính %YoY (MỨC theo ngành/nhóm hàng không so trực tiếp được, phải quy về YoY mới so được).
    print("[dulieukinhte.com — GDP theo ngành (quý) + CPI theo nhóm hàng (tháng), cho bảng heatmap]")
    try:
        with open(SECTOR_DETAIL_PATH, "r", encoding="utf-8") as f:
            sector_detail = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        sector_detail = {}

    # THEM (user 2026-10-01, "xuất khẩu theo mặt hàng... lập bảng màu để biết mặt hàng nào đang
    # xuất khẩu tốt lên" / "tương tự với nhập khẩu theo mặt hàng" / "đây là bảng dữ liệu giá nhập
    # khẩu, hãy tạo bảng để biết giá đang căng không" / "tương tự với giá xuất khẩu"): cùng cơ chế
    # merge-dồn-theo-kỳ như GDP/CPI ở trên — 6 bộ dùng CHUNG 1 vòng lặp (tránh lặp code 6 lần).
    _SECTOR_DETAIL_SOURCES = [
        ("gdp_sector_levels", "GDP theo ngành (quý)", fetch_dulieukinhte_gdp_sector),
        ("cpi_group_levels", "CPI theo nhóm hàng (tháng)", fetch_dulieukinhte_cpi_group),
        ("export_commodity_levels", "Xuất khẩu theo mặt hàng (tháng)", fetch_dulieukinhte_export_commodity),
        ("import_commodity_levels", "Nhập khẩu theo mặt hàng (tháng)", fetch_dulieukinhte_import_commodity),
        ("export_price_levels", "Giá xuất khẩu bình quân theo mặt hàng (tháng)", fetch_dulieukinhte_export_price),
        ("import_price_levels", "Giá nhập khẩu bình quân theo mặt hàng (tháng)", fetch_dulieukinhte_import_price),
    ]
    for file_key, label_vn, fetch_fn in _SECTOR_DETAIL_SOURCES:
        new_data = fetch_fn()
        sector_detail.setdefault(file_key, {})
        for label, series in new_data.items():
            sector_detail[file_key].setdefault(label, {}).update(series)
        print(f"  -> {label_vn}: {len(new_data)}/{len(sector_detail[file_key])} hàng (mới/tổng đã lưu)")

    with open(SECTOR_DETAIL_PATH, "w", encoding="utf-8") as f:
        json.dump(sector_detail, f, ensure_ascii=False, indent=2)
    print(f"  -> Đã ghi {SECTOR_DETAIL_PATH}")

    # Rủi ro lãi suất/thanh khoản TOÀN HỆ THỐNG ngân hàng (2026-08) — ghi RIÊNG vào
    # data/bank_alm/<TICKER>.json (KHÔNG phải vimo_raw.json ở trên), template_vimo.py đọc lại các
    # file này để tổng hợp. Kiểm tra RẺ cho cả 26 ngân hàng mỗi tuần (vài giây, chỉ gọi API liệt kê
    # BCTC) — CHỈ ngân hàng nào thực sự có kỳ đã kiểm toán/soát xét MỚI hơn store hiện có mới bị OCR
    # lại (đắt); các tuần thường (chưa ngân hàng nào công bố kỳ mới) gần như miễn phí. Bọc try/except
    # để lỗi phần này KHÔNG làm hỏng toàn bộ phần vĩ mô đã fetch xong ở trên.
    try:
        from bank_system_risk import refresh_quarterly_balance_sheet_all_banks, refresh_bank_alm_data
        print("\n[He thong ALM ngan hang] Cap nhat bang can doi quy (khong can OCR)...")
        refresh_quarterly_balance_sheet_all_banks()
        print("[He thong ALM ngan hang] Kiem tra do moi 2 bang gap (OCR chi khi thuc su can)...")
        result = refresh_bank_alm_data()
        print(f"  -> Ky muc tieu: {result.get('target_period')}, thay doi: {result.get('changed')}")
        for ticker, action in (result.get("actions") or {}).items():
            if action not in ("skip_da_co", "skip_da_va_dung_nguon", "skip_da_danh_dau_thieu"):
                print(f"     {ticker}: {action}")
    except Exception as e:
        print(f"  [WARN] He thong ALM ngan hang: bo qua ({e})")


if __name__ == "__main__":
    update_vimo_raw()
