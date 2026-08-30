#!/usr/bin/env python3
"""
bank_risk_notes.py — Tự động dò, tải, OCR nhẹ và phân tích 2 thuyết minh BCTC "RỦI RO LÃI SUẤT" và
"RỦI RO THANH KHOẢN" cho cổ phiếu NGÂN HÀNG. Chỉ gọi từ template_banking.py — ticker đã được phân
loại ngành ngân hàng từ trước (run_analysis.py), nên module này không tự kiểm tra lại ngành.

Cơ chế (đã verify bằng dữ liệu THẬT — BCTC hợp nhất kiểm toán MBB năm 2025, 2026-08-30, xem chi tiết
trong PR/commit message đi kèm):
- 2 thuyết minh này chỉ có ĐẦY ĐỦ trong BCTC HỢP NHẤT ĐÃ KIỂM TOÁN (Quarter=5 theo quy ước CafeF) —
  báo cáo quý/soát xét bán niên thường rút gọn thuyết minh, không có 2 bảng chi tiết này. Dùng lại
  đúng cơ chế dò+tải PDF đã có ở bctc_pdf_tool.py (CafeF + 24hmoney, generic cho mọi ticker).
- Vị trí trong tài liệu: nằm trong khoảng ~72-100% cuối tài liệu (verify MBB 2025: trang 90-91/103 và
  95-96/103, tức 87-93%) — KHÔNG dùng số thứ tự thuyết minh cố định (số này đổi theo từng ngân
  hàng/năm — MBB 2025 là note 49/51) mà tìm theo TIÊU ĐỀ chữ.
- Cả 2 bảng đều đã có SẴN dòng tổng hợp do chính ngân hàng/kiểm toán tính:
    "Mức chênh nhạy cảm với lãi suất nội, ngoại bảng (5) = (3) + (4)"  (rủi ro lãi suất)
    "Mức chênh thanh khoản ròng (3) = (1) - (2)"                        (rủi ro thanh khoản)
  → CHỈ cần định vị và trích ĐÚNG dòng này (không tự cộng trừ lại từng dòng tài sản/nợ — đáng tin cậy
  hơn nhiều so với tái dựng toàn bộ bảng qua OCR, vốn dễ sai từng ô lẻ).
- OCR dùng pytesseract (nhẹ, vài giây/trang) — ĐÃ THỬ và loại bỏ 2 phương án khác: (1)
  opendataloader-pdf-hybrid (docling+easyocr, cùng cơ chế dùng cho HPG/MWG) bị treo >70 phút không
  xong với 1 file chỉ 25-34 trang khi chạy trên máy không có GPU; (2) easyocr gọi trực tiếp (không
  qua hybrid server) mất ~78 giây/trang, quá chậm để quét dò nhiều trang. pytesseract cần binary
  tesseract-ocr cài sẵn trên máy (đã có trong requirements.txt + .github/workflows/analyze_stock.yml
  cài qua apt cho CI; máy Windows cá nhân cần tự cài qua `winget install UB-Mannheim.TesseractOCR`
  — lệnh này cần chạy trong phiên có thể xác nhận UAC, không chạy được từ tiến trình nền không tương
  tác). Module KHÔNG BAO GIỜ crash nếu thiếu tesseract — trả về None, template_banking.py tự bỏ qua
  phần này và không in cảnh báo gây nhiễu (in 1 dòng [SKIP] rõ ràng để user biết cần cài gì).
"""
import os
import re
import sys
import unicodedata
import tempfile

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

import requests

from bctc_pdf_tool import fetch_cafef_list, fetch_24hmoney_list, select_best_reports, HEADERS

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
CACHE_DIR = os.path.join(PROJECT_ROOT, ".cache")

# Cột (bucket) theo ĐÚNG thứ tự trái->phải trong 2 bảng — verify thật trên BCTC MBB 2025. Một số ngân
# hàng có thể đặt "Quá hạn" thành 1 cột duy nhất thay vì tách "Trên 3 tháng"/"Đến 3 tháng" như MBB —
# labels chỉ dùng để hiển thị, KHÔNG dùng để suy luận, số bucket thật lấy từ chính số lượng giá trị
# OCR trích được (xem _extract_number_row).
INTEREST_RATE_BUCKETS = ["qua_han", "khong_anh_huong_lai_suat", "den_1_thang", "tu_1_3_thang",
                          "tu_3_6_thang", "tu_6_12_thang", "tu_1_5_nam", "tren_5_nam"]
LIQUIDITY_BUCKETS = ["qua_han_tren_3t", "qua_han_den_3t", "den_1_thang", "tu_1_3_thang",
                      "tu_3_12_thang", "tu_1_5_nam", "tren_5_nam"]


def _strip_accents(s):
    s = unicodedata.normalize("NFD", s or "")
    return "".join(c for c in s if unicodedata.category(c) != "Mn").lower()


def _ocr_page_text(pdf_path, page_index, dpi=300):
    """OCR 1 trang (0-based index) bằng pytesseract. Trả về "" nếu thiếu pytesseract/tesseract binary
    hoặc lỗi bất kỳ bước nào — KHÔNG BAO GIỜ raise."""
    try:
        import pytesseract
        import pypdfium2 as pdfium
    except ImportError:
        return None  # None = "không có công cụ", khác "" = "có công cụ nhưng OCR ra rỗng"
    try:
        doc = pdfium.PdfDocument(pdf_path)
        if page_index >= len(doc):
            return ""
        img = doc[page_index].render(scale=dpi / 72).to_pil()
        return pytesseract.image_to_string(img, lang="vie")
    except Exception as e:
        print(f"  [WARN] OCR trang {page_index+1} loi: {e}")
        return ""


def _find_note_pages(pdf_path, start_frac=0.70, max_pages=40):
    """Quét từ start_frac*tổng_số_trang tới hết tài liệu, OCR TỪNG TRANG (dừng ngay khi đã tìm đủ cả
    2 tiêu đề — không OCR speculative cả vùng). Trả về dict {"lai_suat": page_idx|None,
    "thanh_khoan": page_idx|None} (0-based). Trả về {} nếu thiếu pytesseract."""
    try:
        import pypdfium2 as pdfium
    except ImportError:
        return {}
    doc = pdfium.PdfDocument(pdf_path)
    total = len(doc)
    start = max(0, int(total * start_frac))
    pages_to_scan = list(range(start, total))[:max_pages]

    found = {}
    for idx in pages_to_scan:
        text = _ocr_page_text(pdf_path, idx)
        if text is None:
            return {}  # thiếu pytesseract - dừng hẳn, không quét tiếp vô ích
        flat = _strip_accents(text)
        if "lai_suat" not in found and "rui ro lai suat" in flat:
            found["lai_suat"] = idx
        if "thanh_khoan" not in found and "rui ro thanh khoan" in flat:
            found["thanh_khoan"] = idx
        if "lai_suat" in found and "thanh_khoan" in found:
            break
    return found


_ROW_LABEL_FLAT = {
    "lai_suat": "muc chenh nhay cam",
    "thanh_khoan": "muc chenh thanh khoan rong",
}


def _extract_number_row(text, label_flat, n_buckets, window=600):
    """Tìm dòng có nhãn `label_flat` (đã strip dấu) trong `text`, trích các số VND dạng Việt Nam
    (dấu chấm phân cách nghìn, số âm trong ngoặc) trong `window` ký tự sau nhãn. Chỉ trả về kết quả
    nếu số lượng số trích được >= n_buckets (lấy đúng n_buckets số ĐẦU TIÊN — cột cuối "Tổng cộng"
    nếu dư sẽ bị bỏ qua vì không cần). Trả về None nếu không tìm thấy nhãn hoặc thiếu số — KHÔNG ĐOÁN
    số liệu thiếu.

    Dùng `rfind` (khớp CUỐI CÙNG) chứ không phải `find` (khớp ĐẦU TIÊN): riêng thuyết minh rủi ro lãi
    suất có 2 dòng cùng nhãn "Mức chênh nhạy cảm..." — dòng (3) nội bảng và dòng (5) nội+ngoại bảng
    (đầy đủ hơn, verify thật MBB 2025: 2 dòng cho ra CÙNG SỐ vì (4) ngoại bảng bằng 0, nhưng ngân hàng
    khác có phái sinh lãi suất lớn có thể (4) khác 0 — lấy dòng CUỐI để luôn ra số ĐẦY ĐỦ NHẤT)."""
    flat = _strip_accents(text)
    pos = flat.rfind(label_flat)
    if pos == -1:
        return None
    window_text = text[pos:pos + window] if len(flat) == len(text) else text  # fallback nếu lệch index
    # Số VN: tuỳ chọn ngoặc (âm), chữ số + nhiều nhóm ".xxx", KHÔNG khớp số lẻ 1-2 chữ số đứng riêng
    # (tránh bắt nhầm số thứ tự dòng kiểu "(3)", "(5)" ngay trong chính nhãn dòng).
    nums = re.findall(r"\(?-?[\d]{1,3}(?:\.[\d]{3})+\)?", window_text)
    if len(nums) < n_buckets:
        return None
    vals = []
    for tok in nums[:n_buckets]:
        neg = tok.startswith("(") and tok.endswith(")")
        clean = tok.strip("()").replace(".", "")
        try:
            v = float(clean)
        except ValueError:
            return None
        vals.append(-v if neg else v)
    return vals


def fetch_bank_risk_gaps(ticker):
    """Trả về dict {"interest_rate_gap": {bucket: value}, "liquidity_gap": {bucket: value},
    "source_title":, "source_url":, "fetched_year":} hoặc None nếu bất kỳ bước nào thất bại (thiếu
    tesseract, không tìm thấy BCTC năm, không định vị được note, OCR không đọc đủ số...). KHÔNG BAO
    GIỜ raise — template_banking.py gọi hàm này trong try/except nhưng bản thân hàm đã tự an toàn."""
    ticker = ticker.upper()
    try:
        items = fetch_cafef_list(ticker) + fetch_24hmoney_list(ticker)
    except Exception as e:
        print(f"  [SKIP] Rui ro lai suat/thanh khoan: khong lay duoc danh sach BCTC ({e})")
        return None
    annual = [x for x in select_best_reports(items) if x.get("Quarter") == 5]
    if not annual:
        print("  [SKIP] Rui ro lai suat/thanh khoan: khong tim thay BCTC nam (kiem toan) hop nhat")
        return None
    latest = max(annual, key=lambda x: x["Year"])

    os.makedirs(CACHE_DIR, exist_ok=True)
    pdf_path = os.path.join(CACHE_DIR, f"{ticker}_{latest['Year']}_CN_full.pdf")
    try:
        if not os.path.exists(pdf_path):
            r = requests.get(latest["Link"].replace(" ", "%20"), headers=HEADERS, timeout=60)
            r.raise_for_status()
            with open(pdf_path, "wb") as f:
                f.write(r.content)
    except Exception as e:
        print(f"  [SKIP] Rui ro lai suat/thanh khoan: tai BCTC that bai ({e})")
        return None

    pages = _find_note_pages(pdf_path)
    if not pages:
        print("  [SKIP] Rui ro lai suat/thanh khoan: thieu pytesseract/tesseract-ocr binary "
              "(cai qua 'winget install UB-Mannheim.TesseractOCR' hoac 'apt install tesseract-ocr')")
        return None

    result = {"source_title": latest["Name"], "source_url": latest["Link"], "fetched_year": latest["Year"]}
    for key, bucket_list, n in (("lai_suat", INTEREST_RATE_BUCKETS, len(INTEREST_RATE_BUCKETS)),
                                 ("thanh_khoan", LIQUIDITY_BUCKETS, len(LIQUIDITY_BUCKETS))):
        page_idx = pages.get(key)
        if page_idx is None:
            continue
        # Bảng số thường nằm ở TRANG SAU trang tiêu đề/phương pháp luận (verify MBB: heading trang
        # 90, bảng số trang 91) — OCR cả 2 trang, ưu tiên trang có nhãn dòng tổng hợp.
        vals = None
        for p in (page_idx, page_idx + 1):
            text = _ocr_page_text(pdf_path, p)
            if not text:
                continue
            vals = _extract_number_row(text, _ROW_LABEL_FLAT[key], n)
            if vals:
                break
        if vals:
            gap_key = "interest_rate_gap" if key == "lai_suat" else "liquidity_gap"
            result[gap_key] = dict(zip(bucket_list, vals))

    if "interest_rate_gap" not in result and "liquidity_gap" not in result:
        print("  [SKIP] Rui ro lai suat/thanh khoan: dinh vi duoc trang nhung khong doc du so lieu "
              "(OCR chat luong kem hoac dinh dang bang khac chuan)")
        return None
    return result


# ── Tính chỉ số từ gap đã trích ──────────────────────────────────────────────────────────────────

def compute_interest_rate_risk_metrics(gap, total_assets, shock_bps=100):
    """gap: dict theo INTEREST_RATE_BUCKETS. Trả về dict chỉ số — xem docstring module cho phương
    pháp (static gap, trọng số theo điểm giữa mỗi bucket, chỉ 4 bucket <=12 thang moi anh huong NII
    NAM NAY, bucket >1 nam khong lap lai trong nam danh gia)."""
    b = gap
    horizon_1y_keys = ["den_1_thang", "tu_1_3_thang", "tu_3_6_thang", "tu_6_12_thang"]
    weights = {"den_1_thang": 11.5 / 12, "tu_1_3_thang": 10 / 12, "tu_3_6_thang": 7.5 / 12, "tu_6_12_thang": 3 / 12}
    cum_1y = sum(b.get(k, 0.0) for k in horizon_1y_keys)
    nii_sens = sum(b.get(k, 0.0) * weights[k] for k in horizon_1y_keys) * (shock_bps / 10000)
    ratio_1y = cum_1y / total_assets if total_assets else None
    if ratio_1y is None:
        level = "Khong xac dinh"
    elif abs(ratio_1y) < 0.02:
        level = "Thap (gan trung tinh)"
    elif abs(ratio_1y) < 0.05:
        level = "Trung binh"
    else:
        level = "Cao"
    return {
        "gap_by_bucket": b,
        "gap_ratio_by_bucket": {k: (v / total_assets if total_assets else None) for k, v in b.items()},
        "cumulative_gap_1y": cum_1y,
        "cumulative_gap_1y_ratio": ratio_1y,
        "sensitive_type": "Asset-sensitive" if cum_1y > 0 else ("Liability-sensitive" if cum_1y < 0 else "Trung tinh"),
        "nii_sensitivity_per_shock": nii_sens,
        "shock_bps": shock_bps,
        "sensitivity_level": level,
    }


def compute_liquidity_risk_metrics(gap, total_assets):
    """gap: dict theo LIQUIDITY_BUCKETS."""
    b = gap
    cum_1m = b.get("den_1_thang", 0.0)
    cum_1y = cum_1m + b.get("tu_1_3_thang", 0.0) + b.get("tu_3_12_thang", 0.0)
    ratio_1m = cum_1m / total_assets if total_assets else None
    ratio_1y = cum_1y / total_assets if total_assets else None
    if ratio_1m is None:
        level = "Khong xac dinh"
    elif ratio_1m > -0.03:
        level = "Thap"
    elif ratio_1m > -0.08:
        level = "Trung binh"
    else:
        level = "Cao"
    return {
        "gap_by_bucket": b,
        "gap_ratio_by_bucket": {k: (v / total_assets if total_assets else None) for k, v in b.items()},
        "cumulative_gap_1m": cum_1m,
        "cumulative_gap_1m_ratio": ratio_1m,
        "cumulative_gap_1y": cum_1y,
        "cumulative_gap_1y_ratio": ratio_1y,
        "risk_level": level,
    }


def build_risk_narrative(ir_metrics, liq_metrics, ticker):
    """Trả về đoạn văn tiếng Việt tóm tắt 2 rủi ro, dùng cho PDF/JSON. Không tự bịa số nào ngoài các
    dict đầu vào đã tính."""
    lines = []
    if ir_metrics:
        pct = ir_metrics["cumulative_gap_1y_ratio"]
        pct_s = f"{pct*100:+.2f}%" if pct is not None else "?"
        nii = ir_metrics["nii_sensitivity_per_shock"]
        bps = ir_metrics["shock_bps"]
        lines.append(
            f"Rui ro lai suat: gap luy ke <=1 nam = {pct_s} tong tai san "
            f"({ir_metrics['sensitive_type']}, muc do nhay cam: {ir_metrics['sensitivity_level']}). "
            f"Neu lai suat TANG {bps} diem co ban, NII uoc doi {nii:,.0f} trieu VND; "
            f"neu GIAM {bps} diem, NII uoc doi {-nii:,.0f} trieu VND (dau am = giam, dau duong = tang)."
        )
    if liq_metrics:
        r1m = liq_metrics["cumulative_gap_1m_ratio"]
        r1y = liq_metrics["cumulative_gap_1y_ratio"]
        r1m_s = f"{r1m*100:+.2f}%" if r1m is not None else "?"
        r1y_s = f"{r1y*100:+.2f}%" if r1y is not None else "?"
        lines.append(
            f"Rui ro thanh khoan: gap rong <=1 thang = {r1m_s} tong tai san, luy ke <=1 nam = {r1y_s} "
            f"(muc do: {liq_metrics['risk_level']}). Gap am nghia la ngan hang can tiep tuc tai tai "
            f"tro/huy dong de bu dap phan chenh lech ky han — ap luc nay can duoc phan anh vao COE "
            f"(phan bu rui ro dac thu) hoac P/B muc tieu neu muc do Cao."
        )
    return " ".join(lines)
