#!/usr/bin/env python3
"""
bank_risk_notes.py — Tự động dò, tải, OCR nhẹ và phân tích 2 thuyết minh BCTC "RỦI RO LÃI SUẤT" và
"RỦI RO THANH KHOẢN" cho cổ phiếu NGÂN HÀNG. Chỉ gọi từ template_banking.py — ticker đã được phân
loại ngành ngân hàng từ trước (run_analysis.py), nên module này không tự kiểm tra lại ngành.

Cơ chế (đã verify bằng dữ liệu THẬT — BCTC hợp nhất kiểm toán MBB năm 2025, 2026-08-30, xem chi tiết
trong PR/commit message đi kèm):
- 2 thuyết minh này chỉ có ĐẦY ĐỦ trong BCTC HỢP NHẤT ĐÃ KIỂM TOÁN/SOÁT XÉT (Quarter=5 theo quy ước
  CafeF) — báo cáo quý thường (không soát xét) mới rút gọn thuyết minh. Báo cáo BÁN NIÊN đã soát xét
  CÓ ĐẦY ĐỦ 2 bảng này y hệt báo cáo năm (verify ảnh chụp thật TCB "Tại 30/6/2026", 2026-08 — ban đầu
  từng nhầm loại hẳn báo cáo bán niên, xem fetch_bank_risk_gaps()). Dùng lại đúng cơ chế dò+tải PDF đã
  có ở bctc_pdf_tool.py (CafeF + 24hmoney, generic cho mọi ticker).
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
import difflib
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

from bctc_pdf_tool import fetch_cafef_list, fetch_24hmoney_list, select_ranked_reports, HEADERS

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
    # "Đ"/"đ" KHÔNG phải ký tự tổ hợp (không có dấu phụ NFD tách được) — là 1 CHỮ CÁI RIÊNG trong
    # bảng chữ cái Việt, nên vòng lặp lọc category "Mn" bên dưới bỏ sót, giữ nguyên "đ" thay vì rút
    # gọn về "d" như mọi chữ có dấu khác. Bug thật phát hiện 2026-09-18 khi thêm dò tiêu đề cột bảng
    # "Rủi ro tiền tệ": nhãn "Đô la Mỹ" (bắt đầu bằng "Đ") không khớp được với biến thể cố định viết
    # thường "do la my" dù rõ ràng cùng 1 từ — phải thay thủ công TRƯỚC khi NFD.
    s = (s or "").replace("Đ", "D").replace("đ", "d")
    s = unicodedata.normalize("NFD", s)
    return "".join(c for c in s if unicodedata.category(c) != "Mn").lower()


def _autorotate(img, pytesseract):
    """Dò và sửa hướng trang (90/180/270 độ) bằng OSD của tesseract TRƯỚC khi OCR chính — 1 vài trang
    BCTC (thường là bảng dàn ngang rộng — vd bảng khe hở lãi suất/thanh khoản) bị nhúng/scan LỘN NGƯỢC
    180 độ so với phần còn lại của tài liệu (dù các trang chữ xuôi khác trong CÙNG file vẫn đọc bình
    thường), khiến OCR đọc ra chuỗi ký tự vô nghĩa/lộn ngược thay vì số thật thay vì báo lỗi rõ ràng —
    bug thật phát hiện 2026-08 qua log CI TCB: 1 trang bảng gap đọc ra rác dạng "9S/ téE ¿LE S06...",
    verify bằng cách so chuỗi rác với chính footer trang (LUÔN có ở mọi trang) — "NH-q191/€S08" khi
    đọc ngược lại đúng là "B05a/TCTD-HN", tức trang bị lộn ngược chứ không phải OCR kém đơn thuần.
    Cần cài thêm gói apt "tesseract-ocr-osd" (dữ liệu ngôn ngữ riêng cho OSD, KHÁC "tesseract-ocr-vie")
    — nếu thiếu, image_to_osd() lỗi, hàm này lặng lẽ trả về ảnh gốc KHÔNG xoay (không đoán mù góc quay
    khi không chắc chắn — thà bỏ sót còn hơn xoay sai làm hỏng 1 trang vốn đã đúng chiều)."""
    try:
        from pytesseract import Output
        osd = pytesseract.image_to_osd(img, output_type=Output.DICT)
        angle = osd.get("rotate", 0)
        if angle:
            return img.rotate(-angle, expand=True)
    except Exception:
        pass
    return img


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
        img = _autorotate(img, pytesseract)
        return pytesseract.image_to_string(img, lang="vie")
    except Exception as e:
        print(f"  [WARN] OCR trang {page_index+1} loi: {e}")
        return ""


def _ocr_page_words(pdf_path, page_index, dpi=300):
    """OCR 1 trang, trả về list dict {"text","left","top","width","height"} — 1 phần tử/từ nhận diện
    được (tọa độ pixel, gốc trên-trái). None nếu thiếu pytesseract, [] nếu lỗi OCR/hết trang — KHÔNG
    BAO GIỜ raise. Dùng cho _extract_number_row_by_position() — xem docstring hàm đó lý do cần tọa độ
    thay vì text tuyến tính từ image_to_string()."""
    try:
        import pytesseract
        from pytesseract import Output
        import pypdfium2 as pdfium
    except ImportError:
        return None
    try:
        doc = pdfium.PdfDocument(pdf_path)
        if page_index >= len(doc):
            return []
        img = doc[page_index].render(scale=dpi / 72).to_pil()
        img = _autorotate(img, pytesseract)
        data = pytesseract.image_to_data(img, lang="vie", output_type=Output.DICT)
        words = []
        for i in range(len(data["text"])):
            t = (data["text"][i] or "").strip()
            if not t:
                continue
            words.append({"text": t, "left": data["left"][i], "top": data["top"][i],
                          "width": data["width"][i], "height": data["height"][i]})
        return words
    except Exception as e:
        print(f"  [WARN] OCR (toa do) trang {page_index+1} loi: {e}")
        return []


_HEADING_PHRASE_EN = {"lai_suat": "interest rate risk", "thanh_khoan": "liquidity risk",
                      "tien_te": "currency risk"}


def _find_note_pages(pdf_path, start_frac=0.70, max_pages=40):
    """Quét từ start_frac*tổng_số_trang tới hết tài liệu, OCR TỪNG TRANG (dừng ngay khi đã tìm đủ cả
    2 tiêu đề — không OCR speculative cả vùng). Trả về dict {"lai_suat": (page_idx, lang)|None,
    "thanh_khoan": (page_idx, lang)|None} (0-based; lang="vi"|"en") — CÓ THỂ rỗng/thiếu 1 trong 2 key
    nếu tài liệu này thật sự không có mục đó (vd BCTC quý không soát xét, rút gọn thuyết minh —
    không phải lỗi, gọi nơi dùng cần thử BẢN KHÁC). Trả về None (khác {} rỗng — phân biệt RÕ 2 tình
    huống) NẾU THIẾU pytesseract/tesseract-ocr binary — lúc đó dừng hẳn toàn bộ tính năng, thử bản
    khác cũng vô ích.

    SỬA (user 2026-09-17, sau khi xác nhận qua ảnh chụp trực tiếp OCB Quý 1+2/2025): một số ngân
    hàng (xác nhận OCB) công bố BCTC hoàn toàn bằng TIẾNG ANH tùy kỳ — không cố định 1 ngôn ngữ,
    KHÔNG đoán trước được kỳ nào sẽ là tiếng gì. Tìm kiếm CHỈ tiếng Việt trước đây khiến những kỳ
    tiếng Anh này LUÔN LUÔN thất bại dù bảng gap tồn tại rõ ràng — không phải lỗi tìm nhầm trang, mà
    là không tìm bằng đúng ngôn ngữ tài liệu. Giờ thử CẢ 2 ngôn ngữ cho mỗi trang, ghi nhận rõ ngôn
    ngữ nào khớp (dùng ở _extract_gaps_from_pdf để quyết định đơn vị tiền — bản tiếng Anh của OCB
    xác nhận dùng VND thực, không phải "Triệu đồng" như bản tiếng Việt chuẩn, xem ghi chú ở đó)."""
    try:
        import pypdfium2 as pdfium
    except ImportError:
        return None
    doc = pdfium.PdfDocument(pdf_path)
    total = len(doc)
    start = max(0, int(total * start_frac))
    pages_to_scan = list(range(start, total))[:max_pages]

    # Chỉ khớp khi cụm từ nằm trên 1 DÒNG NGẮN đứng riêng (kiểu tiêu đề mục, vd "45.1   Rủi ro lãi
    # suất") — KHÔNG khớp khi cụm từ chỉ xuất hiện giữa 1 câu văn xuôi dài (vd ghi chú tổng quan liệt
    # kê "...rủi ro tín dụng, rủi ro thanh khoản và rủi ro thị trường"). Bug thật phát hiện 2026-08 qua
    # log TCB Q2/2026: ghi chú "43. CHÍNH SÁCH QUẢN LÝ RỦI RO TÀI CHÍNH" liệt kê "rủi ro thanh khoản"
    # trong 1 câu tổng quan Ở TRANG SỚM HƠN 7 TRANG so với tiêu đề mục "45.3 Rủi ro thanh khoản" thật
    # — khớp nhầm câu văn xuôi này làm mốc trang, khiến vùng quét bảng số (+8 trang từ mốc) dừng lại
    # đúng 1 trang TRƯỚC khi tới được bảng thật. Ngưỡng 50 ký tự đủ rộng cho tiêu đề có số mục + tên
    # ("45.1   Rủi ro lãi suất" ~21 ký tự) nhưng đủ hẹp để loại câu văn xuôi (luôn dài hơn nhiều).
    _HEADING_MAX_LEN = 50
    # Mẫu B05a/TCTD-HN (BIDV xác nhận qua ảnh chụp thật 2026-08-31, rất có thể LPB/HDB/OCB — vốn
    # cùng thiếu dai dẳng cả 3 kỳ đã backfill — cũng dùng chung mẫu chuẩn NHNN này) KHÔNG dùng tiêu
    # đề số mục ngắn mà viết thành 1 câu dẫn ĐẦY ĐỦ đứng ngay TRÊN bảng, vd "Bảng sau trình bày rủi ro
    # thanh khoản của Ngân hàng tại ngày 30 tháng 6 năm 2025:" (~82 ký tự) — dài hơn hẳn ngưỡng 50 nên
    # bị bộ lọc trên loại bỏ dù ĐÚNG LÀ tiêu đề bảng thật (không phải câu văn xuôi tổng quan như bug
    # TCB). Nhận diện riêng mẫu câu này qua tiền tố "bang sau trinh bay" — CHỈ xuất hiện ở đúng câu dẫn
    # bảng thật (không lẫn với câu tổng quan liệt kê nhiều loại rủi ro như bug TCB), nên an toàn để bỏ
    # giới hạn độ dài cho riêng trường hợp này.
    _TABLE_CAPTION_PREFIX = "bang sau trinh bay"

    def _fuzzy_phrase_in_line(line_flat, phrase, min_ratio=0.82):
        """OCR đôi khi đọc sai 1-2 ký tự ngay trong CỤM TỪ TIÊU ĐỀ NGẮN (không phải lỗi cắt dòng/vị
        trí như 2 bug trên) — bug thật phát hiện 2026-08-31 qua BID Quý 3/2025: dùng easyocr đối
        chiếu (proxy cho tesseract, không cài được tesseract thật trong sandbox) đọc chính xác dòng
        "23.1. Rủi ro lãi suất" thành "23.1 Rủi ro lui suẩt" — dấu "ã" (thanh ngã) bị đọc nhầm thành
        "u", khiến so khớp CHÍNH XÁC "rui ro lai suat" thất bại hoàn toàn dù tiêu đề THẬT SỰ có mặt
        rõ ràng trên trang (verify bằng ảnh chụp trực tiếp trang PDF, không phải suy đoán) — đây
        chính là điều user cảnh báo trước (2026-08-31): khớp cứng 1 cụm từ chính xác rất dễ trượt vì
        OCR luôn có khả năng đọc sai vài ký tự. So khớp GẦN ĐÚNG (tỷ lệ giống ký tự, không cần khớp
        tuyệt đối) CHỈ cho bước TÌM TRANG này (không áp dụng cho bước trích SỐ LIỆU sau đó — nơi cần
        độ chính xác cao hơn, đã có 3 tầng dự phòng riêng) — ngưỡng 0.82 đủ để chấp nhận 1-2 ký tự sai
        lệch nhưng vẫn đủ hẹp để không khớp nhầm cụm từ khác."""
        words = line_flat.split()
        phrase_words = phrase.split()
        n = len(phrase_words)
        for i in range(len(words) - n + 1):
            window = " ".join(words[i:i + n])
            if difflib.SequenceMatcher(None, window, phrase).ratio() >= min_ratio:
                return True
        return False

    def _has_heading_line(text, phrase):
        for line in text.split("\n"):
            line_flat = _strip_accents(line).strip()
            if not _fuzzy_phrase_in_line(line_flat, phrase):
                continue
            if len(line_flat) <= _HEADING_MAX_LEN or line_flat.startswith(_TABLE_CAPTION_PREFIX):
                return True
        return False

    found = {}
    for idx in pages_to_scan:
        text = _ocr_page_text(pdf_path, idx)
        if text is None:
            return None  # thiếu pytesseract - dừng hẳn, không quét tiếp vô ích
        for key, phrase_vi in (("lai_suat", "rui ro lai suat"), ("thanh_khoan", "rui ro thanh khoan"),
                                ("tien_te", "rui ro tien te")):
            if key in found:
                continue
            if _has_heading_line(text, phrase_vi):
                found[key] = (idx, "vi")
            elif _has_heading_line(text, _HEADING_PHRASE_EN[key]):
                found[key] = (idx, "en")
        if "lai_suat" in found and "thanh_khoan" in found:
            break

    # "Rủi ro tiền tệ" (FX) là mục BỔ SUNG, không bắt buộc để coi tài liệu là "ok" (xem
    # _extract_gaps_from_pdf) — KHÔNG bắt vòng lặp chính ở trên phải quét hết cả tài liệu chỉ để tìm
    # riêng mục này (sẽ làm chậm toàn bộ pipeline cho đa số ngân hàng KHÔNG công bố mục này). Mục này
    # đã được kiểm tra MIỄN PHÍ (dùng lại text đã OCR) trên mọi trang vòng lặp trên vừa quét qua — chỉ
    # quét THÊM khi chưa thấy, và CHỈ khi đã xác định được ít nhất lãi suất/thanh khoản (FX luôn nằm
    # GẦN 1 trong 2 mục đó, xác nhận qua nhiều ảnh chụp thật BIDV/ACB/TCB, user 2026-09-18) — giới hạn
    # phạm vi quét thêm (+15 trang từ điểm vòng lặp chính vừa dừng) để chặn chi phí OCR phát sinh.
    if "tien_te" not in found and (found.get("lai_suat") or found.get("thanh_khoan")):
        extra_start = idx + 1
        for eidx in range(extra_start, min(extra_start + 15, total)):
            text = _ocr_page_text(pdf_path, eidx)
            if text is None:
                return None
            if _has_heading_line(text, "rui ro tien te"):
                found["tien_te"] = (eidx, "vi")
                break
            if _has_heading_line(text, _HEADING_PHRASE_EN["tien_te"]):
                found["tien_te"] = (eidx, "en")
                break
    return found


_ROW_LABEL_FLAT = {
    "lai_suat": "muc chenh nhay cam",
    # SỬA (user 2026-09-18, phát hiện qua ảnh chụp thật VIB Quý 1/2024): bỏ chữ "ròng" khỏi cụm nhãn
    # — VIB dùng "Mức chênh thanh khoản THUẦN" (đồng nghĩa "ròng" nhưng khác chữ), khiến khớp CHÍNH
    # XÁC "...rong" thất bại hoàn toàn dù dòng thật sự có mặt rõ ràng. "Mức chênh thanh khoản" (không
    # kèm tính từ) vẫn đủ riêng biệt để không khớp nhầm chỗ khác trên trang.
    "thanh_khoan": "muc chenh thanh khoan",
}
# Cặp từ neo (đã strip dấu) để định vị dòng gap theo TỌA ĐỘ — xem _extract_number_row_by_position().
# Chọn 2 từ khá riêng biệt trong cụm nhãn (không dùng "muc"/"chenh" vì quá phổ biến, dễ trùng chỗ
# khác trên trang) để giảm khớp nhầm: "nhạy"+"cảm" (rủi ro lãi suất). Rủi ro thanh khoản ĐỔI sang
# "thanh"+"khoản" (user 2026-09-18, cùng lý do đổi _ROW_LABEL_FLAT ở trên — "khoản"+"ròng" thất bại
# với VIB dùng "thuần") — hàm chọn khớp CUỐI CÙNG trên trang nên không lo trùng với chính tiêu đề mục
# ("Rủi ro thanh khoản" ở đầu trang, luôn đứng TRƯỚC dòng số liệu).
_ROW_ANCHOR_WORDS = {
    "lai_suat": ("nhay", "cam"),
    "thanh_khoan": ("thanh", "khoan"),
}
# Dòng "Tổng nợ phải trả" nằm NGAY TRÊN dòng "Mức chênh thanh khoản ròng" trong CÙNG bảng thanh khoản
# (đã verify ảnh chụp thật TCB) — trích thêm dòng này (tận dụng lại đúng text/tọa độ trang ĐÃ OCR cho
# dòng gap, không tốn thêm lượt OCR nào) để tính "Nợ phải trả ngắn hạn" phục vụ tỷ lệ Liquid
# Assets/Nợ phải trả ngắn hạn (bản ĐƠN GIẢN HÓA của LCR mà user yêu cầu — KHÁC LCR Basel chuẩn, chỉ
# cần tổng nợ phải trả theo kỳ hạn, không cần phân loại HQLA/outflow rate như LCR thật).
_ROW_LABEL_FLAT_LIAB = "tong no phai tra"
_ROW_ANCHOR_WORDS_LIAB = ("no", "phai")
# "Tổng tài sản" — dòng TỔNG của phần Tài sản, đứng NGAY TRÊN "Tổng nợ phải trả" trong cùng bảng, LUÔN
# in đậm/gạch chân đứng RIÊNG 1 dòng (kết thúc phần liệt kê từng khoản mục tài sản) — user (2026-09)
# chỉ ra đúng: đi tìm dòng "Mức chênh..." (nằm sát 1-2 dòng ghi chú/dòng Tổng nợ, dễ bị OCR gom nhầm
# — xem bug BID/STB 2025-Q1/Q2 đã fix qua guard trùng khớp) kém ổn định hơn 2 dòng TỔNG này, vốn tách
# biệt rõ ràng khỏi các dòng xung quanh bằng viền kẻ. Dùng làm phương án CUỐI: TỰ TÍNH gap = Tổng tài
# sản - Tổng nợ phải trả thay vì cố đọc đúng dòng "Mức chênh" đã có sẵn.
_ROW_LABEL_FLAT_ASSETS = "tong tai san"
_ROW_ANCHOR_WORDS_ASSETS = ("tong", "tai")
# Bản tiếng Anh (OCB xác nhận 2026-09-17) dùng "Total assets"/"Total liabilities" — GIỐNG HỆT vai trò
# 2 dòng trên nhưng viết tiếng Anh. Chỉ cần 2 dòng TỔNG này (không cần bản tiếng Anh của dòng "Mức
# chênh..." — tên dòng đó đổi khác nhau tùy ngân hàng, "Net liquidity gap"/"Total interest
# sensitivity gap"... kém ổn định hơn) vì lớp dự phòng thứ 4 (tự tính Tài sản - Nợ phải trả) đã đủ
# dùng cho CẢ 2 ngôn ngữ.
_ROW_LABEL_FLAT_LIAB_EN = "total liabilities"
_ROW_ANCHOR_WORDS_LIAB_EN = ("total", "liabilities")
_ROW_LABEL_FLAT_ASSETS_EN = "total assets"
_ROW_ANCHOR_WORDS_ASSETS_EN = ("total", "assets")

# 3 dòng TÀI SẢN/NỢ chi tiết trong CHÍNH bảng thanh khoản (user 2026-09-18, xác nhận qua ảnh chụp
# thật TCB — nằm SẴN trên trang đã OCR cho bảng gap, không cần tìm trang mới) — phục vụ 2 chỉ số theo
# "danh gia rui ro thanh khoan.docx": "tài sản gần tiền" ≤1 tháng (mục 9, dùng Tiền mặt+NHNN — CHƯA
# thêm Tiền gửi TCTD, xem lý do dưới) và cơ cấu kỳ hạn tiền gửi khách hàng (mục 12). CHỈ trích qua
# khớp CHỮ TUYẾN TÍNH (KHÔNG có tầng dự phòng theo tọa độ như 2 dòng Tổng) — đây là dữ liệu BỔ SUNG
# (không bắt buộc để có gap chính), chấp nhận tỷ lệ trích thành công thấp hơn thay vì rủi ro khớp
# nhầm. CỐ Ý bỏ dòng "Tiền gửi/cấp tín dụng cho các TCTD khác" (dù mục 9 có nhắc) — nhãn dòng NÀY quá
# giống dòng NỢ PHẢI TRẢ tương ứng ("Tiền gửi và vay các TCTD khác") ở nhiều ngân hàng, rủi ro khớp
# nhầm sang phía Nợ cao hơn giá trị dữ liệu thêm được — thà thiếu còn hơn sai.
_ROW_LABELS_CASH = ["tien mat"]
_ROW_LABELS_SBV_DEP = ["tien gui tai nhnn", "tien gui tai ngan hang nha nuoc"]
_ROW_LABELS_CUST_DEP = ["tien gui cua khach hang"]

_COMMA_NUM_RE = re.compile(r"\d{1,3}(?:,\d{3})+")
_PERIOD_NUM_RE = re.compile(r"\d{1,3}(?:\.\d{3})+")
# Ngân hàng thật KHÔNG BAO GIỜ có tổng nợ phải trả VƯỢT QUÁ mức này nếu số liệu THẬT SỰ đã ở đơn vị
# "triệu đồng" — ngân hàng lớn nhất Việt Nam (BIDV) cũng chỉ ~2,7 tỷ TRIỆU đồng (2,7 triệu tỷ đồng)
# — cách ngưỡng này hơn 100 lần. Ngược lại, CÙNG 1 ngân hàng biểu diễn bằng VND THỰC (không rút gọn)
# sẽ vượt ngưỡng này ít nhất 1.000.000 lần. Khoảng cách 2 tình huống rất xa (không có vùng xám), nên
# 1 ngưỡng đơn giản dựa vào ĐỘ LỚN SỐ THẬT đáng tin cậy hơn NHIỀU so với đoán qua chữ trên trang.
# CHỈ dùng làm PHƯƠNG ÁN DỰ PHÒNG khi không có bs_total_assets_ty (xem _detect_unit_divisor) — chỉ
# phân biệt được 2/4 trường hợp (VND thực vs còn lại), không phát hiện được "tỷ đồng"/"nghìn đồng".
_RAW_VND_MAGNITUDE_THRESHOLD = 1_000_000_000_000  # 1.000 tỷ (10^12) nếu hiểu lầm là "triệu đồng"


def _detect_unit_divisor(raw_sum, bs_total_assets_ty):
    """Xác định hệ số quy đổi về "triệu đồng" (quy ước lưu trữ chuẩn của hệ thống) — user (2026-09-18)
    lo ngại: nếu ngân hàng lớn báo cáo bằng triệu đồng còn ngân hàng nhỏ báo cáo bằng VND thực/nghìn
    đồng/tỷ đồng mà không quy đổi đồng bộ, TỔNG HỢP toàn hệ thống (cộng số tuyệt đối của nhiều ngân
    hàng) sẽ sai nghiêm trọng — 1 ngân hàng lệch đơn vị đủ làm méo cả tổng hệ thống.

    Có 4 khả năng, mỗi khả năng cách nhau ĐÚNG 1.000 lần so với quy ước "triệu đồng" đang dùng: tỷ
    đồng (raw NHỎ hơn 1.000 lần — cần NHÂN 1.000), triệu đồng (raw khớp — quy ước chuẩn, không đổi),
    nghìn đồng (raw LỚN hơn 1.000 lần — cần CHIA 1.000), VND thực (raw LỚN hơn 1.000.000 lần — cần
    CHIA 1.000.000). Ưu tiên so sánh trực tiếp với `bs_total_assets_ty` (tổng tài sản THẬT từ Vietcap,
    ĐỘC LẬP với OCR, luôn đáng tin) — phân biệt được CẢ 4 trường hợp vì khoảng cách quá xa nhau (không
    có vùng xám, sai lệch tự nhiên giữa "tổng nợ phải trả" trích được và "tổng tài sản" tham chiếu chỉ
    ~5-15%, không bao giờ nhầm sang bucket kề bên cách 1.000 lần). CHỈ rơi về ngưỡng cố định CŨ (chỉ
    phát hiện được VND thực) khi chưa có snapshot bảng cân đối (vd ngân hàng mới niêm yết)."""
    if raw_sum and bs_total_assets_ty:
        bs_reference_trieu = bs_total_assets_ty * 1000
        if bs_reference_trieu:
            ratio = raw_sum / bs_reference_trieu
            if ratio >= 31_623:
                return 1_000_000  # VND thuc
            if ratio >= 31.6:
                return 1_000  # nghin dong
            if ratio < 0.0316:
                return 0.001  # ty dong (chia 0.001 tuong duong nhan 1.000)
            return 1  # trieu dong - dung quy uoc, khong doi
    if raw_sum and raw_sum > _RAW_VND_MAGNITUDE_THRESHOLD:
        return 1_000_000
    return 1


def _detect_numfmt(text):
    """Xác định dấu phân cách nghìn THỰC SỰ dùng trên trang (phẩy kiểu Mỹ hay chấm kiểu Việt Nam) —
    đếm số lần mỗi kiểu xuất hiện, chọn kiểu PHỔ BIẾN HƠN. Việc này ĐỘC LẬP với ngôn ngữ nhãn dòng và
    đơn vị tiền — bug thật phát hiện 2026-09-18 qua VAB (nhãn dòng tiếng Việt nhưng số theo kiểu phẩy
    Mỹ, y hệt OCB tiếng Anh) chứng minh không thể gộp chung 3 quyết định (chữ nhãn/dấu tách số/đơn vị
    tiền) vào cùng 1 cờ ngôn ngữ như bản sửa trước (130aef4) đã làm."""
    if len(_COMMA_NUM_RE.findall(text)) > len(_PERIOD_NUM_RE.findall(text)):
        return "en"
    return "vi"


# Ô số: số VN chuẩn (dấu chấm phân cách nghìn, ngoặc = âm) HOẶC dấu gạch ngang đơn (= 0) HOẶC — dự
# phòng — 1 dãy ≥4 chữ số THUẦN không dấu chấm (OCR thỉnh thoảng làm mất dấu chấm ở 1 vài ô riêng lẻ,
# bug thật 2026-08: "34.144.640" bị đọc ra "34144640"). BẮT BUỘC dùng CHUNG 1 regex duy nhất (không
# tách "thử chặt trước, lỏng sau" như bản đầu) — thử tách riêng đã gây bug KHÁC: khi 1 ô giữa hàng bị
# mất dấu chấm, phần "chặt" vẫn đủ ĐẾM ra n_buckets token (vì tình cờ nhặt được cột "Tổng cộng" ở cuối
# hàng bù vào chỗ trống), lấy nhầm cột forecast/tổng thay cho ô lỗi — 1 regex duy nhất giữ đúng THỨ TỰ
# trái→phải nên ô mất dấu chấm được điền lại ĐÚNG VỊ TRÍ của nó, không bị cột sau đó nhảy vào thế chỗ.
_CELL_RE = re.compile(r"\(?-?[\d]{1,3}(?:\.[\d]{3})+\)?|(?<![\w.])-(?![\w.])|\(?[\d]{4,}\)?")
# Bản tiếng Anh (OCB xác nhận 2026-09-17) dùng dấu PHẨY phân cách nghìn (chuẩn US/UK, vd
# "5,962,953,436,151") — NGƯỢC với chuẩn Việt Nam (dấu chấm). Cần regex + hàm parse RIÊNG (không thể
# dùng chung 1 regex cho cả 2 dấu phân cách khác nhau) — chọn qua tham số `lang`, xác định từ ngôn
# ngữ của TIÊU ĐỀ mục đã tìm thấy (_find_note_pages), không đoán mù từ nội dung số.
_CELL_RE_EN = re.compile(r"\(?-?[\d]{1,3}(?:,[\d]{3})+\)?|(?<![\w,])-(?![\w,])|\(?[\d]{4,}\)?")


def _extract_cell_tokens(text, n_buckets, lang="vi"):
    """Trích đúng `n_buckets` "ô số" đầu tiên từ `text` theo ĐÚNG thứ tự trái→phải (xem _CELL_RE ở
    trên). Trả về None nếu không đủ số lượng — KHÔNG ĐOÁN số liệu thiếu."""
    cell_re = _CELL_RE if lang == "vi" else _CELL_RE_EN
    toks = [m.group(0) for m in cell_re.finditer(text)]
    return toks[:n_buckets] if len(toks) >= n_buckets else None


def _tokens_to_values(toks, lang="vi"):
    """Chuyển list token chuỗi (từ _extract_cell_tokens) sang list float — "-" = 0.0, ngoặc = âm,
    bỏ dấu phân cách nghìn (chấm cho tiếng Việt, phẩy cho tiếng Anh). Trả về None nếu có token không
    parse được (không nên xảy ra vì toks đã qua 2 regex ở trên, nhưng phòng hờ)."""
    sep = "." if lang == "vi" else ","
    vals = []
    for tok in toks:
        if tok == "-":
            vals.append(0.0)
            continue
        neg = tok.startswith("(") and tok.endswith(")")
        clean = tok.strip("()").replace(sep, "")
        try:
            v = float(clean)
        except ValueError:
            return None
        vals.append(-v if neg else v)
    return vals


def _extract_number_row(text, label_flat, n_buckets, lines_after=6, debug_tag=None, lang="vi"):
    """Tìm DÒNG có nhãn `label_flat` (đã strip dấu) trong `text`, trích các "ô số" trong nhãn đó +
    `lines_after` dòng kế tiếp. Chỉ trả về kết quả nếu số lượng ô trích được >= n_buckets (lấy đúng
    n_buckets ô ĐẦU TIÊN — cột "Tổng cộng" dư nếu có sẽ bị bỏ qua vì không cần). Trả về None nếu
    không tìm thấy nhãn hoặc thiếu số — KHÔNG ĐOÁN số liệu thiếu.

    Xử lý THEO DÒNG (không phải theo vị trí ký tự trong text gốc): tránh bug lệch index giữa bản đã
    strip dấu (`flat`, NFD rồi bỏ dấu tổ hợp LUÔN làm text NGẮN ĐI so với bản gốc) và text gốc — dùng
    vị trí ký tự trong `flat` để cắt `text` gốc là SAI vì 2 chuỗi lệch độ dài, sẽ cắt trúng đoạn không
    liên quan trên trang (bug thật, phát hiện 2026-08 khi xem lại log TCB: SKIP "OCR chất lượng kém"
    trong khi rất có thể nhãn+bảng đọc ĐÚNG nhưng bị cắt sai vị trí do bug này).

    "Ô số" chấp nhận CẢ dấu gạch ngang "-" (băng 0 — rất phổ biến trong thuyết minh BCTC cho cột
    không phát sinh giao dịch, vd cột "Quá hạn" của nhiều ngân hàng) — nếu chỉ khớp số thật sẽ làm
    lệch cột so với nhãn (dòng có 1 dấu "-" thì các cột sau đó bị đếm lùi 1 vị trí)."""
    lines = text.split("\n")
    flat_lines = [_strip_accents(l) for l in lines]
    # Nhãn có thể trải dài qua NHIỀU dòng (vd MBB thật: "Mức chênh nhạy cảm với lãi suất" rồi mới
    # đến "nội, ngoại bảng (5) = (3) + (4)" ở dòng sau) — ghép cửa sổ trượt 3 dòng liên tiếp để tìm.
    label_line_idx = None
    for i in range(len(lines)):
        joined = " ".join(flat_lines[i:i + 3])
        if label_flat in joined:
            label_line_idx = i  # lấy dòng ĐẦU tiên của cửa sổ khớp cuối cùng tìm được (label ổn định)
    if label_line_idx is None:
        if debug_tag:
            # In thêm các dòng có chứa "chenh" (từ khoá chung của MỌI biến thể nhãn dòng gap: "muc
            # chenh", "chenh lech"...) để biết ngay năm nay ngân hàng viết nhãn khác đi thế nào, thay
            # vì chỉ báo "không thấy" — tránh phải đoán mù rồi chờ 1 vòng chạy CI nữa mới biết (bug
            # thật, 2026-08: TCB dùng "nội bảng" thay vì "nội, ngoại bảng" cho nhãn lãi suất, phát
            # hiện được nhờ preview tương tự; nhãn thanh khoản có thể cũng đổi cách viết tương tự).
            candidates = [lines[i] for i in range(len(lines)) if "chenh" in flat_lines[i]]
            preview = " | ".join(c.strip() for c in candidates[:3])[:250] if candidates else \
                " ".join(lines).strip()[:200]
            print(f"  [DIAG] {debug_tag}: khong tim thay nhan '{label_flat}' tren trang nay. "
                  f"Dong co 'chenh' tren trang (neu co): \"{preview}\"")
        return None
    window_text = "\n".join(lines[label_line_idx:label_line_idx + 1 + lines_after])
    # Nhãn dòng gap LUÔN kèm công thức tham chiếu kiểu "(3) = (1) - (2)" hoặc "(5) = (3) + (4)" ngay
    # trong/cạnh chính nó — dấu "-" hoặc "+" ở đây là TOÁN TỬ, không phải ô dữ liệu, phải loại trước
    # khi tìm "ô số" (bug thật, 2026-08: dấu "-" của công thức bị đếm nhầm thành cột đầu tiên, làm
    # lệch toàn bộ các cột phía sau 1 vị trí).
    window_text = re.sub(r"\(\s*\d\s*\)\s*=\s*\(\s*\d\s*\)(?:\s*[+\-]\s*\(\s*\d\s*\))*", " ", window_text)
    toks = _extract_cell_tokens(window_text, n_buckets, lang=lang)
    if toks is None:
        if debug_tag:
            preview = window_text.replace("\n", " | ")[:300]
            print(f"  [DIAG] {debug_tag}: tim thay nhan nhung khong du so lieu. Cua so OCR: \"{preview}\"")
        return None
    return _tokens_to_values(toks, lang=lang)


def _extract_number_row_loose(text, n_buckets, debug_tag=None, lang="vi"):
    """Phương án CUỐI CÙNG khi cả _extract_number_row (khớp đúng cụm nhãn "mức chênh nhạy cảm.../mức
    chênh thanh khoản ròng") lẫn _extract_number_row_by_position (khớp đúng cụm nhãn + tọa độ) đều
    thất bại. User (2026-08-31) chỉ ra: bắt cứng đúng 1 cụm nhãn dòng SỐ LIỆU rất dễ trượt (ngân hàng
    viết tắt/khác cách hành văn 1 chút, hoặc OCR đọc lệch vài ký tự trong cụm dài) — trong khi MỤC
    LỚN (rủi ro lãi suất/rủi ro thanh khoản) đã được xác định CHẮC CHẮN từ bước tìm trang
    (_find_note_pages, khớp qua tiêu đề mục ổn định hơn nhiều). Nới lỏng: CHỈ cần dòng chứa "chênh"
    (từ khoá chung mọi biến thể: "mức chênh", "chênh lệch"...) rồi thử trích số trên MỌI dòng ứng
    viên — không đòi khớp đúng chữ nữa. CHỈ được gọi SAU KHI đã ở đúng trang mục cần tìm (page_idx
    từ _find_note_pages), giảm rủi ro khớp nhầm dòng "chênh lệch tỷ giá"/"chênh lệch đánh giá lại
    tài sản" ở mục khác trên cùng trang — vẫn có rủi ro khớp nhầm cao hơn 2 phương án chính xác ở
    trên (đây là lý do dùng SAU CÙNG, không thay thế), nên LUÔN in rõ dòng đã khớp để tự soát lại
    được qua log nếu nghi ngờ.

    Thử theo thứ tự TỪ DƯỚI LÊN (dòng xuất hiện sau trước) — nhất quán với _extract_number_row (ưu
    tiên dòng "nội, ngoại bảng" gộp, luôn nằm SAU dòng "nội bảng" riêng, nếu bảng có cả 2)."""
    lines = text.split("\n")
    flat_lines = [_strip_accents(l) for l in lines]
    candidate_idxs = [i for i in range(len(lines)) if "chenh" in flat_lines[i]]
    for idx in reversed(candidate_idxs):
        window_text = "\n".join(lines[idx:idx + 3])
        window_text = re.sub(r"\(\s*\d\s*\)\s*=\s*\(\s*\d\s*\)(?:\s*[+\-]\s*\(\s*\d\s*\))*", " ", window_text)
        toks = _extract_cell_tokens(window_text, n_buckets, lang=lang)
        if toks is not None:
            if debug_tag:
                print(f"  [DIAG] {debug_tag}: khop qua fallback noi long, dong \"{lines[idx].strip()[:80]}\"")
            return _tokens_to_values(toks, lang=lang)
    return None


def _extract_number_row_multi_label(text, label_candidates, n_buckets, lang="vi", debug_tag=None):
    """Thử LẦN LƯỢT từng cụm nhãn trong `label_candidates` qua _extract_number_row() (khớp tuyến
    tính, KHÔNG có tầng dự phòng theo tọa độ) — dùng cho các dòng CHI TIẾT bổ sung (Tiền mặt/Tiền gửi
    NHNN/Tiền gửi khách hàng, user 2026-09-18) mà cách viết nhãn khác nhau tùy ngân hàng/kỳ (vd
    "Tiền gửi tại NHNN" hay "Tiền gửi tại Ngân hàng Nhà nước"). Đây là dữ liệu BỔ SUNG (không bắt
    buộc để có gap chính) nên chấp nhận tỷ lệ trích thành công thấp hơn 3 tầng chính, không cố thêm
    tầng tọa độ (khó chọn từ neo đủ riêng biệt cho các nhãn này — dễ khớp nhầm sang dòng khác có từ
    tương tự, vd dòng nợ phải trả cũng nhắc "NHNN"/"các TCTD khác")."""
    for label in label_candidates:
        vals = _extract_number_row(text, label, n_buckets, lang=lang, debug_tag=debug_tag)
        if vals:
            return vals
    return None


def _extract_number_row_by_position(words, anchor_words, n_buckets, debug_tag=None, lang="vi"):
    """Định vị dòng gap bằng TỌA ĐỘ PIXEL thay vì thứ tự đọc tuyến tính của image_to_string() — BUG
    THẬT phát hiện 2026-08 qua log CI thật của TCB: với bảng rộng 9 cột trải hết bề ngang trang,
    tesseract đọc lộn xộn — nhãn dòng gap bị nối liền với footnote/tiêu đề cột nằm Ở VỊ TRÍ KHÁC trên
    trang thay vì 8 số thật nằm NGAY SAU nó theo hàng ngang. `_extract_number_row` (dựa vào
    image_to_string) chỉ đáng tin cho các đoạn văn xuôi hẹp, KHÔNG đáng tin cho bảng số rộng — hàm
    này thay thế bằng cách tự dựng lại đúng 1 HÀNG theo tọa độ:
    1. Tìm 2 "từ neo" (`anchor_words`, đã strip dấu, vd "nhay"+"cam") đứng gần nhau theo trục dọc
       (cùng 1 hàng) — xác nhận đây đúng là cụm nhãn cần tìm, không phải từ trùng ngẫu nhiên ở chỗ khác.
    2. Lấy tọa độ Y của từ neo ĐẦU làm mốc hàng, gom TẤT CẢ các từ khác trên trang có Y lệch trong
       ngưỡng chiều cao 1 dòng chữ (cùng hàng ngang thật, không phụ thuộc thứ tự OCR trả về).
    3. Sắp xếp lại các từ đó theo X (trái→phải) — đây chính là thứ tái tạo ĐÚNG thứ tự cột của bảng.
    4. Trích số VN từ chuỗi đã sắp xếp lại, giống `_extract_number_row`.
    Nếu có NHIỀU cặp neo khớp trên trang (nhãn dòng gap có thể lặp — VD (3) nội bảng và (5) nội+ngoại
    bảng), lấy cặp CUỐI (gần cuối trang hơn, giống lý do dùng rfind trước đây)."""
    a1, a2 = anchor_words
    flat = [(_strip_accents(w["text"]), w) for w in words]
    anchor = None
    for t1, w1 in flat:
        if t1 != a1:
            continue
        for t2, w2 in flat:
            if t2 == a2 and abs(w2["top"] - w1["top"]) < max(w1["height"], w2["height"], 1) * 1.5:
                anchor = w1  # ghi đè nếu tìm thấy khớp SAU (gần cuối trang hơn) — tương đương rfind
                break
    if anchor is None:
        if debug_tag:
            # Tương tự preview "chenh" ở _extract_number_row: in các từ OCR đọc được có chứa 1 phần
            # của từ neo (vd "khoan"/"rong" hoặc biến thể gần đúng) để biết OCR đọc nhãn thành gì thay
            # vì chỉ báo "không thấy" — không đoán mù, chờ preview này ở log lần chạy sau.
            near = [w["text"] for t, w in flat if a1[:3] in t or a2[:3] in t]
            preview = ", ".join(near[:10]) if near else "(khong co tu nao gan giong)"
            print(f"  [DIAG] {debug_tag}: (toa do) khong tim thay tu neo '{a1}'+'{a2}' gan nhau tren "
                  f"trang nay. Tu OCR gan giong: {preview}")
        return None
    row_top, row_h = anchor["top"], max(anchor["height"], 1)
    # Bảng trải rất rộng hết bề ngang trang (~1700px) — 1 vài ô ở xa neo có thể lệch Y vài pixel do
    # scan hơi nghiêng/rung nét, không có 1 ngưỡng dung sai duy nhất đúng cho mọi lần quét. Thử tăng
    # dần dung sai, DÙNG NGƯỠNG NHỎ NHẤT đã đủ đọc ra n_buckets ô để giảm rủi ro gom nhầm hàng bên
    # cạnh (bug thật, 2026-08: log TCB v1 thiếu 3/8 ô dù đã định vị đúng hàng neo).
    toks, row_text = None, ""
    for tol_mult in (1.2, 2.0, 3.0, 4.5):
        row_words = sorted((w for w in words if abs(w["top"] - row_top) < row_h * tol_mult),
                            key=lambda w: w["left"])
        row_text = " ".join(w["text"] for w in row_words)
        toks = _extract_cell_tokens(row_text, n_buckets, lang=lang)
        if toks is not None:
            break
    if toks is None:
        if debug_tag:
            preview = row_text[:300]
            print(f"  [DIAG] {debug_tag}: (toa do) tim thay hang neo nhung khong du so lieu du moi nguong "
                  f"dung sai da thu. Hang rong nhat: \"{preview}\"")
        return None
    return _tokens_to_values(toks, lang=lang)


_CURRENCY_ROW_RE = re.compile(
    r"\b(USD|VND|EUR|GBP|JPY|AUD|CHF|CNY|SGD|HKD)\b\s+([\d]+,[\d]+)\s*%\s*"
    r"(\(?-?[\d]{1,3}(?:\.[\d]{3})*\)?)\s+(\(?-?[\d]{1,3}(?:\.[\d]{3})*\)?)"
)


def _clean_vn_number(tok):
    if tok in ("-", ""):
        return 0.0
    neg = tok.startswith("(") and tok.endswith(")")
    v = float(tok.strip("()").replace(".", ""))
    return -v if neg else v


def _extract_rate_sensitivity_table(text):
    """Trích bảng "Độ nhạy đối với lãi suất" do CHÍNH ngân hàng công bố (không phải khe hở kỳ hạn) —
    dạng "USD  1,50%  (189.357)  (151.486)" mỗi dòng 1 loại tiền: % lãi suất tăng giả định, ảnh hưởng
    LNTT (triệu đồng), ảnh hưởng VCSH (triệu đồng). KHÔNG PHẢI mọi ngân hàng có mục này riêng (MBB
    2025 không có, TCB 2025 có ở "45.1 Rủi ro lãi suất" ngay trước bảng khe hở) — trả về {} nếu không
    tìm thấy dòng nào khớp, KHÔNG suy diễn. Đáng tin hơn NII sensitivity tự ước tính (static gap) vì
    là số ngân hàng tự tính có tính đến hành vi/giả định nội bộ, không phải xấp xỉ tuyến tính đơn giản.
    """
    result = {}
    for m in _CURRENCY_ROW_RE.finditer(text):
        ccy, pct_str, pbt_tok, equity_tok = m.groups()
        if ccy in result:
            continue
        try:
            pct = float(pct_str.replace(",", "."))
            pbt = _clean_vn_number(pbt_tok)
            equity = _clean_vn_number(equity_tok)
        except ValueError:
            continue
        result[ccy] = {"rate_increase_pct": pct, "pbt_impact": pbt, "equity_impact": equity}
    return result


def _is_half_year(name):
    n = _strip_accents(name)
    return "6 thang" in n or "ban nien" in n or "soat xet" in n


# Tháng kết thúc kỳ báo cáo theo Quarter (quy ước CafeF: 1-4 = quý thường, 5 = năm kiểm toán, 6 = bán
# niên đã soát xét) — dùng để so sánh ĐỘ MỚI THẬT giữa các loại kỳ khác nhau (vd Quý 3 (tháng 9) mới
# hơn bán niên (tháng 6) dù cùng 1 năm).
_QUARTER_END_MONTH = {1: 3, 2: 6, 3: 9, 4: 12, 5: 12, 6: 6}


# ── Rủi ro tiền tệ (FX Position) — user 2026-09-18 cung cấp 4 ảnh chụp thật (bảng chưa rõ tên NH,
# BIDV, ACB, TCB) làm mẫu. Bảng này KHÁC HẲN cấu trúc 2 bảng lãi suất/thanh khoản: CỘT là ĐỒNG TIỀN
# (USD/EUR/Vàng/...) thay vì kỳ hạn cố định — số lượng VÀ thứ tự cột khác nhau tùy ngân hàng (vd
# BIDV xếp EUR TRƯỚC USD, ACB có tới 7 đồng tiền kể cả Vàng/JPY/AUD/CAD) nên không thể dùng bucket
# cố định như INTEREST_RATE_BUCKETS/LIQUIDITY_BUCKETS — phải ĐỌC dòng tiêu đề cột để biết chính xác
# số lượng + thứ tự đồng tiền trước khi trích số.
#
# Mỗi mục (ma_chuan, [biến thể nhãn đã strip dấu]) — biến thể DÀI xếp trước biến thể NGẮN để tránh
# khớp nhầm khi 1 biến thể ngắn là tập con của 1 tên khác (vd "eur" là tập con của "euro duoc quy
# doi"). "OTHER"/"khac" LUÔN là cột "Ngoại tệ khác"/"Các loại ngoại tệ khác" gộp nhiều đồng tiền nhỏ
# ngân hàng không tách riêng, KHÔNG phải 1 đồng tiền cụ thể.
_FX_CCY_PATTERNS = [
    ("USD", ["do la my duoc quy doi", "do la my", "usd duoc quy doi", "usd"]),
    ("EUR", ["euro duoc quy doi", "euro", "eur duoc quy doi", "eur"]),
    ("GOLD", ["vang"]),
    ("JPY", ["yen nhat", "jpy"]),
    ("GBP", ["bang anh", "gbp"]),
    ("AUD", ["aud"]),
    ("CAD", ["cad"]),
    ("OTHER", ["cac loai ngoai te khac da qd", "cac loai ngoai te khac", "cac ngoai te khac duoc quy doi",
               "ngoai te khac", "khac"]),
]
_ROW_LABEL_FLAT_FX_ONBALANCE = "trang thai tien te noi bang"
_ROW_LABEL_FLAT_FX_OFFBALANCE = "trang thai tien te ngoai bang"
# ABB (anh chup thuc, user 2026-09-19) viet "noi ngoai bang" KHONG co dau phay - khac BIDV/TCB dung
# "noi, ngoai bang" co phay - thu ca 2 bien the (giong nhieu cach viet khac nhau da gap o cac nhan
# dong khac trong file nay, vd VIB "thuan"/"rong").
_ROW_LABEL_FLAT_FX_NET_CANDIDATES = ["trang thai tien te noi, ngoai bang", "trang thai tien te noi ngoai bang"]


def _find_fx_currency_header(text):
    """Dò dòng tiêu đề cột của bảng "Rủi ro tiền tệ" (vd "Đô la Mỹ | Euro | Vàng | Ngoại tệ khác |
    Tổng cộng") để xác định DANH SÁCH + THỨ TỰ đồng tiền THẬT SỰ xuất hiện trên trang này — số lượng
    và thứ tự khác nhau tùy ngân hàng nên không đoán trước được. Trả về list mã tiền theo đúng thứ
    tự trái->phải (KHÔNG gồm cột "Tổng cộng" — không cần, đã có tổng theo dòng ở nơi khác), hoặc None
    nếu không tìm được cửa sổ nào có >=2 tên đồng tiền khác nhau (ngưỡng >=2 để loại các dòng văn
    xuôi tình cờ chỉ nhắc 1 đồng tiền, vd "...quy đổi ra VNĐ theo tỷ giá USD...").

    BUG THẬT phát hiện 2026-09-19 (user cung cấp ảnh chụp thật ABB Quý 1/2026 — hệ thống vẫn báo
    "missing" dù bảng rõ ràng có mặt): tên cột "EUR được quy đổi"/"Các ngoại tệ khác được quy đổi"
    thường quá dài, bị NGẮT XUỐNG DÒNG trong chính ô tiêu đề (vd "EUR được" / "quy đổi" 2 dòng hiển
    thị) — OCR theo dải ngang (image_to_string đọc theo Y-coordinate) trả về ĐÚNG 2 dòng text tách
    biệt cho 2 "dải" đó (dòng 1 gộp phần đầu MỌI cột, dòng 2 gộp phần "quy đổi"/"khác được..." còn
    lại), khiến tìm trên TỪNG DÒNG RIÊNG LẺ chỉ khớp được các đồng tiền có tên KHÔNG bị ngắt (vd chỉ
    thấy "EUR"/"USD" ở dòng 1, bỏ sót hẳn "Ngoại tệ khác" vì chữ "khác" nằm ở dòng 2) — trích thiếu
    cột, làm SAI toàn bộ ánh xạ đồng tiền->giá trị. Giờ dùng CỬA SỔ TRƯỢT 3 dòng liên tiếp (giống
    kỹ thuật _extract_number_row dùng cho nhãn dòng bị ngắt) thay vì chỉ xét 1 dòng — thứ tự trái-
    phải vẫn đúng vì mỗi dòng thành phần đều giữ ĐÚNG thứ tự cột của chính nó, ghép nối tuần tự
    không làm xáo trộn.

    Chọn cửa sổ khớp NHIỀU đồng tiền nhất nếu có nhiều cửa sổ ứng viên (cửa sổ tiêu đề cột thật luôn
    liệt kê ĐẦY ĐỦ mọi đồng tiền của bảng, nhiều hơn hẳn bất kỳ câu văn xuôi nào tình cờ nhắc vài
    đồng tiền)."""
    lines = text.split("\n")
    flat_lines = [_strip_accents(l).strip() for l in lines]
    best = None
    for i in range(len(lines)):
        flat = " ".join(flat_lines[i:i + 3])
        if not flat or len(flat) > 250:
            continue
        hits = []
        for ccy, variants in _FX_CCY_PATTERNS:
            pos = None
            for v in variants:
                p = flat.find(v)
                if p != -1 and (pos is None or p < pos):
                    pos = p
            if pos is not None:
                hits.append((pos, ccy))
        if len(hits) >= 2:
            hits.sort()
            ccys = [c for _, c in hits]
            if best is None or len(ccys) > len(best):
                best = ccys
    return best


def _extract_fx_position(pdf_path, page_idx, unit_divisor=1):
    """Trích bảng "Rủi ro tiền tệ" bắt đầu từ `page_idx` (trang tiêu đề, từ _find_note_pages) — quét
    tối đa 6 trang kế tiếp (thường bảng số nằm ngay dưới tiêu đề, không xa như 2 bảng lãi suất/thanh
    khoản). Trả về dict {ma_tien: {"assets","liabilities","onbalance","offbalance","net"}} (đơn vị
    TRIỆU đồng, đã áp dụng `unit_divisor` — xem _detect_unit_divisor, nên truyền từ đơn vị ĐÃ xác
    định được của 2 bảng lãi suất/thanh khoản trong CÙNG tài liệu thay vì tự đoán lại — bảng FX chỉ
    là 1 tập con nhỏ của tổng tài sản nên so sánh độ lớn trực tiếp với tổng tài sản thật (cách
    _detect_unit_divisor dùng cho 2 bảng kia) KHÔNG đáng tin ở đây) hoặc None nếu không đọc đủ.

    Bắt buộc đọc được CẢ "Tổng tài sản" và "Tổng nợ phải trả" theo từng đồng tiền (2 dòng luôn có
    mặt, tách biệt rõ khỏi các dòng chi tiết xung quanh bằng viền kẻ — giống 2 dòng TỔNG dùng làm
    lớp dự phòng cho bảng thanh khoản/lãi suất). 3 dòng "Trạng thái tiền tệ nội bảng/ngoại bảng/nội,
    ngoại bảng" cố đọc trực tiếp trước, tự tính bù nếu thiếu 1 trong 3 (nội bảng = tài sản - nợ; nội,
    ngoại bảng = nội bảng + ngoại bảng nếu có, ngược lại lấy tạm bằng nội bảng)."""
    for p in range(page_idx, page_idx + 6):
        text = _ocr_page_text(pdf_path, p)
        if not text:
            continue
        ccys = _find_fx_currency_header(text)
        if not ccys:
            continue
        n = len(ccys)
        numfmt = _detect_numfmt(text)
        assets_vals = _extract_number_row(text, _ROW_LABEL_FLAT_ASSETS, n,
                                           debug_tag=f"fx_tong_tai_san trang {p+1}", lang=numfmt)
        liab_vals = _extract_number_row(text, _ROW_LABEL_FLAT_LIAB, n,
                                         debug_tag=f"fx_tong_no trang {p+1}", lang=numfmt)
        if not (assets_vals and liab_vals):
            continue  # co the day chua dung dong/du so - thu trang ke tiep
        onbalance_vals = _extract_number_row(text, _ROW_LABEL_FLAT_FX_ONBALANCE, n,
                                              debug_tag=f"fx_noi_bang trang {p+1}", lang=numfmt)
        offbalance_vals = _extract_number_row(text, _ROW_LABEL_FLAT_FX_OFFBALANCE, n,
                                               debug_tag=f"fx_ngoai_bang trang {p+1}", lang=numfmt)
        net_vals = _extract_number_row_multi_label(text, _ROW_LABEL_FLAT_FX_NET_CANDIDATES, n,
                                                    lang=numfmt, debug_tag=f"fx_noi_ngoai_bang trang {p+1}")
        if not onbalance_vals:
            onbalance_vals = [a - l for a, l in zip(assets_vals, liab_vals)]
        if not net_vals:
            net_vals = [o + b for o, b in zip(onbalance_vals, offbalance_vals)] if offbalance_vals \
                else list(onbalance_vals)
        if not offbalance_vals:
            offbalance_vals = [nv - ob for nv, ob in zip(net_vals, onbalance_vals)]
        for vals in (assets_vals, liab_vals, onbalance_vals, offbalance_vals, net_vals):
            if unit_divisor != 1:
                vals[:] = [v / unit_divisor for v in vals]
        result = {}
        for i, ccy in enumerate(ccys):
            result[ccy] = {
                "assets": assets_vals[i], "liabilities": liab_vals[i],
                "onbalance": onbalance_vals[i], "offbalance": offbalance_vals[i],
                "net": net_vals[i],
            }
        return result
    return None


def _extract_gaps_from_pdf(pdf_path, bs_total_assets_ty=None):
    """Định vị + trích 2 bảng gap từ 1 file PDF cụ thể đã tải sẵn. Trả về ("missing_tool", None) nếu
    thiếu pytesseract/tesseract-ocr binary (dừng hẳn, thử file khác cũng vô ích), ("no_note", None)
    nếu tài liệu này THẬT SỰ không có 1 trong 2 tiêu đề/không đọc đủ số (vd BCTC quý không soát xét,
    rút gọn thuyết minh — không phải lỗi, bên gọi nên thử bản BCTC khác), hoặc ("ok", partial_dict)
    với partial_dict chứa các key đã trích được trong {"interest_rate_gap", "liquidity_gap",
    "interest_rate_sensitivity_disclosed", "liabilities_by_bucket", "fx_position"} (key áp chót — tổng
    nợ phải trả theo kỳ hạn từ bảng thanh khoản — chỉ có nếu trích được, dùng tính Liquid Assets/Nợ
    phải trả ngắn hạn; "fx_position" — trạng thái ngoại tệ theo đồng tiền, xem _extract_fx_position —
    BỔ SUNG, không bắt buộc để tài liệu được coi là "ok", chỉ "interest_rate_gap"/"liquidity_gap" mới
    quyết định "ok" hay "no_note").

    `bs_total_assets_ty`: tổng tài sản THẬT (tỷ đồng, từ Vietcap, ĐỘC LẬP với OCR) của đúng ngân hàng/
    kỳ này nếu bên gọi có sẵn — dùng làm mốc để _detect_unit_divisor() nhận diện CHÍNH XÁC đơn vị tiền
    thật sự (tỷ/triệu/nghìn đồng/VND thực), thay vì chỉ đoán qua 1 ngưỡng cố định (user 2026-09-18)."""
    pages = _find_note_pages(pdf_path)
    if pages is None:
        return "missing_tool", None

    result = {}
    doc_unit_divisor = None  # dung chung cho bang FX phia duoi (xem _extract_fx_position) - bang FX
    # chi la 1 tap con nho cua tong tai san nen KHONG doi chieu duoc voi bs_total_assets_ty nhu 2
    # bang gap; dung lai don vi da xac dinh duoc tu 2 bang gap TRONG CUNG tai lieu dang tin hon.
    for key, bucket_list, n in (("lai_suat", INTEREST_RATE_BUCKETS, len(INTEREST_RATE_BUCKETS)),
                                 ("thanh_khoan", LIQUIDITY_BUCKETS, len(LIQUIDITY_BUCKETS))):
        found_key = pages.get(key)
        if found_key is None:
            continue
        page_idx, lang = found_key
        # Bảng số nằm CÁCH trang tiêu đề/phương pháp luận 1 SỐ trang KHÔNG CỐ ĐỊNH — verify thật: MBB
        # 2025 cách đúng 1 trang (heading trang 90 -> bảng trang 91), nhưng TCB 2025 cách 2 trang vì
        # có thêm 1 trang phụ "Độ nhạy đối với lãi suất" (bảng ảnh hưởng LNTT/VCSH theo % lãi suất
        # tăng — khác bảng khe hở lãi suất theo kỳ hạn) chen giữa (heading trang 92 -> bảng trang 94).
        # Mục thanh khoản TCB 2025 KHÔNG tìm thấy nhãn/từ neo trong +5 trang đầu (verify qua log CI
        # thật) — có thể còn nhiều trang phụ hơn xen giữa so với mục lãi suất; nới lên +7 trang, vẫn
        # DỪNG NGAY khi đọc đủ số, không OCR speculative quá xa.
        vals = None
        for p in range(page_idx, page_idx + 8):
            text = _ocr_page_text(pdf_path, p)
            if not text:
                continue
            # Bảng "Độ nhạy đối với lãi suất" (do NH tự công bố, USD/VND... % LS -> LNTT/VCSH) hay nằm
            # NGAY TRƯỚC bảng khe hở lãi suất trong cùng mục 45.1 — tranh thủ đọc luôn trên đường quét
            # tìm bảng khe hở, không tốn thêm lượt OCR trang nào. Không phải NH nào cũng công bố mục
            # này (MBB 2025 không có) nên chỉ ghi nếu tìm thấy, không bắt buộc.
            if key == "lai_suat" and "interest_rate_sensitivity_disclosed" not in result:
                sens = _extract_rate_sensitivity_table(text)
                if sens:
                    result["interest_rate_sensitivity_disclosed"] = sens
            # SỬA (user 2026-09-18, phát hiện qua VAB Quý 1/2026): ban đầu gắn định dạng số (dấu
            # phẩy/chấm phân cách nghìn) VÀ đơn vị tiền (VND thực/Triệu đồng) THEO NGÔN NGỮ tiêu đề
            # (lang) — SAI, vì VAB công bố bằng TIẾNG VIỆT nhưng số lại dùng dấu PHẨY (giống OCB tiếng
            # Anh). Định dạng tách số giờ đoán từ BẰNG CHỨNG TRỰC TIẾP trên trang (_detect_numfmt) —
            # độc lập hoàn toàn với `lang` (giờ CHỈ còn dùng để chọn CHỮ nhãn dòng cần tìm).
            #
            # SỬA THÊM (cùng ngày, phát hiện tiếp qua chính VAB): đoán ĐƠN VỊ TIỀN (VND thực hay Triệu
            # đồng) qua CHỮ "VNĐ" lặp lại trên trang CŨNG SAI — VAB in "VNĐ" lặp lại y hệt OCB nhưng
            # dữ liệu THỰC RA đã là "Triệu đồng" (đối chiếu Tổng tài sản suy ra ăn khớp với
            # quarterly_balance_sheet độc lập từ Vietcap: ~143.740 tỷ vs ~142.390 tỷ, hợp lý; nếu chia
            # thêm 1.000.000 sẽ sai lệch cả triệu lần). Đơn vị KHÔNG đoán được đáng tin từ chữ trên
            # trang — chỉ đáng tin từ ĐỘ LỚN của số liệu trích được (xem _RAW_VND_MAGNITUDE_THRESHOLD),
            # tính SAU khi đã trích được `liab_vals` phía dưới, không tính trước ở đây.
            numfmt = _detect_numfmt(text)
            wwords = None
            vals = None
            if lang == "vi":
                # 3 tang chinh xac cao chi thu duoc khi tieu de la TIENG VIET - nhan dong "Muc
                # chenh..." tieng Anh doi ten tuy ngan hang ("Net liquidity gap"/"Total interest
                # sensitivity gap"...) kem on dinh hon, nen ban tieng Anh bo qua thang 3 tang nay,
                # di thang toi lop du phong thu 4 (Tong tai san - Tong no phai tra, on dinh ca 2
                # ngon ngu) - vua nhanh hon (khong tom ocr toa do vo ich) vua tranh khop nham.
                vals = _extract_number_row(text, _ROW_LABEL_FLAT[key], n, debug_tag=f"{key} trang {p+1}", lang=numfmt)
                if not vals:
                    # Fallback theo TỌA ĐỘ (xem docstring _extract_number_row_by_position) — CHỈ chạy
                    # khi cách đọc tuyến tính ở trên thất bại, vì OCR theo tọa độ tốn thêm 1 lượt OCR
                    # trang (chậm hơn) — hầu hết trang không phải bảng gap sẽ bị loại ngay ở bước
                    # tuyến tính (rẻ) phía trên mà không cần OCR lại theo tọa độ.
                    wwords = _ocr_page_words(pdf_path, p)
                    if wwords:
                        vals = _extract_number_row_by_position(wwords, _ROW_ANCHOR_WORDS[key], n,
                                                                debug_tag=f"{key} trang {p+1}", lang=numfmt)
                if not vals:
                    # Lớp dự phòng thứ 3 (xem docstring _extract_number_row_loose) — chỉ chạy khi cả
                    # 2 phương án khớp đúng chữ ở trên đều thất bại, dùng LẠI đúng `text` đã OCR sẵn
                    # (không tốn thêm lượt OCR nào).
                    vals = _extract_number_row_loose(text, n, debug_tag=f"{key} trang {p+1}", lang=numfmt)
            # Tranh thủ trích luôn dòng "Tổng nợ phải trả"/"Total liabilities" trên CÙNG trang/text/
            # wwords đã OCR cho dòng gap (KHÔNG tốn thêm lượt OCR nào) — dùng để suy ra TÀI SẢN theo
            # bucket (= gap + nợ phải trả) cho CẢ 2 bảng: bảng thanh khoản -> Liquid Assets/Cumulative
            # Liquidity Gap Ratio; bảng lãi suất -> RSA (Rate Sensitive Assets) để tính RSA/RSL (user
            # 2026-08-31 yêu cầu bộ chỉ số rủi ro lãi suất đầy đủ, xem plan). Lưu 2 KEY KHÁC NHAU
            # (liabilities_by_bucket cho thanh_khoan giữ NGUYÊN tên cũ — tương thích dữ liệu đã
            # backfill; interest_rate_liabilities_by_bucket cho lai_suat là key MỚI) vì 2 bảng có cấu
            # trúc bucket khác nhau.
            # SỬA (user 2026-09-17): trích "Tổng nợ phải trả" LUÔN LUÔN (không chỉ khi `vals` đã có) —
            # cần sẵn cho lớp dự phòng thứ 4 ngay dưới đây (tự tính gap = Tổng tài sản - Tổng nợ phải
            # trả khi đọc thẳng dòng "Mức chênh..." thất bại, hoặc bản tiếng Anh luôn cần đường này).
            liab_result_key = "liabilities_by_bucket" if key == "thanh_khoan" else "interest_rate_liabilities_by_bucket"
            liab_label = _ROW_LABEL_FLAT_LIAB if lang == "vi" else _ROW_LABEL_FLAT_LIAB_EN
            liab_anchor = _ROW_ANCHOR_WORDS_LIAB if lang == "vi" else _ROW_ANCHOR_WORDS_LIAB_EN
            liab_vals = None
            if liab_result_key not in result:
                liab_vals = _extract_number_row(text, liab_label, n, debug_tag=f"no_phai_tra trang {p+1}", lang=numfmt)
                if not liab_vals:
                    if wwords is None:
                        wwords = _ocr_page_words(pdf_path, p)
                    if wwords:
                        liab_vals = _extract_number_row_by_position(wwords, liab_anchor, n,
                                                                     debug_tag=f"no_phai_tra trang {p+1}", lang=numfmt)
                if liab_vals and liab_vals == vals:
                    # Bug that phat hien 2026-08-31 (ABB/STB 2025-Q1): khi 2 dong "Tong no phai tra"
                    # va "Muc chenh..." nam qua gan nhau theo truc doc, buoc dung sai tang dan cua
                    # _extract_number_row_by_position co the gom NHAM dung dong gap lam dong no phai
                    # tra (7/7 gia tri trung khop tuyet doi voi dong gap — khong the la trung hop that
                    # voi so lieu ngan hang thuc). Coi nhu CHUA trich duoc, KHONG luu du lieu hong.
                    print(f"  [DIAG] no_phai_tra trang {p+1}: bo qua vi trung khop tuyet doi voi dong "
                          f"gap (nghi ngo gom nham hang)")
                    liab_vals = None
            # Doan don vi tien (ty/trieu/nghin dong/VND thuc) qua DO LON so lieu trich duoc DOI CHIEU
            # voi tong tai san THAT tu Vietcap (neu co) - xem _detect_unit_divisor(), KHONG qua chu
            # tren trang (da chung minh khong dang tin - xem ghi chu VAB o tren). Uu tien liab_vals
            # (luon co san, tinh truoc ca vals qua lop du phong thu 4); dung tam vals neu vi ly do nao
            # do liab_vals chua co (hiem).
            reference_sum = sum(abs(v) for v in liab_vals) if liab_vals else \
                (sum(abs(v) for v in vals) if vals else None)
            unit_divisor = _detect_unit_divisor(reference_sum, bs_total_assets_ty)
            if reference_sum:
                doc_unit_divisor = unit_divisor
            if not vals and liab_vals:
                # Lớp dự phòng THỨ 4 (user 2026-09-17 đề xuất, sau khi thấy dòng "Mức chênh..." của
                # BID/STB thất bại dù dòng nằm rõ ràng trên trang, VÀ là đường DUY NHẤT cho bản tiếng
                # Anh — xem nhánh lang=="en" ở trên): thay vì cố đọc đúng dòng TỔNG HỢP "Mức chênh..."
                # (thường nằm sát dòng "Tổng nợ phải trả" + 1 dòng ghi chú (*), dễ bị OCR gom nhầm —
                # xem bug BID/STB ở guard phía trên; tên dòng tiếng Anh cũng đổi khác nhau tùy ngân
                # hàng), đọc dòng "Tổng tài sản"/"Total assets" (TÁCH BIỆT rõ ràng khỏi các dòng xung
                # quanh — kết thúc phần liệt kê tài sản, luôn in đậm/gạch chân) rồi TỰ TÍNH gap = Tổng
                # tài sản - Tổng nợ phải trả. Đã có `liab_vals` sẵn (vừa trích ở trên) nên chỉ cần
                # trích thêm đúng 1 dòng "Tổng tài sản".
                assets_label = _ROW_LABEL_FLAT_ASSETS if lang == "vi" else _ROW_LABEL_FLAT_ASSETS_EN
                assets_anchor = _ROW_ANCHOR_WORDS_ASSETS if lang == "vi" else _ROW_ANCHOR_WORDS_ASSETS_EN
                assets_vals = _extract_number_row(text, assets_label, n, debug_tag=f"tong_tai_san trang {p+1}", lang=numfmt)
                if not assets_vals:
                    if wwords is None:
                        wwords = _ocr_page_words(pdf_path, p)
                    if wwords:
                        assets_vals = _extract_number_row_by_position(wwords, assets_anchor, n,
                                                                       debug_tag=f"tong_tai_san trang {p+1}", lang=numfmt)
                if assets_vals:
                    vals = [a - l for a, l in zip(assets_vals, liab_vals)]
                    print(f"  [DIAG] {key} trang {p+1}: tinh gap = Tong tai san - Tong no phai tra "
                          f"(lop du phong thu 4, khong doc truc tiep duoc dong Muc chenh)")
            if vals and liab_result_key not in result and liab_vals:
                if unit_divisor != 1:
                    liab_vals = [v / unit_divisor for v in liab_vals]
                result[liab_result_key] = dict(zip(bucket_list, liab_vals))
            # 3 dong chi tiet bo sung (xem _ROW_LABELS_CASH/...) — CHI trich tu bang THANH KHOAN (muc
            # 9/12 file huong dan la khai niem thanh khoan theo ky han, khac ban chat voi bang lai
            # suat tai dinh gia) — dung LAI chinh `text` da OCR cho dong gap, khong ton them OCR nao.
            if vals and key == "thanh_khoan":
                for result_key, labels in (("cash_by_bucket", _ROW_LABELS_CASH),
                                            ("sbv_dep_by_bucket", _ROW_LABELS_SBV_DEP),
                                            ("customer_deposits_by_bucket", _ROW_LABELS_CUST_DEP)):
                    if result_key in result:
                        continue
                    extra_vals = _extract_number_row_multi_label(text, labels, n, lang=numfmt,
                                                                  debug_tag=f"{result_key} trang {p+1}")
                    if extra_vals:
                        if unit_divisor != 1:
                            extra_vals = [v / unit_divisor for v in extra_vals]
                        result[result_key] = dict(zip(bucket_list, extra_vals))
            if vals:
                if unit_divisor != 1:
                    vals = [v / unit_divisor for v in vals]
                break
        if vals:
            gap_key = "interest_rate_gap" if key == "lai_suat" else "liquidity_gap"
            result[gap_key] = dict(zip(bucket_list, vals))

    # Rui ro tien te (FX) - BO SUNG, khong bat buoc de tai lieu duoc coi la "ok" (rat nhieu ngan
    # hang khong cong bo muc nay) - chi thu khi da tim thay trang tieu de (xem _find_note_pages).
    # Dung lai doc_unit_divisor da xac dinh tu 2 bang gap TRONG CUNG tai lieu (mac dinh 1 neu ca 2
    # bang gap kia deu khong thanh cong - hiem khi FX thanh cong ma ca 2 bang kia deu that bai).
    fx_page = pages.get("tien_te")
    if fx_page is not None:
        fx_page_idx, _fx_lang = fx_page
        fx_result = _extract_fx_position(pdf_path, fx_page_idx, unit_divisor=doc_unit_divisor or 1)
        if fx_result:
            result["fx_position"] = fx_result

    if "interest_rate_gap" not in result and "liquidity_gap" not in result:
        return "no_note", None
    # Danh dau tai lieu nay DA duoc thu voi bo dot OCR co nhan dien FX (bat ke co tim thay bang FX
    # thuc su hay khong - "khong tim thay" cung la 1 ket qua co gia tri, vd ngan hang khong cong bo
    # muc nay) - user 2026-09-19 phat hien qua ABB/ACB: nhieu ky da "reported" (du ca lai suat +
    # thanh khoan) TU TRUOC KHI co tinh nang FX nay, is_fully_reported() (chi xet lai suat+thanh
    # khoan) khien backfill_period() BO QUA NGAY nhung ky nay vi da coi la "xong", FX se KHONG BAO
    # GIO duoc thu du co bang thuc su ton tai trong BCTC. Co field rieng nay de bank_system_risk.py
    # phan biet duoc "chua tung thu FX" (can thu lai 1 lan) voi "da thu roi nhung xac nhan khong co"
    # (khong can thu lai vo ich moi lan backfill).
    result["fx_checked"] = True
    return "ok", result


def _download_report_pdf(ticker, cand):
    """Tải 1 báo cáo (dict từ select_ranked_reports) về cache, dùng lại nếu đã tải trước đó. Tên file
    cache phân biệt theo (Year, loại kỳ) — KHÔNG chỉ theo Year — để 1 báo cáo bán niên và báo cáo cả
    năm CÙNG NĂM (vd bán niên 2026 rồi cuối năm có thêm báo cáo năm 2026) không bị đè/dùng nhầm cache
    của nhau.

    BUG THẬT phát hiện 2026-09-17 (qua OCB Quý 1/2025 — user chụp ảnh chứng minh bảng gap TIẾNG VIỆT
    tồn tại rõ ràng, trong khi hệ thống báo "thiếu" mãi không sửa được): tên file cache TRƯỚC ĐÂY chỉ
    dựa vào (ticker, năm, loại kỳ) — KHÔNG phân biệt theo NGUỒN (cafef vs 24hmoney có thể là 2 tài
    liệu HOÀN TOÀN KHÁC NHAU cho cùng 1 kỳ, verify thật: candidate cafef của OCB là bản TIẾNG ANH 79
    trang, candidate 24hmoney là bản TIẾNG VIỆT 41 trang). fetch_bank_risk_gaps_for_period() thử
    candidate[0] thất bại rồi chuyển sang candidate[1] — nhưng vì CÙNG 1 tên file cache, bước tải
    candidate[1] thấy file "đã tồn tại" (từ candidate[0]) nên BỎ QUA TẢI LUÔN, tái sử dụng NHẦM đúng
    file candidate[0] đã thất bại — toàn bộ cơ chế "thử nguồn thay thế" (select_ranked_reports, thêm
    2026-08-31) bị VÔ HIỆU HÓA hoàn toàn bởi bug này bất cứ khi nào 2 nguồn thực sự khác nội dung.
    Fix: thêm 8 ký tự đầu mã băm MD5 của URL vào tên file — mỗi URL nguồn khác nhau chắc chắn có file
    cache RIÊNG, không còn đụng độ."""
    import hashlib
    q = cand.get("Quarter")
    if q in (5, 6):
        period_tag = "H1" if (q == 6 or _is_half_year(cand.get("Name", ""))) else "FY"
    else:
        period_tag = f"Q{q}"
    link_hash = hashlib.md5(cand["Link"].encode("utf-8")).hexdigest()[:8]
    pdf_path = os.path.join(CACHE_DIR, f"{ticker}_{cand['Year']}_{period_tag}_{link_hash}_CN_full.pdf")
    if os.path.exists(pdf_path):
        return pdf_path
    r = requests.get(cand["Link"].replace(" ", "%20"), headers=HEADERS, timeout=60)
    r.raise_for_status()
    with open(pdf_path, "wb") as f:
        f.write(r.content)
    return pdf_path


def _is_reviewed_candidate(cand):
    q = cand.get("Quarter")
    return q in (5, 6) or _is_half_year(cand.get("Name", ""))


def _period_key_for_candidate(cand):
    """Trả về "YYYY-Qn" (n=1-4) hoặc "YYYY-FY" cho MỌI báo cáo hợp nhất hợp lệ — KỂ CẢ báo cáo quý
    thường không soát xét — None nếu không map được vào 1 kỳ rõ ràng. Dùng LÀM KHÓA cho toàn bộ hệ
    thống lưu trữ theo kỳ (bank_alm_store.py) — CÙNG quy ước "YYYY-Qn" với dữ liệu bảng cân đối theo
    quý (không cần OCR), để 1 kỳ dù đến từ nguồn nào cũng gộp về đúng 1 khóa.

    SỬA LỚN 2026-08 (user cung cấp ảnh chụp thật báo cáo Quý 3/2025 CỦA CẢ MBB và TCB): trước đây
    hàm này chỉ coi báo cáo đã kiểm toán/soát xét là đáng tin cho 2 bảng khe hở lãi suất/thanh
    khoản, dựa trên giả định BCTC quý thường không có thuyết minh này. Giả định đó SAI — verify
    trực tiếp bằng cách tải THẬT 2 báo cáo Quý 3/2025 (KHÔNG soát xét) và render ẢNH từng trang
    (KHÔNG dùng pdfplumber.extract_text() — phương pháp cũ từng cho kết luận sai vì các trang bảng
    dạng này thường không trích xuất được bằng text, phải xem bằng ảnh/OCR): cả 2 ngân hàng đều CÓ
    ĐẦY ĐỦ mục "RỦI RO THỊ TRƯỜNG" (lãi suất/tiền tệ/thanh khoản) ngay trong báo cáo quý thường. Vì
    vậy giờ MỌI Quarter hợp lệ (1-6) đều là candidate khả dĩ — _extract_gaps_from_pdf tự trả về
    "no_note" an toàn nếu 1 báo cáo cụ thể thực sự không có, không giả định trước theo loại kỳ nữa."""
    q = cand.get("Quarter")
    name = cand.get("Name", "")
    is_half = (q == 6) or (q in (2, 5) and _is_half_year(name))
    if is_half:
        return f"{cand['Year']}-Q2"
    if q == 5:
        return f"{cand['Year']}-FY"
    if q in (1, 2, 3, 4):
        return f"{cand['Year']}-Q{q}"
    return None


def _select_candidate_reports(ticker):
    """Lấy + xếp hạng danh sách BCTC của 1 ticker — logic CHỌN dùng chung cho cả fetch_bank_risk_gaps
    (mới nhất), fetch_bank_risk_gaps_for_period (1 kỳ lịch sử cụ thể), và latest_reviewed_period
    (kiểm tra rẻ, không tải/OCR). Trả về (newest_overall, newest_reviewed, reviewed_reports):
    - newest_overall: BCTC MỚI NHẤT bất kể loại kỳ (None nếu không lấy được danh sách/không có gì).
    - newest_reviewed: BCTC MỚI NHẤT có period_key xác định được (tên biến giữ nguyên từ trước khi
      sửa 2026-08 — KHÔNG còn nghĩa "đã soát xét" nữa, giờ bao gồm CẢ báo cáo quý thường, xem
      _period_key_for_candidate).
    - reviewed_reports: list TẤT CẢ bản CÓ period_key xác định được (mọi Quarter 1-6), mỗi phần tử
      có thêm key "period_key", ĐÃ KHỬ TRÙNG theo period_key — khi 1 kỳ có CẢ bản đã kiểm toán/soát
      xét LẪN bản quý thường (hiếm, nhưng có thể xảy ra), ưu tiên giữ bản đã kiểm toán/soát xét
      (đáng tin cậy hơn — số liệu chính thức, ít khả năng sai lệch do soát xét lại); nếu cùng mức độ
      soát xét, ưu tiên Quarter==6 (gắn đúng từ đầu, không phải do 24hmoney gắn nhầm Quarter=5)."""
    try:
        items = fetch_cafef_list(ticker) + fetch_24hmoney_list(ticker)
    except Exception:
        return None, None, []
    # select_ranked_reports() (KHÔNG phải select_best_reports()) — GIỮ LẠI mọi ứng viên hợp lệ cho
    # mỗi (Year, Quarter) thay vì chỉ 1 bản thắng cuộc, để fetch_bank_risk_gaps_for_period() có bản
    # THAY THẾ nếu bản ưu tiên nhất tải/đọc thất bại (xem docstring select_ranked_reports — user
    # 2026-08-31 chứng minh bằng ảnh chụp thật ACB/BID/HDB rằng dữ liệu THỰC SỰ tồn tại dù trước đó
    # bị báo "thiếu" chỉ vì thử đúng 1 nguồn rồi bỏ cuộc).
    reports = [x for x in select_ranked_reports(items) if x.get("Quarter") in _QUARTER_END_MONTH]
    if not reports:
        return None, None, []

    def _recency_key(x):
        is_half = _is_half_year(x.get("Name", ""))
        # fetch_24hmoney_list()/_parse_24hmoney_period() gắn Quarter=5 CHO CẢ báo cáo bán niên đã
        # kiểm toán/soát xét (kỳ THẬT kết thúc tháng 6) lẫn báo cáo năm (kỳ kết thúc tháng 12) — nếu
        # cứ tra thẳng _QUARTER_END_MONTH[5]=12 cho case bán niên bị gắn nhầm này, nó sẽ trông "mới"
        # ngang báo cáo NĂM dù thật ra chỉ mới tới giữa năm, có thể lấn át 1 báo cáo quý 3 thật sự mới
        # hơn (bug phát hiện qua unit test khi mô phỏng kịch bản Quý 3 xuất hiện sau bán niên).
        end_month = 6 if (x["Quarter"] == 5 and is_half) else _QUARTER_END_MONTH[x["Quarter"]]
        reviewed = 1 if _is_reviewed_candidate(x) else 0
        return (x["Year"], end_month, reviewed)

    newest_overall = max(reports, key=_recency_key)

    def _cand_priority(x):
        # Độ ưu tiên trong SỐ CÁC BẢN CÙNG 1 KỲ (period_key) — bản đã soát xét/kiểm toán hơn bản
        # chưa, cùng mức soát xét thì Quarter==6 (gắn đúng từ đầu) hơn Quarter khác — giữ NGUYÊN
        # thứ tự ưu tiên cũ (trước đây dùng để CHỌN 1 bản duy nhất), giờ dùng để SẮP XẾP list cho
        # fetch_bank_risk_gaps_for_period() thử lần lượt.
        return (1 if _is_reviewed_candidate(x) else 0, 1 if x.get("Quarter") == 6 else 0)

    by_period = {}
    for cand in reports:
        pk = _period_key_for_candidate(cand)
        if pk is None:
            continue
        by_period.setdefault(pk, []).append(cand)

    reviewed_reports = []
    for pk, cands in by_period.items():
        cands.sort(key=_cand_priority, reverse=True)
        for cand in cands:
            tagged = dict(cand)
            tagged["period_key"] = pk
            reviewed_reports.append(tagged)
    if not reviewed_reports:
        return newest_overall, None, []
    newest_reviewed = max(reviewed_reports, key=_recency_key)
    return newest_overall, newest_reviewed, reviewed_reports


def latest_reviewed_period(ticker):
    """Kiểm tra RẺ (chỉ gọi API liệt kê danh sách BCTC, KHÔNG tải PDF/KHÔNG OCR — vài giây) xem kỳ
    MỚI NHẤT hiện có của ticker này là gì (vd "2026-Q3") — tên hàm giữ nguyên từ trước khi sửa
    2026-08 (không còn nghĩa "đã soát xét" nữa, xem _period_key_for_candidate: BCTC quý thường CŨNG
    có 2 bảng gap, verify thật qua ảnh chụp MBB/TCB). Dùng cho bước kiểm tra "có dữ liệu mới hơn dữ
    liệu đã lưu chưa" trước khi quyết định có cần OCR lại hay không (xem bank_system_risk.py). Trả
    về None nếu không lấy được danh sách BCTC hoặc không có bản nào map được vào 1 kỳ rõ ràng."""
    ticker = ticker.upper()
    _, newest_reviewed, _ = _select_candidate_reports(ticker)
    return newest_reviewed.get("period_key") if newest_reviewed else None


def _lookup_bs_total_assets_ty(ticker, period_key):
    """Tra tổng tài sản THẬT (tỷ đồng, từ Vietcap qua bank_alm_store.quarterly_balance_sheet — ĐỘC
    LẬP với OCR, luôn cập nhật sẵn hàng tuần bất kể pipeline gap) của đúng (ticker, period_key) —
    dùng làm mốc nhận diện đơn vị tiền cho _detect_unit_divisor() (user 2026-09-18). Trả về None nếu
    chưa có snapshot (vd ngân hàng mới niêm yết, hoặc period_key không map được sang kỳ quý) — bên
    gọi tự rơi về ngưỡng cố định cũ khi đó, KHÔNG BAO GIỜ raise (giống mọi hàm fetch khác trong file
    này — lỗi tra cứu phụ không được làm hỏng luồng OCR chính)."""
    try:
        import bank_alm_store
        qkey = bank_alm_store.gap_period_to_quarter(period_key)
        if not qkey:
            return None
        store = bank_alm_store.load_bank_store(ticker)
        snap = store.get("quarterly_balance_sheet", {}).get(qkey)
        return snap.get("total_assets") if snap else None
    except Exception:
        return None


def fetch_bank_risk_gaps(ticker):
    """Trả về dict {"interest_rate_gap": {bucket: value}, "liquidity_gap": {bucket: value},
    "source_title":, "source_url":, "fetched_year":} hoặc None nếu bất kỳ bước nào thất bại (thiếu
    tesseract, không tìm thấy BCTC nào, không định vị được note, OCR không đọc đủ số...). KHÔNG BAO
    GIỜ raise — template_banking.py gọi hàm này trong try/except nhưng bản thân hàm đã tự an toàn.

    SỬA 2026-08: trước đây giả định chỉ BCTC đã KIỂM TOÁN/SOÁT XÉT (Quarter=5 năm, Quarter=6 bán
    niên) mới có đủ 2 bảng gap, báo cáo quý thường bị coi là "rút gọn thuyết minh" — giả định này
    SAI, verify thật bằng ảnh chụp báo cáo Quý 3/2025 (không soát xét) của CẢ MBB và TCB: đều có đầy
    đủ mục "RỦI RO THỊ TRƯỜNG". Giờ MỌI loại kỳ đều là candidate hợp lệ (xem
    _period_key_for_candidate) — vẫn ưu tiên thử bản MỚI NHẤT theo kỳ dữ liệu thật sự trước, chỉ rơi
    về bản đã kiểm toán/soát xét gần nhất (đáng tin hơn nếu có nghi ngờ chất lượng OCR) khi bản mới
    nhất đó thật sự không đọc được.

    Chỉ lấy được kỳ MỚI NHẤT hiện có — xem fetch_bank_risk_gaps_for_period() để lấy 1 kỳ lịch sử cụ
    thể (dùng cho backfill hệ thống ngân hàng, bank_system_risk.py)."""
    ticker = ticker.upper()
    newest_overall, newest_reviewed, _ = _select_candidate_reports(ticker)
    if newest_overall is None:
        print("  [SKIP] Rui ro lai suat/thanh khoan: khong lay duoc danh sach BCTC hoac khong tim thay BCTC nao")
        return None
    if newest_reviewed is None:
        print("  [SKIP] Rui ro lai suat/thanh khoan: khong xac dinh duoc ky bao cao nao")
        return None

    # Thử bản MỚI NHẤT (bất kể loại kỳ) trước — thường TRÙNG với newest_reviewed (cùng 1 candidate,
    # vì giờ mọi loại kỳ đều hợp lệ), lúc đó chỉ tốn đúng 1 lượt thử. Chỉ thử THÊM bản kiểm toán/
    # soát xét khi bản mới nhất khác nó VÀ đọc không ra (hiếm — vd lỗi OCR trang bị lật/mờ).
    os.makedirs(CACHE_DIR, exist_ok=True)
    candidates_to_try = [newest_overall]
    if newest_reviewed["Link"] != newest_overall["Link"]:
        candidates_to_try.append(newest_reviewed)

    for cand in candidates_to_try:
        try:
            pdf_path = _download_report_pdf(ticker, cand)
        except Exception as e:
            print(f"  [WARN] Rui ro lai suat/thanh khoan: tai '{cand['Name']}' that bai ({e})")
            continue
        cand_period_key = cand.get("period_key") or _period_key_for_candidate(cand)
        bs_total_assets_ty = _lookup_bs_total_assets_ty(ticker, cand_period_key) if cand_period_key else None
        status, partial = _extract_gaps_from_pdf(pdf_path, bs_total_assets_ty=bs_total_assets_ty)
        if status == "missing_tool":
            print("  [SKIP] Rui ro lai suat/thanh khoan: thieu pytesseract/tesseract-ocr binary "
                  "(cai qua 'winget install UB-Mannheim.TesseractOCR' hoac 'apt install tesseract-ocr')")
            return None
        if status == "no_note":
            print(f"  [DIAG] Rui ro lai suat/thanh khoan: '{cand['Name']}' khong doc du 2 bang nay"
                  + (" - thu ban khac" if cand is newest_overall and len(candidates_to_try) > 1 else ""))
            continue
        partial["source_title"] = cand["Name"]
        partial["source_url"] = cand["Link"]
        partial["fetched_year"] = cand["Year"]
        # period_key ("YYYY-H1"/"YYYY-FY") — None nếu cand tình cờ là 1 báo cáo quý thường đọc được
        # (hiếm, không đúng quy ước lưu trữ per-period) — nơi gọi (fetch_bank_risk_gaps_cached) tự
        # bỏ qua việc lưu vào bank_alm_store khi None, không suy diễn kỳ.
        partial["period_key"] = cand.get("period_key") or _period_key_for_candidate(cand)
        return partial

    print("  [SKIP] Rui ro lai suat/thanh khoan: da thu (các) BCTC gan nhat, khong ban nao doc du so lieu "
          "(OCR chat luong kem hoac dinh dang bang khac chuan)")
    return None


def fetch_bank_risk_gaps_for_period(ticker, period_key):
    """Giống fetch_bank_risk_gaps() nhưng lấy ĐÚNG 1 KỲ LỊCH SỬ cụ thể (vd "2025-H1") thay vì luôn
    lấy kỳ mới nhất — dùng cho backfill lịch sử hệ thống ngân hàng (bank_system_risk.py). Trả về
    None nếu ticker này KHÔNG CÓ báo cáo đã kiểm toán/soát xét đúng kỳ đó (KHÔNG PHẢI lỗi — ngân
    hàng có thể chưa niêm yết lúc đó, hoặc CafeF/24hmoney không còn lưu bản cũ) — không thử "kỳ gần
    đúng nhất", chỉ khớp CHÍNH XÁC period_key hoặc trả về None, để backfill không tự đoán số liệu.
    KHÔNG BAO GIỜ raise, giống fetch_bank_risk_gaps().

    SỬA 2026-08-31 (user cung cấp ảnh chụp thật ACB/BID chứng minh dữ liệu tồn tại dù trước đó báo
    "thiếu"): _select_candidate_reports() giờ trả về NHIỀU ứng viên cho cùng 1 period_key (sắp xếp
    ưu tiên giảm dần) thay vì chỉ 1 — THỬ LẦN LƯỢT từng ứng viên (tải + đọc), dùng bản ĐẦU TIÊN
    thành công, chỉ bỏ cuộc khi TẤT CẢ đều thất bại. Lỗi tải (link hỏng/404) hay "no_note" (tài liệu
    này thật sự không đọc đủ 2 bảng) đều đáng thử bản khác; "missing_tool" (thiếu tesseract) là lỗi
    MÔI TRƯỜNG — thử bản khác cũng vô ích, dừng ngay.

    SỬA (user 2026-09-18): tra sẵn tổng tài sản THẬT (Vietcap, độc lập OCR) của đúng kỳ này, truyền
    xuống _extract_gaps_from_pdf() để nhận diện đơn vị tiền chính xác hơn (xem _detect_unit_divisor)
    — tránh ngân hàng nhỏ báo cáo bằng VND thực/tỷ/nghìn đồng làm SAI LỆCH khi tổng hợp toàn hệ thống
    cùng các ngân hàng khác đang đúng đơn vị triệu đồng."""
    ticker = ticker.upper()
    _, _, reviewed_reports = _select_candidate_reports(ticker)
    candidates = [c for c in reviewed_reports if c.get("period_key") == period_key]
    if not candidates:
        print(f"  [SKIP] Rui ro lai suat/thanh khoan ({period_key}): khong tim thay BCTC dung ky nay cho {ticker}")
        return None

    bs_total_assets_ty = _lookup_bs_total_assets_ty(ticker, period_key)
    os.makedirs(CACHE_DIR, exist_ok=True)
    for i, cand in enumerate(candidates):
        n_left = len(candidates) - i - 1
        try:
            pdf_path = _download_report_pdf(ticker, cand)
        except Exception as e:
            print(f"  [WARN] Rui ro lai suat/thanh khoan ({period_key}): tai '{cand['Name']}' that bai ({e})"
                  + (f" - thu ban thay the ({n_left} con lai)" if n_left else ""))
            continue
        status, partial = _extract_gaps_from_pdf(pdf_path, bs_total_assets_ty=bs_total_assets_ty)
        if status == "missing_tool":
            print("  [SKIP] Rui ro lai suat/thanh khoan: thieu pytesseract/tesseract-ocr binary "
                  "(cai qua 'winget install UB-Mannheim.TesseractOCR' hoac 'apt install tesseract-ocr')")
            return None
        if status == "no_note":
            print(f"  [DIAG] Rui ro lai suat/thanh khoan ({period_key}): '{cand['Name']}' khong doc du 2 bang"
                  + (f" - thu ban thay the ({n_left} con lai)" if n_left else ""))
            continue
        partial["source_title"] = cand["Name"]
        partial["source_url"] = cand["Link"]
        partial["fetched_year"] = cand["Year"]
        partial["period_key"] = period_key
        return partial
    return None


def fetch_bank_risk_gaps_cached(ticker):
    """Wrapper tái sử dụng dữ liệu GIỮA 2 pipeline: phân tích 1 mã lẻ (template_banking.py) và tổng
    hợp toàn hệ thống ngân hàng (bank_system_risk.py) — cả 2 nên gọi hàm này thay vì
    fetch_bank_risk_gaps() trực tiếp, để 1 lần OCR phục vụ được cả 2 nơi.

    Với ticker thuộc bank_universe.BANKING_TICKERS: kiểm tra RẺ (latest_reviewed_period, không OCR)
    xem kỳ mới nhất hiện có là gì; nếu data/bank_alm/<TICKER>.json ĐÃ CÓ đúng kỳ đó với
    status="reported" thì trả thẳng từ store (không tốn mạng/OCR); nếu không, gọi
    fetch_bank_risk_gaps() bình thường rồi LƯU LẠI kết quả vào store cho lần gọi sau (của chính
    pipeline này hoặc pipeline kia). Ticker ngoài bank_universe rơi thẳng về fetch_bank_risk_gaps()
    không qua store (không nên xảy ra vì module này chỉ được gọi cho cổ phiếu ngân hàng, nhưng
    phòng hờ). KHÔNG BAO GIỜ raise — lỗi đọc/ghi store bị nuốt, không làm hỏng kết quả OCR đã fetch
    được."""
    ticker = ticker.upper()
    try:
        from bank_universe import BANKING_TICKERS
    except Exception:
        BANKING_TICKERS = frozenset()
    if ticker not in BANKING_TICKERS:
        return fetch_bank_risk_gaps(ticker)

    import bank_alm_store
    cheap_period = latest_reviewed_period(ticker)
    if cheap_period:
        try:
            bank_alm_store.record_cheap_check(ticker, cheap_period)
            stored = bank_alm_store.get_period_entry(ticker, cheap_period)
        except Exception:
            stored = None
        if stored and stored.get("status") == "reported" and \
                (stored.get("interest_rate_gap") or stored.get("liquidity_gap")):
            result = {"period_key": cheap_period}
            for k in ("interest_rate_gap", "liquidity_gap", "liabilities_by_bucket",
                      "interest_rate_sensitivity_disclosed"):
                if stored.get(k) is not None:
                    result[k] = stored[k]
            src = stored.get("source") or {}
            result["source_title"] = src.get("title")
            result["source_url"] = src.get("url")
            result["fetched_year"] = src.get("fetched_year")
            return result

    partial = fetch_bank_risk_gaps(ticker)
    if partial and partial.get("period_key"):
        source = {"title": partial.get("source_title"), "url": partial.get("source_url"),
                  "fetched_year": partial.get("fetched_year")}
        try:
            bank_alm_store.upsert_reported_period(ticker, partial["period_key"], partial, source)
        except Exception:
            pass  # cache la "co thi tot" — loi ghi khong lam hong ket qua da fetch duoc
    return partial


# ── Tính chỉ số từ gap đã trích ──────────────────────────────────────────────────────────────────

def compute_interest_rate_risk_metrics(gap, total_assets, equity=None, disclosed_sensitivity=None,
                                        stress_bps=(100, 200), nii=None):
    """gap: dict theo INTEREST_RATE_BUCKETS. Trả về dict chỉ số — xem docstring module cho phương
    pháp (static gap, trọng số theo điểm giữa mỗi bucket, chỉ 4 bucket <=12 thang moi anh huong NII
    NAM NAY, bucket >1 nam khong lap lai trong nam danh gia).

    `equity` (VCSH, cùng đơn vị `total_assets`) — tuỳ chọn, nếu có sẽ thêm Gap/Equity (quy mô lệch kỳ
    hạn so với vốn chủ, KHÔNG phải so với tổng tài sản — 1 lệch nhỏ so với tổng tài sản có thể lớn so
    với vốn chủ vì đòn bẩy ngân hàng rất cao).
    `disclosed_sensitivity` — dict {ccy: {...}} từ _extract_rate_sensitivity_table() nếu ngân hàng
    CÓ tự công bố (đáng tin hơn nii_sensitivity_per_shock tự ước tính vì có tính hành vi/giả định nội
    bộ của ngân hàng) — chỉ đính kèm vào output, không dùng để ghi đè số tự tính.
    `stress_bps` — các mức sốc lãi suất (điểm cơ bản) muốn tính kịch bản, mặc định +100/+200bp.
    `nii` (Thu nhập lãi thuần thực tế kỳ gần nhất, cùng đơn vị `total_assets`) — tuỳ chọn, nếu có sẽ
    thêm stress_scenarios_nii_ratio (ΔNII/NII thực tế) — mức độ NGHIÊM TRỌNG của cú sốc so với chính
    dòng thu nhập lãi thuần của ngân hàng, đáng đọc hơn nhiều so với chỉ nhìn số tỷ VND tuyệt đối."""
    b = gap
    horizon_1y_keys = ["den_1_thang", "tu_1_3_thang", "tu_3_6_thang", "tu_6_12_thang"]
    weights = {"den_1_thang": 11.5 / 12, "tu_1_3_thang": 10 / 12, "tu_3_6_thang": 7.5 / 12, "tu_6_12_thang": 3 / 12}
    cum_1y = sum(b.get(k, 0.0) for k in horizon_1y_keys)
    weighted_gap = sum(b.get(k, 0.0) * weights[k] for k in horizon_1y_keys)
    ratio_1y = cum_1y / total_assets if total_assets else None
    gap_to_equity = cum_1y / equity if equity else None
    # Gap LŨY KẾ theo TỪNG mốc kỳ hạn (không chỉ điểm cuối <=12 tháng) — 1 gap <=12 tháng dương có thể
    # che giấu 1 gap âm rất lớn ở mốc <=1 tháng (xem [[feedback_bank_alm_risk_framework]] — nguyên tắc
    # "never sum all buckets into one number", phải đọc CẢ chuỗi lũy kế theo từng mốc).
    _horizon_keys_progressive = [
        ("1m", ["den_1_thang"]),
        ("3m", ["den_1_thang", "tu_1_3_thang"]),
        ("6m", ["den_1_thang", "tu_1_3_thang", "tu_3_6_thang"]),
        ("12m", horizon_1y_keys),
    ]
    cumulative_gap_by_horizon = {}
    for hz_label, hz_keys in _horizon_keys_progressive:
        hz_val = sum(b.get(k, 0.0) for k in hz_keys)
        cumulative_gap_by_horizon[hz_label] = {
            "value": hz_val,
            "ratio": (hz_val / total_assets if total_assets else None),
        }
    # "level"/"sensitivity_level" CHỈ là MỐC THAM CHIẾU để định hướng đọc số liệu (0-2% cân bằng, 2-5%
    # theo dõi, >5% đáng chú ý) — KHÔNG phải ngưỡng đạt/không đạt kiểu quy định (không có ngưỡng pháp
    # lý nào cho gap nội bộ ngân hàng). Diễn giải đúng đắn LUÔN cần đi kèm: gap này có thực sự ảnh
    # hưởng đáng kể đến NII/LNTT không (xem stress_scenarios_nii), so với biến động NIM lịch sử ra sao,
    # và xu hướng qua các kỳ (cần dữ liệu nhiều năm, hiện tại BCTC risk-note chỉ lấy được kỳ gần nhất).
    # Thang tham chiếu ĐÚNG như user hướng dẫn (0-2% khá cân bằng, 2-5% cần theo dõi, 5-10% khá lớn,
    # >10% đáng chú ý) — chỉ là "ngưỡng phân tích để so sánh các ngân hàng", KHÔNG phải ngưỡng pháp lý.
    if ratio_1y is None:
        level = "Không xác định"
    elif abs(ratio_1y) < 0.02:
        level = "Khá cân bằng"
    elif abs(ratio_1y) < 0.05:
        level = "Cần theo dõi"
    elif abs(ratio_1y) < 0.10:
        level = "Khá lớn"
    else:
        level = "Đáng chú ý"
    # Kich ban stress: +bps (lai suat tang -> gap am se AN, gap duong se LOI) va -bps (nguoc lai) cho
    # MOI muc trong stress_bps — cong thuc tuyen tinh (static gap) nen +200bp = 2x +100bp, van tinh
    # rieng tung muc (khong chi nhan 2) de neu sau nay doi sang mo hinh phi tuyen thi khong phai sua
    # ca cho goi.
    scenarios = {}
    for bps in stress_bps:
        scenarios[f"+{bps}bp"] = round(weighted_gap * (bps / 10000))
        scenarios[f"-{bps}bp"] = round(weighted_gap * (-bps / 10000))
    scenarios_ratio = ({k: (v / nii if nii else None) for k, v in scenarios.items()} if nii else None)
    return {
        "gap_by_bucket": b,
        "gap_ratio_by_bucket": {k: (v / total_assets if total_assets else None) for k, v in b.items()},
        "cumulative_gap_by_horizon": cumulative_gap_by_horizon,
        "cumulative_gap_1y": cum_1y,
        "cumulative_gap_1y_ratio": ratio_1y,
        "gap_to_equity_1y": gap_to_equity,
        "sensitive_type": "Asset-sensitive" if cum_1y > 0 else ("Liability-sensitive" if cum_1y < 0 else "Trung tính"),
        "nii_sensitivity_per_shock": scenarios.get("+100bp", round(weighted_gap * 0.01)),
        "shock_bps": 100,
        "stress_scenarios_nii": scenarios,
        "stress_scenarios_nii_ratio": scenarios_ratio,
        "sensitivity_level": level,
        "disclosed_sensitivity": disclosed_sensitivity or None,
        # Không tự tính ΔEVE (Economic Value of Equity): cần duration/thời lượng còn lại của TỪNG
        # bucket, trong khi thuyết minh BCTC VN chỉ công bố SỐ DƯ gap theo bucket, không có duration —
        # ước lượng duration sẽ là bịa số, không đáng tin hơn việc không tính (xem
        # [[feedback_bank_alm_risk_framework]]).
        "not_computed": ["ΔEVE (cần duration từng bucket, BCTC không công bố)"],
        # Rủi ro định tính KHÔNG thấy được từ riêng bảng gap số — chỉ là lưu ý đi kèm, không phải chỉ
        # số tính toán: (1) basis risk — tài sản/nguồn vốn có thể tham chiếu lãi suất KHÁC NHAU không
        # di chuyển cùng nhịp dù cùng 1 bucket kỳ hạn; (2) rủi ro hành vi — CASA ghi nhận kỳ hạn "không
        # nhạy cảm lãi suất"/qua đêm theo hợp đồng nhưng hành vi thực tế của khách hàng thường ổn định
        # (sticky) hơn nhiều so với kỳ hạn hợp đồng.
        "qualitative_caveats": [
            "Basis risk: tài sản/nguồn vốn có thể tham chiếu lãi suất khác nhau, không di chuyển "
            "cùng nhịp dù cùng 1 bucket kỳ hạn.",
            "Rủi ro hành vi: CASA/tiền gửi không kỳ hạn ghi nhận theo hợp đồng (qua đêm) nhưng hành "
            "vi thực tế thường ổn định (sticky) hơn nhiều.",
        ],
    }


def compute_liquidity_risk_metrics(gap, total_assets, equity=None, liquid_assets=None,
                                    customer_deposits=None, deposit_stress_pct=(5, 10, 20),
                                    liabilities_gap=None):
    """gap: dict theo LIQUIDITY_BUCKETS.

    `liquid_assets` (tiền mặt + tiền gửi NHNN + tiền gửi/cho vay TCTD — LẤY THẲNG từ dữ liệu BCTC đã
    fetch sẵn qua Vietcap trong template_banking.py, KHÔNG cần OCR thêm) — tuỳ chọn, nếu có sẽ tính
    Liquid Assets/Tổng tài sản, dùng làm "đạn" so sánh với kịch bản rút tiền gửi.
    `customer_deposits` — tuỳ chọn, để chạy stress test rút X% tiền gửi khách hàng (deposit run) so
    với `liquid_assets`: thiếu hụt = X%×tiền gửi − liquid_assets (dương = liquid assets KHÔNG đủ bù,
    cần huy động/vay thêm — đúng chuỗi nhân quả rủi ro thanh khoản → áp lực huy động bạn đã nêu).
    `liabilities_gap` — dict theo LIQUIDITY_BUCKETS từ dòng "Tổng nợ phải trả" (nếu trích được, xem
    _extract_gaps_from_pdf) — dùng tính Liquid Assets/Nợ phải trả ngắn hạn, bản ĐƠN GIẢN HÓA của LCR
    (KHÁC LCR Basel chuẩn — không phân loại HQLA/outflow rate — nhưng vẫn hữu ích để phân tích BCTC)."""
    b = gap
    cum_1m = b.get("den_1_thang", 0.0)
    cum_1y = cum_1m + b.get("tu_1_3_thang", 0.0) + b.get("tu_3_12_thang", 0.0)
    ratio_1m = cum_1m / total_assets if total_assets else None
    ratio_1y = cum_1y / total_assets if total_assets else None
    if ratio_1m is None:
        level = "Không xác định"
    elif ratio_1m > -0.03:
        level = "Thấp"
    elif ratio_1m > -0.08:
        level = "Trung bình"
    else:
        level = "Cao"
    # Gap lũy kế theo TỪNG mốc (không chỉ <=1 tháng) — "quá hạn" tính vào mốc <=1 tháng vì thực chất
    # còn cấp bách hơn (đã trễ hạn), khác cumulative_gap_1m ở trên (giữ nguyên định nghĩa cũ để không
    # đổi ý nghĩa của field đã dùng chỗ khác — field mới này bổ sung, không thay thế).
    _st_keys = ["qua_han_tren_3t", "qua_han_den_3t", "den_1_thang"]
    _horizon_keys_progressive = [
        ("1m", _st_keys),
        ("3m", _st_keys + ["tu_1_3_thang"]),
        ("12m", _st_keys + ["tu_1_3_thang", "tu_3_12_thang"]),
    ]
    cumulative_gap_by_horizon = {}
    for hz_label, hz_keys in _horizon_keys_progressive:
        hz_val = sum(b.get(k, 0.0) for k in hz_keys)
        cumulative_gap_by_horizon[hz_label] = {
            "value": hz_val,
            "ratio": (hz_val / total_assets if total_assets else None),
        }
    result = {
        "gap_by_bucket": b,
        "gap_ratio_by_bucket": {k: (v / total_assets if total_assets else None) for k, v in b.items()},
        "cumulative_gap_by_horizon": cumulative_gap_by_horizon,
        "cumulative_gap_1m": cum_1m,
        "cumulative_gap_1m_ratio": ratio_1m,
        "cumulative_gap_1y": cum_1y,
        "cumulative_gap_1y_ratio": ratio_1y,
        "gap_to_equity_1m": (cum_1m / equity if equity else None),
        "risk_level": level,
        "liquid_assets_ratio": (liquid_assets / total_assets if liquid_assets and total_assets else None),
        # Không tự tính LCR/NSFR chuẩn Basel: cần phân loại HQLA (tài sản thanh khoản chất lượng cao)
        # và nguồn vốn ổn định theo đúng trọng số quy định — BCTC VN không công bố đủ chi tiết để phân
        # loại đúng, một số ước lượng "gần đúng" sẽ SAI LỆCH và tự tin giả tạo hơn là không tính (xem
        # [[feedback_bank_alm_risk_framework]]). liquid_assets_to_st_liabilities_simple bên dưới LÀ 1
        # bản đơn giản hoá KHÁC (không cần phân loại HQLA/outflow rate), không mâu thuẫn với việc này.
        "not_computed": ["LCR chuẩn Basel (cần phân loại HQLA chi tiết)",
                          "NSFR (cần phân loại nguồn vốn ổn định chi tiết)"],
    }
    if liabilities_gap:
        short_term_liab = sum(liabilities_gap.get(k, 0.0) for k in _st_keys + ["tu_1_3_thang", "tu_3_12_thang"])
        result["short_term_liabilities"] = short_term_liab
        result["liquid_assets_to_st_liabilities_simple"] = (
            liquid_assets / short_term_liab if liquid_assets is not None and short_term_liab else None)
    if liquid_assets is not None and customer_deposits:
        result["deposit_run_stress"] = {
            f"-{p}%": round(customer_deposits * (p / 100) - liquid_assets) for p in deposit_stress_pct
        }
        # Ty le "dan" con lai sau stress — vd 65% nghia la liquid assets van con du 65% so voi phan
        # tien gui bi rut, KHONG phai "con thieu 35%" (do CON tai san khac ngoai liquid assets co the
        # xoay xo further, chi la khong "ngay lap tuc" bang liquid assets).
        result["deposit_run_coverage"] = {
            f"-{p}%": (liquid_assets / (customer_deposits * (p / 100)) if customer_deposits * (p / 100) else None)
            for p in deposit_stress_pct
        }
    return result


def compute_balance_sheet_alm_ratios(total_assets, loans, interbank_liab, bonds_issued):
    """2 tỷ lệ cơ cấu bảng cân đối phục vụ ALM tính THẲNG từ dữ liệu BCTC đã fetch sẵn qua Vietcap
    trong template_banking.py (loans_hist/interbank_hist/bonds_hist/total_assets_hist) — KHÔNG cần
    OCR thuyết minh, khác 2 hàm gap ở trên. Mẫu số dùng total_assets (= tổng nguồn vốn, vì
    Assets = Nợ phải trả + VCSH) thay vì tự cộng lại từng khoản mục nguồn vốn — tránh rủi ro
    thiếu/dư 1 khoản mục nào đó khi tính tổng nguồn vốn độc lập.
    - loan_to_assets: tài sản càng dồn vào cho vay (kém thanh khoản hơn) thì dư địa xoay xở càng ít.
    - wholesale_funding_ratio: phụ thuộc vốn liên NH/GTCG (nhạy cảm tâm lý thị trường) càng cao thì
      thanh khoản càng dễ biến động theo khẩu vị thị trường thay vì tiền gửi khách hàng ổn định."""
    return {
        "loan_to_assets": (loans / total_assets if total_assets else None),
        "wholesale_funding_ratio": (((interbank_liab or 0) + (bonds_issued or 0)) / total_assets
                                     if total_assets else None),
    }


def build_risk_narrative_lines(ir_metrics, liq_metrics, ticker, bs_ratios=None, ldr=None, casa=None,
                                nim_current=None):
    """Trả về LIST các câu tiếng Việt CÓ DẤU tóm tắt 2 rủi ro (+ cơ cấu bảng cân đối nếu có) — mỗi
    phần tử là 1 ý riêng, dùng để hiển thị dạng gạch đầu dòng trên web (đọc dễ hơn 1 đoạn văn dài);
    nối lại bằng " ".join() để ra bản văn xuôi cho PDF/narrative cũ. Không tự bịa số nào ngoài các
    dict đầu vào đã tính. Theo đúng phương pháp [[feedback_bank_alm_risk_framework]] — 5 tầng đánh
    giá rủi ro lãi suất (Repricing Gap/Assets, Cumulative Gap theo từng mốc, NII sensitivity, ΔEVE,
    NIM thực tế), và bộ chỉ số thanh khoản (LDR đúng cách, CASA, Liquid Assets/Assets, Liquid
    Assets/Nợ ngắn hạn (LCR đơn giản hoá), Liquidity Gap theo mốc, Wholesale Funding): đọc gap LŨY KẾ
    theo từng mốc (không chỉ điểm cuối), mức tham chiếu chỉ để định hướng chứ không phải ngưỡng
    đạt/không đạt, luôn đọc CẶP chỉ số cùng nhau, và nêu rõ những gì KHÔNG tính được thay vì bỏ qua
    im lặng.

    `ldr`, `casa` (%, 0-100) — LDR/CASA đã tính SẴN ở nơi khác trong template_banking.py (LDR theo
    đúng công thức tín dụng/huy động user đã hướng dẫn, KHÔNG phải Loans/Deposits đơn thuần — công
    thức đó sai vì không phản ánh đúng khác biệt cơ cấu nguồn vốn) — chỉ ĐÍNH KÈM vào đánh giá ALM để
    đọc cặp cùng Liquid Assets/Assets, KHÔNG tính lại.
    `nim_current` (%, kỳ gần nhất) — Tầng 5 của khung 5 tầng: Gap chỉ là mô hình lý thuyết, NIM thực
    tế mới là thứ THỰC SỰ xảy ra."""
    lines = []
    if ir_metrics:
        pct = ir_metrics["cumulative_gap_1y_ratio"]
        pct_s = f"{pct*100:+.2f}%" if pct is not None else "?"
        toe = ir_metrics.get("gap_to_equity_1y")
        toe_s = f", {toe*100:+.0f}% VCSH" if toe is not None else ""
        stype = ir_metrics["sensitive_type"]
        # Diễn giải TRỰC TIẾP ý nghĩa dấu gap (RSA-RSL) theo đúng nguyên lý user đưa ra — không chỉ
        # dán nhãn "Asset/Liability-sensitive" mà nói rõ NẾU lãi suất tăng/giảm thì NII bị ảnh hưởng
        # theo chiều nào, vì đây là phần "đáng tiền" nhất của repricing gap.
        if stype == "Asset-sensitive":
            direction = ("tài sản được định lại lãi suất NHANH HƠN nguồn vốn — lãi suất TĂNG có lợi "
                         "cho NII, lãi suất GIẢM bất lợi cho NII")
        elif stype == "Liability-sensitive":
            direction = ("chi phí vốn được định lại lãi suất NHANH HƠN tài sản — lãi suất TĂNG gây "
                         "áp lực lên NII/NIM, lãi suất GIẢM lại có lợi")
        else:
            direction = "gần như trung tính với biến động lãi suất"
        lines.append(
            f"Repricing Gap (Tầng 1+2): gap lũy kế ≤1 năm = {pct_s} tổng tài sản{toe_s} — Ngân hàng "
            f"đang {stype} ({direction}). Mức tham chiếu: {ir_metrics['sensitivity_level']} (thang so "
            f"sánh 0-2%/2-5%/5-10%/>10%, KHÔNG phải ngưỡng đạt/không đạt)."
        )
        hz = ir_metrics.get("cumulative_gap_by_horizon") or {}
        hz_parts = [f"≤{lbl}: {(hz.get(lbl) or {}).get('ratio')*100:+.2f}%"
                    for lbl in ("1m", "3m", "6m", "12m") if (hz.get(lbl) or {}).get("ratio") is not None]
        if hz_parts:
            lines.append(
                "Cumulative Gap theo từng mốc (không chỉ nhìn điểm cuối ≤12 tháng — 1 gap dương ở "
                "mốc cuối vẫn có thể che 1 gap âm lớn ở mốc gần hơn, cú sốc lãi suất xảy ra HÔM NAY "
                "chứ không phải sau 1 năm): " + "; ".join(hz_parts) + ".")
        sc = ir_metrics.get("stress_scenarios_nii") or {}
        sc_ratio = ir_metrics.get("stress_scenarios_nii_ratio")
        if sc:
            sc_s = "; ".join(
                f"{k}: {v:+,.0f} tỷ VND" + (f" ({sc_ratio[k]*100:+.1f}% NII)" if sc_ratio and sc_ratio.get(k) is not None else "")
                for k, v in sc.items())
            lines.append(f"NII sensitivity (Tầng 3, ước tính static gap): {sc_s}.")
        disclosed = ir_metrics.get("disclosed_sensitivity")
        if disclosed:
            parts = "; ".join(
                f"{ccy} +{v['rate_increase_pct']:.2f}%: LNTT {v['pbt_impact']:+,.0f} tỷ, "
                f"VCSH {v['equity_impact']:+,.0f} tỷ" for ccy, v in disclosed.items()
            )
            lines.append(f"Độ nhạy lãi suất do NGÂN HÀNG TỰ CÔNG BỐ (đáng tin cậy hơn ước tính static gap ở trên): {parts}.")
        if nim_current is not None:
            lines.append(
                f"NIM thực tế kỳ gần nhất (Tầng 5): {nim_current:.2f}% — đây là con số THỰC SỰ xảy "
                f"ra, cần theo dõi khi môi trường lãi suất thay đổi, trong khi Repricing Gap ở trên "
                f"chỉ là mô hình lý thuyết (static gap, chưa tính hành vi khách hàng/prepayment)."
            )
    if liq_metrics:
        r1m = liq_metrics["cumulative_gap_1m_ratio"]
        r1y = liq_metrics["cumulative_gap_1y_ratio"]
        r1m_s = f"{r1m*100:+.2f}%" if r1m is not None else "?"
        r1y_s = f"{r1y*100:+.2f}%" if r1y is not None else "?"
        la_ratio = liq_metrics.get("liquid_assets_ratio")
        la_s = f" Liquid Assets/Tổng TS = {la_ratio*100:.1f}%." if la_ratio is not None else ""
        lines.append(
            f"Liquidity Gap: gap ròng ≤1 tháng = {r1m_s} tổng tài sản, luỹ kế ≤1 năm = {r1y_s} (mức "
            f"tham chiếu: {liq_metrics['risk_level']} — chỉ là mốc định hướng).{la_s} Gap âm nghĩa là "
            f"ngân hàng đang có mismatch kỳ hạn (bình thường với mô hình ngân hàng — huy động ngắn "
            f"cho vay dài), quan trọng là có nguồn thanh khoản thay thế hay không — cần đọc CÙNG "
            f"Liquid Assets buffer ở trên, không đọc riêng lẻ."
        )
        st_ratio = liq_metrics.get("liquid_assets_to_st_liabilities_simple")
        if st_ratio is not None:
            lines.append(
                f"Liquid Assets/Nợ phải trả ngắn hạn (LCR đơn giản hoá, KHÁC LCR Basel chuẩn — không "
                f"phân loại HQLA/outflow rate nhưng vẫn hữu ích để tham khảo): {st_ratio*100:.1f}% — "
                f"tài sản thanh khoản ngay có thể cover được khoảng đó nghĩa vụ ngắn hạn."
            )
        stress = liq_metrics.get("deposit_run_stress")
        coverage = liq_metrics.get("deposit_run_coverage")
        if stress:
            parts = "; ".join(
                f"rút {k} tiền gửi KH: thiếu hụt {v:+,.0f} tỷ VND"
                + (f" (Liquid Assets che phủ {coverage[k]*100:.0f}%)" if coverage and coverage.get(k) is not None else "")
                for k, v in stress.items())
            lines.append(f"Stress test rút tiền gửi (bank run): {parts}.")
    if bs_ratios:
        la = bs_ratios.get("loan_to_assets")
        wf = bs_ratios.get("wholesale_funding_ratio")
        ldr_s = f", LDR = {ldr:.1f}%" if ldr is not None else ""
        casa_s = f", CASA = {casa:.1f}%" if casa is not None else ""
        if la is not None or wf is not None or ldr is not None or casa is not None:
            lines.append(
                f"Cơ cấu bảng cân đối (đọc CÙNG Liquid Assets/Tổng TS ở trên, không đọc riêng lẻ): "
                f"Cho vay/Tổng tài sản = {la*100:.1f}%{ldr_s}{casa_s}" + (
                    f", Wholesale Funding/Tổng tài sản = {wf*100:.1f}%." if wf is not None else "."
                )
            )
    not_computed = []
    for m in (ir_metrics, liq_metrics):
        if m and m.get("not_computed"):
            not_computed.extend(m["not_computed"])
    if not_computed:
        lines.append("KHÔNG tính (BCTC không công bố đủ chi tiết, ước lượng sẽ sai lệch và tự tin "
                      "giả tạo hơn là không tính): " + "; ".join(not_computed) + ".")
    caveats = []
    for m in (ir_metrics, liq_metrics):
        if m and m.get("qualitative_caveats"):
            for c in m["qualitative_caveats"]:
                if c not in caveats:
                    caveats.append(c)
    if caveats:
        lines.append("Lưu ý định tính (không thể hiện qua bảng gap số): " + " ".join(caveats))
    lines.append(
        "Các mức tham chiếu ở trên chỉ dựa trên 1 kỳ báo cáo gần nhất, KHÔNG phải ngưỡng đạt/không "
        "đạt cố định — mỗi ngân hàng có mô hình kinh doanh, tỷ trọng bán lẻ/doanh nghiệp, cơ cấu kỳ "
        "hạn khác nhau; đánh giá đầy đủ cần so sánh với peer cùng ngành và xu hướng qua nhiều kỳ."
    )
    return lines


def build_risk_narrative(ir_metrics, liq_metrics, ticker, bs_ratios=None):
    """Bản văn xuôi (nối các câu của build_risk_narrative_lines bằng khoảng trắng) — giữ lại cho
    tương thích ngược (PDF/JSON cũ dùng 1 chuỗi), dùng build_risk_narrative_lines() trực tiếp ở nơi
    cần hiển thị dạng danh sách (vd web)."""
    return " ".join(build_risk_narrative_lines(ir_metrics, liq_metrics, ticker, bs_ratios=bs_ratios))


def build_risk_summary(ir_metrics, liq_metrics):
    """Trả về dict {"interest_rate_level", "liquidity_level", "summary_text"} hoặc None — bản TÓM TẮT
    2-3 CÂU đặt Ở ĐẦU phần đánh giá ALM chi tiết (build_risk_narrative_lines) VÀ lặp lại trong khối
    "Nhận định nhanh Earning Release" của phần cập nhật quý — người đọc cần biết NGAY rủi ro lãi suất/
    thanh khoản có cao không và khả năng CHỐNG CHỊU biến cố ra sao trước khi đọc chi tiết từng
    tầng/chỉ số dài dòng bên dưới. Khả năng chống chịu dựa trên chính KẾT QUẢ STRESS TEST (NII sensitivity,
    deposit run coverage) — đáng tin hơn nhiều so với chỉ nhìn mức tham chiếu gap thô (1 gap "Cần theo
    dõi" có thể vẫn CHỐNG CHỊU TỐT nếu tác động NII thực tế nhỏ, xem [[feedback_bank_alm_risk_framework]]
    — nguyên tắc luôn đọc cặp Gap + mức độ ảnh hưởng thực tế cùng nhau). Không tự bịa số, chỉ diễn giải
    lại các dict đầu vào đã tính ở compute_interest_rate_risk_metrics/compute_liquidity_risk_metrics."""
    if not ir_metrics and not liq_metrics:
        return None
    parts = []
    ir_level = liq_level = None
    if ir_metrics:
        ir_level = ir_metrics["sensitivity_level"]
        gap_ratio = ir_metrics.get("cumulative_gap_1y_ratio")
        sc_ratio = ir_metrics.get("stress_scenarios_nii_ratio") or {}
        worst = max((abs(v) for v in sc_ratio.values() if v is not None), default=None)
        gap_s = f" (gap {gap_ratio*100:+.2f}% tổng tài sản)" if gap_ratio is not None else ""
        if worst is not None:
            if worst < 0.05:
                resil = "TỐT — NII ít biến động ngay cả khi lãi suất sốc mạnh"
            elif worst < 0.15:
                resil = "VỪA PHẢI — NII biến động đáng kể nhưng chưa ở mức nghiêm trọng"
            else:
                resil = "CẦN LƯU Ý — NII biến động mạnh khi lãi suất sốc"
            parts.append(f"Rủi ro lãi suất: mức tham chiếu {ir_level}{gap_s}, khả năng chống chịu {resil} "
                          f"(kịch bản sốc mạnh nhất làm NII đổi khoảng {worst*100:.1f}%).")
        else:
            parts.append(f"Rủi ro lãi suất: mức tham chiếu {ir_level}{gap_s}.")
    if liq_metrics:
        liq_level = liq_metrics["risk_level"]
        coverage = liq_metrics.get("deposit_run_coverage") or {}
        worst_cov = min((v for v in coverage.values() if v is not None), default=None)
        if worst_cov is not None:
            if worst_cov >= 1.0:
                resil = "TỐT — tài sản thanh khoản đủ bù ngay cả kịch bản rút tiền gửi mạnh nhất đã test"
            elif worst_cov >= 0.5:
                resil = "VỪA PHẢI — có thể cần huy động thêm nếu bị rút tiền gửi mạnh"
            else:
                resil = "CẦN LƯU Ý — tài sản thanh khoản không đủ bù kịch bản rút tiền gửi mạnh đã test"
            parts.append(f"Rủi ro thanh khoản: mức tham chiếu {liq_level}, khả năng chống chịu {resil} "
                          f"(che phủ {worst_cov*100:.0f}% ở kịch bản rút tiền gửi mạnh nhất).")
        else:
            parts.append(f"Rủi ro thanh khoản: mức tham chiếu {liq_level}.")
    return {
        "interest_rate_level": ir_level,
        "liquidity_level": liq_level,
        "summary_text": " ".join(parts),
    }
