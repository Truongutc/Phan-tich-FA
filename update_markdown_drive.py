#!/usr/bin/env python3
"""
update_markdown_drive.py — Nhận markdown BCTC người dùng dán thủ công (khi PDF quá xấu, OCR không
đọc đủ), tự dò kỳ báo cáo (quý/năm) từ nội dung, rồi upload lên đúng folder Google Drive Ngành/Mã
của ticker đó — CÙNG folder chứa Excel/PDF báo cáo, KHÔNG lưu vào git (BCTC_PDF/ không được commit).

Dùng bởi workflow GitHub Actions "Cập nhật Markdown BCTC" (2 ô nhập: mã cổ phiếu + markdown dán vào,
mỗi lần chạy = 1 kỳ báo cáo). run_kcn_analysis() (template_kcn.py) sẽ tự tải các file này về từ Drive
khi phát hiện kỳ tương ứng vẫn thiếu/lệch sau khi đã thử tự động tải+OCR (xem check_segment_consistency).

Usage:
    python update_markdown_drive.py <TICKER>
    (nội dung markdown đọc từ biến môi trường MARKDOWN_CONTENT, hoặc từ stdin nếu không có)
"""
import os
import re
import sys
import json
import tempfile
from collections import Counter

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, PROJECT_ROOT)
import google_drive_uploader as gdrive

_MONTH_TO_Q = {3: 1, 6: 2, 9: 3, 12: 4}


def detect_period_from_markdown(md_content):
    """Dò period_key ('YYYY(CN)' hoặc 'YYYYQN') từ nội dung markdown BCTC — ưu tiên tìm ngày kết
    thúc kỳ dạng 'ngày DD tháng MM năm YYYY' (xuất hiện lặp lại nhiều lần trong MỌI BCTC ở đầu mỗi
    bảng số liệu — đáng tin cậy hơn nhiều so với chỉ tìm 'năm YYYY' đơn lẻ, vốn có thể là ngày ký/
    ngày phát hành chứ không phải kỳ báo cáo). Chuẩn hoá theo tháng kết thúc kỳ: 3->Q1, 6->Q2, 9->Q3;
    12 với ngày >=25 (gần cuối năm) -> báo cáo NĂM (CN) thay vì Q4 (vì BCTC quý 4 hiếm khi công bố
    riêng, thường gộp luôn vào báo cáo năm kiểm toán). Trả về None nếu không dò được kỳ nào — KHÔNG
    đoán bừa để tránh gán nhầm dữ liệu vào sai kỳ.

    Dùng lớp ký tự [àa]/[áa]/[ăa] thay vì chữ có dấu cố định vì markdown do OCR/paste thủ công đôi
    khi mất dấu tiếng Việt ("ngay"/"thang"/"nam" thay vì "ngày"/"tháng"/"năm") — chữ không dấu vẫn
    phải khớp được, nếu không toàn bộ việc dò kỳ sẽ fail dù nội dung ngày tháng vẫn đọc được bằng mắt.
    Có thêm fallback ngày dạng số "DD/MM/YYYY" hoặc "DD-MM-YYYY" cho các bảng không viết chữ.
    """
    dates = re.findall(
        r"ng[àa]y\s*(\d{1,2})\s*th[áa]ng\s*(\d{1,2})\s*n[ăa]m\s*(\d{4})",
        md_content, re.IGNORECASE,
    )
    if not dates:
        dates = re.findall(r"(\d{1,2})[/\.\-](\d{1,2})[/\.\-](\d{4})", md_content)
        # Lọc các match vô lý (không phải ngày/tháng thật) để tránh nhầm với số liệu tài chính
        # dạng "1.234.567" (dấu chấm phân cách nghìn) bị regex bắt nhầm thành ngày.
        dates = [(d, m, y) for d, m, y in dates if 1 <= int(d) <= 31 and 1 <= int(m) <= 12]

    if dates:
        # Ngày xuất hiện NHIỀU NHẤT trong văn bản đáng tin cậy hơn ngày đầu tiên gặp (ngày đầu có thể
        # là ngày ký/phát hành báo cáo, không phải ngày kết thúc kỳ kế toán).
        best = Counter((d, m, y) for d, m, y in dates).most_common(1)[0][0]
        day, month, year = int(best[0]), int(best[1]), int(best[2])
        if month == 12 and day >= 25:
            return f"{year}(CN)"
        if month in _MONTH_TO_Q:
            return f"{year}Q{_MONTH_TO_Q[month]}"

    # "Quý II/2025", "Quý 2 năm 2025", "Q2 2025"...
    m = re.search(
        r"qu[ýy]\s*(I{1,3}|IV|[1-4])\s*[/\-]?\s*(?:n[ăa]m)?\s*(\d{4})",
        md_content, re.IGNORECASE,
    )
    if m:
        roman = {"I": 1, "II": 2, "III": 3, "IV": 4}
        qraw = m.group(1).upper()
        q = roman.get(qraw) or int(qraw)
        return f"{m.group(2)}Q{q}"

    m = re.search(r"n[ăa]m\s*(\d{4})", md_content, re.IGNORECASE)
    if m and re.search(r"ki[eể]m\s*to[áa]n", md_content, re.IGNORECASE):
        return f"{m.group(1)}(CN)"
    return None


def main():
    if len(sys.argv) < 2:
        print("Usage: python update_markdown_drive.py <TICKER>")
        print("  Nội dung markdown đọc từ biến môi trường MARKDOWN_CONTENT hoặc stdin.")
        sys.exit(1)

    ticker = sys.argv[1].upper()
    md_content = os.environ.get("MARKDOWN_CONTENT")
    if not md_content:
        print("[INFO] Không có biến môi trường MARKDOWN_CONTENT — đang đọc từ stdin...")
        md_content = sys.stdin.read()
    if not md_content or not md_content.strip():
        print("[ERROR] Không có nội dung markdown để xử lý.")
        sys.exit(1)

    period_key = detect_period_from_markdown(md_content)
    if not period_key:
        print("[ERROR] Không dò được kỳ báo cáo (quý/năm) từ nội dung markdown. Hãy đảm bảo trong "
              "nội dung có câu dạng 'ngày DD tháng MM năm YYYY' (thường có sẵn ở đầu mọi bảng BCTC) "
              "hoặc 'năm YYYY (đã kiểm toán)' cho báo cáo năm.")
        preview = md_content[:400].replace("\n", " ⏎ ")
        print(f"[DEBUG] Độ dài nội dung nhận được: {len(md_content)} ký tự.")
        print(f"[DEBUG] 400 ký tự đầu tiên: {preview}")
        sys.exit(1)
    print(f"[INFO] Dò được kỳ báo cáo: {period_key}")

    # Sector lấy từ data/<TICKER>.json (đã có sau lần phân tích trước — xem save_json_kcn) để file
    # markdown lưu ĐÚNG folder Ngành/Mã trên Drive, khớp nơi Excel/PDF báo cáo đang được lưu.
    ticker_json_path = os.path.join(PROJECT_ROOT, "data", f"{ticker}.json")
    sector = None
    if os.path.exists(ticker_json_path):
        try:
            with open(ticker_json_path, "r", encoding="utf-8") as f:
                sector = json.load(f).get("sector")
        except Exception:
            pass
    if not sector:
        print(f"[WARN] Không tìm thấy data/{ticker}.json (chưa phân tích {ticker} lần nào) — hãy chạy "
              f"'Stock Analysis and Report Pipeline' cho {ticker} ít nhất 1 lần trước khi bổ sung markdown.")

    normalized_period = period_key.replace("(", "_").replace(")", "")
    file_name = f"{ticker}_{normalized_period}_manual.md"
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = os.path.join(tmpdir, file_name)
        with open(tmp_path, "w", encoding="utf-8") as f:
            f.write(md_content)
        file_id, link = gdrive.upload_file(tmp_path, sector=sector, ticker=ticker)

    if file_id:
        print(f"[OK] Đã lưu {file_name} (kỳ {period_key}) lên Google Drive: {link}")
    else:
        print("[ERROR] Upload lên Google Drive thất bại — kiểm tra lại secret "
              "GDRIVE_SERVICE_ACCOUNT_JSON/GDRIVE_WEBAPP_URL.")
        sys.exit(1)


if __name__ == "__main__":
    main()
