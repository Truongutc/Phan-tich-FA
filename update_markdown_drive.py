#!/usr/bin/env python3
"""
update_markdown_drive.py — Nhận markdown BCTC người dùng dán thủ công (khi PDF quá xấu, OCR không
đọc đủ), tự dò kỳ báo cáo (quý/năm) từ nội dung, rồi upload lên đúng folder Google Drive Ngành/Mã
của ticker đó — CÙNG folder chứa Excel/PDF báo cáo, KHÔNG lưu vào git (BCTC_PDF/ không được commit).

Dùng bởi workflow GitHub Actions "Cập nhật Markdown BCTC" (2 ô nhập: mã cổ phiếu + markdown dán vào).
Có thể dán NHIỀU kỳ trong 1 lần chạy bằng marker "Start <kỳ>" ở đầu mỗi khối (vd. "Start 2025",
"Start Q2 2024", "Start 2024(CN)") — dòng "End ..." hoặc "---" ở cuối khối là tuỳ chọn, không bắt
buộc, chỉ để dễ đọc. Nếu không dùng marker (dán 1 khối trơn), hệ thống tự dò kỳ từ nội dung như cũ.

run_kcn_analysis() (template_kcn.py) sẽ tự tải các file này về từ Drive khi phát hiện kỳ tương ứng
vẫn thiếu/lệch sau khi đã thử tự động tải+OCR (xem check_segment_consistency).

Usage:
    python update_markdown_drive.py <TICKER>
    (nội dung markdown đọc từ biến môi trường MARKDOWN_CONTENT, hoặc từ stdin nếu không có)

Ví dụ nội dung nhiều kỳ:
    Start 2025
    ... (toàn bộ markdown BCTC năm 2025, thường có sẵn cột so sánh năm 2024) ...
    End 2025
    Start Q2 2024
    ... (markdown BCTC quý 2/2024) ...
    End Q2 2024

Nhãn LŨY KẾ (số đứng TRƯỚC q/Q, khác với "Q2 2024" là quý riêng lẻ):
    Start 2q2022    -> lũy kế Q1+Q2 gộp chung (không phải số riêng của Q2)
    Start 3q 2022   -> lũy kế Q1+Q2+Q3 gộp chung
Hệ thống (xem segments_kcn_parser.py::_derive_standalone_quarter) sẽ tự trừ các quý trước đã có
trong kho để suy ra đúng số của riêng quý đó trước khi lưu — mảng nào kho chưa đủ dữ liệu để trừ thì
bị bỏ qua (báo WARN) thay vì lưu nhầm số lũy kế thành số 1 quý.
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
_ROMAN_Q = {"I": 1, "II": 2, "III": 3, "IV": 4}

# Ô nhập "markdown" của workflow_dispatch trên giao diện web GitHub render thành Ô 1 DÒNG — dán nội
# dung nhiều dòng vào đó bị trình duyệt XOÁ SẠCH ký tự xuống dòng, gộp toàn bộ thành 1 dòng duy nhất
# (xác nhận thực tế qua log lỗi thật: "Start 2022  # 6.1 DOANH THU...  | Chỉ tiêu | Năm 2022 ..." toàn
# bộ dính liền). Vì vậy KHÔNG được dò marker dựa vào ranh giới đầu/cuối dòng (^/$) như trước — phải
# tìm nhãn kỳ với ĐỘ DÀI GIỚI HẠN RÕ RÀNG ngay sau "Start", nếu không phần capture (.+?)$ sẽ ăn lan
# hết toàn bộ phần còn lại của "dòng" (giờ là toàn bộ nội dung) làm nhãn, để lại thân khối rỗng.
_LABEL_PATTERN = (
    r"\d{4}\s*\(\s*[A-Za-z]{2}\s*\)"                                    # 2022(CN)
    r"|\d\s*[Qq]\s*\d{4}"                                               # 2q2022 / 3q 2022 (lũy kế)
    r"|\d{4}\s*[Qq]\s*[1-4]"                                            # 2024Q2
    r"|[Qq]\s*[1-4]\s*[/\s\-]*\s*\d{4}"                                 # Q2 2024 / q2/2024
    r"|qu[ýy]?\s*(?:I{1,3}|IV|[1-4])\s*[/\s\-]*\s*(?:n[ăa]m\s*)?\d{4}"  # quý 2 năm 2024
    r"|\d{4}"                                                           # 2022
)
_START_MARKER_RE = re.compile(
    r"\b(?:start|bắt\s*đầu|bat\s*dau)\b[:\-]?\s*(" + _LABEL_PATTERN + r")",
    re.IGNORECASE,
)
_END_MARKER_RE = re.compile(
    r"\b(?:end|kết\s*thúc|ket\s*thuc)\b[:\-]?\s*(?:" + _LABEL_PATTERN + r")?",
    re.IGNORECASE,
)


def _reflow_table_rows(text):
    """Khôi phục lại xuống dòng cho các HÀNG BẢNG markdown bị gộp thành 1 dòng (xem lý do ở
    _START_MARKER_RE). Trong markdown BCTC, mỗi hàng bảng dạng '| a | b | c |' — khi bị gộp dòng,
    ranh giới giữa 2 hàng liên tiếp luôn là '|' + ĐÚNG 2 khoảng trắng + '|' (khác với khoảng trắng
    đơn quanh dấu '|' phân cách CỘT trong cùng 1 hàng) — xác nhận qua log lỗi thật. Chèn lại xuống
    dòng tại đúng ranh giới đó để parse_markdown_tables() (segments_kcn_parser.py) nhận diện được
    bảng. An toàn để chạy cả khi nội dung đã có xuống dòng thật (không đổi gì thêm)."""
    return re.sub(r"\|(\s{2,})\|", "|\n|", text)


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
        qraw = m.group(1).upper()
        q = _ROMAN_Q.get(qraw) or int(qraw)
        return f"{m.group(2)}Q{q}"

    m = re.search(r"n[ăa]m\s*(\d{4})", md_content, re.IGNORECASE)
    if m and re.search(r"ki[eể]m\s*to[áa]n", md_content, re.IGNORECASE):
        return f"{m.group(1)}(CN)"
    return None


def parse_period_label(label):
    """Chuyển nhãn kỳ gõ tay ở dòng 'Start ...' (vd. '2025', '2025(CN)', 'Q2 2024', 'quý 2 năm 2024',
    '2024 Q2') thành (period_key, cum_quarters).

    cum_quarters=None  -> dữ liệu kỳ ĐƠN như thường lệ (đúng 1 quý riêng lẻ, hoặc cả năm).
    cum_quarters=N>1   -> dữ liệu LŨY KẾ N quý đầu năm (vd. 'Start 2q2022' = lũy kế Q1+Q2 gộp chung,
                          KHÔNG PHẢI số riêng của quý 2) — nhận biết bằng SỐ đứng TRƯỚC chữ q/Q
                          ('2q2022', '3q 2022'), khác với 'q2 2022'/'Q2 2024' (chữ q đứng trước số)
                          nghĩa là quý ĐƠN thứ 2 như bình thường. run_parse_and_merge() (xem
                          segments_kcn_parser.py) sẽ tự trừ các quý trước đã có trong kho để suy ra
                          đúng số của riêng quý N, tránh gán nhầm số lũy kế (gộp nhiều quý) thành số
                          của 1 quý — sẽ thổi phồng sai (vd. lũy kế 2 quý gán nhầm thành Q2 sẽ cộng
                          luôn cả Q1 vào).

    Trả về (None, None) nếu không hiểu được nhãn — khi đó main() sẽ tự dò kỳ trong nội dung khối như
    đường dự phòng."""
    label = label.strip()

    # LŨY KẾ: số đứng NGAY TRƯỚC q/Q (vd. '2q2022', '3q 2022') — PHẢI kiểm tra trước các pattern
    # "quý đơn" bên dưới để không bị nhầm.
    m = re.fullmatch(r"(\d)\s*[Qq]\s*(\d{4})", label)
    if m:
        n, year = int(m.group(1)), m.group(2)
        if 1 <= n <= 4:
            return f"{year}Q{n}", (n if n > 1 else None)

    m = re.fullmatch(r"(\d{4})\s*\(\s*CN\s*\)", label, re.IGNORECASE)
    if m:
        return f"{m.group(1)}(CN)", None
    m = re.fullmatch(r"(\d{4})\s*[Qq]\s*([1-4])", label)
    if m:
        return f"{m.group(1)}Q{m.group(2)}", None
    m = re.search(r"qu[ýy]?\s*(I{1,3}|IV|[1-4])\s*[/\s\-]*\s*(\d{4})", label, re.IGNORECASE)
    if not m:
        m = re.search(r"\bQ\s*([1-4])\s*[/\s\-]*\s*(\d{4})", label, re.IGNORECASE)
    if m:
        qraw = m.group(1).upper()
        q = _ROMAN_Q.get(qraw) or int(qraw)
        return f"{m.group(2)}Q{q}", None
    m = re.fullmatch(r"(\d{4})", label)
    if m:
        return f"{m.group(1)}(CN)", None
    m = re.search(r"(\d{4})", label)
    if m:
        return f"{m.group(1)}(CN)", None
    return None, None


def split_markdown_blocks(content):
    """Tách nội dung dán vào theo marker 'Start <kỳ>' — mỗi khối chạy từ sau 1 marker Start tới
    trước marker Start kế tiếp (hoặc hết nội dung). KHÔNG dựa vào ranh giới đầu/cuối dòng (xem lý do
    ở _START_MARKER_RE) nên vẫn hoạt động đúng dù nội dung bị gộp thành 1 dòng. Marker 'End ...'
    (nếu có) bị loại bỏ khỏi thân khối — chỉ mang tính trang trí, không bắt buộc, không cần khớp
    đúng nhãn với Start. Nếu KHÔNG có marker Start nào -> trả về 1 khối duy nhất (label=None) chứa
    toàn bộ nội dung, giữ tương thích ngược với cách dán 1-kỳ/1-lần-chạy trước đây."""
    starts = list(_START_MARKER_RE.finditer(content))
    if not starts:
        return [(None, _reflow_table_rows(content))]
    blocks = []
    for i, m in enumerate(starts):
        label = m.group(1).strip()
        body_start = m.end()
        body_end = starts[i + 1].start() if i + 1 < len(starts) else len(content)
        body = _END_MARKER_RE.sub("", content[body_start:body_end]).strip()
        body = _reflow_table_rows(body)
        blocks.append((label, body))
    return blocks


def upload_one_period(ticker, sector, period_key, content):
    normalized_period = period_key.replace("(", "_").replace(")", "")
    file_name = f"{ticker}_{normalized_period}_manual.md"
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = os.path.join(tmpdir, file_name)
        with open(tmp_path, "w", encoding="utf-8") as f:
            f.write(content)
        file_id, link = gdrive.upload_file(tmp_path, sector=sector, ticker=ticker)

    if file_id:
        print(f"[OK] Đã lưu {file_name} (kỳ {period_key}) lên Google Drive: {link}")
        return True
    print(f"[ERROR] Upload {file_name} (kỳ {period_key}) lên Google Drive thất bại — kiểm tra lại "
          "secret GDRIVE_SERVICE_ACCOUNT_JSON/GDRIVE_WEBAPP_URL.")
    return False


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

    blocks = split_markdown_blocks(md_content)
    if len(blocks) > 1:
        print(f"[INFO] Phát hiện {len(blocks)} khối (marker 'Start ...') trong nội dung dán vào.")

    any_ok = False
    for label, body in blocks:
        if not body.strip():
            print(f"[WARN] Bỏ qua khối '{label}' vì rỗng.")
            continue

        period_key, cum_n = (None, None)
        if label:
            period_key, cum_n = parse_period_label(label)
        if label and not period_key:
            print(f"[WARN] Không hiểu nhãn kỳ '{label}' ở dòng Start — thử tự dò trong nội dung khối...")
        if not period_key:
            period_key = detect_period_from_markdown(body)

        if not period_key:
            print(f"[ERROR] Không dò được kỳ báo cáo cho khối '{label or '(không có marker Start)'}'. "
                  "Hãy ghi rõ dòng 'Start <kỳ>' (vd. 'Start 2025' hoặc 'Start Q2 2024') ở đầu khối, "
                  "hoặc đảm bảo nội dung có câu 'ngày DD tháng MM năm YYYY'.")
            preview = body[:400].replace("\n", " ⏎ ")
            print(f"[DEBUG] Độ dài khối: {len(body)} ký tự. 400 ký tự đầu: {preview}")
            continue

        if cum_n and cum_n > 1:
            # Đánh dấu ngay trong nội dung (không phải tên file) để run_parse_and_merge() nhận biết
            # đây là số LŨY KẾ chứ không phải số riêng của kỳ — xem segments_kcn_parser.py.
            body = f"<!-- CUMULATIVE_QUARTERS: {cum_n} -->\n" + body
            print(f"[INFO] Khối '{label}' -> kỳ {period_key}, đánh dấu LŨY KẾ {cum_n} quý đầu năm "
                  f"(hệ thống sẽ tự trừ các quý trước đã có trong kho để suy ra đúng số quý này).")
        else:
            print(f"[INFO] Khối '{label or '(tự dò)'}' -> kỳ báo cáo: {period_key}")

        if upload_one_period(ticker, sector, period_key, body):
            any_ok = True

    if not any_ok:
        sys.exit(1)


if __name__ == "__main__":
    main()
