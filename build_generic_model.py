#!/usr/bin/env python3
"""
build_generic_model.py — Generic Stock Model and PDF Report Generator
"""
import os
import sys
import json
import math
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, black, white, grey
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                 TableStyle, PageBreak, KeepTogether)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY

# Adjust import path
sys.path.append(os.path.dirname(__file__))
from fetch_data import fetch_all, section_to_years, get_field_map
import google_drive_uploader

def build_generic(ticker):
    print(f"\n=== Running Automated Generic Analysis for {ticker} ===")
    
    # 1. Fetch Data
    try:
        data = fetch_all(ticker, use_cache=True)
    except Exception as e:
        print(f"Error fetching data for {ticker}: {e}")
        return False
        
    # Check details
    details_url = f"https://trading.vietcap.com.vn/api/iq-insight-service/v1/company/details?ticker={ticker}"
    try:
        r = google_drive_uploader.get_drive_service() # just checking if we can initialize to verify setup, but we make API request manually
        session = fetch_all.__globals__.get('_session')()
        det_r = session.get(details_url, timeout=10)
        det_data = det_r.json().get("data", {})
    except Exception as e:
        print(f"Could not fetch company details: {e}")
        det_data = {}
        
    company_name = det_data.get("enOrganName", det_data.get("organName", f"Công ty Cổ phần {ticker}"))
    current_price = det_data.get("currentPrice", 10000)
    shares = det_data.get("numberOfSharesMktCap", 1000000)
    market_cap = det_data.get("marketCap", current_price * shares)
    sector = det_data.get("sector", "General")
    
    is_recs = section_to_years(data, "INCOME_STATEMENT")
    bs_recs = section_to_years(data, "BALANCE_SHEET")
    cf_recs = section_to_years(data, "CASH_FLOW")
    
    # Determine years available
    available_years = sorted(list(set([r.get("yearReport") for r in is_recs if r.get("yearReport")])))
    if not available_years:
        print("No historical data years found!")
        return False
        
    # We want up to 5 historical years
    hist_years = available_years[-5:]
    print(f"Historical years: {hist_years}")
    
    # Helper to retrieve data safely
    def get_val(recs, year, fields):
        if not isinstance(fields, list):
            fields = [fields]
        for r in recs:
            if r.get("yearReport") == year:
                for f in fields:
                    val = r.get(f)
                    if val is not None:
                        return val
        return 0

    # Fields Mapping
    # Rev: isa3 or isa1, COGS: isa4 or isa2, GP: isa5, NPAT: isa22 or isa20, EPS: isa23
    rev_fields = ["isa3", "isa1"]
    cogs_fields = ["isa4", "isa2"]
    gp_fields = ["isa5"]
    npat_fields = ["isa22", "isa20"]
    eps_fields = ["isa23"]
    
    # BS Fields: Assets: bsa53 or bsa50, Equity: bsa78 or bsa75, Cash: bsa2 or bsa1
    assets_fields = ["bsa53", "bsa50"]
    equity_fields = ["bsa78", "bsa75"]
    cash_fields = ["bsa2", "bsa1"]
    debt_fields = ["bsa56", "bsa71"] # short + long debt
    
    # Extracted data arrays
    rev_hist = [get_val(is_recs, y, rev_fields) / 1e9 for y in hist_years]
    gp_hist = [get_val(is_recs, y, gp_fields) / 1e9 for y in hist_years]
    npat_hist = [get_val(is_recs, y, npat_fields) / 1e9 for y in hist_years]
    eps_hist = [get_val(is_recs, y, eps_fields) for y in hist_years]
    
    assets_hist = [get_val(bs_recs, y, assets_fields) / 1e9 for y in hist_years]
    equity_hist = [get_val(bs_recs, y, equity_fields) / 1e9 for y in hist_years]
    cash_hist = [get_val(bs_recs, y, cash_fields) / 1e9 for y in hist_years]
    debt_hist = [(get_val(bs_recs, y, "bsa56") + get_val(bs_recs, y, "bsa71")) / 1e9 for y in hist_years]
    
    # 2. Simple Forecast Model
    # Forecast next 3 years
    fc_years = [hist_years[-1] + 1, hist_years[-1] + 2, hist_years[-1] + 3]
    
    # Calculate simple growth averages
    rev_growths = []
    for i in range(1, len(rev_hist)):
        if rev_hist[i-1] > 0:
            rev_growths.append((rev_hist[i] / rev_hist[i-1]) - 1)
    avg_rev_growth = sum(rev_growths) / len(rev_growths) if rev_growths else 0.08
    # Clamp growth to realistic [5%, 15%]
    avg_rev_growth = max(0.05, min(0.15, avg_rev_growth))
    
    # Average GP Margin
    gp_margins = [gp_hist[i] / rev_hist[i] if rev_hist[i] > 0 else 0 for i in range(len(rev_hist))]
    avg_gp_margin = sum(gp_margins) / len(gp_margins) if gp_margins else 0.15
    
    # Average NP Margin
    np_margins = [npat_hist[i] / rev_hist[i] if rev_hist[i] > 0 else 0 for i in range(len(rev_hist))]
    avg_np_margin = sum(np_margins) / len(np_margins) if np_margins else 0.05
    
    # Forecast projections
    rev_fc = []
    gp_fc = []
    npat_fc = []
    eps_fc = []
    equity_fc = []
    
    last_rev = rev_hist[-1]
    last_equity = equity_hist[-1]
    
    for y in fc_years:
        next_rev = last_rev * (1 + avg_rev_growth)
        next_gp = next_rev * avg_gp_margin
        next_npat = next_rev * avg_np_margin
        # Basic equity expansion: Retain 70% of earnings
        next_equity = last_equity + (next_npat * 0.7)
        
        rev_fc.append(next_rev)
        gp_fc.append(next_gp)
        npat_fc.append(next_npat)
        equity_fc.append(next_equity)
        
        # Estimate EPS assuming constant share count
        est_eps = (next_npat * 1e9) / (shares if shares else 1)
        eps_fc.append(est_eps)
        
        last_rev = next_rev
        last_equity = next_equity
        
    all_years = hist_years + fc_years
    all_rev = rev_hist + rev_fc
    all_npat = npat_hist + npat_fc
    all_eps = eps_hist + eps_fc
    
    # Simple Valuation (Target P/E and P/B)
    latest_eps = eps_hist[-1] if eps_hist[-1] else 1000
    current_pe = current_price / latest_eps if latest_eps > 0 else 10
    target_pe = 10.0 # Standard defensive P/E
    target_pb = 1.2  # Standard P/B
    
    est_fair_pe = eps_fc[0] * target_pe
    
    # 3. Create Excel Model
    wb = openpyxl.Workbook()
    
    # Cover Sheet
    ws_cover = wb.active
    ws_cover.title = "Cover"
    ws_cover.views.sheetView[0].showGridLines = True
    
    ws_cover["A2"] = f"BÁO CÁO PHÂN TÍCH CỔ PHIẾU {ticker}"
    ws_cover["A2"].font = Font(name="Calibri", size=18, bold=True)
    ws_cover["A3"] = company_name
    ws_cover["A3"].font = Font(name="Calibri", size=12, italic=True)
    
    ws_cover["A5"] = "Giá hiện tại:"
    ws_cover["B5"] = current_price
    ws_cover["B5"].number_format = "#,##0"
    
    ws_cover["A6"] = "Vốn hóa (tỷ):"
    ws_cover["B6"] = market_cap / 1e9
    ws_cover["B6"].number_format = "#,##0"
    
    # P&L Sheet
    ws_pl = wb.create_sheet(title="P&L")
    ws_pl.views.sheetView[0].showGridLines = True
    ws_pl.append(["Chỉ tiêu"] + [f"{y}A" for y in hist_years] + [f"{y}E" for y in fc_years])
    ws_pl.append(["Doanh thu thuần (tỷ)"] + rev_hist + rev_fc)
    ws_pl.append(["Lợi nhuận gộp (tỷ)"] + gp_hist + gp_fc)
    ws_pl.append(["Lợi nhuận sau thuế (tỷ)"] + npat_hist + npat_fc)
    ws_pl.append(["EPS (VND)"] + eps_hist + eps_fc)
    
    for row in ws_pl.iter_rows(min_row=1, max_row=5):
        for cell in row:
            if isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0"
                
    # Style and save
    out_dir = os.path.join(os.path.dirname(__file__), "Bao cao", ticker)
    os.makedirs(out_dir, exist_ok=True)
    
    month_str = datetime.now().strftime("%Y-%m")
    excel_path = os.path.join(out_dir, f"{ticker}_Model_{month_str}.xlsx")
    pdf_path = os.path.join(out_dir, f"{ticker}_Phan_Tich_{month_str}.pdf")
    
    wb.save(excel_path)
    print(f"[OK] Excel saved locally at: {excel_path}")
    
    # 4. Create PDF Report using ReportLab
    doc = SimpleDocTemplate(pdf_path, pagesize=A4,
                            rightMargin=20*mm, leftMargin=20*mm,
                            topMargin=20*mm, bottomMargin=20*mm)
                            
    styles = getSampleStyleSheet()
    
    # Modify default styles for safer rendering
    title_style = ParagraphStyle(
        'DocTitle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=24,
        leading=28,
        textColor=HexColor("#1A365D"),
        spaceAfter=15
    )
    
    h1_style = ParagraphStyle(
        'SectionHeading',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=16,
        leading=20,
        textColor=HexColor("#2B6CB0"),
        spaceBefore=15,
        spaceAfter=10,
        keepWithNext=True
    )
    
    body_style = ParagraphStyle(
        'BodyTextCustom',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=10.5,
        leading=14,
        textColor=HexColor("#2D3748"),
        spaceAfter=10
    )
    
    table_header_style = ParagraphStyle(
        'TableHeader',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=9,
        leading=11,
        textColor=white,
        alignment=TA_CENTER
    )
    
    table_cell_style = ParagraphStyle(
        'TableCell',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=9,
        leading=11,
        textColor=HexColor("#2D3748"),
        alignment=TA_CENTER
    )

    story = []
    
    # Title Page / Header
    story.append(Paragraph(f"BÁO CÁO PHÂN TÍCH DOANH NGHIỆP: {ticker}", title_style))
    story.append(Paragraph(company_name, ParagraphStyle('Subtitle', parent=body_style, fontName='Helvetica-Oblique', fontSize=12, textColor=HexColor("#4A5568"))))
    story.append(Spacer(1, 15))
    
    # Summary Table
    summary_data = [
        [Paragraph("Mã cổ phiếu", table_header_style), Paragraph("Giá hiện tại (VND)", table_header_style), Paragraph("Vốn hóa thị trường", table_header_style), Paragraph("Ngành", table_header_style)],
        [Paragraph(ticker, table_cell_style), Paragraph(f"{current_price:,.0f}", table_cell_style), Paragraph(f"{market_cap/1e9:,.1f} tỷ VND", table_cell_style), Paragraph(sector, table_cell_style)]
    ]
    t = Table(summary_data, colWidths=[35*mm, 45*mm, 45*mm, 45*mm])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), HexColor("#1A365D")),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('GRID', (0,0), (-1,-1), 0.5, HexColor("#CBD5E0")),
    ]))
    story.append(t)
    story.append(Spacer(1, 20))
    
    # 1. Luận điểm đầu tư
    story.append(Paragraph("1. Luận điểm đầu tư chính", h1_style))
    story.append(Paragraph(f"Cổ phiếu {ticker} ({company_name}) đang giao dịch tại mức giá {current_price:,.0f} VND. Dựa trên phân tích kết quả kinh doanh lịch sử và các giả định tăng trưởng thận trọng, chúng tôi đánh giá doanh nghiệp có tiềm năng duy trì đà tăng trưởng doanh thu ở mức {avg_rev_growth*100:.1f}% mỗi năm trong giai đoạn tới nhờ vào sự phục hồi chung của ngành {sector}.", body_style))
    
    # 2. Tóm tắt tài chính
    story.append(Paragraph("2. Tóm tắt số liệu tài chính", h1_style))
    
    # Financial Table Headers
    fin_headers = [Paragraph("Chỉ tiêu", table_header_style)] + [Paragraph(f"{y}A", table_header_style) for y in hist_years] + [Paragraph(f"{y}E", table_header_style) for y in fc_years]
    
    row_rev = [Paragraph("Doanh thu (tỷ)", table_cell_style)] + [Paragraph(f"{v:,.1f}", table_cell_style) for v in all_rev]
    row_npat = [Paragraph("LNST (tỷ)", table_cell_style)] + [Paragraph(f"{v:,.1f}", table_cell_style) for v in all_npat]
    row_eps = [Paragraph("EPS (VND)", table_cell_style)] + [Paragraph(f"{v:,.0f}", table_cell_style) for v in all_eps]
    
    fin_table_data = [fin_headers, row_rev, row_npat, row_eps]
    
    # Calculate column widths: first col is wider, rest are equal
    num_cols = len(all_years) + 1
    col_widths = [45*mm] + [125*mm / len(all_years)] * len(all_years)
    
    t_fin = Table(fin_table_data, colWidths=col_widths)
    t_fin.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), HexColor("#2B6CB0")),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('GRID', (0,0), (-1,-1), 0.5, HexColor("#E2E8F0")),
    ]))
    story.append(t_fin)
    story.append(Spacer(1, 20))
    
    # 3. Định giá & Rủi ro
    story.append(Paragraph("3. Định giá & Khuyến nghị", h1_style))
    story.append(Paragraph(f"Áp dụng phương pháp định giá P/E mục tiêu ở mức {target_pe}x đối với EPS dự phóng năm {fc_years[0]}, giá trị hợp lý của cổ phiếu {ticker} ước tính đạt khoảng {est_fair_pe:,.0f} VND/CP.", body_style))
    
    story.append(Paragraph("4. Rủi ro đầu tư", h1_style))
    story.append(Paragraph("- Rủi ro biến động tỷ giá và lãi suất tăng ảnh hưởng chi phí vốn.", body_style))
    story.append(Paragraph("- Rủi ro cạnh tranh gia tăng trong phân khúc thị trường nội địa.", body_style))
    
    doc.build(story)
    print(f"[OK] PDF saved locally at: {pdf_path}")
    
    # 5. Upload to Google Drive
    gdrive_excel_url = None
    gdrive_pdf_url = None
    
    try:
        print("[GDrive] Starting upload to Google Drive...")
        _, gdrive_excel_url = google_drive_uploader.upload_file(excel_path)
        _, gdrive_pdf_url = google_drive_uploader.upload_file(pdf_path)
    except Exception as e:
        print(f"[GDrive] Upload failed: {e}")
        
    # 6. Save JSON Summary data
    summary_json = {
        "ticker": ticker,
        "companyName": company_name,
        "sector": sector,
        "currentPrice": current_price,
        "marketCap": market_cap,
        "shares": shares,
        "gdriveExcelUrl": gdrive_excel_url,
        "gdrivePdfUrl": gdrive_pdf_url,
        "data": {
            "years": all_years,
            "revenue": all_rev,
            "npat": all_npat,
            "eps": all_eps,
            "equity": equity_hist + equity_fc
        }
    }
    
    data_dir = os.path.join(os.path.dirname(__file__), "data")
    os.makedirs(data_dir, exist_ok=True)
    json_path = os.path.join(data_dir, f"{ticker}.json")
    
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(summary_json, f, ensure_ascii=False, indent=2)
    print(f"[OK] JSON Summary saved at: {json_path}")
    
    return True

if __name__ == "__main__":
    if len(sys.argv) > 1:
        ticker = sys.argv[1].upper()
        build_generic(ticker)
    else:
        print("Please provide a stock ticker. Example: python build_generic_model.py TCB")
