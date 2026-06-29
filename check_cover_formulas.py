import openpyxl
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

wb = openpyxl.load_workbook(r'E:\1. Projects\4. AIC - FA\Bao cao\VIB\VIB_Model_2026-06.xlsx')

ws_cov = wb['01_Cover']
print("=== 01_Cover Formula check ===")
print("C5 (Gia hien tai):", ws_cov['C5'].value)
print("C6 (Gia muc tieu):", ws_cov['C6'].value)
print("C7 (Upside):", ws_cov['C7'].value)
print("C8 (Khuyen nghi):", ws_cov['C8'].value)

ws_ass = wb['02_Assumptions']
print("\n=== 02_Assumptions values ===")
print("B2 (Gia hien tai):", ws_ass['B2'].value)

ws_val = wb['07_Valuation']
print("\n=== 07_Valuation check ===")
print("B21 (Target Price):", ws_val['B21'].value)
print("B22 (Gia hien tai):", ws_val['B22'].value)
print("B23 (Upside):", ws_val['B23'].value)
