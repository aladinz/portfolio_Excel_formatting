import sys
import os
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def format_portfolio_universal(filepath):
    """
    Universal formatter for both portfolio file structures:
    - Type A: Traditional structure (Executive Summary + Monthly Performance sheets)
    - Type B: Data structure (single Data sheet with raw monthly data)
    """
    
    print(f"\nProcessing: {filepath}")
    wb = load_workbook(filepath)
    
    # Define color scheme (consistent across both types)
    header_fill = PatternFill(start_color="1F4788", end_color="1F4788", fill_type="solid")
    subheader_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    metric_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    highlight_fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
    data_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    section_fills = {
        'profit': PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"),
        'trading': PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"),
        'cash': PatternFill(start_color="F4B084", end_color="F4B084", fill_type="solid"),
        'market': PatternFill(start_color="F1DCDB", end_color="F1DCDB", fill_type="solid"),
    }
    
    # Define fonts
    header_font = Font(bold=True, size=12, color="FFFFFF")
    title_font = Font(bold=True, size=16, color="FFFFFF")
    subheader_font = Font(bold=True, size=11, color="FFFFFF")
    bold_font = Font(bold=True, size=10)
    regular_font = Font(size=10)
    
    # Define borders
    thin_border = Border(
        left=Side(style='thin', color="000000"),
        right=Side(style='thin', color="000000"),
        top=Side(style='thin', color="000000"),
        bottom=Side(style='thin', color="000000")
    )
    thick_border = Border(
        left=Side(style='medium', color="000000"),
        right=Side(style='medium', color="000000"),
        top=Side(style='medium', color="000000"),
        bottom=Side(style='medium', color="000000")
    )
    
    # ========== DETECT FILE STRUCTURE ==========
    sheets = wb.sheetnames
    is_type_a = 'Executive Summary' in sheets and 'Monthly Performance' in sheets
    is_type_b = 'Data' in sheets and len(sheets) == 1
    
    if is_type_a:
        print("  → Detected: Type A (Executive Summary + Monthly Performance structure)")
        format_type_a(wb, header_fill, subheader_fill, metric_fill, highlight_fill,
                     header_font, title_font, subheader_font, bold_font, regular_font,
                     thin_border, thick_border)
    
    elif is_type_b:
        print("  → Detected: Type B (Data sheet structure)")
        format_type_b(wb, header_fill, subheader_fill, metric_fill, data_fill,
                     section_fills, header_font, title_font, subheader_font, bold_font,
                     regular_font, thin_border)
    
    else:
        print("  ⚠ Warning: Unknown file structure. Attempting basic formatting...")
        format_data_sheet_only(wb, metric_fill, header_font, bold_font, thin_border)
    
    # Save the workbook
    wb.save(filepath)
    print(f"✓ File saved successfully!\n")


def format_type_a(wb, header_fill, subheader_fill, metric_fill, highlight_fill,
                  header_font, title_font, subheader_font, bold_font, regular_font,
                  thin_border, thick_border):
    """Format Type A files (Executive Summary + Monthly Performance)"""
    
    # ========== FORMAT EXECUTIVE SUMMARY ==========
    ws_exec = wb['Executive Summary']
    
    # Title
    ws_exec['A1'].font = title_font
    ws_exec['A1'].fill = header_fill
    ws_exec['A1'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    ws_exec.merge_cells('A1:E1')
    ws_exec.row_dimensions[1].height = 30
    
    # Date
    ws_exec['A2'].font = Font(italic=True, size=10)
    ws_exec['A2'].alignment = Alignment(horizontal='left', vertical='center')
    ws_exec.row_dimensions[2].height = 18
    ws_exec.row_dimensions[3].height = 8
    
    # KPI Section
    ws_exec['A4'].font = subheader_font
    ws_exec['A4'].fill = subheader_fill
    ws_exec['A4'].alignment = Alignment(horizontal='left', vertical='center')
    ws_exec.merge_cells('A4:E4')
    ws_exec.row_dimensions[4].height = 22
    
    for col_letter in ['A', 'B', 'C']:
        cell = ws_exec[col_letter + '5']
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
    ws_exec.row_dimensions[5].height = 20
    
    for row in range(6, 15):
        ws_exec['A' + str(row)].font = bold_font
        ws_exec['A' + str(row)].fill = metric_fill
        ws_exec['A' + str(row)].border = thin_border
        ws_exec['A' + str(row)].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        
        ws_exec['B' + str(row)].font = Font(bold=True, size=11)
        ws_exec['B' + str(row)].fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        ws_exec['B' + str(row)].border = thin_border
        ws_exec['B' + str(row)].alignment = Alignment(horizontal='right', vertical='center')
        
        ws_exec['C' + str(row)].font = regular_font
        ws_exec['C' + str(row)].fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
        ws_exec['C' + str(row)].border = thin_border
        ws_exec['C' + str(row)].alignment = Alignment(horizontal='left', vertical='center')
        ws_exec.row_dimensions[row].height = 18
    
    ws_exec.row_dimensions[15].height = 8
    
    # Trading Activity Section
    ws_exec['A16'].font = subheader_font
    ws_exec['A16'].fill = subheader_fill
    ws_exec['A16'].alignment = Alignment(horizontal='left', vertical='center')
    ws_exec.merge_cells('A16:E16')
    ws_exec.row_dimensions[16].height = 22
    
    for col_letter in ['A', 'B', 'C', 'D', 'E']:
        cell = ws_exec[col_letter + '17']
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
    ws_exec.row_dimensions[17].height = 20
    
    for row in range(18, 22):
        for col_letter in ['A', 'B', 'C', 'D', 'E']:
            cell = ws_exec[col_letter + str(row)]
            if row == 21:
                cell.font = bold_font
                cell.fill = highlight_fill
                cell.border = thick_border
            else:
                if col_letter == 'A':
                    cell.font = bold_font
                    cell.fill = metric_fill
                else:
                    cell.fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='center', vertical='center')
        ws_exec.row_dimensions[row].height = 18
    
    # Set column widths
    ws_exec.column_dimensions['A'].width = 28
    ws_exec.column_dimensions['B'].width = 18
    ws_exec.column_dimensions['C'].width = 28
    ws_exec.column_dimensions['D'].width = 15
    ws_exec.column_dimensions['E'].width = 15
    
    print("  ✓ Executive Summary formatted")
    
    # ========== FORMAT MONTHLY PERFORMANCE ==========
    ws_monthly = wb['Monthly Performance']
    
    ws_monthly['A1'].font = Font(bold=True, size=14, color="FFFFFF")
    ws_monthly['A1'].fill = header_fill
    ws_monthly['A1'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    ws_monthly.merge_cells('A1:M1')
    ws_monthly.row_dimensions[1].height = 25
    ws_monthly.row_dimensions[2].height = 8
    
    # Headers
    for col in range(1, 14):
        cell = ws_monthly.cell(row=3, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
    ws_monthly.row_dimensions[3].height = 22
    
    # Portfolio Values
    for row in [4, 5]:
        for col in range(1, 14):
            cell = ws_monthly.cell(row=row, column=col)
            if col == 1:
                cell.font = bold_font
                cell.fill = metric_fill
            else:
                cell.fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='right', vertical='center')
        ws_monthly.row_dimensions[row].height = 18
    
    ws_monthly.row_dimensions[6].height = 8
    
    # Profit Metrics Section
    ws_monthly['A7'].font = subheader_font
    ws_monthly['A7'].fill = subheader_fill
    ws_monthly.merge_cells('A7:M7')
    ws_monthly['A7'].alignment = Alignment(horizontal='left', vertical='center')
    ws_monthly.row_dimensions[7].height = 20
    
    for row in range(8, 13):
        for col in range(1, 14):
            cell = ws_monthly.cell(row=row, column=col)
            if col == 1:
                cell.font = bold_font
                cell.fill = metric_fill
            else:
                cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='right', vertical='center')
        ws_monthly.row_dimensions[row].height = 18
    
    ws_monthly.row_dimensions[13].height = 8
    
    # Trading Activity Section
    ws_monthly['A14'].font = subheader_font
    ws_monthly['A14'].fill = subheader_fill
    ws_monthly.merge_cells('A14:M14')
    ws_monthly['A14'].alignment = Alignment(horizontal='left', vertical='center')
    ws_monthly.row_dimensions[14].height = 20
    
    for row in range(15, 21):
        for col in range(1, 14):
            cell = ws_monthly.cell(row=row, column=col)
            if col == 1:
                cell.font = bold_font
                cell.fill = metric_fill
            else:
                cell.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='right', vertical='center')
        ws_monthly.row_dimensions[row].height = 18
    
    ws_monthly.row_dimensions[21].height = 8
    
    # Cash Position Section
    ws_monthly['A22'].font = subheader_font
    ws_monthly['A22'].fill = subheader_fill
    ws_monthly.merge_cells('A22:M22')
    ws_monthly['A22'].alignment = Alignment(horizontal='left', vertical='center')
    ws_monthly.row_dimensions[22].height = 20
    
    for row in range(23, 26):
        for col in range(1, 14):
            cell = ws_monthly.cell(row=row, column=col)
            if col == 1:
                cell.font = bold_font
                cell.fill = metric_fill
            else:
                cell.fill = PatternFill(start_color="F4B084", end_color="F4B084", fill_type="solid")
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='right', vertical='center')
        ws_monthly.row_dimensions[row].height = 18
    
    ws_monthly.row_dimensions[26].height = 8
    
    # Market Comparison Section
    ws_monthly['A27'].font = subheader_font
    ws_monthly['A27'].fill = subheader_fill
    ws_monthly.merge_cells('A27:M27')
    ws_monthly['A27'].alignment = Alignment(horizontal='left', vertical='center')
    ws_monthly.row_dimensions[27].height = 20
    
    for row in range(28, 31):
        for col in range(1, 14):
            cell = ws_monthly.cell(row=row, column=col)
            if col == 1:
                cell.font = bold_font
                cell.fill = metric_fill
            else:
                cell.fill = PatternFill(start_color="F1DCDB", end_color="F1DCDB", fill_type="solid")
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='right', vertical='center')
        ws_monthly.row_dimensions[row].height = 18
    
    ws_monthly.column_dimensions['A'].width = 25
    for col in range(2, 14):
        ws_monthly.column_dimensions[get_column_letter(col)].width = 14
    
    print("  ✓ Monthly Performance formatted")


def format_type_b(wb, header_fill, subheader_fill, metric_fill, data_fill,
                  section_fills, header_font, title_font, subheader_font, bold_font,
                  regular_font, thin_border):
    """Format Type B files (Data sheet with raw monthly data)"""
    
    ws = wb['Data']
    
    # Format the title (usually first row)
    ws['A1'].font = title_font
    ws['A1'].fill = header_fill
    ws['A1'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    ws.row_dimensions[1].height = 25
    
    # Find month headers (typically in row 4)
    month_row = None
    for row in range(1, 10):
        cell_val = ws[f'B{row}'].value
        if cell_val and str(cell_val).strip() and any(month in str(cell_val) for month in ['25', '26']):
            month_row = row
            break
    
    if month_row:
        # Format month headers
        for col in range(2, ws.max_column + 1):
            cell = ws.cell(row=month_row, column=col)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = thin_border
        ws.row_dimensions[month_row].height = 20
        
        # Format metric rows
        last_section = None
        for row in range(month_row + 1, ws.max_row + 1):
            metric_name = ws[f'A{row}'].value
            if metric_name:
                # Determine which section this metric belongs to
                metric_str = str(metric_name).lower()
                if 'profit' in metric_str or 'dividend' in metric_str or 'price' in metric_str or 'sales' in metric_str:
                    section_fill = section_fills['profit']
                    last_section = 'profit'
                elif 'trade' in metric_str or 'buy' in metric_str or 'sell' in metric_str or 'purchase' in metric_str or 'turnover' in metric_str:
                    section_fill = section_fills['trading']
                    last_section = 'trading'
                elif 'cash' in metric_str or 'fund' in metric_str or 'deposit' in metric_str or 'withdraw' in metric_str:
                    section_fill = section_fills['cash']
                    last_section = 'cash'
                else:
                    section_fill = section_fills.get(last_section, data_fill)
                
                # Format metric label
                ws[f'A{row}'].font = bold_font
                ws[f'A{row}'].fill = metric_fill
                ws[f'A{row}'].border = thin_border
                ws[f'A{row}'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                
                # Format data cells
                for col in range(2, ws.max_column + 1):
                    cell = ws.cell(row=row, column=col)
                    cell.fill = section_fill
                    cell.border = thin_border
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                
                ws.row_dimensions[row].height = 16
    
    # Set column widths
    ws.column_dimensions['A'].width = 28
    for col in range(2, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(col)].width = 14
    
    print("  ✓ Data sheet formatted with professional styling")


def format_data_sheet_only(wb, metric_fill, header_font, bold_font, thin_border):
    """Fallback: Apply basic formatting to any data sheet"""
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        
        # Format first row as header
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=1, column=col)
            cell.font = header_font
            cell.fill = metric_fill
            cell.border = thin_border
        
        # Format data rows
        for row in range(2, min(ws.max_row + 1, 100)):
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row=row, column=col)
                cell.border = thin_border
        
        print(f"  ✓ {sheet_name} formatted with basic styling")


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("\n" + "=" * 70)
        print("UNIVERSAL PORTFOLIO FORMATTER - Auto-detects file structure")
        print("=" * 70)
        print("\nUsage: python format_all.py <filename.xlsx> [filename2.xlsx ...]")
        print("\nExamples:")
        print("  python format_all.py Portfolio1.xlsx")
        print("  python format_all.py *.xlsx  (formats all .xlsx files)")
        print("\nSupported Structures:")
        print("  • Type A: Executive Summary + Monthly Performance sheets")
        print("  • Type B: Single Data sheet with monthly metrics")
        print("=" * 70 + "\n")
        sys.exit(0)
    
    # Get all files to process
    files_to_process = []
    for arg in sys.argv[1:]:
        if '*' in arg:
            # Wildcard - expand it
            import glob
            files_to_process.extend(glob.glob(arg))
        else:
            files_to_process.append(arg)
    
    print("\n" + "=" * 70)
    print(f"Processing {len(files_to_process)} file(s)...")
    print("=" * 70)
    
    success_count = 0
    error_count = 0
    
    for filename in files_to_process:
        try:
            format_portfolio_universal(filename)
            success_count += 1
        except Exception as e:
            print(f"✗ Error processing {filename}: {e}\n")
            error_count += 1
    
    print("=" * 70)
    print(f"COMPLETE: {success_count} file(s) formatted successfully")
    if error_count > 0:
        print(f"ERRORS: {error_count} file(s) failed")
    print("=" * 70 + "\n")
