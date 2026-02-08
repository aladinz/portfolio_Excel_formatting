# -*- coding: utf-8 -*-
import sys
import os

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, BarChart, Reference

# Import restructuring function from restructure_type_b.py
try:
    from restructure_type_b import restructure_type_b_to_type_a
except ImportError:
    # If import fails, define a fallback
    def restructure_type_b_to_type_a(filepath):
        print("  ! Warning: Could not import restructure function")
        return False

def safe_merge_cells(ws, cell_range):
    """
    Safely merge cells by checking for conflicts first.
    This prevents Excel corruption from overlapping or conflicting merges.
    """
    try:
        # Simply try to merge - if it fails, continue anyway
        # The merged_cells registry will handle it
        ws.merge_cells(cell_range)
    except Exception as e:
        # If merge fails, silently continue - data is more important than formatting
        pass

def format_portfolio_universal(filepath):
    """
    Universal formatter for portfolio files with extended Executive Summary sections.
    Handles all structure types:
    - Type A with extended sections (Exec Summary + Monthly + Trading/Insights/Actions)
    - Type A without extended sections (original structure)
    - Type B (Data only)
    """
    
    print(f"\nProcessing: {filepath}")
    wb = load_workbook(filepath)
    
    # Define color scheme
    header_fill = PatternFill(start_color="1F4788", end_color="1F4788", fill_type="solid")
    subheader_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    metric_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    highlight_fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
    
    # Section fills
    section_fills = {
        'trading': PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"),
        'trading_activity': PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"),
        'insights': PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"),
        'actions': PatternFill(start_color="F4B084", end_color="F4B084", fill_type="solid"),
        'cash': PatternFill(start_color="F4B084", end_color="F4B084", fill_type="solid"),
        'market': PatternFill(start_color="F1DCDB", end_color="F1DCDB", fill_type="solid"),
    }
    
    data_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    
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
        print("  -> Detected: Type A (Executive Summary + Monthly Performance)")
        format_type_a_extended(wb, header_fill, subheader_fill, metric_fill, highlight_fill,
                              header_font, title_font, subheader_font, bold_font, regular_font,
                              thin_border, thick_border, section_fills)
    
    elif is_type_b:
        print("  -> Detected: Type B (Data sheet structure)")
        print("  -> Restructuring to Type A format...")
        
        # Restructure Type B to Type A
        if restructure_type_b_to_type_a(filepath):
            # Reload the restructured file
            wb = load_workbook(filepath)
            print("  -> Applying Type A formatting to restructured file...")
            format_type_a_extended(wb, header_fill, subheader_fill, metric_fill, highlight_fill,
                                  header_font, title_font, subheader_font, bold_font, regular_font,
                                  thin_border, thick_border, section_fills)
        else:
            print("  ! Warning: Restructuring failed, skipping file")
            return
    
    else:
        print("  ! Warning: Unknown file structure. Attempting basic formatting...")
    
    # Save the workbook
    wb.save(filepath)
    print(f"[OK] File saved successfully!\n")


def format_type_a_extended(wb, header_fill, subheader_fill, metric_fill, highlight_fill,
                           header_font, title_font, subheader_font, bold_font, regular_font,
                           thin_border, thick_border, section_fills):
    """Format Type A files with extended Executive Summary sections"""
    
    # ========== FORMAT EXECUTIVE SUMMARY ==========
    ws_exec = wb['Executive Summary']
    
    # Title
    ws_exec['A1'].font = title_font
    ws_exec['A1'].fill = header_fill
    ws_exec['A1'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    safe_merge_cells(ws_exec, 'A1:E1')
    ws_exec.row_dimensions[1].height = 30
    
    # Date
    ws_exec['A2'].font = Font(italic=True, size=10)
    ws_exec.row_dimensions[2].height = 18
    ws_exec.row_dimensions[3].height = 8
    
    # KPI Section
    ws_exec['A4'].font = subheader_font
    ws_exec['A4'].fill = subheader_fill
    ws_exec['A4'].alignment = Alignment(horizontal='left', vertical='center')
    safe_merge_cells(ws_exec, 'A4:E4')
    ws_exec.row_dimensions[4].height = 22
    
    for col_letter in ['A', 'B', 'C']:
        cell = ws_exec[col_letter + '5']
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
    ws_exec.row_dimensions[5].height = 20
    
    # Format KPI data rows (6-14)
    for row in range(6, 15):
        ws_exec[f'A{row}'].font = bold_font
        ws_exec[f'A{row}'].fill = metric_fill
        ws_exec[f'A{row}'].border = thin_border
        ws_exec[f'A{row}'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        
        ws_exec[f'B{row}'].font = Font(bold=True, size=11)
        ws_exec[f'B{row}'].fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        ws_exec[f'B{row}'].border = thin_border
        ws_exec[f'B{row}'].alignment = Alignment(horizontal='right', vertical='center')
        
        ws_exec[f'C{row}'].font = regular_font
        ws_exec[f'C{row}'].fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
        ws_exec[f'C{row}'].border = thin_border
        ws_exec[f'C{row}'].alignment = Alignment(horizontal='left', vertical='center')
        ws_exec.row_dimensions[row].height = 18
    
    ws_exec.row_dimensions[15].height = 8
    
    # ========== FORMAT EXTENDED SECTIONS ==========
    
    # Trading Activity Summary (rows 16-20)
    # Extract trading data from Monthly Performance sheet
    if 'Monthly Performance' in wb.sheetnames and ws_exec['A16'].value and 'TRADING' in str(ws_exec['A16'].value).upper():
        ws_monthly = wb['Monthly Performance']
        
        # Find trading activity section in Monthly Performance
        total_trades = 0
        buy_trades = 0
        sell_trades = 0
        
        for row in range(1, ws_monthly.max_row + 1):
            cell_val = str(ws_monthly[f'A{row}'].value or '').upper()
            if 'TOTAL TRADES' in cell_val and 'TURNOVER' not in cell_val:
                # Sum all values in this row
                for col in range(2, ws_monthly.max_column + 1):
                    val = ws_monthly.cell(row=row, column=col).value
                    if val:
                        try:
                            total_trades += float(val) if isinstance(val, (int, float)) else 0
                        except:
                            pass
            elif 'BUY TRADES' in cell_val or 'BUY TRANSACTIONS' in cell_val:
                for col in range(2, ws_monthly.max_column + 1):
                    val = ws_monthly.cell(row=row, column=col).value
                    if val:
                        try:
                            buy_trades += float(val) if isinstance(val, (int, float)) else 0
                        except:
                            pass
            elif 'SELL TRADES' in cell_val or 'SELL TRANSACTIONS' in cell_val:
                for col in range(2, ws_monthly.max_column + 1):
                    val = ws_monthly.cell(row=row, column=col).value
                    if val:
                        try:
                            sell_trades += float(val) if isinstance(val, (int, float)) else 0
                        except:
                            pass
        
        # Populate trading activity data
        if total_trades > 0:
            ws_exec['B17'].value = f"{int(total_trades)}"
            ws_exec['B18'].value = f"{int(buy_trades)}"
            ws_exec['B19'].value = f"{int(sell_trades)}"
    
    if ws_exec['A16'].value and 'TRADING' in str(ws_exec['A16'].value).upper():
        ws_exec['A16'].font = subheader_font
        ws_exec['A16'].fill = subheader_fill
        ws_exec['A16'].alignment = Alignment(horizontal='left', vertical='center')
        safe_merge_cells(ws_exec, 'A16:E16')
        ws_exec.row_dimensions[16].height = 22
        
        for row in range(17, 21):
            if ws_exec[f'A{row}'].value:
                ws_exec[f'A{row}'].font = bold_font
                ws_exec[f'A{row}'].fill = metric_fill
                ws_exec[f'A{row}'].border = thin_border
                ws_exec[f'A{row}'].alignment = Alignment(horizontal='left', vertical='center')
                
                for col in ['B', 'C', 'D', 'E']:
                    cell = ws_exec[f'{col}{row}']
                    cell.fill = section_fills['trading']
                    cell.border = thin_border
                    cell.alignment = Alignment(horizontal='left', vertical='center')
                
                ws_exec.row_dimensions[row].height = 18
    
    ws_exec.row_dimensions[21].height = 8
    
    # Key Insights & Recommendations (rows 22-28)
    if ws_exec['A22'].value and 'KEY INSIGHTS' in str(ws_exec['A22'].value).upper():
        ws_exec['A22'].font = subheader_font
        ws_exec['A22'].fill = subheader_fill
        ws_exec['A22'].alignment = Alignment(horizontal='left', vertical='center')
        safe_merge_cells(ws_exec, 'A22:E22')
        ws_exec.row_dimensions[22].height = 22
        
        for row in range(23, 29):
            if ws_exec[f'A{row}'].value:
                ws_exec[f'A{row}'].font = regular_font
                ws_exec[f'A{row}'].fill = section_fills['insights']
                ws_exec[f'A{row}'].border = thin_border
                ws_exec[f'A{row}'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                
                for col in ['B', 'C', 'D', 'E']:
                    cell = ws_exec[f'{col}{row}']
                    cell.fill = section_fills['insights']
                    cell.border = thin_border
                
                ws_exec.row_dimensions[row].height = 20
    
    ws_exec.row_dimensions[29].height = 8
    
    # Action Items & Strategy (rows 30-36)
    if ws_exec['A30'].value and 'ACTION ITEMS' in str(ws_exec['A30'].value).upper():
        ws_exec['A30'].font = subheader_font
        ws_exec['A30'].fill = subheader_fill
        ws_exec['A30'].alignment = Alignment(horizontal='left', vertical='center')
        safe_merge_cells(ws_exec, 'A30:E30')
        ws_exec.row_dimensions[30].height = 22
        
        for row in range(31, 37):
            if ws_exec[f'A{row}'].value:
                ws_exec[f'A{row}'].font = regular_font
                ws_exec[f'A{row}'].fill = section_fills['actions']
                ws_exec[f'A{row}'].border = thin_border
                ws_exec[f'A{row}'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                
                for col in ['B', 'C', 'D', 'E']:
                    cell = ws_exec[f'{col}{row}']
                    cell.fill = section_fills['actions']
                    cell.border = thin_border
                
                ws_exec.row_dimensions[row].height = 20
    
    # Set column widths
    ws_exec.column_dimensions['A'].width = 28
    ws_exec.column_dimensions['B'].width = 45
    ws_exec.column_dimensions['C'].width = 20
    ws_exec.column_dimensions['D'].width = 15
    ws_exec.column_dimensions['E'].width = 15
    
    print("  [OK] Executive Summary formatted (with extended sections)")
    
    # ========== FORMAT MONTHLY PERFORMANCE ==========
    ws_monthly = wb['Monthly Performance']
    
    ws_monthly['A1'].font = Font(bold=True, size=14, color="FFFFFF")
    ws_monthly['A1'].fill = header_fill
    ws_monthly['A1'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    safe_merge_cells(ws_monthly, 'A1:M1')
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
    safe_merge_cells(ws_monthly, 'A7:M7')
    ws_monthly['A7'].alignment = Alignment(horizontal='left', vertical='center')
    ws_monthly.row_dimensions[7].height = 20
    
    for row in range(8, 13):
        for col in range(1, 14):
            cell = ws_monthly.cell(row=row, column=col)
            if col == 1:
                cell.font = bold_font
                cell.fill = metric_fill
            else:
                cell.fill = section_fills['trading']
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='right', vertical='center')
        ws_monthly.row_dimensions[row].height = 18
    
    ws_monthly.row_dimensions[13].height = 8
    
    # Trading Activity Section
    ws_monthly['A14'].font = subheader_font
    ws_monthly['A14'].fill = subheader_fill
    safe_merge_cells(ws_monthly, 'A14:M14')
    ws_monthly['A14'].alignment = Alignment(horizontal='left', vertical='center')
    ws_monthly.row_dimensions[14].height = 20
    
    for row in range(15, 21):
        for col in range(1, 14):
            cell = ws_monthly.cell(row=row, column=col)
            if col == 1:
                cell.font = bold_font
                cell.fill = metric_fill
            else:
                cell.fill = section_fills['trading_activity']
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='right', vertical='center')
        ws_monthly.row_dimensions[row].height = 18
    
    ws_monthly.row_dimensions[21].height = 8
    
    # Cash Position Section
    ws_monthly['A22'].font = subheader_font
    ws_monthly['A22'].fill = subheader_fill
    safe_merge_cells(ws_monthly, 'A22:M22')
    ws_monthly['A22'].alignment = Alignment(horizontal='left', vertical='center')
    ws_monthly.row_dimensions[22].height = 20
    
    for row in range(23, 27):
        for col in range(1, 14):
            cell = ws_monthly.cell(row=row, column=col)
            if col == 1:
                cell.font = bold_font
                cell.fill = metric_fill
            else:
                cell.fill = section_fills['cash']
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='right', vertical='center')
        ws_monthly.row_dimensions[row].height = 18
    
    ws_monthly.row_dimensions[27].height = 8
    
    # Market Comparison Section
    ws_monthly['A28'].font = subheader_font
    ws_monthly['A28'].fill = subheader_fill
    safe_merge_cells(ws_monthly, 'A28:M28')
    ws_monthly['A28'].alignment = Alignment(horizontal='left', vertical='center')
    ws_monthly.row_dimensions[28].height = 20
    
    for row in range(29, 33):
        for col in range(1, 14):
            cell = ws_monthly.cell(row=row, column=col)
            if col == 1:
                cell.font = bold_font
                cell.fill = metric_fill
            else:
                cell.fill = section_fills['market']
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='right', vertical='center')
        ws_monthly.row_dimensions[row].height = 18
    
    # Format remaining rows
    for row in range(33, ws_monthly.max_row + 1):
        for col in range(1, 14):
            cell = ws_monthly.cell(row=row, column=col)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='right', vertical='center')
            if col == 1:
                cell.font = bold_font
                cell.fill = metric_fill
            else:
                cell.fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    
    # Set column widths
    ws_monthly.column_dimensions['A'].width = 28
    for col in range(2, 14):
        ws_monthly.column_dimensions[get_column_letter(col)].width = 14
    
    print("  [OK] Monthly Performance formatted")
    
    # ========== ADD CHARTS ==========
    add_charts_to_executive_summary(wb)
    print("  [OK] Charts added")




def add_charts_to_executive_summary(wb):
    """Add Portfolio Growth and Monthly Returns charts to Executive Summary"""
    try:
        ws_exec = wb['Executive Summary']
        ws_monthly = wb['Monthly Performance']
        
        # Extract months
        months = []
        for col in range(2, 14):
            col_letter = get_column_letter(col)
            month = ws_monthly[f'{col_letter}3'].value
            if month:
                months.append(str(month).strip())
        
        # Extract portfolio values (row 5)
        portfolio_values = []
        for col in range(2, 14):
            col_letter = get_column_letter(col)
            val = ws_monthly[f'{col_letter}5'].value
            if val:
                try:
                    if isinstance(val, str):
                        val = float(val.replace('$', '').replace(',', ''))
                    else:
                        val = float(val)
                    portfolio_values.append(val)
                except:
                    portfolio_values.append(0)
        
        # Extract monthly profits (row 8)
        monthly_profits = []
        for col in range(2, 14):
            col_letter = get_column_letter(col)
            val = ws_monthly[f'{col_letter}8'].value
            if val:
                try:
                    if isinstance(val, str):
                        val = float(val.replace('$', '').replace(',', ''))
                    else:
                        val = float(val)
                    monthly_profits.append(val)
                except:
                    monthly_profits.append(0)
        
        if months and portfolio_values and monthly_profits:
            # Add data to Executive Summary for chart references
            ws_exec['G3'].value = "Month"
            ws_exec['H3'].value = "Portfolio Value"
            
            for i, (month, value) in enumerate(zip(months, portfolio_values)):
                ws_exec[f'G{4+i}'].value = month
                ws_exec[f'H{4+i}'].value = value
            
            ws_exec['J3'].value = "Month"
            ws_exec['K3'].value = "Monthly Profit"
            
            for i, (month, profit) in enumerate(zip(months, monthly_profits)):
                ws_exec[f'J{4+i}'].value = month
                ws_exec[f'K{4+i}'].value = profit
            
            # Portfolio Growth Line Chart
            portfolio_chart = LineChart()
            portfolio_chart.title = "Portfolio Growth - 12 Month Progression"
            portfolio_chart.style = 10
            portfolio_chart.y_axis.title = "Portfolio Value ($)"
            portfolio_chart.x_axis.title = "Month"
            portfolio_chart.height = 10
            portfolio_chart.width = 16
            
            data = Reference(ws_exec, min_col=8, min_row=3, max_row=3+len(portfolio_values))
            categories = Reference(ws_exec, min_col=7, min_row=4, max_row=3+len(portfolio_values))
            portfolio_chart.add_data(data, titles_from_data=True)
            portfolio_chart.set_categories(categories)
            
            portfolio_chart.series[0].graphicalProperties.line.solidFill = "1F4788"
            portfolio_chart.series[0].graphicalProperties.line.width = 25000
            
            ws_exec.add_chart(portfolio_chart, "F1")
            
            # Monthly Returns Bar Chart
            returns_chart = BarChart()
            returns_chart.type = "col"
            returns_chart.title = "Monthly Returns - 12 Month Performance"
            returns_chart.style = 10
            returns_chart.y_axis.title = "Profit ($)"
            returns_chart.x_axis.title = "Month"
            returns_chart.height = 10
            returns_chart.width = 16
            
            data = Reference(ws_exec, min_col=11, min_row=3, max_row=3+len(monthly_profits))
            categories = Reference(ws_exec, min_col=10, min_row=4, max_row=3+len(monthly_profits))
            returns_chart.add_data(data, titles_from_data=True)
            returns_chart.set_categories(categories)
            
            returns_chart.series[0].graphicalProperties.solidFill = "4472C4"
            
            ws_exec.add_chart(returns_chart, "F20")
    
    except Exception as e:
        # Charts are optional - don't fail if they can't be added
        pass


def format_type_b(wb, header_fill, subheader_fill, metric_fill, data_fill,
                 section_fills, header_font, title_font, subheader_font, bold_font,
                 regular_font, thin_border):
    """Format Type B files (Data sheet only)"""
    
    ws = wb['Data']
    
    # Apply header formatting to first row (months)
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = title_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
    
    ws.row_dimensions[1].height = 25
    
    # Apply formatting to data rows
    for row in range(2, ws.max_row + 1):
        # First column (metric names) gets darker formatting
        cell_a = ws.cell(row=row, column=1)
        cell_a.fill = subheader_fill
        cell_a.font = Font(bold=True, size=10, color="FFFFFF")
        cell_a.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        cell_a.border = thin_border
        
        # Data columns get light formatting with borders
        for col in range(2, ws.max_column + 1):
            cell = ws.cell(row=row, column=col)
            cell.fill = data_fill
            cell.font = regular_font
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.border = thin_border
            
            # Format numbers with currency if they contain currency symbols
            if cell.value and isinstance(cell.value, str) and '$' in str(cell.value):
                cell.number_format = '$#,##0.00'
    
    # Optimize column widths
    for col in range(1, ws.max_column + 1):
        max_length = 0
        column_letter = get_column_letter(col)
        
        for row in range(1, ws.max_row + 1):
            cell = ws.cell(row=row, column=col)
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        
        adjusted_width = min(max_length + 2, 30)
        ws.column_dimensions[column_letter].width = adjusted_width
    
    print("  [OK] Data sheet formatting applied")


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("\n" + "="*70)
        print("UNIVERSAL PORTFOLIO FORMATTER (EXTENDED WITH CHARTS)")
        print("Professional formatting for all portfolio file types")
        print("="*70)
        print("\nUsage: python format_all.py <filename.xlsx> [file2.xlsx ...]")
        print("\nFeatures:")
        print("  [OK] Professional color-coded formatting")
        print("  [OK] Portfolio Growth Line Charts")
        print("  [OK] Monthly Returns Bar Charts")
        print("  [OK] Executive Summary with KPIs")
        print("  [OK] Trading Activity Summary")
        print("  [OK] Key Insights & Recommendations")
        print("  [OK] Action Items & Strategy")
        print("\nSupports:")
        print("  * Type A: Executive Summary + Monthly Performance + Data")
        print("  * Type A Extended: + Trading Activity + Key Insights + Action Items")
        print("  * Type B: Single Data sheet")
        print("\nExample:")
        print("  python format_all.py Portfolio1.xlsx")
        print("  python format_all.py *.xlsx")
        print("="*70 + "\n")
        sys.exit(0)
    
    # Get all files to process
    files_to_process = []
    for arg in sys.argv[1:]:
        if '*' in arg:
            import glob
            files_to_process.extend(glob.glob(arg))
        else:
            files_to_process.append(arg)
    
    print("\n" + "="*70)
    print(f"Processing {len(files_to_process)} file(s)...")
    print("="*70)
    
    for filename in files_to_process:
        try:
            format_portfolio_universal(filename)
        except Exception as e:
            print(f"[ERROR] {e}\n")
    
    print("="*70)
    print("FORMATTING COMPLETE")
    print("="*70 + "\n")
