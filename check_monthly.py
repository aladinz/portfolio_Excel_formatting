from openpyxl import load_workbook

files = [
    'Portfolio report_Investment_07.02.2026.xlsx',
    'Portfolio report_Rollover IRA_07.02.2026.xlsx',
    'Portfolio report_Roth IRA_07.02.2026.xlsx'
]

print('\n' + '='*70)
print('MONTHLY PERFORMANCE FORMATTING CHECK')
print('='*70)

for filename in files:
    print(f'\n{filename}:')
    
    wb = load_workbook(filename)
    ws_monthly = wb['Monthly Performance']
    
    # Check Monthly Performance formatting
    checks = {
        'A1 (Title)': 'A1',
        'A3 (Header Row)': 'A3',
        'A7 (Profit Section)': 'A7',
        'A14 (Trading Section)': 'A14',
    }
    
    for check_name, cell_addr in checks.items():
        cell = ws_monthly[cell_addr]
        has_color = False
        actual_color = 'NONE'
        
        if cell.fill and hasattr(cell.fill.start_color, 'rgb'):
            actual_color = cell.fill.start_color.rgb.upper() if cell.fill.start_color.rgb else 'NONE'
            actual_color = actual_color.replace('00', '', 1)
            has_color = actual_color != '000000' and actual_color != 'NONE'
        
        status = '✓' if has_color else '✗'
        cell_value = str(cell.value)[:40] if cell.value else 'empty'
        print(f'  {status} {check_name}: Color={actual_color}, Value="{cell_value}"')
    
    # Check data cells have colors too
    data_cell = ws_monthly['B4']
    data_color = 'NONE'
    if data_cell.fill and hasattr(data_cell.fill.start_color, 'rgb'):
        data_color = data_cell.fill.start_color.rgb.upper() if data_cell.fill.start_color.rgb else 'NONE'
    
    print(f'  Data cell B4 color: {data_color}')
    
    wb.close()

print('\n' + '='*70)
print('All formatting should show ✓')
print('='*70 + '\n')
