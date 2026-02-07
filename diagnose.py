from openpyxl import load_workbook

files = [
    'Portfolio report_Investment_07.02.2026.xlsx',
    'Portfolio report_Rollover IRA_07.02.2026.xlsx',
    'Portfolio report_Roth IRA_07.02.2026.xlsx'
]

print('\n' + '='*70)
print('DETAILED FORMATTING CHECK')
print('='*70)

for filename in files:
    print(f'\n{filename}:')
    
    wb = load_workbook(filename)
    ws_exec = wb['Executive Summary']
    
    # Check key cells for formatting
    checks = {
        'A1 (Title)': ('A1', ['1F4788', '4472C4']),  # Should be dark blue
        'A4 (KPI Header)': ('A4', ['4472C4']),  # Should be medium blue
        'A6 (Label)': ('A6', ['D9E1F2']),  # Should be light blue
        'A5 (Column Header)': ('A5', ['1F4788']),  # Should be dark blue
    }
    
    for check_name, (cell_addr, expected_colors) in checks.items():
        cell = ws_exec[cell_addr]
        has_color = False
        actual_color = 'NONE'
        
        if cell.fill and hasattr(cell.fill.start_color, 'rgb'):
            actual_color = cell.fill.start_color.rgb.upper() if cell.fill.start_color.rgb else 'NONE'
            # Remove leading zeros
            actual_color = actual_color.replace('00', '', 1)
            
            for expected in expected_colors:
                if expected.upper() in actual_color:
                    has_color = True
                    break
        
        status = '✓' if has_color else '✗'
        print(f'  {status} {check_name}: {actual_color}')
    
    wb.close()

print('\n' + '='*70)
print('If you see ✗, colors are missing')
print('='*70 + '\n')
