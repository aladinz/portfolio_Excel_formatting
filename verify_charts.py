from openpyxl import load_workbook

wb = load_workbook('test_chart.xlsx')
ws = wb['Executive Summary']

print('CHART ENHANCEMENTS VERIFICATION')
print('=' * 60)

for i, chart in enumerate(ws._charts, 1):
    print(f'\nChart {i}:')
    try:
        title_text = chart.title.tx.rich.p[0].r[0].t
    except:
        title_text = str(chart.title)
    print(f'  Title: {title_text}')
    print(f'  Legend position: {chart.legend.position}')
    print(f'  Number of series: {len(chart.series)}')
    
    for j, series in enumerate(chart.series, 1):
        print(f'  Series {j}:')
        print(f'    Has data labels: {hasattr(series, "dLbls") and series.dLbls is not None}')
        if hasattr(series, 'dLbls') and series.dLbls:
            print(f'      Show values: {series.dLbls.showVal}')
        
        if hasattr(series, 'graphicalProperties') and series.graphicalProperties:
            gp = series.graphicalProperties
            if hasattr(gp, 'solidFill') and gp.solidFill:
                print(f'    Color: {gp.solidFill}')
            elif hasattr(gp, 'line') and gp.line:
                color = gp.line.solidFill if hasattr(gp.line, 'solidFill') else 'default'
                print(f'    Line color: {color}')
    
    # Check for trendlines
    if hasattr(chart, 'series') and len(chart.series) > 0:
        if hasattr(chart.series[0], 'trendlines') and len(chart.series[0].trendlines) > 0:
            ttype = chart.series[0].trendlines[0].trendlineType
            print(f'  Trend line: YES (type: {ttype})')
