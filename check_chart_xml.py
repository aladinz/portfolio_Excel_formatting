import zipfile
import xml.etree.ElementTree as ET

# Extract chart XML from Excel file
with zipfile.ZipFile('test_chart.xlsx', 'r') as zip_ref:
    with zip_ref.open('xl/charts/chart1.xml') as chart_file:
        chart_content = chart_file.read().decode('utf-8')

print("CHART XML ANALYSIS")
print("=" * 60)

# Check for data labels
if 'dLbls' in chart_content:
    print("✓ Data labels (dLbls) found in chart XML")
else:
    print("✗ Data labels (dLbls) NOT found in chart XML")

# Count series
series_count = chart_content.count('<c:ser>')
print(f"Series count: {series_count}")

# Check for specific elements
checks = [
    ('Trend line', 'trendline'),
    ('Title styling', 'rPr'),
    ('Series colors', 'solidFill'),
]

for name, element in checks:
    count = chart_content.count(element)
    print(f"{name} ({element}): {count} occurrence(s)")

# Check chart2 (Monthly Returns)
with zipfile.ZipFile('test_chart.xlsx', 'r') as zip_ref:
    with zip_ref.open('xl/charts/chart2.xml') as chart_file:
        chart2_content = chart_file.read().decode('utf-8')

print("\nCHART2 (Monthly Returns) ANALYSIS")
print("=" * 60)
series2_count = chart2_content.count('<c:ser>')
print(f"Series count: {series2_count}")

if 'dLbls' in chart2_content:
    print("✓ Data labels found")
else:
    print("✗ Data labels NOT found")
