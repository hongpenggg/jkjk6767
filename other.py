import pandas as pd
import numpy as np

# 1. Setup Operating Assumptions
# ------------------------------
# NOTE: 2022 Revenue/COGS/SG&A Actuals are MISSING from the upload. 
# I have estimated Revenue based on the depreciation hint (787 / 1.5% = ~52,467) to make the model functional.
revenue_2022_est = 52467 
cogs_2022_est = revenue_2022_est * 0.77
sga_2022_est = revenue_2022_est * 0.165

years = ['2022A', '2023E', '2024E', '2025E', '2026E', '2027E']
growth_rate = 0.10
tax_rate = 0.20

# Drivers (Assumptions)
# COGS %: 77% (2023) -> 76% (2027) linearly
cogs_pct = [0.77, 0.77, 0.7675, 0.7650, 0.7625, 0.76] # 2022 placeholder, then declining

# SG&A %: 16.5% (2023) -> 16.0% (2027) linearly
sga_pct = [0.165, 0.165, 0.16375, 0.1625, 0.16125, 0.16]

# Fixed/Variable Drivers
depreciation_actual = 787
depreciation_rate = 0.015

amortization_actual = 82
amortization_proj = 100

interest_expense = 25

# 2. Build the Data Dictionary
# ----------------------------
data = {
    'Item': [
        'Revenue Growth', 'Revenues', 'Less: COGS', 'Gross Profit', 'Margin %',
        'Less: SG&A', 'EBITDA', 'Margin %',
        'Less: Depreciation', 'Less: Amortization', 'EBIT', 'Margin %',
        'Less: Interest', 'Pre-Tax Income', 'Less: Taxes', 'Net Income'
    ]
}

# Initialize columns
for year in years:
    data[year] = 0.0

df = pd.DataFrame(data)
df.set_index('Item', inplace=True)

# 3. Calculate Projections
# ------------------------

# 2022 Actuals (Populate with estimates/knowns)
df.loc['Revenues', '2022A'] = revenue_2022_est
df.loc['Less: COGS', '2022A'] = -cogs_2022_est
df.loc['Less: SG&A', '2022A'] = -sga_2022_est
df.loc['Less: Depreciation', '2022A'] = -depreciation_actual
df.loc['Less: Amortization', '2022A'] = -amortization_actual
df.loc['Less: Interest', '2022A'] = -interest_expense

# Loop for Projections (2023-2027)
prev_col = '2022A'
for i, col in enumerate(years[1:], 1): # Start at index 1 (2023)
    
    # Revenue
    df.loc['Revenue Growth', col] = growth_rate
    df.loc['Revenues', col] = df.loc['Revenues', prev_col] * (1 + growth_rate)
    
    # COGS (Using specific year assumption)
    df.loc['Less: COGS', col] = -(df.loc['Revenues', col] * cogs_pct[i])
    
    # SG&A
    df.loc['Less: SG&A', col] = -(df.loc['Revenues', col] * sga_pct[i])
    
    # Depreciation (1.5% of Revenue)
    df.loc['Less: Depreciation', col] = -(df.loc['Revenues', col] * depreciation_rate)
    
    # Amortization (Constant $100)
    df.loc['Less: Amortization', col] = -amortization_proj
    
    # Interest (Constant $25)
    df.loc['Less: Interest', col] = -interest_expense
    
    prev_col = col

# Calculate Subtotals (Vectorized for simplicity in display, though Excel will use formulas)
df.loc['Gross Profit'] = df.loc['Revenues'] + df.loc['Less: COGS']
df.loc['EBITDA'] = df.loc['Gross Profit'] + df.loc['Less: SG&A']
df.loc['EBIT'] = df.loc['EBITDA'] + df.loc['Less: Depreciation'] + df.loc['Less: Amortization']
df.loc['Pre-Tax Income'] = df.loc['EBIT'] + df.loc['Less: Interest']
df.loc['Less: Taxes'] = -(df.loc['Pre-Tax Income'] * tax_rate)
df.loc['Net Income'] = df.loc['Pre-Tax Income'] + df.loc['Less: Taxes']

# Calculate Margins
df.loc['Margin %'] = 0.0 # Clear rows
# We will rely on Excel formulas for the margins in the final file, 
# but for the dataframe display we calculate them:
gross_margin = (df.loc['Gross Profit'] / df.loc['Revenues'])
ebitda_margin = (df.loc['EBITDA'] / df.loc['Revenues'])
ebit_margin = (df.loc['EBIT'] / df.loc['Revenues'])

# Map them back strictly to the display rows (using row index positions is safer with duplicate names)
# But here we will just export to Excel with formulas.

# 4. Export to Excel with Formulas
# --------------------------------
output_file = 'Projected_Income_Statement.xlsx'
writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
df.to_excel(writer, sheet_name='Projections', startrow=1)

workbook = writer.book
worksheet = writer.sheets['Projections']

# Formats
money_fmt = workbook.add_format({'num_format': '#,##0', 'font_name': 'Arial', 'font_size': 10})
pct_fmt = workbook.add_format({'num_format': '0.0%', 'font_name': 'Arial', 'font_size': 10})
bold_fmt = workbook.add_format({'bold': True, 'font_name': 'Arial', 'font_size': 10})
header_fmt = workbook.add_format({'bold': True, 'bottom': 1, 'font_name': 'Arial', 'font_size': 10})

# Apply formatting and Formulas
# Columns B to G (Indices 1 to 6)
cols = ['B', 'C', 'D', 'E', 'F', 'G']

for col_idx, col_letter in enumerate(cols):
    row_start = 3 # Data starts at row 3 (Excel index 1-based, header is row 2)
    
    # Revenue Growth (Row 3)
    if col_idx > 0:
        worksheet.write_formula(f'{col_letter}3', '0.10', pct_fmt)
    else:
        worksheet.write(f'{col_letter}3', "n/a", pct_fmt)

    # Revenue (Row 4)
    if col_idx > 0: # Projected
        prev_col = cols[col_idx-1]
        worksheet.write_formula(f'{col_letter}4', f'{prev_col}4*(1+{col_letter}3)', money_fmt)
    else:
        worksheet.write(f'{col_letter}4', revenue_2022_est, money_fmt) # 2022 Hardcode

    # COGS (Row 5)
    # Logic: Revenue * Assumed %. 
    current_cogs_pct = cogs_pct[col_idx]
    if col_idx > 0:
        worksheet.write_formula(f'{col_letter}5', f'-{col_letter}4*{current_cogs_pct}', money_fmt)
    else:
        worksheet.write(f'{col_letter}5', -cogs_2022_est, money_fmt)

    # Gross Profit (Row 6)
    worksheet.write_formula(f'{col_letter}6', f'SUM({col_letter}4:{col_letter}5)', bold_fmt)

    # GP Margin (Row 7)
    worksheet.write_formula(f'{col_letter}7', f'{col_letter}6/{col_letter}4', pct_fmt)

    # SG&A (Row 8)
    current_sga_pct = sga_pct[col_idx]
    if col_idx > 0:
        worksheet.write_formula(f'{col_letter}8', f'-{col_letter}4*{current_sga_pct}', money_fmt)
    else:
        worksheet.write(f'{col_letter}8', -sga_2022_est, money_fmt)

    # EBITDA (Row 9)
    worksheet.write_formula(f'{col_letter}9', f'{col_letter}6+{col_letter}8', bold_fmt)
    
    # EBITDA Margin (Row 10)
    worksheet.write_formula(f'{col_letter}10', f'{col_letter}9/{col_letter}4', pct_fmt)

    # Depreciation (Row 11)
    if col_idx > 0:
        worksheet.write_formula(f'{col_letter}11', f'-{col_letter}4*0.015', money_fmt)
    else:
        worksheet.write(f'{col_letter}11', -depreciation_actual, money_fmt)

    # Amortization (Row 12)
    if col_idx > 0:
        worksheet.write_formula(f'{col_letter}12', '-100', money_fmt)
    else:
        worksheet.write(f'{col_letter}12', -amortization_actual, money_fmt)

    # EBIT (Row 13)
    worksheet.write_formula(f'{col_letter}13', f'SUM({col_letter}9,{col_letter}11,{col_letter}12)', bold_fmt)

    # EBIT Margin (Row 14)
    worksheet.write_formula(f'{col_letter}14', f'{col_letter}13/{col_letter}4', pct_fmt)

    # Interest (Row 15)
    worksheet.write_formula(f'{col_letter}15', '-25', money_fmt)

    # Pre-Tax (Row 16)
    worksheet.write_formula(f'{col_letter}16', f'{col_letter}13+{col_letter}15', bold_fmt)

    # Tax (Row 17)
    worksheet.write_formula(f'{col_letter}17', f'-{col_letter}16*0.20', money_fmt)

    # Net Income (Row 18)
    worksheet.write_formula(f'{col_letter}18', f'SUM({col_letter}16:{col_letter}17)', bold_fmt)

worksheet.set_column('A:A', 30)
worksheet.set_column('B:G', 15)

writer.close()
print("Model generated successfully.")