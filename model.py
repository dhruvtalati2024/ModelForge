import requests
import pandas as pd
import os
import re
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.dimensions import ColumnDimension
import numpy as np # Import numpy for np.nan

# === CONFIG ===
# It's highly recommended to load API_KEY from an environment variable for security
# For now, keeping it here as per your original script, but be mindful in production
API_KEY = "fAFye0LEj92N3uvQAV54r5SRPgFtG2MX"
SAVE_PATH = os.path.expanduser("~/Desktop/Model")
os.makedirs(SAVE_PATH, exist_ok=True)

# === CLEAN LABELS ===
def clean_label(s):
    """Cleans and formats financial statement labels for readability."""
    s = re.sub(r'([a-z])([A-Z])', r'\1 \2', s).replace('_', ' ').title()
    replacements = {
        'Ebitda': 'EBITDA', 'Ebit': 'EBIT', 'Opex': 'Operating Expenses', 'R D': 'R&D',
        'Cogs': 'Cost of Goods Sold', 'Capex': 'Capital Expenditures',
        'Netincome': 'Net Income', 'Depreciationandamortization': 'Depreciation & Amortization',
        'Deferredincometax': 'Deferred Income Tax', 'Deferredrevenue': 'Deferred Revenue',
        'Totalcurrentassets': 'Total Current Assets', 'Totalcurrentliabilities': 'Total Current Liabilities',
        'Totalstockholdersequity': 'Total Stockholders Equity', 'Totalassets': 'Total Assets',
        'Totaldebt': 'Total Debt', 'Shortterminvestments': 'Short Term Investments',
        'Propertyplantequipmentnet': 'Property Plant & Equipment Net',
        'Cashandcashequivalents': 'Cash & Cash Equivalents',
        'Operatingcashflow': 'Operating Cash Flow', 'Freecashflow': 'Free Cash Flow',
        'Interestexpense': 'Interest Expense', 'Grossprofit': 'Gross Profit',
        'Operatingincome': 'Operating Income', 'Revenue': 'Revenue',
        'Researchanddevelopmentexpenses': 'Research & Development Expenses'
    }
    for old, new in replacements.items():
        s = s.replace(old, new)
    return s.strip()

# === FETCH DATA ===
def fetch_fmp_data(statement, ticker, years, api_key):
    """Fetches financial data from FMP API, cleans, and formats it."""
    url = f"https://financialmodelingprep.com/api/v3/{statement}/{ticker}?limit={years}&apikey={api_key}"
    r = requests.get(url )

    if r.status_code != 200:
        print(f"Error fetching {statement} for {ticker}: {r.status_code} - {r.text}")
        return pd.DataFrame()

    try:
        data = r.json()
        if not isinstance(data, list) or not data:
            print(f"No data found for {statement} for {ticker}.")
            return pd.DataFrame()

        df = pd.DataFrame(data)
        drop_cols = ['symbol', 'cik', 'reportedCurrency', 'fillingDate', 'acceptedDate',
                     'calendarYear', 'period', 'link', 'finalLink']
        df.drop(columns=[col for col in drop_cols if col in df.columns], inplace=True)

        df['date'] = pd.to_datetime(df['date'])
        df.set_index("date", inplace=True)
        df.sort_index(inplace=True) # Sort by date ascending

        df = df.transpose()
        df.columns = [col.strftime('%Y') for col in df.columns]
        df.index = [clean_label(i) for i in df.index]

        # Convert to numeric, coerce errors to NaN, and convert to millions
        df = df.apply(pd.to_numeric, errors='coerce') / 1_000_000
        return df.round(1) # Round to 1 decimal place for millions
    except Exception as e:
        print(f"An error occurred while processing {statement} data for {ticker}: {e}")
        return pd.DataFrame()

# === GROWTH INDEX ===
def growth_index(df):
    """Calculates the growth index for a given financial statement DataFrame."""
    if df.empty:
        return pd.DataFrame()

    # Ensure the DataFrame is sorted by columns (years) in ascending order
    df_sorted = df.sort_index(axis=1)

    # Initialize growth_df with float dtype to handle np.nan
    growth_df = pd.DataFrame(index=df_sorted.index, columns=df_sorted.columns, dtype=float)

    for index, row in df_sorted.iterrows():
        # Find the first non-zero, non-NaN value in the row to use as base
        base_value = None
        for val in row:
            if pd.notna(val) and val != 0:
                base_value = val
                break

        if base_value is not None:
            # Calculate growth relative to the base value
            growth_df.loc[index] = (row / base_value).round(2)
        else:
            # If no valid base, all growth values are NaN for that row
            growth_df.loc[index] = np.nan # Use np.nan for missing values

    return growth_df

# === FINANCIAL RATIOS ===
def calculate_ratios(is_df, bs_df, cf_df):
    """Calculates various financial ratios from Income Statement, Balance Sheet, and Cash Flow data."""
    # Get common years and sort them
    common_years = sorted(list(set(is_df.columns) & set(bs_df.columns) & set(cf_df.columns)))
    if not common_years:
        print("No common years found across all financial statements for ratio calculation. Skipping ratios.")
        return pd.DataFrame()

    # Filter dataframes to only common years and ensure consistent column order
    is_df_filtered = is_df[common_years]
    bs_df_filtered = bs_df[common_years]
    cf_df_filtered = cf_df[common_years]

    # Helper to safely get a row from a DataFrame, handling missing rows
    def get_row(df, key):
        if key in df.index:
            return df.loc[key]
        # Return a Series of NaNs with the correct index (years) if key is not found
        return pd.Series([np.nan] * len(df.columns), index=df.columns)

    # Helper for safe division, handling zero or NaN denominators
    def safe_div(numerator, denominator):
        denominator_numeric = pd.to_numeric(denominator, errors='coerce')
        # Replace 0 with NaN to avoid ZeroDivisionError, then perform division
        return numerator / denominator_numeric.replace(0, np.nan)

    ratios = pd.DataFrame(index=common_years) # Ratios will be indexed by year

    # --- Extract necessary components ---
    revenue = get_row(is_df_filtered, "Revenue")
    gross_profit = get_row(is_df_filtered, "Gross Profit")
    operating_income = get_row(is_df_filtered, "Operating Income")
    ebitda = get_row(is_df_filtered, "EBITDA")
    net_income = get_row(is_df_filtered, "Net Income")
    ebit = get_row(is_df_filtered, "EBIT")
    interest_expense = get_row(is_df_filtered, "Interest Expense")

    total_assets = get_row(bs_df_filtered, "Total Assets")
    total_stockholders_equity = get_row(bs_df_filtered, "Total Stockholders Equity")
    total_current_assets = get_row(bs_df_filtered, "Total Current Assets")
    total_current_liabilities = get_row(bs_df_filtered, "Total Current Liabilities")
    inventory = get_row(bs_df_filtered, "Inventory")
    cash_and_equivalents = get_row(bs_df_filtered, "Cash & Cash Equivalents")
    short_term_investments = get_row(bs_df_filtered, "Short Term Investments")
    total_liabilities = get_row(bs_df_filtered, "Total Liabilities")
    # Prefer 'Total Debt' if available, otherwise fall back to 'Total Liabilities'
    total_debt = get_row(bs_df_filtered, "Total Debt")
    # Use total_liabilities if total_debt is all NaN
    debt_for_ratio = total_debt if not total_debt.isnull().all() else total_liabilities

    property_plant_equipment_net = get_row(bs_df_filtered, "Property Plant & Equipment Net")

    free_cash_flow = get_row(cf_df_filtered, "Free Cash Flow")
    operating_cash_flow = get_row(cf_df_filtered, "Operating Cash Flow")
    capital_expenditure = get_row(cf_df_filtered, "Capital Expenditure")

    # --- Profitability Ratios ---
    ratios["Gross Margin"] = safe_div(gross_profit, revenue)
    ratios["Operating Margin"] = safe_div(operating_income, revenue)
    ratios["EBITDA Margin"] = safe_div(ebitda, revenue)
    ratios["Net Profit Margin"] = safe_div(net_income, revenue)
    ratios["Return on Assets (ROA)"] = safe_div(net_income, total_assets)
    ratios["Return on Equity (ROE)"] = safe_div(net_income, total_stockholders_equity)
    # ROCE = EBIT / (Total Assets - Current Liabilities) or EBIT / Capital Employed
    ratios["Return on Capital Employed (ROCE)"] = safe_div(ebit, total_assets - total_current_liabilities)

    # --- Liquidity Ratios ---
    ratios["Current Ratio"] = safe_div(total_current_assets, total_current_liabilities)
    ratios["Quick Ratio"] = safe_div(total_current_assets - inventory, total_current_liabilities)
    ratios["Cash Ratio"] = safe_div(cash_and_equivalents + short_term_investments, total_current_liabilities)

    # --- Leverage Ratios ---
    ratios["Debt to Equity"] = safe_div(debt_for_ratio, total_stockholders_equity)
    ratios["Debt to Assets"] = safe_div(debt_for_ratio, total_assets)
    ratios["Interest Coverage Ratio"] = safe_div(ebit, interest_expense) # EBIT / Interest Expense

    # --- Efficiency Ratios ---
    ratios["Asset Turnover"] = safe_div(revenue, total_assets)
    ratios["Fixed Asset Turnover"] = safe_div(revenue, property_plant_equipment_net)
    # Working Capital = Current Assets - Current Liabilities
    working_capital = total_current_assets - total_current_liabilities
    ratios["Working Capital Turnover"] = safe_div(revenue, working_capital)

    # --- Cash Flow Ratios ---
    ratios["FCF to Net Income"] = safe_div(free_cash_flow, net_income)
    ratios["FCF to Revenue"] = safe_div(free_cash_flow, revenue)
    # CapEx is often negative in CF statement, use absolute value for ratio
    ratios["CapEx to Operating Cash Flow"] = safe_div(capital_expenditure.abs(), operating_cash_flow)

    # Transpose to have ratios as rows and years as columns, then multiply by 100 for percentages
    return ratios.transpose().multiply(100).round(2) # Round to 2 decimal places for percentages

# === WRITE TO EXCEL ===
def write_excel(file_name, income_df, cash_df, balance_df, ratios_df):
    """Writes the financial data, growth indices, and ratios to an Excel workbook."""
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Financial Model"
    ws2 = wb.create_sheet("Growth Index")
    ws3 = wb.create_sheet("Financial Ratios")

    # Define styles
    thin = Side(border_style="thin", color="000000")
    border = Border(top=thin, left=thin, right=thin, bottom=thin)
    bold_font = Font(bold=True)
    header_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")

    def write_block(ws, df, title, start_row, percent_format=False, is_ratio_sheet=False):
        """Helper function to write a DataFrame block to a worksheet with styling."""
        if df.empty:
            ws.cell(row=start_row, column=1, value=f"No data available for {title.split('(')[0].strip()}.").font = Font(italic=True)
            return start_row + 2

        # Title row
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=len(df.columns) + 1)
        title_cell = ws.cell(row=start_row, column=1, value=title)
        title_cell.font = Font(bold=True, size=14, color="000080") # Dark blue title
        title_cell.alignment = align_left
        start_row += 1

        # Convert all NaN/NA values to None for openpyxl compatibility
        # Use .where to replace NaN/NA with None, preserving original dtype where possible
        df_to_write = df.where(pd.notna(df), None)

        # Write DataFrame to sheet
        for r_idx, row_data in enumerate(dataframe_to_rows(df_to_write, index=True, header=True), start_row):
            for c_idx, val in enumerate(row_data, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                cell.border = border

                # Apply bold font to headers (first row and first column)
                if r_idx == start_row or c_idx == 1:
                    cell.font = bold_font
                    if r_idx == start_row: # Header row fill
                        cell.fill = header_fill
                        cell.alignment = align_center # Center year headers

                # Apply number formatting
                if isinstance(val, (int, float)):
                    if percent_format:
                        cell.number_format = '0.00%' # Two decimal places for percentages (e.g., 1.00 -> 100.00%)
                    elif is_ratio_sheet:
                        cell.number_format = '0.00' # Ratios are already multiplied by 100, so just show as number
                    else:
                        cell.number_format = '#,##0.0' # One decimal place for millions (e.g., 1234.5)
                    cell.alignment = align_center
                else:
                    cell.alignment = align_left # Align row labels to left

        # Adjust column widths dynamically
        for col_idx in range(1, df.shape[1] + 2): # +1 for index column, +1 for 1-based indexing
            max_len = 0
            # Iterate through cells in the current column to find max content length
            for r_idx in range(start_row, ws.max_row + 1):
                cell_value = ws.cell(row=r_idx, column=col_idx).value
                if cell_value is not None:
                    # Estimate length based on formatted string for numbers
                    if isinstance(cell_value, (int, float)):
                        if percent_format:
                            formatted_val = f"{cell_value:.2f}%"
                        elif is_ratio_sheet:
                            formatted_val = f"{cell_value:.2f}"
                        else:
                            formatted_val = f"{cell_value:,.1f}"
                        max_len = max(max_len, len(formatted_val))
                    else:
                        max_len = max(max_len, len(str(cell_value)))
            # Set a reasonable min/max width
            adjusted_width = max(15, min(max_len + 2, 50)) # Min 15, Max 50 characters wide
            ws.column_dimensions[get_column_letter(col_idx)].width = adjusted_width # Direct assignment

        return ws.max_row + 2 # Return next starting row for the next block

    # Write to "Financial Model" sheet
    r = 1
    r = write_block(ws1, income_df, "INCOME STATEMENT (Millions USD)", r)
    r = write_block(ws1, cash_df, "CASH FLOW STATEMENT (Millions USD)", r)
    write_block(ws1, balance_df, "BALANCE SHEET (Millions USD)", r)

    # Write to "Growth Index" sheet
    r = 1
    r = write_block(ws2, growth_index(income_df), "INCOME STATEMENT GROWTH INDEX", r, percent_format=True)
    r = write_block(ws2, growth_index(cash_df), "CASH FLOW STATEMENT GROWTH INDEX", r, percent_format=True)
    write_block(ws2, growth_index(balance_df), "BALANCE SHEET GROWTH INDEX", r, percent_format=True)

    # Write to "Financial Ratios" sheet
    write_block(ws3, calculate_ratios(income_df, balance_df, cash_df), "FINANCIAL RATIOS (%)", 1, is_ratio_sheet=True)

    wb.save(file_name)
    return file_name # Return the path to the saved file

# === EXECUTION PIPELINE ===
def run_financial_analysis(ticker_symbol, years_of_data):
    """Main function to fetch data, calculate, and write to Excel."""
    print(f"Fetching data for {ticker_symbol} for {years_of_data} years...")
    income_df = fetch_fmp_data("income-statement", ticker_symbol, years_of_data, API_KEY)
    cash_df = fetch_fmp_data("cash-flow-statement", ticker_symbol, years_of_data, API_KEY)
    balance_df = fetch_fmp_data("balance-sheet-statement", ticker_symbol, years_of_data, API_KEY)

    # Check if all necessary dataframes were fetched successfully
    if income_df.empty or cash_df.empty or balance_df.empty:
        print(f"Could not fetch complete financial data for {ticker_symbol}. Please check the ticker or try again later.")
        return None

    ratios_df = calculate_ratios(income_df, balance_df, cash_df)

    output_file = f"{SAVE_PATH}/Model_{ticker_symbol}_{years_of_data}Y.xlsx"
    try:
        written_file = write_excel(output_file, income_df, cash_df, balance_df, ratios_df)
        if written_file:
            print(f"\n✅ Done! Excel file saved to:\n{written_file}")
        return written_file
    except Exception as e:
        print(f"An error occurred while writing the Excel file: {e}")
        return None

# --- Main execution when script is run directly ---
if __name__ == "__main__":
    ticker = input("Enter ticker symbol (e.g., AAPL): ").strip().upper()
    try:
        num_years = int(input("Enter number of years of data: ").strip())
    except ValueError:
        print("Invalid number of years. Please enter an integer.")
        exit()

    run_financial_analysis(ticker, num_years)
