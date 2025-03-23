import pandas as pd
from datetime import datetime

# Load the Excel file
file_path = "D:/wildfire/wildfire/RegressionTesting/rebates_deals_api1.3.xlsx"  # Replace with actual file path
xls = pd.ExcelFile(file_path)  # Load all sheets

# Define required columns for commission validation
commission_columns = [
    "merchantName", "merchantFullName", "merchantId",
    "affiliateNetwork", "commissionType", "commissionValue"
]

# Dictionary to store DataFrames for invalid commissions, missing exclusions, and duplicate merchants
invalid_commissions_data = {}
missing_exclusions_data = {}
duplicate_merchants_data = {}

# Iterate over all sheets in the input file
for sheet_name in xls.sheet_names:
    df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str)  # Read as string to preserve formatting

    # --- Part 1: Check for Invalid Commission Values ---
    if all(col in df.columns for col in commission_columns):
        df["commissionValue"] = pd.to_numeric(df["commissionValue"], errors='coerce')  # Convert commissionValue to numeric

        invalid_commission = df[
            ((df["commissionType"] == "%") & (df["commissionValue"] < 1)) |  
            ((df["commissionType"] == "flat") & (df["commissionValue"] < 5))  
        ]

        # Keep only required columns
        invalid_commission = invalid_commission[commission_columns]

        if not invalid_commission.empty:
            invalid_commissions_data[f"Invalid_{sheet_name}"] = invalid_commission

    # --- Part 2: Check for Missing Exclusions (Column L is empty) ---
    if len(df.columns) > 11:  # Ensure Column L exists
        missing_exclusions = df[df.iloc[:, 11].isna()]  # Column L (zero-based index 11)

        if not missing_exclusions.empty:
            new_df = missing_exclusions[["merchantName", "merchantFullName", "merchantId", "affiliateNetwork"]]
            missing_exclusions_data[f"MissingExclusions_{sheet_name}"] = new_df

    # --- Part 3: Identify Duplicate Merchants ---
    if all(col in df.columns for col in ["merchantName", "merchantFullName", "affiliateNetwork"]):
        duplicate_merchants = df[
            df.duplicated(subset=['merchantName'], keep=False) |
            df.duplicated(subset=['merchantFullName'], keep=False)
        ]

        if not duplicate_merchants.empty:
            duplicate_merchants_data[f"DuplicateMerchants_{sheet_name}"] = duplicate_merchants[
                ["merchantName", "merchantFullName", "merchantId", "affiliateNetwork"]
            ]

# --- Write Output File Only if Data Exists ---
if invalid_commissions_data or missing_exclusions_data or duplicate_merchants_data:
    today_date = datetime.today().strftime("%Y-%m-%d")
    output_file = f"Exclusion_Commission_Duplicate_Report_{today_date}.xlsx"

    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        # Write invalid commission sheets
        for sheet, data in invalid_commissions_data.items():
            data.to_excel(writer, sheet_name=sheet, index=False)

        # Write missing exclusions sheets
        for sheet, data in missing_exclusions_data.items():
            data.to_excel(writer, sheet_name=sheet, index=False)

        # Write duplicate merchant sheets
        for sheet, data in duplicate_merchants_data.items():
            data.to_excel(writer, sheet_name=sheet, index=False)

    print(f"New file created with multiple sheet: {output_file}")
else:
    print("No invalid commissions, missing exclusions, or duplicate merchants found.")
