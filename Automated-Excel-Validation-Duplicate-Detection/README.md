Automated Merchant Data Validation and Processing Script

Overview:
This Python script automates the validation and processing of merchant data stored in an Excel file. It scans all sheets in the input file and identifies potential data inconsistencies based on predefined conditions. The processed data is then categorized and saved in a new Excel file with multiple sheets for easy review.

Features & Functionality:

1.Invalid Commission Detection:

The script checks commission values based on two conditions:
If the commissionType is "%", the commissionValue should be at least 1.
If the commissionType is "flat", the commissionValue should be at least 5.
Any records that violate these conditions are marked as invalid commissions and saved in a separate sheet.

2.Missing Exclusions Check:

The script verifies whether Column L (exclusion data) is empty for any merchant.
If exclusions are missing, the merchant details (such as merchantName, merchantFullName, merchantId, and affiliateNetwork) are extracted and saved in a separate sheet for further review.

3.Duplicate Merchant Identification:

The script detects duplicate merchant entries by comparing values in the merchantName and merchantFullName columns.
If a merchant appears multiple times, all occurrences are extracted along with their merchantId and affiliateNetwork.
The duplicate records are saved in a separate sheet for analysis.

Output:
If any issues are found, a new Excel report is generated with a name format:
Exclusion_Commission_Duplicate_Report_YYYY-MM-DD.xlsx

This report contains separate sheets for:
	1.Invalid Commissions
	2.Missing Exclusions
	3.Duplicate Merchants
If no issues are detected, the script simply prints:
"No invalid commissions, missing exclusions, or duplicate merchants found."

Automation & Efficiency:
The script efficiently processes multiple sheets in a single run.
All necessary data transformations and checks are performed using Pandas, ensuring high performance.
The output report makes it easy for stakeholders to review and take corrective action.
This script is particularly useful for teams handling large merchant datasets and looking to automate data validation and quality assurance.
