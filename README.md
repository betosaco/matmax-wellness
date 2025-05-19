# MATMAX WELLNESS Financial Model Export

This repository contains a comprehensive export script for the MATMAX WELLNESS financial model, which exports all model sheets to Google Sheets.

## Complete Export Script

The `full_export.py` script exports 37 different sheets from the financial model to Google Sheets, providing a complete view of your financial projections:

- **Main summary sheets**: Dashboard, Venue Characteristics, Pricing Table
- **Financial statements**: Income Statement, Balance Sheet, Cash Flow
- **Revenue projections**: Revenue Summary, Membership Revenue, Punch Pass Revenue, Additional Services
- **Marketing revenue**: Marketing Revenue, Content Revenue, Sponsorship Revenue, Media Revenue, PR Value
- **Retail analysis**: Retail Revenue, Retail Inventory, Retail Space Analysis, Bestselling Products
- **Expense breakdowns**: Expense Summary, Teacher Expenses, Admin Expenses, Facility Expenses, Operating Expenses
- **Capital & financing**: Capital Expenditures, Loan Details
- **Customer metrics**: Customer Acquisition, Customer Segmentation, Churn Analysis, Customer Lifetime Value, Retention Strategies
- **Financial analysis**: Financial Ratios, Landlord Analysis, Break-even Analysis, Sensitivity Analysis, Occupancy Analysis, Scenario Analysis

## Setup

1. Install required packages:
   ```
   pip install google-api-python-client google-auth pandas numpy
   ```

2. Configure credentials:
   - Option 1: Set environment variable `GOOGLE_CREDS_JSON` with your Google service account credentials
   - Option 2: Create a file at `credentials/google_service_account.json` with your service account credentials
   - See `credentials/README.md` for detailed instructions

3. Update the `SPREADSHEET_ID` variable in `full_export.py` to your Google Sheet ID

## Usage

Run the script:
```
python full_export.py
```

The script automatically:
1. Loads your financial model data
2. Processes all calculations
3. Exports all sheets to Google Sheets
4. Handles API rate limits to ensure successful exports

## Security

This implementation prioritizes security:
- No hardcoded credentials in code
- Secure credential management
- Temporary files are automatically cleaned up
- Clear documentation of security practices 