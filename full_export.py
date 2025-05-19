#!/usr/bin/env python3
"""
Full Export script for MATMAX WELLNESS Financial Model
This exports all the most important sheets to Google Sheets directly
"""

import os
import json
import sys
import time
import pandas as pd
import numpy as np
from googleapiclient.discovery import build
from google.oauth2 import service_account

print("Script starting...")

# Add financial_model directory to Python path
sys.path.append(os.path.join(os.path.dirname(__file__), 'financial_model'))

# Import model modules
import model_config
from revenue_projections import (
    calculate_membership_revenue, 
    calculate_punch_pass_revenue, 
    calculate_additional_services_revenue,
    consolidate_revenue
)
from expense_projections import (
    calculate_teacher_expenses,
    calculate_admin_expenses,
    calculate_facility_expenses,
    calculate_operating_expenses,
    calculate_capex,
    calculate_loan_payments,
    consolidate_expenses
)
from financial_statements import (
    create_income_statement,
    create_balance_sheet,
    create_cash_flow_statement
)
# Import additional modules
from customer_metrics import (
    calculate_customer_acquisition_metrics,
    calculate_customer_segmentation,
    calculate_churn_analysis,
    calculate_clv_by_channel,
    calculate_retention_strategies
)
from marketing_revenue import (
    calculate_content_revenue,
    calculate_sponsorship_revenue,
    calculate_paid_media_revenue,
    calculate_pr_value,
    consolidate_marketing_revenue
)
from financial_ratios import (
    create_financial_ratios,
    create_landlord_analysis,
    create_breakeven_analysis,
    create_sensitivity_analysis,
    create_occupancy_analysis,
    create_scenario_analysis
)
from retail_revenue import (
    calculate_retail_revenue,
    calculate_retail_inventory,
    calculate_retail_space_analysis,
    calculate_bestsellers_analysis
)

# Set spreadsheet ID - replace with your Google Sheet ID
SPREADSHEET_ID = "1ZjIpQtGYNwxOnP-f_DK7e36qBGIw37jkVYCMUK9xupM"

# Load credentials from environment or file
def load_credentials():
    """Load Google API credentials from environment variables or a file"""
    # Check if credentials are provided via environment variables
    if os.environ.get('GOOGLE_CREDS_JSON'):
        return json.loads(os.environ.get('GOOGLE_CREDS_JSON'))
    
    # Check for credentials file
    creds_path = os.path.join(os.path.dirname(__file__), 'credentials', 'google_service_account.json')
    if os.path.exists(creds_path):
        with open(creds_path, 'r') as f:
            return json.load(f)
    
    # If no credentials found, inform the user
    print("⚠️ No Google API credentials found.")
    print("Please either:")
    print("1. Set the GOOGLE_CREDS_JSON environment variable with your service account JSON")
    print("2. Place your service account JSON file at: credentials/google_service_account.json")
    return None

# Create credentials file
CREDENTIALS = load_credentials()

def transpose_dataframe(df, index_col=None):
    """
    Transpose a dataframe to have years as columns and metrics as rows
    """
    # Make a copy to avoid modifying the original
    df_copy = df.copy()
    
    # Determine the index column if not specified
    if index_col is None:
        # Try common index column names
        if 'Year' in df_copy.columns:
            index_col = 'Year'
        elif 'Item' in df_copy.columns:
            index_col = 'Item'
        elif 'Variable' in df_copy.columns:
            index_col = 'Variable'
        elif 'Customer Segment' in df_copy.columns:
            index_col = 'Customer Segment'
        elif 'Churn Factor' in df_copy.columns:
            index_col = 'Churn Factor'
        elif 'Channel' in df_copy.columns:
            index_col = 'Channel'
        elif 'Retention Strategy' in df_copy.columns:
            index_col = 'Retention Strategy'
        elif 'Room' in df_copy.columns:
            index_col = 'Room'
        elif 'Scenario' in df_copy.columns:
            index_col = 'Scenario'
        elif 'Product' in df_copy.columns:
            index_col = 'Product'
        else:
            # If no suitable index column is found, use the first column
            index_col = df_copy.columns[0]
    
    # Set the index column
    df_copy.set_index(index_col, inplace=True)
    
    # Transpose the DataFrame
    transposed = df_copy.transpose()
    
    # Convert index to a column called 'Item'
    transposed.reset_index(inplace=True)
    transposed.rename(columns={'index': 'Item'}, inplace=True)
    
    return transposed

def add_descriptions(df, descriptions=None):
    """
    Add descriptions to a dataframe
    """
    if descriptions is None:
        descriptions = {}
    
    # Add a description column at the second position
    df.insert(1, 'Description', '')
    
    # Fill in descriptions where available
    for idx, row in df.iterrows():
        item = row['Item']  # Get the name of the metric/item
        if item in descriptions:
            df.at[idx, 'Description'] = descriptions[item]
        else:
            df.at[idx, 'Description'] = 'Financial metric'
    
    return df

def create_dashboard():
    """Create a dashboard with main parameters"""
    dashboard = []
    
    # General Parameters
    dashboard.append(["MODEL PARAMETERS", "", ""])
    dashboard.append(["Parameter", "Value", "Description"])
    dashboard.append(["MODEL_YEARS", model_config.MODEL_YEARS, "Number of years to project"])
    dashboard.append(["CURRENCY", model_config.CURRENCY, "Currency"])
    dashboard.append(["INFLATION_RATE", model_config.INFLATION_RATE * 100, "Annual inflation rate (%)"])
    dashboard.append(["DISCOUNT_RATE", model_config.DISCOUNT_RATE * 100, "Discount rate for DCF (%)"])
    
    # Membership Parameters
    dashboard.append(["", "", ""])
    dashboard.append(["MEMBERSHIP PARAMETERS", "", ""])
    dashboard.append(["Membership Type", "Monthly Price", "Annual Price"])
    for membership_type, details in model_config.MEMBERSHIP_TYPES.items():
        dashboard.append([
            membership_type, 
            details['monthly_price'], 
            details['annual_price']
        ])
    
    # Create DataFrame from the dashboard
    return pd.DataFrame(dashboard)

def create_pricing_table():
    """Create pricing table"""
    pricing_data = []
    
    # Header
    pricing_data.append(["PRICING TABLE", "", "", ""])
    
    # Membership Section
    pricing_data.append(["MEMBERSHIPS", "", "", ""])
    pricing_data.append(["Type", "Monthly Price", "Annual Price", "Savings"])
    for membership, details in model_config.MEMBERSHIP_TYPES.items():
        monthly = details['monthly_price']
        annual = details['annual_price']
        savings = (monthly * 12 - annual) / (monthly * 12) * 100
        
        pricing_data.append([
            membership,
            f"{model_config.CURRENCY} {monthly}",
            f"{model_config.CURRENCY} {annual}",
            f"{savings:.1f}%"
        ])
    
    # Punch Pass Section
    pricing_data.append(["", "", "", ""])
    pricing_data.append(["PUNCH PASSES", "", "", ""])
    pricing_data.append(["Type", "Price", "Price Per Class", "Discount"])
    for pass_type, details in model_config.PUNCH_PASS_TYPES.items():
        num_passes = int(pass_type.split()[0])
        price = details['price']
        price_per_class = price / num_passes
        discount = details['discount'] * 100
        
        pricing_data.append([
            pass_type,
            f"{model_config.CURRENCY} {price}",
            f"{model_config.CURRENCY} {price_per_class:.2f}",
            f"{discount:.1f}%"
        ])
    
    return pd.DataFrame(pricing_data)

def create_venue_characteristics():
    """Create venue characteristics table"""
    venue_data = []
    
    # Header
    venue_data.append(["VENUE CHARACTERISTICS", "", "", ""])
    
    # Rooms and Capacity
    venue_data.append(["ROOMS", "", "", ""])
    venue_data.append(["Room Type", "Capacity", "Setup Cost", "Maintenance"])
    for room, details in model_config.ROOMS.items():
        venue_data.append([
            room,
            details['capacity'],
            f"{model_config.CURRENCY} {details['setup_cost']:,}",
            f"{model_config.CURRENCY} {details['maintenance_annual']:,}/year"
        ])
    
    # Facility Details
    total_capacity = sum(room['capacity'] for room in model_config.ROOMS.values())
    venue_data.append(["", "", "", ""])
    venue_data.append(["CAPACITY SUMMARY", "", "", ""])
    venue_data.append(["Total Room Capacity", total_capacity, "", ""])
    venue_data.append(["Classes Per Day", len(model_config.ROOMS) * model_config.CLASSES_PER_ROOM_PER_DAY, "", ""])
    venue_data.append(["Days Open Per Week", model_config.DAYS_OPEN_PER_WEEK, "", ""])
    
    return pd.DataFrame(venue_data)

def export_to_google_sheets():
    """Export essential sheets directly to Google Sheets"""
    print("\nMATMAX WELLNESS Financial Model Export")
    print("=====================================")
    
    try:
        # Check if credentials are available
        if not CREDENTIALS:
            print("❌ Error: Google API credentials not found. Cannot proceed with export.")
            return False
            
        # Save credentials to temporary file
        creds_file = "full_export_creds.json"
        with open(creds_file, "w") as f:
            json.dump(CREDENTIALS, f)
        print(f"Saved credentials to {creds_file}")
        
        # Set up credentials for Google API
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        creds = service_account.Credentials.from_service_account_file(
            creds_file, scopes=SCOPES)
        
        # Build the Sheets API service
        service = build('sheets', 'v4', credentials=creds)
        sheets = service.spreadsheets()
        
        print("Generating financial model data...")
        
        # Generate all core model data
        membership_revenue = calculate_membership_revenue()
        punch_pass_revenue = calculate_punch_pass_revenue()
        additional_services = calculate_additional_services_revenue()
        consolidated_revenue = consolidate_revenue()
        
        expenses, capex_df, loan_df, teacher_expenses, admin_expenses, facility_expenses, operating_expenses = consolidate_expenses(membership_revenue, punch_pass_revenue)
        
        income_statement = create_income_statement()
        balance_sheet = create_balance_sheet()
        cash_flow = create_cash_flow_statement()
        
        # Transpose data for Google Sheets format
        dashboard_df = create_dashboard()
        pricing_table = create_pricing_table()
        venue_characteristics = create_venue_characteristics()
        
        # Add descriptions to all sheets
        # Define a basic set of metric descriptions
        descriptions = {
            'Total Revenue': 'Sum of all revenue streams',
            'Total Expenses': 'Sum of all expenses',
            'Net Profit': 'Profit after all expenses and taxes',
            'Operating Profit (EBIT)': 'Earnings before interest and taxes',
            'Cash': 'Cash and cash equivalents',
            'Total Assets': 'Sum of all assets'
        }
        
        # Convert financial statements to transposed format
        revenue_transposed = transpose_dataframe(consolidated_revenue)
        expenses_transposed = transpose_dataframe(expenses)
        income_statement_transposed = transpose_dataframe(income_statement)
        
        # Create balance sheet with year mapping
        year_mapping = {i: f"Y{i}" for i in range(model_config.MODEL_YEARS + 1)}
        balance_sheet_copy = balance_sheet.copy()
        balance_sheet_copy['Year'] = balance_sheet_copy['Year'].apply(lambda x: year_mapping.get(x, x))
        balance_sheet_transposed = transpose_dataframe(balance_sheet_copy, index_col='Year')
        
        cash_flow_transposed = transpose_dataframe(cash_flow)
        
        # Add descriptions to transposed dataframes
        revenue_transposed = add_descriptions(revenue_transposed, descriptions)
        expenses_transposed = add_descriptions(expenses_transposed, descriptions)
        income_statement_transposed = add_descriptions(income_statement_transposed, descriptions)
        balance_sheet_transposed = add_descriptions(balance_sheet_transposed, descriptions)
        cash_flow_transposed = add_descriptions(cash_flow_transposed, descriptions)
        
        # Process additional detailed sheets
        membership_revenue_transposed = transpose_dataframe(membership_revenue)
        punch_pass_revenue_transposed = transpose_dataframe(punch_pass_revenue)
        additional_services_transposed = transpose_dataframe(additional_services)
        
        teacher_expenses_transposed = transpose_dataframe(teacher_expenses)
        admin_expenses_transposed = transpose_dataframe(admin_expenses)
        facility_expenses_transposed = transpose_dataframe(facility_expenses)
        operating_expenses_transposed = transpose_dataframe(operating_expenses)
        capex_transposed = transpose_dataframe(capex_df)
        loan_transposed = transpose_dataframe(loan_df)
        
        # Customer metrics
        customer_metrics = calculate_customer_acquisition_metrics()
        customer_metrics_transposed = transpose_dataframe(customer_metrics)
        customer_segmentation = calculate_customer_segmentation()
        churn_analysis = calculate_churn_analysis()
        clv_by_channel = calculate_clv_by_channel()
        retention_strategies = calculate_retention_strategies()
        
        # Marketing revenue
        content_revenue = calculate_content_revenue()
        content_revenue_transposed = transpose_dataframe(content_revenue)
        sponsorship_revenue = calculate_sponsorship_revenue()
        sponsorship_revenue_transposed = transpose_dataframe(sponsorship_revenue)
        media_revenue = calculate_paid_media_revenue()
        media_revenue_transposed = transpose_dataframe(media_revenue)
        pr_value = calculate_pr_value()
        pr_value_transposed = transpose_dataframe(pr_value)
        marketing_revenue = consolidate_marketing_revenue()
        marketing_revenue_transposed = transpose_dataframe(marketing_revenue)
        
        # Financial ratios
        financial_ratios = create_financial_ratios(income_statement, balance_sheet, cash_flow)
        financial_ratios_transposed = transpose_dataframe(financial_ratios)
        landlord_analysis = create_landlord_analysis(income_statement)
        landlord_analysis_transposed = transpose_dataframe(landlord_analysis)
        breakeven_analysis = create_breakeven_analysis(income_statement, expenses)
        breakeven_analysis_transposed = transpose_dataframe(breakeven_analysis)
        sensitivity_analysis = create_sensitivity_analysis(income_statement)
        occupancy_analysis = create_occupancy_analysis()
        scenario_analysis = create_scenario_analysis(income_statement)
        
        # Retail revenue
        retail_revenue = calculate_retail_revenue()
        retail_revenue_transposed = transpose_dataframe(retail_revenue)
        retail_inventory = calculate_retail_inventory()
        retail_inventory_transposed = transpose_dataframe(retail_inventory)
        retail_space = calculate_retail_space_analysis()
        retail_space_transposed = transpose_dataframe(retail_space)
        bestsellers = calculate_bestsellers_analysis()
        
        # Add descriptions to detailed sheets
        membership_revenue_transposed = add_descriptions(membership_revenue_transposed)
        punch_pass_revenue_transposed = add_descriptions(punch_pass_revenue_transposed)
        additional_services_transposed = add_descriptions(additional_services_transposed)
        teacher_expenses_transposed = add_descriptions(teacher_expenses_transposed)
        admin_expenses_transposed = add_descriptions(admin_expenses_transposed)
        facility_expenses_transposed = add_descriptions(facility_expenses_transposed)
        operating_expenses_transposed = add_descriptions(operating_expenses_transposed)
        capex_transposed = add_descriptions(capex_transposed)
        loan_transposed = add_descriptions(loan_transposed)
        customer_metrics_transposed = add_descriptions(customer_metrics_transposed)
        content_revenue_transposed = add_descriptions(content_revenue_transposed)
        sponsorship_revenue_transposed = add_descriptions(sponsorship_revenue_transposed)
        media_revenue_transposed = add_descriptions(media_revenue_transposed)
        pr_value_transposed = add_descriptions(pr_value_transposed)
        marketing_revenue_transposed = add_descriptions(marketing_revenue_transposed)
        financial_ratios_transposed = add_descriptions(financial_ratios_transposed)
        landlord_analysis_transposed = add_descriptions(landlord_analysis_transposed)
        breakeven_analysis_transposed = add_descriptions(breakeven_analysis_transposed)
        retail_revenue_transposed = add_descriptions(retail_revenue_transposed)
        retail_inventory_transposed = add_descriptions(retail_inventory_transposed)
        retail_space_transposed = add_descriptions(retail_space_transposed)
        
        # Get the existing sheets in the spreadsheet
        spreadsheet = sheets.get(spreadsheetId=SPREADSHEET_ID).execute()
        existing_sheets = [sheet['properties']['title'] for sheet in spreadsheet['sheets']]
        print(f"Existing sheets: {existing_sheets}")
        
        # Collection of sheets to export
        all_sheets = {
            # Main summary sheets
            'Dashboard': dashboard_df,
            'Venue Characteristics': venue_characteristics,
            'Pricing Table': pricing_table,
            
            # Financial statements
            'Income Statement': income_statement_transposed,
            'Balance Sheet': balance_sheet_transposed,
            'Cash Flow': cash_flow_transposed,
            
            # Revenue sheets
            'Revenue Summary': revenue_transposed,
            'Membership Revenue': membership_revenue_transposed,
            'Punch Pass Revenue': punch_pass_revenue_transposed,
            'Additional Services': additional_services_transposed,
            
            # Marketing and content revenue
            'Marketing Revenue': marketing_revenue_transposed,
            'Content Revenue': content_revenue_transposed,
            'Sponsorship Revenue': sponsorship_revenue_transposed,
            'Media Revenue': media_revenue_transposed,
            'PR Value': pr_value_transposed,
            
            # Retail revenue
            'Retail Revenue': retail_revenue_transposed,
            'Retail Inventory': retail_inventory_transposed,
            'Retail Space Analysis': retail_space_transposed,
            'Bestselling Products': bestsellers,
            
            # Expense sheets
            'Expense Summary': expenses_transposed,
            'Teacher Expenses': teacher_expenses_transposed,
            'Admin Expenses': admin_expenses_transposed,
            'Facility Expenses': facility_expenses_transposed,
            'Operating Expenses': operating_expenses_transposed,
            
            # Capital and financing
            'Capital Expenditures': capex_transposed,
            'Loan Details': loan_transposed,
            
            # Customer metrics
            'Customer Acquisition': customer_metrics_transposed,
            'Customer Segmentation': customer_segmentation,
            'Churn Analysis': churn_analysis,
            'Customer Lifetime Value': clv_by_channel,
            'Retention Strategies': retention_strategies,
            
            # Financial analysis
            'Financial Ratios': financial_ratios_transposed,
            'Landlord Analysis': landlord_analysis_transposed,
            'Break-even Analysis': breakeven_analysis_transposed,
            'Sensitivity Analysis': sensitivity_analysis,
            'Occupancy Analysis': occupancy_analysis,
            'Scenario Analysis': scenario_analysis
        }
        
        successful_exports = 0
        
        # Export sheets
        for sheet_name, df in all_sheets.items():
            print(f"\nExporting sheet: {sheet_name}...")
            
            # Convert dataframe to values, replacing NaN with empty string
            df_clean = df.fillna('')
            values = [df_clean.columns.tolist()] + df_clean.values.tolist()
            
            # Create or update sheet
            if sheet_name not in existing_sheets:
                print(f"  Creating new sheet: {sheet_name}")
                request_body = {
                    'requests': [{
                        'addSheet': {
                            'properties': {
                                'title': sheet_name
                            }
                        }
                    }]
                }
                sheets.batchUpdate(
                    spreadsheetId=SPREADSHEET_ID,
                    body=request_body
                ).execute()
                existing_sheets.append(sheet_name)
                # Add a delay to avoid hitting API rate limits
                time.sleep(1)
            else:
                print(f"  Updating existing sheet: {sheet_name}")
            
            # Update sheet data
            sheets.values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=f'{sheet_name}!A1',
                valueInputOption='RAW',
                body={'values': values}
            ).execute()
            
            successful_exports += 1
            print(f"  ✓ Sheet '{sheet_name}' exported successfully")
            
            # Add a delay between API calls to avoid rate limits
            time.sleep(1)
        
        print(f"\n✅ Export succeeded! Exported {successful_exports} sheets:")
        for sheet_name in all_sheets.keys():
            print(f"- {sheet_name}")
        
        print(f"\nGoogle Sheet URL: https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}")
        
        # Clean up
        if os.path.exists(creds_file):
            os.remove(creds_file)
            print(f"Removed temporary credentials file: {creds_file}")
        
        return True
        
    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return False

# Run the export if script is executed directly
if __name__ == "__main__":
    export_to_google_sheets() 