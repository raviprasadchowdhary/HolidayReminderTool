import pandas as pd
from datetime import datetime, timedelta
import calendar
import os
import unicodedata # <--- NEW: For robust string cleaning

# --- Helper function for robust string cleaning ---
def clean_string(text):
    if isinstance(text, str):
        # Normalize Unicode characters (e.g., convert non-breaking spaces, ligatures to simpler forms)
        # and strip leading/trailing whitespace.
        try:
            text = unicodedata.normalize('NFKC', text)
            return text.strip()
        except Exception: # In case of any weird string that unicodedata can't handle
            return text.replace('\xa0', ' ').encode('ascii', 'ignore').decode('utf-8').strip()
    return text

def get_holiday_data(holiday_file='holidays.csv'):
    """Reads holiday data from a CSV file with 'Shore' and 'Locations' columns."""
    try:
        df = pd.read_csv(holiday_file)
        
        # Ensure required columns exist
        required_columns = ['Date', 'HolidayName', 'Shore', 'Locations']
        for col in required_columns:
            if col not in df.columns:
                print(f"Error: Required column '{col}' not found in {holiday_file}. Please check your CSV headers.")
                # logging.error(f"Required column '{col}' not found in {holiday_file}.") # Assuming logging is set up in main
                return pd.DataFrame()

        # Convert 'Date' column with specific format, coercing errors
        try:
            df['Date'] = pd.to_datetime(df['Date'], format='%m/%d/%Y', errors='coerce')
        except Exception as e: # Catch any unexpected error during conversion
            print(f"Error converting 'Date' column in '{holiday_file}'. Ensure dates are in MM/DD/YYYY format. Details: {e}")
            # logging.error(f"Error converting 'Date' column in '{holiday_file}': {e}")
            return pd.DataFrame()

        # Drop rows where date conversion failed (NaT)
        original_row_count = len(df)
        df.dropna(subset=['Date'], inplace=True)
        if len(df) < original_row_count:
            print(f"Warning: {original_row_count - len(df)} row(s) in '{holiday_file}' were dropped due to invalid date formats.")
            # logging.warning(f"{original_row_count - len(df)} row(s) in '{holiday_file}' dropped due to invalid dates.")


        # Apply cleaning to relevant text columns immediately after reading
        for col in ['HolidayName', 'Shore', 'Locations']:
            if col in df.columns: # Check if column exists before applying
                 df[col] = df[col].apply(clean_string)

        return df
    except FileNotFoundError:
        print(f"Error: Holiday file '{holiday_file}' not found. Please ensure it's in the same directory.")
        # logging.error(f"Holiday file '{holiday_file}' not found.")
        return pd.DataFrame()
    except Exception as e:
        print(f"An unexpected error occurred while reading '{holiday_file}': {e}")
        # logging.error(f"Unexpected error reading '{holiday_file}': {e}")
        return pd.DataFrame()


def generate_modern_holiday_email_html(holidays_df, company_name_footer="Your Company", signature_name="HR Team"):
    """
    Generates a modern, good-looking HTML content for the holiday reminder email,
    including 2x2 table and colored calendars, based on new holiday data structure.
    Accepts company_name_footer and signature_name for customization.
    """
    current_date = datetime.now()
    # For consistent testing output as per your screenshot, you can uncomment the line below:
    # current_date = datetime(2025, 5, 24) # Ensure this is May for the dummy data to show up as "current"

    current_month_name = current_date.strftime('%B %Y')
    next_month_date = (current_date.replace(day=1) + timedelta(days=32)).replace(day=1)
    next_month_name = next_month_date.strftime('%B %Y')

    current_month_holidays_all = holidays_df[
        (holidays_df['Date'].dt.month == current_date.month) &
        (holidays_df['Date'].dt.year == current_date.year)
    ].sort_values(by='Date')

    next_month_holidays_all = holidays_df[
        (holidays_df['Date'].dt.month == next_month_date.month) &
        (holidays_df['Date'].dt.year == next_month_date.year)
    ].sort_values(by='Date')

    def get_filtered_holidays_for_table(month_df, shore_type):
        locations_str = ""
        if shore_type == 'Onshore':
            filtered = month_df[
                (month_df['Shore'] == 'Onshore') | (month_df['Shore'] == 'Both')
            ]
            return ''.join([
                f'<li style="margin-bottom: 8px; font-size: 14px; color: #555555;"><strong>{h["Date"].strftime("%b %d")}</strong>: {h["HolidayName"]} '
                f'<span style="color: #888888;">({h["Locations"]})</span></li>' # <--- ADDED Locations for Onshore
                for idx, h in filtered.iterrows()
            ]) if not filtered.empty else '<li style="font-size: 14px; color: #777777;">No Onshore Holidays</li>'
        else: # Offshore
            filtered = month_df[
                (month_df['Shore'] == 'Offshore') | (month_df['Shore'] == 'Both')
            ]
            return ''.join([
                f'<li style="margin-bottom: 8px; font-size: 14px; color: #555555;"><strong>{h["Date"].strftime("%b %d")}</strong>: {h["HolidayName"]} '
                f'<span style="color: #888888;">({h["Locations"]})</span></li>'
                for idx, h in filtered.iterrows()
            ]) if not filtered.empty else '<li style="font-size: 14px; color: #777777;">No Offshore Holidays</li>'

    table_html = f"""
    <p style="font-size: 16px; color: #333333; margin-bottom: 20px;">Hi Team,</p>
    <p style="font-size: 16px; color: #333333; margin-bottom: 30px;">Here are the upcoming holidays for <strong style="color: #007bff;">{current_month_name}</strong> and <strong style="color: #007bff;">{next_month_name}</strong>:</p>

    <table border="0" cellpadding="0" cellspacing="0" style="width:100%; max-width: 600px; margin: 0 auto 30px auto; border-radius: 8px; overflow: hidden; background-color: #ffffff; box-shadow: 0 4px 8px rgba(0,0,0,0.05);">
        <tr>
            <th colspan="2" style="background-color:#e9ecef; padding: 15px; font-size: 18px; color: #333333; text-align: center; border-bottom: 1px solid #dee2e6;">{current_month_name} Holidays</th>
            <th colspan="2" style="background-color:#e9ecef; padding: 15px; font-size: 18px; color: #333333; text-align: center; border-bottom: 1px solid #dee2e6;">{next_month_name} Holidays</th>
        </tr>
        <tr>
            <td style="vertical-align: top; padding: 20px; border-right: 1px solid #dee2e6; border-bottom: 1px solid #dee2e6;">
                <h4 style="margin-top:0; color: #007bff; font-size: 16px;">Onshore</h4>
                <ul style="list-style-type:none; padding:0; margin:0;">
                    {get_filtered_holidays_for_table(current_month_holidays_all, 'Onshore')}
                </ul>
            </td>
            <td style="vertical-align: top; padding: 20px; border-bottom: 1px solid #dee2e6;">
                <h4 style="margin-top:0; color: #28a745; font-size: 16px;">Offshore</h4>
                <ul style="list-style-type:none; padding:0; margin:0;">
                    {get_filtered_holidays_for_table(current_month_holidays_all, 'Offshore')}
                </ul>
            </td>
            <td style="vertical-align: top; padding: 20px; border-right: 1px solid #dee2e6; border-bottom: 1px solid #dee2e6;">
                <h4 style="margin-top:0; color: #007bff; font-size: 16px;">Onshore</h4>
                <ul style="list-style-type:none; padding:0; margin:0;">
                    {get_filtered_holidays_for_table(next_month_holidays_all, 'Onshore')}
                </ul>
            </td>
            <td style="vertical-align: top; padding: 20px; border-bottom: 1px solid #dee2e6;">
                <h4 style="margin-top:0; color: #28a745; font-size: 16px;">Offshore</h4>
                <ul style="list-style-type:none; padding:0; margin:0;">
                    {get_filtered_holidays_for_table(next_month_holidays_all, 'Offshore')}
                </ul>
            </td>
        </tr>
    </table>
    """

    calendar_html = """
    <p style="font-size: 16px; color: #333333; margin-top: 40px; text-align: center;">A quick look at your holiday calendars:</p>
    <div style="display: flex; justify-content: center; align-items: flex-start; gap: 20px; flex-wrap: wrap; max-width: 800px; margin: 0 auto 30px auto;">
    """

    for month_date, month_holidays_df in [(current_date, current_month_holidays_all), (next_month_date, next_month_holidays_all)]:
        year = month_date.year
        month = month_date.month

        onshore_dates = set(month_holidays_df[
            (month_holidays_df['Shore'] == 'Onshore') |
            (month_holidays_df['Shore'] == 'Both')
        ]['Date'].dt.day.tolist())

        offshore_dates = set(month_holidays_df[
            (month_holidays_df['Shore'] == 'Offshore') |
            (month_holidays_df['Shore'] == 'Both')
        ]['Date'].dt.day.tolist())

        both_dates = onshore_dates.intersection(offshore_dates)
        only_onshore = onshore_dates - both_dates
        only_offshore = offshore_dates - both_dates

        cal = calendar.HTMLCalendar(calendar.SUNDAY) # Start week on Sunday
        month_cal_html = cal.formatmonth(year, month)

        month_cal_html = month_cal_html.replace('border="0"', '') \
                                       .replace('cellpadding="0"', '') \
                                       .replace('cellspacing="0"', '') \
                                       .replace('class="month"', '')

        for day in range(1, 32): # Iterate up to 31 to cover all possible days
            day_str_exact = f'>{day}<'
            day_str_with_attr = f'">{day}<' # To handle cases where calendar might add attributes to <td>

            if day in both_dates:
                replace_with = f'><span style="background-color: #90EE90; color: #333; font-weight: bold; padding: 4px 6px; border-radius: 4px; display: inline-block;">{day}</span><'
                month_cal_html = month_cal_html.replace(day_str_exact, replace_with)
                month_cal_html = month_cal_html.replace(day_str_with_attr, f'"{replace_with}')
            elif day in only_onshore:
                replace_with = f'><span style="background-color: #FFD700; color: #333; font-weight: bold; padding: 4px 6px; border-radius: 4px; display: inline-block;">{day}</span><'
                month_cal_html = month_cal_html.replace(day_str_exact, replace_with)
                month_cal_html = month_cal_html.replace(day_str_with_attr, f'"{replace_with}')
            elif day in only_offshore:
                replace_with = f'><span style="background-color: #ADD8E6; color: #333; font-weight: bold; padding: 4px 6px; border-radius: 4px; display: inline-block;">{day}</span><'
                month_cal_html = month_cal_html.replace(day_str_exact, replace_with)
                month_cal_html = month_cal_html.replace(day_str_with_attr, f'"{replace_with}')
        
        month_cal_html = month_cal_html.replace( # Make month name larger and centered
            f'<th colspan="7" class="month">{month_date.strftime("%B %Y")}</th>',
            f'<th colspan="7" style="text-align:center; font-size: 18px; padding:10px 0; background-color: #f0f0f0; color: #333;">{month_date.strftime("%B %Y")}</th>'
        )


        month_cal_html = month_cal_html.replace(
            '<table',
            '<table style="width:100%; border-collapse: collapse; font-size: 14px; text-align: center; border: 0;"'
        ).replace(
            '<th>', # For day names (Mon, Tue, etc.)
            '<th style="background-color: #f8f9fa; padding: 8px; color: #555555; text-align: center; border-bottom: 1px solid #e0e0e0; border-top: 1px solid #e0e0e0;">'
        ).replace(
            '<td>',
            '<td style="padding: 8px; border: 1px solid #e0e0e0; text-align: center; vertical-align: middle; height: 35px;">' # Added height for consistency
        )


        calendar_html += f"""
        <div style="background-color: #ffffff; border-radius: 8px; box-shadow: 0 4px 8px rgba(0,0,0,0.05); padding: 20px; flex: 1 1 calc(50% - 40px); min-width: 300px; box-sizing: border-box; margin:10px;">
            {month_cal_html}
        </div>
        """
    calendar_html += "</div>"
    calendar_html += """
    <p style="margin-top: 30px; text-align: center; font-size: 14px; color: #555555;">
        <span style="background-color: #ADD8E6; padding: 4px 10px; border-radius: 4px; display: inline-block; margin-right: 5px; vertical-align: middle;">&nbsp;</span> <span style="vertical-align: middle;">Offshore</span>
        &nbsp;&nbsp;
        <span style="background-color: #FFD700; padding: 4px 10px; border-radius: 4px; display: inline-block; margin-right: 5px; vertical-align: middle;">&nbsp;</span> <span style="vertical-align: middle;">Onshore</span>
        &nbsp;&nbsp;
        <span style="background-color: #90EE90; padding: 4px 10px; border-radius: 4px; display: inline-block; margin-right: 5px; vertical-align: middle;">&nbsp;</span> <span style="vertical-align: middle;">Both</span>
    </p>
    """

    full_html_content = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Upcoming Holidays</title>
        <style>
            body {{
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                line-height: 1.6;
                color: #333333;
                background-color: #f4f7f6;
                margin: 0;
                padding: 0;
            }}
            .email-container {{
                max-width: 700px; /* Increased width slightly for better calendar display */
                margin: 30px auto;
                background-color: #ffffff;
                padding: 30px;
                border-radius: 12px;
                box-shadow: 0 8px 16px rgba(0,0,0,0.1);
                border: 1px solid #e0e0e0;
                box-sizing: border-box;
            }}
            .header {{
                text-align: center;
                padding-bottom: 20px;
                border-bottom: 1px solid #eeeeee;
                margin-bottom: 30px;
            }}
            .header h1 {{
                color: #007bff;
                font-size: 28px;
                margin: 0;
            }}
            .header img {{ /* Basic style for a logo if you add one */
                max-width: 150px; 
                margin-bottom: 10px;
            }}
            .footer {{
                text-align: center;
                padding-top: 30px;
                margin-top: 30px;
                border-top: 1px solid #eeeeee;
                font-size: 14px;
                color: #777777;
            }}
            /* Ensure calendar cells don't break words awkwardly */
            .calendar-table td span {{
                word-break: normal;
                white-space: nowrap;
            }}
        </style>
    </head>
    <body>
        <div class="email-container">
            <div class="header">
                <h1>🎉 Holiday Reminder! 🎉</h1>
            </div>

            {table_html}
            {calendar_html}

            <p style="font-size: 16px; color: #333333; margin-top: 40px;">Wishing you restful and joyful holidays! <br>
            Let's plan deliverables accordingly without affecting Holidays!</p>
            <p style="font-size: 16px; color: #333333;">Best regards,<br/><strong style="color: #007bff;">{signature_name}</strong></p>

            <div class="footer">
                <p>This is an automated reminder. Please do not reply to this email.</p>
                <p>&copy; {datetime.now().year} {company_name_footer}</p>
            </div>
        </div>
    </body>
    </html>
    """
    return full_html_content

# --- Example Usage (for direct testing of this file) ---
if __name__ == "__main__":
    # Dummy holiday data for testing (ensure dates are MM/DD/YYYY for the parser)
    dummy_holidays_data = {
        'Date': [
            '01/01/2025', '01/26/2025', '05/26/2025', '05/27/2025', # May for current_date test
            '06/27/2025', '06/28/2025', '08/15/2025', '05/invalid_date/2025' # Test invalid date
        ],
        'HolidayName': [
            'New Year', 'Republic Day', 'Memorial Day Test', 'TwentySeventh May Test',
            'TwentySeventh June', 'TwentyEighth June', 'Independence Day', 'Bad Date Holiday'
        ],
        'Shore': [
            'Both', 'Offshore', 'Onshore', 'Offshore',
            'Offshore', 'Both', 'Offshore', 'Both'
        ],
        'Locations': [
            'All Offshore & Onshore', 'Bangalore, Mumbai', 'All USA Locations', 'Bangalore, Pune',
            'Bangalore, Pune, Mumbai', 'All Offshore & Onshore', 'Pune', 'All'
        ]
    }
    dummy_holidays_df = pd.DataFrame(dummy_holidays_data)
    # No need to convert 'Date' here, get_holiday_data will do it.

    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        dummy_csv_path = os.path.join(script_dir, 'dummy_holidays.csv')
        dummy_holidays_df.to_csv(dummy_csv_path, index=False)
        print(f"Temporary dummy_holidays.csv created at: {dummy_csv_path}")

        # Test get_holiday_data
        print("\n--- Testing get_holiday_data ---")
        holidays_df_from_csv = get_holiday_data(dummy_csv_path)
        print("Loaded holidays_df_from_csv head:")
        print(holidays_df_from_csv.head())
        print(f"Shape of loaded data: {holidays_df_from_csv.shape}")
        print("--- End testing get_holiday_data ---\n")


        if not holidays_df_from_csv.empty:
            # If testing generate_modern_holiday_email_html, ensure current_date is set to May 2025
            # by uncommenting `# current_date = datetime(2025, 5, 24)` inside the function
            # or pass a specific date to it for testing if you modify the function signature.
            print("Generating HTML with default company/signature names...")
            email_html = generate_modern_holiday_email_html(
                holidays_df_from_csv,
                company_name_footer="Test Corp",
                signature_name="Test HR"
            )
            output_html_file = "modern_holiday_reminder_enhanced.html"
            with open(output_html_file, "w", encoding="utf-8") as f:
                f.write(email_html)
            print(f"Generated {output_html_file} using data from CSV. Open it in a browser to preview.")
        else:
            print("Could not load holiday data from dummy CSV for HTML generation test.")

    except Exception as e:
        print(f"Error during dummy CSV creation/reading for test: {e}")
    finally:
        if 'dummy_csv_path' in locals() and os.path.exists(dummy_csv_path):
            # os.remove(dummy_csv_path) # Comment out to inspect dummy_holidays.csv
            # print("Cleaned up dummy_holidays.csv.")
            print(f"Dummy file {dummy_csv_path} was NOT removed for inspection.")
            pass