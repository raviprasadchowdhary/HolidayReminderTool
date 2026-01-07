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
        """Optimized: Use itertuples() instead of iterrows() for better performance"""
        if shore_type == 'Onshore':
            filtered = month_df[
                (month_df['Shore'] == 'Onshore') | (month_df['Shore'] == 'Both')
            ]
            if filtered.empty:
                return '<li style="font-size: 14px; color: #777777;">No Onshore Holidays</li>'
            return ''.join([
                f'<li style="margin-bottom: 8px; font-size: 14px; color: #555555;"><strong>{row.Date.strftime("%b %d")}</strong>: {row.HolidayName} '
                f'<span style="color: #888888;">({row.Locations})</span></li>'
                for row in filtered.itertuples()
            ])
        else:  # Offshore
            filtered = month_df[
                (month_df['Shore'] == 'Offshore') | (month_df['Shore'] == 'Both')
            ]
            if filtered.empty:
                return '<li style="font-size: 14px; color: #777777;">No Offshore Holidays</li>'
            return ''.join([
                f'<li style="margin-bottom: 8px; font-size: 14px; color: #555555;"><strong>{row.Date.strftime("%b %d")}</strong>: {row.HolidayName} '
                f'<span style="color: #888888;">({row.Locations})</span></li>'
                for row in filtered.itertuples()
            ])

    table_html = f"""
    <p style="font-size: 16px; color: #333333; margin-bottom: 20px; text-align: left;">Hi Team,</p>
    <p style="font-size: 16px; color: #333333; margin-bottom: 30px; text-align: left;">Here are the upcoming holidays for <strong style="color: #007bff;">{current_month_name}</strong> and <strong style="color: #007bff;">{next_month_name}</strong>:</p>

    <div style="width: 100%; text-align: center;">
    <table border="0" cellpadding="0" cellspacing="0" style="display: inline-block; width: auto; min-width: 600px; max-width: 90%; margin: 0 auto 30px auto; border-radius: 8px; overflow: hidden; background-color: #ffffff; box-shadow: 0 4px 8px rgba(0,0,0,0.05);">
        <tr>
            <th colspan="2" style="background-color:#e9ecef; padding: 15px; font-size: 18px; color: #333333; text-align: center; border-bottom: 1px solid #dee2e6;">{current_month_name} Holidays</th>
            <th colspan="2" style="background-color:#e9ecef; padding: 15px; font-size: 18px; color: #333333; text-align: center; border-bottom: 1px solid #dee2e6;">{next_month_name} Holidays</th>
        </tr>
        <tr>
            <td style="vertical-align: top; padding: 20px; border-right: 1px solid #dee2e6; border-bottom: 1px solid #dee2e6; width: 25%;">
                <h4 style="margin-top:0; color: #007bff; font-size: 16px; text-align: center;">Onshore</h4>
                <ul style="list-style-type:none; padding:0; margin:0;">
                    {get_filtered_holidays_for_table(current_month_holidays_all, 'Onshore')}
                </ul>
            </td>
            <td style="vertical-align: top; padding: 20px; border-bottom: 1px solid #dee2e6; width: 25%;">
                <h4 style="margin-top:0; color: #28a745; font-size: 16px; text-align: center;">Offshore</h4>
                <ul style="list-style-type:none; padding:0; margin:0;">
                    {get_filtered_holidays_for_table(current_month_holidays_all, 'Offshore')}
                </ul>
            </td>
            <td style="vertical-align: top; padding: 20px; border-right: 1px solid #dee2e6; border-bottom: 1px solid #dee2e6; width: 25%;">
                <h4 style="margin-top:0; color: #007bff; font-size: 16px; text-align: center;">Onshore</h4>
                <ul style="list-style-type:none; padding:0; margin:0;">
                    {get_filtered_holidays_for_table(next_month_holidays_all, 'Onshore')}
                </ul>
            </td>
            <td style="vertical-align: top; padding: 20px; border-bottom: 1px solid #dee2e6; width: 25%;">
                <h4 style="margin-top:0; color: #28a745; font-size: 16px; text-align: center;">Offshore</h4>
                <ul style="list-style-type:none; padding:0; margin:0;">
                    {get_filtered_holidays_for_table(next_month_holidays_all, 'Offshore')}
                </ul>
            </td>
        </tr>
    </table>
    </div>
    """

    calendar_html = f"""
    <p style="font-size: 16px; color: #333333; margin-top: 40px; text-align: left;">A quick look at your holiday calendars:</p>
    <div style="width: 100%; text-align: center;">
    <table border="0" cellpadding="0" cellspacing="0" style="display: inline-block; width: auto; min-width: 600px; max-width: 90%; margin: 0 auto 30px auto; border-radius: 8px; overflow: hidden; background-color: #ffffff; box-shadow: 0 4px 8px rgba(0,0,0,0.05);">
        <tr>
            <th colspan="2" style="background-color:#e9ecef; padding: 15px; font-size: 18px; color: #333333; text-align: center; border-bottom: 1px solid #dee2e6;">{current_month_name}</th>
            <th colspan="2" style="background-color:#e9ecef; padding: 15px; font-size: 18px; color: #333333; text-align: center; border-bottom: 1px solid #dee2e6;">{next_month_name}</th>
        </tr>
        <tr>
    """
    
    # Generate calendar content for both months
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

        cal = calendar.HTMLCalendar(calendar.SUNDAY)
        month_cal_html = cal.formatmonth(year, month)

        # Optimize: Clean up calendar HTML in single pass
        replacements = {
            'border="0"': '',
            'cellpadding="0"': '',
            'cellspacing="0"': '',
            'class="month"': ''
        }
        for old, new in replacements.items():
            month_cal_html = month_cal_html.replace(old, new)

        # Optimize: Highlight holiday dates (only process days that exist in the month)
        import calendar as cal_module
        max_day = cal_module.monthrange(year, month)[1]
        
        for day in range(1, max_day + 1):
            day_str_exact = f'>{day}<'
            day_str_with_attr = f'">{day}<'

            color = None
            if day in both_dates:
                color = '#90EE90'
            elif day in only_onshore:
                color = '#FFD700'
            elif day in only_offshore:
                color = '#ADD8E6'
            
            if color:
                replace_with = f'><span style="background-color: {color}; color: #333; font-weight: bold; padding: 4px 6px; border-radius: 4px; display: inline-block;">{day}</span><'
                month_cal_html = month_cal_html.replace(day_str_exact, replace_with, 1)
                month_cal_html = month_cal_html.replace(day_str_with_attr, f'"{replace_with}', 1)
        
        month_cal_html = month_cal_html.replace(
            f'<th colspan="7" class="month">{month_date.strftime("%B %Y")}</th>',
            f'<th colspan="7" style="text-align:center; font-size: 16px; padding:10px 0; background-color: #f0f0f0; color: #333;">{month_date.strftime("%B %Y")}</th>'
        )

        month_cal_html = month_cal_html.replace(
            '<table',
            '<table style="width:100%; border-collapse: collapse; font-size: 15px; text-align: center; border: 0;"'
        ).replace(
            '<th>',
            '<th style="background-color: #f8f9fa; padding: 10px; color: #555555; text-align: center; border-bottom: 1px solid #e0e0e0; border-top: 1px solid #e0e0e0; font-size: 14px; font-weight: bold;"'
        ).replace(
            '<td>',
            '<td style="padding: 10px; border: 1px solid #e0e0e0; text-align: center; vertical-align: middle; height: 40px; font-size: 14px;"'
        )

        # Determine background color based on month
        bg_color = "#f8f9fa" if month_date == current_date else "#e9ecef"
        
        calendar_html += f"""
            <td colspan="2" style="vertical-align: top; padding: 20px; border-bottom: 1px solid #dee2e6;">
                <div style="background-color: {bg_color}; border-radius: 6px; padding: 20px;">
                    {month_cal_html}
                </div>
            </td>
        """
    
    calendar_html += """
        </tr>
    </table>
    </div>
    """
    calendar_html += """
    <div style="margin-top: 30px; text-align: center;">
        <table border="0" cellpadding="0" cellspacing="0" style="display: inline-block; margin: 0 auto;">
            <tr>
                <td style="padding: 8px 5px; text-align: center;">
                    <table border="0" cellpadding="0" cellspacing="0" style="display: inline-block;">
                        <tr>
                            <td bgcolor="#ADD8E6" width="40" height="20" style="border: 1px solid #87CEEB; background-color: #ADD8E6;">&nbsp;</td>
                        </tr>
                    </table>
                </td>
                <td style="padding: 8px 10px 8px 5px; text-align: left;">
                    <span style="font-size: 14px; color: #555555; font-weight: 500;">Offshore</span>
                </td>
                <td style="padding: 8px 5px; text-align: center;">
                    <table border="0" cellpadding="0" cellspacing="0" style="display: inline-block;">
                        <tr>
                            <td bgcolor="#FFD700" width="40" height="20" style="border: 1px solid #DAA520; background-color: #FFD700;">&nbsp;</td>
                        </tr>
                    </table>
                </td>
                <td style="padding: 8px 10px 8px 5px; text-align: left;">
                    <span style="font-size: 14px; color: #555555; font-weight: 500;">Onshore</span>
                </td>
                <td style="padding: 8px 5px; text-align: center;">
                    <table border="0" cellpadding="0" cellspacing="0" style="display: inline-block;">
                        <tr>
                            <td bgcolor="#90EE90" width="40" height="20" style="border: 1px solid #7CCD7C; background-color: #90EE90;">&nbsp;</td>
                        </tr>
                    </table>
                </td>
                <td style="padding: 8px 10px 8px 5px; text-align: left;">
                    <span style="font-size: 14px; color: #555555; font-weight: 500;">Both</span>
                </td>
            </tr>
        </table>
    </div>
    """

    full_html_content = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
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
                max-width: 900px;
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
                text-align: center;
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
            <div class="header" style="text-align: center;">
                <h1 style="color: #007bff; font-size: 28px; margin: 0; text-align: center;">🎉 Holiday Reminder! 🎉</h1>
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
    import configparser
    import sys
    
    # Load configuration from config.ini (same as main_tool.py)
    config = configparser.ConfigParser()
    config_file_path = 'config.ini'
    
    try:
        if not os.path.exists(config_file_path):
            print(f"Error: Configuration file '{config_file_path}' not found.")
            print("Please create config.ini as described in the README.")
            sys.exit(1)
        
        config.read(config_file_path)
        
        # Read file paths from config
        HOLIDAYS_FILE = config.get('FILE_PATHS', 'HOLIDAYS_FILE')
        EMPLOYEES_FILE = config.get('FILE_PATHS', 'EMPLOYEES_FILE')
        
        # Read email content settings from config
        COMPANY_NAME_SUBJECT_SUFFIX = config.get('EMAIL_CONTENT', 'COMPANY_NAME_SUBJECT_SUFFIX', fallback="Upcoming Holiday Reminder!")
        COMPANY_NAME_FOOTER = config.get('EMAIL_CONTENT', 'COMPANY_NAME_FOOTER', fallback="Your Company Name")
        SIGNATURE_NAME = config.get('EMAIL_CONTENT', 'SIGNATURE_NAME', fallback="HR Department")
        
        print("--- Holiday Email Generator (Preview Mode) ---")
        print(f"Loading holidays from: {HOLIDAYS_FILE}")
        print(f"Employee list from: {EMPLOYEES_FILE}")
        print(f"Company: {COMPANY_NAME_FOOTER}")
        print(f"Signature: {SIGNATURE_NAME}\n")
        
        # Load actual holiday data
        holidays_df = get_holiday_data(HOLIDAYS_FILE)
        
        if holidays_df.empty:
            print("Error: No holiday data found or file is empty.")
            sys.exit(1)
        
        print(f"Loaded {len(holidays_df)} holidays from {HOLIDAYS_FILE}")
        print("\nHoliday Data Preview:")
        print(holidays_df.to_string(index=False))
        print()
        
        # Load employee data for preview
        try:
            employees_df = pd.read_csv(EMPLOYEES_FILE)
            if 'Email' in employees_df.columns:
                recipient_count = len(employees_df)
                print(f"Found {recipient_count} recipient(s) in {EMPLOYEES_FILE}")
                print("Recipients:", ', '.join(employees_df['Email'].astype(str).tolist()))
                print()
            else:
                print(f"Warning: 'Email' column not found in {EMPLOYEES_FILE}")
        except FileNotFoundError:
            print(f"Warning: Employee file '{EMPLOYEES_FILE}' not found.")
        except Exception as e:
            print(f"Warning: Could not read employee file: {e}")
        
        # Generate HTML email
        print("Generating HTML email preview...")
        email_html = generate_modern_holiday_email_html(
            holidays_df,
            company_name_footer=COMPANY_NAME_FOOTER,
            signature_name=SIGNATURE_NAME
        )
        
        # Save full HTML to file for browser preview
        output_html_file = "holiday_email_preview.html"
        with open(output_html_file, "w", encoding="utf-8") as f:
            f.write(email_html)
        
        # Extract just the body content
        import re
        body_match = re.search(r'<body>(.*?)</body>', email_html, re.DOTALL)
        if body_match:
            email_body_only = body_match.group(1).strip()
        else:
            email_body_only = email_html
        
        # Create .eml file that can be opened directly in email clients
        from email.mime.multipart import MIMEMultipart
        from email.mime.text import MIMEText
        
        eml_msg = MIMEMultipart('alternative')
        eml_msg['Subject'] = f"Upcoming Holiday Reminder! - {COMPANY_NAME_SUBJECT_SUFFIX}"
        eml_msg['From'] = 'holidays@company.com'  # Placeholder
        eml_msg['To'] = 'recipients@company.com'  # Placeholder
        
        # Attach HTML content
        html_part = MIMEText(email_body_only, 'html', 'utf-8')
        eml_msg.attach(html_part)
        
        # Save as .eml file
        output_eml_file = "holiday_reminder_draft.eml"
        with open(output_eml_file, "w", encoding="utf-8") as f:
            f.write(eml_msg.as_string())
        
        print(f"\n✓ Email preview generated successfully!")
        print(f"\n📧 EASIEST METHOD - Open Draft Email:")
        print(f"   1. Double-click: {output_eml_file}")
        print(f"   2. This will open in your default email client (Outlook/Thunderbird)")
        print(f"   3. Update the To/From addresses and send!")
        print(f"   → All formatting, colors, and shadows will be preserved!")
        print(f"\n📄 Files created:")
        print(f"   - {output_eml_file} (Double-click to open as draft email)")
        print(f"   - {output_html_file} (Browser preview)")
        print(f"\nEmail Subject: Upcoming Holiday Reminder! - {COMPANY_NAME_SUBJECT_SUFFIX}")
        print(f"\nNOTE: If .eml doesn't open, open {output_html_file} in browser,")
        print(f"      then use browser File > Print > Save as PDF and attach to email,")
        
    except configparser.Error as e:
        print(f"Error reading configuration file: {e}")
        print("Please ensure config.ini is correctly formatted.")
        sys.exit(1)
    except FileNotFoundError as e:
        print(f"File not found: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)