import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
from datetime import datetime, timedelta
from apscheduler.schedulers.blocking import BlockingScheduler
import os
import sys
import configparser
import re # <--- NEW: For email validation
import logging # <--- NEW: For logging

# --- Import the email generator module ---
import email_generator

# --- Basic Logging Configuration ---
logging.basicConfig(filename='holiday_tool.log',
                    level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s',
                    datefmt='%Y-%m-%d %H:%M:%S')
logging.info("Holiday Reminder Tool starting up.")

# --- Configuration (READ FROM config.ini) ---
config = configparser.ConfigParser()
config_file_path = 'config.ini'

try:
    if not os.path.exists(config_file_path):
        logging.error(f"Configuration file '{config_file_path}' not found. Please create it.")
        raise FileNotFoundError(f"Configuration file '{config_file_path}' not found. Please create it as described in Step 1.")
    config.read(config_file_path)

    # Read common settings
    SERVICE_PROVIDER = config.get('EMAIL_SETTINGS', 'SERVICE_PROVIDER')
    SENDER_EMAIL = config.get('EMAIL_SETTINGS', 'SENDER_EMAIL')
    SENDER_PASSWORD = config.get('EMAIL_SETTINGS', 'SENDER_PASSWORD')

    # Dynamically set SMTP server and port based on service provider
    if SERVICE_PROVIDER.lower() == 'gmail':
        SMTP_SERVER = 'smtp.gmail.com'
        SMTP_PORT = 587
    elif SERVICE_PROVIDER.lower() == 'outlook':
        SMTP_SERVER = 'smtp.office365.com'
        SMTP_PORT = 587
    else:
        logging.error(f"Unsupported SERVICE_PROVIDER: {SERVICE_PROVIDER}. Must be 'Gmail' or 'Outlook'.")
        raise ValueError(f"Unsupported SERVICE_PROVIDER: {SERVICE_PROVIDER}. Must be 'Gmail' or 'Outlook'.")

    HOLIDAYS_FILE = config.get('FILE_PATHS', 'HOLIDAYS_FILE')
    EMPLOYEES_FILE = config.get('FILE_PATHS', 'EMPLOYEES_FILE')

    # Read email content settings
    COMPANY_NAME_SUBJECT_SUFFIX = config.get('EMAIL_CONTENT', 'COMPANY_NAME_SUBJECT_SUFFIX', fallback="Upcoming Holiday Reminder!")
    COMPANY_NAME_FOOTER = config.get('EMAIL_CONTENT', 'COMPANY_NAME_FOOTER', fallback="Your Company Name")
    SIGNATURE_NAME = config.get('EMAIL_CONTENT', 'SIGNATURE_NAME', fallback="HR Department")


except configparser.Error as e:
    logging.critical(f"Error reading configuration file: {e}")
    print(f"Error reading configuration file: {e}")
    print("Please ensure config.ini is correctly formatted with [EMAIL_SETTINGS], [FILE_PATHS], and [EMAIL_CONTENT] sections.")
    sys.exit(1)
except FileNotFoundError as e:
    logging.critical(f"Configuration file not found: {e}")
    print(e)
    sys.exit(1)
except ValueError as e:
    logging.critical(f"Configuration Error: {e}")
    print(f"Configuration Error: {e}")
    sys.exit(1)
except Exception as e:
    logging.critical(f"An unexpected error occurred while loading configuration: {e}")
    print(f"An unexpected error occurred while loading configuration: {e}")
    sys.exit(1)

def is_valid_email(email):
    """Basic email validation using regex."""
    if not isinstance(email, str):
        return False
    # A common regex for basic email validation
    regex = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    return re.match(regex, email) is not None

def send_email(to_email, subject, html_content):
    """Sends an HTML email."""
    try:
        msg = MIMEMultipart('alternative')
        msg['From'] = SENDER_EMAIL
        msg['To'] = to_email
        msg['Subject'] = subject

        final_html_content = email_generator.clean_string(html_content) # Clean the final HTML just in case
        part = MIMEText(final_html_content, 'html', 'utf-8')
        msg.attach(part)

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.sendmail(SENDER_EMAIL, to_email, msg.as_string())
        print(f"Email sent successfully to {to_email}")
        logging.info(f"Email sent successfully to {to_email}")
    except smtplib.SMTPAuthenticationError as e:
        print(f"Failed to send email to {to_email}. Error: SMTP Authentication failed. Check SENDER_EMAIL and SENDER_PASSWORD (App Password). Details: {e}")
        logging.error(f"SMTP Authentication failed for {to_email}. Check credentials. Details: {e}")
    except smtplib.SMTPServerDisconnected as e:
        print(f"Failed to send email to {to_email}. Error: SMTP server disconnected unexpectedly. Check network or server status. Details: {e}")
        logging.error(f"SMTP server disconnected for {to_email}. Details: {e}")
    except smtplib.SMTPException as e:
        print(f"Failed to send email to {to_email}. Error: An SMTP error occurred. Details: {e}")
        logging.error(f"SMTP error for {to_email}. Details: {e}")
    except Exception as e:
        print(f"Failed to send email to {to_email}. A general error occurred: {e}")
        logging.error(f"General error sending email to {to_email}. Details: {e}")


def send_holiday_reminders():
    """
    Main function to orchestrate reading data, generating email, and sending.
    """
    current_run_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(f"--- Running Holiday Reminder at {current_run_time} ---")
    logging.info(f"--- Running Holiday Reminder scheduled job at {current_run_time} ---")

    # 1. Get holiday data
    holidays_df = email_generator.get_holiday_data(HOLIDAYS_FILE)
    if holidays_df.empty:
        print("No holiday data found or file is empty. Skipping email generation.")
        logging.warning("No holiday data found or file is empty. Skipping email generation.")
        return

    # 2. Generate email HTML content
    email_html_content = email_generator.generate_modern_holiday_email_html(
        holidays_df,
        company_name_footer=COMPANY_NAME_FOOTER,
        signature_name=SIGNATURE_NAME
    )

    # 3. Get employee emails
    try:
        employees_df = pd.read_csv(EMPLOYEES_FILE)
        if 'Email' not in employees_df.columns:
            logging.error(f"Required column 'Email' not found in {EMPLOYEES_FILE}.")
            raise KeyError(f"Required column 'Email' not found in {EMPLOYEES_FILE}. Please check your CSV headers.")
        
        # Clean and validate emails
        valid_recipient_emails = []
        for email in employees_df['Email']:
            cleaned_email = email_generator.clean_string(email)
            if is_valid_email(cleaned_email):
                valid_recipient_emails.append(cleaned_email)
            else:
                print(f"Invalid or empty email format: '{email}'. Skipping.")
                logging.warning(f"Invalid or empty email format: '{email}'. Skipping.")
        
        recipient_emails = valid_recipient_emails

    except FileNotFoundError:
        print(f"Error: Employee file '{EMPLOYEES_FILE}' not found. Cannot send emails.")
        logging.error(f"Employee file '{EMPLOYEES_FILE}' not found.")
        return
    except KeyError as e:
        print(f"Error: Missing expected column in '{EMPLOYEES_FILE}'. {e}")
        logging.error(f"Missing expected column in '{EMPLOYEES_FILE}'. {e}")
        return
    except Exception as e:
        print(f"An unexpected error occurred while reading '{EMPLOYEES_FILE}': {e}")
        logging.error(f"Unexpected error reading '{EMPLOYEES_FILE}': {e}")
        return


    if not recipient_emails:
        print("No valid recipient emails found in employees file. No emails to send.")
        logging.warning("No valid recipient emails found. No emails to send.")
        return

    subject = f"Upcoming Holiday Reminder! - {COMPANY_NAME_SUBJECT_SUFFIX}"
    for email in recipient_emails:
        send_email(email, subject, email_html_content)

    print("--- Holiday Reminder run complete ---")
    logging.info("--- Holiday Reminder run complete ---")

# --- Scheduling the task ---
if __name__ == "__main__":
    print("Starting Holiday Reminder Tool...")
    # Set timezone for India
    scheduler = BlockingScheduler(timezone='Asia/Kolkata')

    # Schedule to run every 14 days at 9:00 AM (for production use)
    # scheduler.add_job(send_holiday_reminders, 'interval', days=14, start_date=datetime(datetime.now().year, datetime.now().month, datetime.now().day, 9, 0, 0))
    # scheduler.add_job(send_holiday_reminders, 'cron', day_of_week='mon', hour=9, minute=0, week='*/2', timezone='Asia/Kolkata')


    # For quick testing (runs 2 seconds after script starts):
    scheduler.add_job(send_holiday_reminders, 'date', run_date=datetime.now() + timedelta(seconds=2))

    print(f"Scheduler initialized. For quick testing, job will run in 2 seconds.")
    print("Press Ctrl+C to exit.")
    logging.info("Scheduler initialized. Test job scheduled.")
    try:
        scheduler.start()
    except (KeyboardInterrupt, SystemExit):
        print("Scheduler stopped.")
        logging.info("Scheduler stopped by user.")
    except Exception as e:
        print(f"Scheduler failed: {e}")
        logging.critical(f"Scheduler failed: {e}", exc_info=True)