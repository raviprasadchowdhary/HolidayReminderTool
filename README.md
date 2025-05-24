# Holiday Reminder Tool 

This tool automatically sends holiday reminders to employees via email. It reads holiday information from a CSV file, generates a nicely formatted HTML email with calendars for the current and next month, and sends it to a list of employees.

## Features

- Reads holiday data from `holidays.csv`.
- Generates attractive HTML emails with current and next month's holidays.
- Displays holidays categorized by "Onshore" and "Offshore" teams, including locations.
- Sends emails to a list of recipients from `employees.csv`.
- Configurable for both Gmail and Outlook/Office 365.
- Email subject, footer company name, and signature name are configurable.
- Automated scheduling (e.g., every 14 days).
- Basic logging to `holiday_tool.log`.
- Basic validation for recipient email addresses.

## Setup Guide

Follow these steps to set up and run the Holiday Reminder Tool:

### 1. Prerequisites

- **Python**: Make sure you have Python installed (version 3.6+ recommended).
- **Required Libraries**: Install dependencies by running:

    ```bash
    pip install pandas apscheduler
    ```

### 2. Configure `config.ini`

This file holds your email settings, file paths, and email content customizations.

- Create a file named `config.ini` in the same folder as `main_tool.py`.
- Copy and paste the following content into `config.ini`:

    ```ini
    [EMAIL_SETTINGS]
    # For Gmail:
    SERVICE_PROVIDER = Gmail
    SENDER_EMAIL = your_gmail_email@gmail.com
    SENDER_PASSWORD = your_gmail_app_password
    # SMTP_SERVER = smtp.gmail.com
    # SMTP_PORT = 587

    # For Outlook/Office 365 (Uncomment these lines and comment out Gmail lines if you want to use Outlook)
    # SERVICE_PROVIDER = Outlook
    # SENDER_EMAIL = your_outlook_email@outlook.com
    # SENDER_PASSWORD = your_outlook_app_password
    # SMTP_SERVER = smtp.office365.com
    # SMTP_PORT = 587

    [FILE_PATHS]
    HOLIDAYS_FILE = holidays.csv
    EMPLOYEES_FILE = employees.csv

    [EMAIL_CONTENT]
    # This suffix will be appended to "Upcoming Holiday Reminder! - "
    COMPANY_NAME_SUBJECT_SUFFIX = Your Company Name Subject
    COMPANY_NAME_FOOTER = Your Full Company Name (for email footer)
    SIGNATURE_NAME = Your Name / Department (for email signature)
    ```

- **Update `SENDER_EMAIL` and `SENDER_PASSWORD`:**
    - For both Gmail and Outlook, use an "App Password" if you have 2-Factor Authentication (2FA) enabled (highly recommended). Your regular email password will likely not work.

#### How to get an App Password

- **Gmail:**
    1. Go to your [Google Account](https://myaccount.google.com/).
    2. Navigate to **Security** → **How you sign in to Google** → **App passwords** (you might need to enable 2-Step Verification first).
    3. Sign in if prompted.
    4. Select **Mail** for the app and **Other** (e.g., "Holiday Tool") for the device.
    5. Click **Generate** and copy the 16-character password. Use this password without spaces.

- **Outlook/Office 365:**
    1. Go to your [Microsoft account security page](https://account.microsoft.com/security).
    2. Select **Advanced security options**.
    3. Under **App passwords**, click **Create a new app password** (if you don't see this, you may need to enable two-step verification).
    4. Copy the generated password.

- **Choose your email provider:**
    - If using Gmail: Ensure `SERVICE_PROVIDER = Gmail` is active and the Outlook lines are commented out (`#`). Fill in your Gmail email and App Password.
    - If using Outlook: Comment out the Gmail lines (`#`). Uncomment the Outlook lines and fill in your Outlook email and App Password.

- **Update Email Content Customizations:**
    - Modify `COMPANY_NAME_SUBJECT_SUFFIX`, `COMPANY_NAME_FOOTER`, and `SIGNATURE_NAME` in the `[EMAIL_CONTENT]` section as needed.

### 3. Prepare `holidays.csv`

This file lists all the holidays.

- Create a file named `holidays.csv` in the same folder as `main_tool.py`.
- Format the file with the following columns: `Date`, `HolidayName`, `Shore`, `Locations`.

    ```csv
    Date,HolidayName,Shore,Locations
    01/01/2025,New Year's Day,Both,"All USA Locations, All India Locations"
    01/26/2025,Republic Day,Offshore,India
    05/26/2025,Memorial Day,Onshore,"USA East, USA West"
    06/15/2025,Company Holiday,Both,All Locations
    12/25/2025,Christmas Day,Both,"USA, India"
    ```

- **Notes:**
    - `Date`: Must be in `MM/DD/YYYY` format (e.g., `05/26/2025` for May 26th, 2025). Rows with invalid dates will be skipped.
    - `HolidayName`: Name of the holiday.
    - `Shore`: Can be `Onshore`, `Offshore`, or `Both`. This determines how holidays are categorized in the email.
    - `Locations`: Specific locations affected by the holiday. If a holiday applies to multiple locations, enclose them in quotes if they contain commas (e.g., `"Bengalore, Pune, Mumbai"`).

### 4. Prepare `employees.csv`

This file lists the email addresses of recipients.

- Create a file named `employees.csv` in the same folder as `main_tool.py`.
- Format the file as follows (Column headers: `Employee ID,Employee Name,Email,Locations` - only `Email` is strictly required by the script for sending, others are for your reference).

    ```csv
    Employee ID,Employee Name,Email,Locations
    100001,Rohan Chowdhary,employee1@example.com,Bangalore
    100002,Rahul Chowdhary,employee2@example.com,Pune
    100003,Raviprasad Chowdhary,another_valid_email@example.com,Onshore
    invalid-email-format,Skipped User,notanemail,Unknown
    100004,Test User,testuser@example.com,Mumbai
    ```

- Ensure each email address is on a new line under the `Email` header. The script will attempt to validate email formats and skip invalid ones.

### 5. Ensure `email_generator.py` is present

Make sure you have the `email_generator.py` file (provided with this tool) in the same directory. This file contains the logic for generating the holiday email's HTML content.

### 6. Run the Tool

- Open your terminal or command prompt.
- Navigate to the directory where you saved all your files (e.g., `main_tool.py`, `config.ini`, etc.):

    ```bash
    cd path\to\your\holiday_tool_folder
    ```

    *(Replace with your actual folder path.)*

- Run the script:

    ```bash
    python main_tool.py
    ```

- The tool is configured by default to run a test email 2 seconds after starting.
- For continuous operation (e.g., every 14 days at 9 AM), you'll need to:
    1.  Comment out the line for quick testing in `main_tool.py`:
        `# scheduler.add_job(send_holiday_reminders, 'date', run_date=datetime.now() + timedelta(seconds=2))`
    2.  Uncomment one of the production scheduler lines, for example:
        `scheduler.add_job(send_holiday_reminders, 'cron', day_of_week='mon', hour=9, minute=0, week='*/2', timezone='Asia/Kolkata')`
        (This example runs the job every other Monday at 9:00 AM Kolkata time. Adjust the cron parameters as needed for your desired schedule.)

---

## Important Notes & Troubleshooting

- **App Passwords**: This is the most common point of failure for email sending. Double-check that you're using an App Password if 2FA is enabled.
- **`config.ini` Format**: Ensure there are no extra spaces or hidden characters, especially around the `=` signs and in the password. The `configparser` module is sensitive to formatting.
- **CSV Headers & Format**:
    - `holidays.csv` must have headers: `Date,HolidayName,Shore,Locations`. Dates must be `MM/DD/YYYY`.
    - `employees.csv` must have an `Email` header.
- **Internet Connection**: The tool needs an active internet connection to send emails.
- **Firewall/Antivirus**: Ensure your firewall or antivirus is not blocking Python from making network connections to SMTP servers (usually ports 587, 465, or 25).
- **Timezone**: The scheduler is set by default to `Asia/Kolkata`. Adjust the `timezone` parameter in `BlockingScheduler(timezone='Your/Timezone')` and in the `cron` job if needed.
- **Log File**: Check `holiday_tool.log` in the same directory as the script for detailed information about its operations and any errors encountered.
- **Testing `email_generator.py`**: You can run `python email_generator.py` directly. It will create a `dummy_holidays.csv`, process it, and generate a `modern_holiday_reminder_enhanced.html` file for previewing the email design. (The `current_date` in `email_generator.py` is fixed to May 2025 for this test to show data; comment this out for real use).

---