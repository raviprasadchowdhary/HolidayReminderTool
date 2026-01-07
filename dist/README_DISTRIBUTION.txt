# Holiday Reminder Tool - Standalone Distribution

## Quick Start (No Python Required!)

This is a **standalone version** of the Holiday Reminder Tool. No Python installation needed!

### What's Included:

- **HolidayReminderTool.exe** - Main GUI application
- **EmailGeneratorPreview.exe** - Console tool to generate email previews
- **config.ini** - Configuration file
- **holidays.csv** - Holiday data
- **Employees.csv** - Employee/recipient list

---

## How to Use

### Option 1: GUI Application (Recommended)

1. **Double-click `HolidayReminderTool.exe`**
2. The Holiday Reminder application window will open
3. Click **"Generate & Open in Email Client"** to create and open the holiday email in Outlook
4. Or click **"Preview in Browser"** to see how the email looks

### Option 2: Console Preview Tool

1. **Double-click `EmailGeneratorPreview.exe`**
2. Email preview files will be generated:
   - `holiday_email_preview.html` - Open in browser
   - `holiday_reminder_draft.eml` - Open in email client

---

## Configuration

### Update Holiday Data

Edit **holidays.csv** to add/modify holidays:
```csv
Date,HolidayName,Shore,Locations
01/01/2026,New Year,Both,All Near On & Offshore locations
01/20/2026,Martin Luther King Jr Day,Onshore,All Near & Onshore locations
```

### Update Recipients

Edit **Employees.csv** to add/modify email recipients:
```csv
Name,Email,Department,Location
John Doe,john.doe@company.com,IT,Bangalore
```

### Update Company Settings

Edit **config.ini** to customize:
- Company name
- Email signature
- SMTP settings (for automated sending)

---

## System Requirements

- **Windows 7 or later**
- **Microsoft Outlook** (for "Open in Email Client" feature)
- Internet connection (for sending emails)

---

## Features

✅ **No Python Installation Required** - Runs as standalone .exe  
✅ **Beautiful HTML Email Templates** - Professional holiday reminders  
✅ **Visual Calendar** - Color-coded holidays (Onshore/Offshore/Both)  
✅ **Outlook Integration** - Opens draft email directly in Outlook  
✅ **Customizable** - Edit CSV files to update holidays and recipients  
✅ **Preview Mode** - See email in browser before sending  

---

## Troubleshooting

### "Windows protected your PC" message
- Click **"More info"** → **"Run anyway"**
- This is normal for unsigned executables

### Outlook doesn't open
- Make sure Microsoft Outlook is installed
- Try using "Preview in Browser" instead
- Or double-click the generated `.eml` file

### Email colors not showing in Outlook
- This is normal - Outlook strips some CSS
- The color blocks should still appear
- Use "Preview in Browser" to see full formatting

---

## File Structure

```
dist/
├── HolidayReminderTool.exe      # Main GUI application
├── EmailGeneratorPreview.exe    # Console preview tool
├── config.ini                   # Configuration settings
├── holidays.csv                 # Holiday data
├── Employees.csv                # Recipient list
└── README_DISTRIBUTION.txt      # This file
```

---

## Support

For issues or questions, contact your IT administrator or the tool developer.

---

## Version Information

- **Version:** 2.0 (Standalone)
- **Build Date:** January 2026
- **Platform:** Windows (64-bit)
- **Built with:** PyInstaller

---

**Enjoy your Holiday Reminder Tool!** 🎉
