import tkinter as tk
from tkinter import ttk, messagebox
import configparser
import sys
import os
import webbrowser
import urllib.parse
import tempfile
import subprocess
import importlib

# Import our existing email generator
import email_generator
# Reload to get latest changes
importlib.reload(email_generator)

class HolidayEmailApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Holiday Reminder Email Generator")
        self.root.geometry("700x500")
        self.root.resizable(True, True)
        
        # Configure style
        style = ttk.Style()
        style.theme_use('clam')  # Use a more modern theme
        
        # Load configuration
        self.load_config()
        
        # Create UI
        self.create_widgets()
        
    def load_config(self):
        """Load configuration from config.ini"""
        config = configparser.ConfigParser()
        config_file_path = 'config.ini'
        
        try:
            if not os.path.exists(config_file_path):
                messagebox.showerror("Error", f"Configuration file '{config_file_path}' not found.")
                sys.exit(1)
            
            config.read(config_file_path)
            
            self.holidays_file = config.get('FILE_PATHS', 'HOLIDAYS_FILE')
            self.employees_file = config.get('FILE_PATHS', 'EMPLOYEES_FILE')
            self.company_name_subject_suffix = config.get('EMAIL_CONTENT', 'COMPANY_NAME_SUBJECT_SUFFIX', 
                                                          fallback="Upcoming Holiday Reminder!")
            self.company_name_footer = config.get('EMAIL_CONTENT', 'COMPANY_NAME_FOOTER', 
                                                  fallback="Your Company Name")
            self.signature_name = config.get('EMAIL_CONTENT', 'SIGNATURE_NAME', 
                                            fallback="HR Department")
            
        except Exception as e:
            messagebox.showerror("Configuration Error", f"Error loading config.ini: {e}")
            sys.exit(1)
    
    def create_widgets(self):
        """Create the GUI widgets"""
        # Header
        header_frame = ttk.Frame(self.root, padding="20")
        header_frame.pack(fill=tk.X)
        
        title_label = ttk.Label(header_frame, text="🎉 Holiday Reminder Email Generator", 
                               font=('Arial', 16, 'bold'))
        title_label.pack()
        
        subtitle_label = ttk.Label(header_frame, 
                                   text="Generate holiday email and open in your email client",
                                   font=('Arial', 10))
        subtitle_label.pack()
        
        # Info Frame
        info_frame = ttk.LabelFrame(self.root, text="Configuration", padding="15")
        info_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Display config info
        ttk.Label(info_frame, text=f"Holiday File: {self.holidays_file}").pack(anchor=tk.W, pady=2)
        ttk.Label(info_frame, text=f"Employee File: {self.employees_file}").pack(anchor=tk.W, pady=2)
        ttk.Label(info_frame, text=f"Company: {self.company_name_footer}").pack(anchor=tk.W, pady=2)
        ttk.Label(info_frame, text=f"Signature: {self.signature_name}").pack(anchor=tk.W, pady=2)
        
        ttk.Separator(info_frame, orient='horizontal').pack(fill=tk.X, pady=10)
        
        subject_label = ttk.Label(info_frame, text="Email Subject:", font=('Arial', 9, 'bold'))
        subject_label.pack(anchor=tk.W, pady=2)
        
        subject_text = ttk.Label(info_frame, 
                                text=f"Upcoming Holiday Reminder! - {self.company_name_subject_suffix}",
                                foreground='blue')
        subject_text.pack(anchor=tk.W, pady=2)
        
        # Status label
        self.status_label = ttk.Label(info_frame, text="", foreground="green")
        self.status_label.pack(pady=10)
        
        # Buttons Frame
        button_frame = ttk.Frame(self.root, padding="20")
        button_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Create a grid layout for buttons
        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=1)
        
        # Generate and Open in Email Client button (main action)
        self.open_email_btn = tk.Button(button_frame, 
                                        text="📧 Generate & Open in Email Client",
                                        command=self.open_in_email_client,
                                        bg='#007bff',
                                        fg='white',
                                        font=('Arial', 11, 'bold'),
                                        height=2,
                                        cursor='hand2',
                                        relief=tk.RAISED,
                                        bd=3)
        self.open_email_btn.grid(row=0, column=0, columnspan=2, sticky='ew', pady=5, padx=5)
        
        # Preview in Browser button
        self.preview_btn = tk.Button(button_frame, 
                                     text="👁️ Preview in Browser",
                                     command=self.preview_in_browser,
                                     bg='#28a745',
                                     fg='white',
                                     font=('Arial', 10),
                                     height=2,
                                     cursor='hand2',
                                     relief=tk.RAISED,
                                     bd=2)
        self.preview_btn.grid(row=1, column=0, sticky='ew', pady=5, padx=5)
        
        # Save Files button
        self.save_btn = tk.Button(button_frame, 
                                  text="💾 Save Files",
                                  command=self.save_files,
                                  bg='#ffc107',
                                  fg='black',
                                  font=('Arial', 10),
                                  height=2,
                                  cursor='hand2',
                                  relief=tk.RAISED,
                                  bd=2)
        self.save_btn.grid(row=1, column=1, sticky='ew', pady=5, padx=5)
        
        # Exit button
        exit_btn = tk.Button(button_frame, 
                             text="❌ Exit",
                             command=self.root.quit,
                             bg='#dc3545',
                             fg='white',
                             font=('Arial', 10),
                             height=2,
                             cursor='hand2',
                             relief=tk.RAISED,
                             bd=2)
        exit_btn.grid(row=2, column=0, columnspan=2, sticky='ew', pady=5, padx=5)
        
    def generate_email_content(self):
        """Generate the email HTML content"""
        try:
            # Load holiday data
            holidays_df = email_generator.get_holiday_data(self.holidays_file)
            
            if holidays_df.empty:
                messagebox.showerror("Error", "No holiday data found or file is empty.")
                return None
            
            # Generate email HTML
            email_html = email_generator.generate_modern_holiday_email_html(
                holidays_df,
                company_name_footer=self.company_name_footer,
                signature_name=self.signature_name
            )
            
            return email_html
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate email content: {e}")
            return None
    
    def open_in_email_client(self):
        """Generate email and open in default email client"""
        self.status_label.config(text="Generating email content...", foreground="blue")
        self.root.update()
        
        email_html = self.generate_email_content()
        if not email_html:
            self.status_label.config(text="Failed to generate email", foreground="red")
            return
        
        # Extract body content
        import re
        body_match = re.search(r'<body>(.*?)</body>', email_html, re.DOTALL)
        if body_match:
            email_body = body_match.group(1).strip()
        else:
            email_body = email_html
        
        # Try Outlook COM automation (Windows only)
        if sys.platform == 'win32':
            try:
                import win32com.client
                
                # Create Outlook instance
                outlook = win32com.client.Dispatch("Outlook.Application")
                
                # Create new mail item
                mail = outlook.CreateItem(0)  # 0 = olMailItem
                
                # Set properties
                mail.Subject = f"Upcoming Holiday Reminder! - {self.company_name_subject_suffix}"
                mail.HTMLBody = email_body
                mail.To = ""  # Empty - user will fill
                mail.CC = ""  # Empty
                mail.BCC = ""  # Empty
                
                # Display the email in compose mode
                mail.Display(True)  # True = non-modal (doesn't block)
                
                self.status_label.config(text="✓ Email opened in Outlook (compose mode)!", foreground="green")
                messagebox.showinfo("Success", 
                                  "Email draft opened in Outlook compose window!\n\n"
                                  "The email is ready to edit:\n"
                                  "- Fill in To: recipient email addresses\n"
                                  "- Add CC/BCC if needed\n"
                                  "- Review the content\n"
                                  "- Click Send when ready!\n\n"
                                  "All formatting (colors, shadows, calendars) is preserved!")
                return
                
            except ImportError:
                messagebox.showerror("Error", 
                                   "pywin32 is not installed.\n\n"
                                   "Please install it by running:\n"
                                   "pip install pywin32\n\n"
                                   "Then restart the application.")
                self.status_label.config(text="pywin32 not installed", foreground="red")
                return
                
            except Exception as outlook_error:
                error_msg = str(outlook_error)
                messagebox.showerror("Outlook Error", 
                                   f"Failed to open Outlook:\n{error_msg}\n\n"
                                   "Make sure Microsoft Outlook is installed.\n\n"
                                   "Using fallback method instead...")
                # Fall through to .eml method
        
        # Fallback: Save as .eml and open
        try:
            temp_dir = tempfile.gettempdir()
            eml_path = os.path.join(temp_dir, "holiday_reminder.eml")
            
            from email.mime.multipart import MIMEMultipart
            from email.mime.text import MIMEText
            
            msg = MIMEMultipart('alternative')
            msg['Subject'] = f"Upcoming Holiday Reminder! - {self.company_name_subject_suffix}"
            msg['From'] = ''
            msg['To'] = ''
            
            html_part = MIMEText(email_body, 'html', 'utf-8')
            msg.attach(html_part)
            
            with open(eml_path, 'w', encoding='utf-8') as f:
                f.write(msg.as_string())
            
            os.startfile(eml_path)
            
            self.status_label.config(text="✓ Email file opened!", foreground="green")
            messagebox.showinfo("Success", 
                              "Email draft file opened!\n\n"
                              "Fill in To/From/CC/BCC and send!")
            
        except Exception as e:
            self.status_label.config(text="Failed to open email", foreground="red")
            messagebox.showerror("Error", 
                               f"Could not open email:\n{e}\n\n"
                               "Use 'Save Files' button instead.")
    
    def preview_in_browser(self):
        """Generate and preview email in browser"""
        self.status_label.config(text="Generating preview...", foreground="blue")
        self.root.update()
        
        email_html = self.generate_email_content()
        if not email_html:
            self.status_label.config(text="Failed to generate email", foreground="red")
            return
        
        try:
            # Save to temp file and open in browser
            temp_dir = tempfile.gettempdir()
            preview_path = os.path.join(temp_dir, "holiday_email_preview.html")
            
            with open(preview_path, 'w', encoding='utf-8') as f:
                f.write(email_html)
            
            webbrowser.open('file://' + preview_path)
            
            self.status_label.config(text="✓ Preview opened in browser!", foreground="green")
            
        except Exception as e:
            self.status_label.config(text="Failed to open preview", foreground="red")
            messagebox.showerror("Error", f"Failed to open preview: {e}")
    
    def save_files(self):
        """Save email files to current directory"""
        self.status_label.config(text="Saving files...", foreground="blue")
        self.root.update()
        
        email_html = self.generate_email_content()
        if not email_html:
            self.status_label.config(text="Failed to generate email", foreground="red")
            return
        
        try:
            # Extract body content
            import re
            body_match = re.search(r'<body>(.*?)</body>', email_html, re.DOTALL)
            if body_match:
                email_body = body_match.group(1).strip()
            else:
                email_body = email_html
            
            # Save HTML preview
            with open("holiday_email_preview.html", 'w', encoding='utf-8') as f:
                f.write(email_html)
            
            # Save .eml file
            from email.mime.multipart import MIMEMultipart
            from email.mime.text import MIMEText
            
            msg = MIMEMultipart('alternative')
            msg['Subject'] = f"Upcoming Holiday Reminder! - {self.company_name_subject_suffix}"
            msg['From'] = ''
            msg['To'] = ''
            
            html_part = MIMEText(email_body, 'html', 'utf-8')
            msg.attach(html_part)
            
            with open("holiday_reminder_draft.eml", 'w', encoding='utf-8') as f:
                f.write(msg.as_string())
            
            self.status_label.config(text="✓ Files saved successfully!", foreground="green")
            messagebox.showinfo("Success", 
                              "Files saved:\n"
                              "- holiday_email_preview.html (preview)\n"
                              "- holiday_reminder_draft.eml (email draft)\n\n"
                              "Double-click the .eml file to open in your email client!")
            
        except Exception as e:
            self.status_label.config(text="Failed to save files", foreground="red")
            messagebox.showerror("Error", f"Failed to save files: {e}")


def main():
    root = tk.Tk()
    app = HolidayEmailApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
