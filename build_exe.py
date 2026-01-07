"""
Build script to create standalone executable for Holiday Reminder Tool
This script uses PyInstaller to bundle the application with all dependencies
"""

import subprocess
import sys
import os

def build_executable():
    """Build the standalone executable using PyInstaller"""
    
    print("=" * 60)
    print("Building Holiday Reminder Tool Standalone Executable")
    print("=" * 60)
    
    # PyInstaller command for email_app.py (GUI version)
    pyinstaller_cmd = [
        sys.executable,
        '-m',
        'PyInstaller',
        '--onefile',  # Create a single executable file
        '--windowed',  # No console window (GUI app)
        '--name=HolidayReminderTool',  # Name of the executable
        '--icon=NONE',  # No icon (you can add one later)
        '--add-data=config.ini;.',  # Include config.ini
        '--add-data=holidays.csv;.',  # Include holidays.csv
        '--add-data=Employees.csv;.',  # Include Employees.csv
        '--hidden-import=pandas',
        '--hidden-import=win32com.client',
        '--hidden-import=email_generator',
        '--collect-all=pandas',
        '--collect-all=pywin32',
        'email_app.py'
    ]
    
    print("\nBuilding GUI application (email_app.py)...")
    print(f"Command: {' '.join(pyinstaller_cmd)}\n")
    
    try:
        result = subprocess.run(pyinstaller_cmd, check=True, capture_output=True, text=True)
        print(result.stdout)
        print("\n✓ GUI executable built successfully!")
        print(f"   Location: dist\\HolidayReminderTool.exe")
    except subprocess.CalledProcessError as e:
        print(f"✗ Error building GUI executable:")
        print(e.stderr)
        return False
    
    # Also build the email generator preview tool
    print("\n" + "=" * 60)
    print("Building Email Generator Preview Tool")
    print("=" * 60)
    
    pyinstaller_cmd_preview = [
        sys.executable,
        '-m',
        'PyInstaller',
        '--onefile',  # Create a single executable file
        '--console',  # Keep console window for preview tool
        '--name=EmailGeneratorPreview',
        '--add-data=config.ini;.',
        '--add-data=holidays.csv;.',
        '--add-data=Employees.csv;.',
        '--hidden-import=pandas',
        '--collect-all=pandas',
        'email_generator.py'
    ]
    
    print("\nBuilding preview tool (email_generator.py)...")
    print(f"Command: {' '.join(pyinstaller_cmd_preview)}\n")
    
    try:
        result = subprocess.run(pyinstaller_cmd_preview, check=True, capture_output=True, text=True)
        print(result.stdout)
        print("\n✓ Preview tool executable built successfully!")
        print(f"   Location: dist\\EmailGeneratorPreview.exe")
    except subprocess.CalledProcessError as e:
        print(f"✗ Error building preview executable:")
        print(e.stderr)
        return False
    
    print("\n" + "=" * 60)
    print("BUILD COMPLETE!")
    print("=" * 60)
    print("\nExecutables created in 'dist' folder:")
    print("  1. HolidayReminderTool.exe - GUI application")
    print("  2. EmailGeneratorPreview.exe - Console preview tool")
    print("\nIMPORTANT: Copy these files to the dist folder:")
    print("  - config.ini")
    print("  - holidays.csv")
    print("  - Employees.csv")
    print("\nOr distribute the entire 'dist' folder to users.")
    print("\nUsers can run HolidayReminderTool.exe without Python installed!")
    
    return True

if __name__ == "__main__":
    # Check if PyInstaller is installed
    try:
        import PyInstaller
    except ImportError:
        print("PyInstaller is not installed. Installing now...")
        subprocess.run([sys.executable, '-m', 'pip', 'install', 'pyinstaller'], check=True)
    
    success = build_executable()
    
    if success:
        print("\n✓ Build process completed successfully!")
        sys.exit(0)
    else:
        print("\n✗ Build process failed. Please check the errors above.")
        sys.exit(1)
