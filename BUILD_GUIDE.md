# Building Standalone Executable - Developer Guide

This guide explains how to build the standalone executable from source code.

## Prerequisites

- Python 3.7 or later installed
- All dependencies installed (pandas, pywin32, APScheduler)
- PyInstaller installed (`pip install pyinstaller`)

## Build Process

### Method 1: Automated Build Script (Recommended)

Simply run:
```bash
python build_exe.py
```

This will:
1. Create two standalone executables
2. Copy required data files to dist folder
3. Generate a complete distribution package

### Method 2: Manual PyInstaller Commands

#### Build GUI Application:
```bash
python -m PyInstaller --onefile --windowed --name=HolidayReminderTool ^
    --add-data=config.ini;. ^
    --add-data=holidays.csv;. ^
    --add-data=Employees.csv;. ^
    --hidden-import=pandas ^
    --hidden-import=win32com.client ^
    --hidden-import=email_generator ^
    --collect-all=pandas ^
    --collect-all=pywin32 ^
    email_app.py
```

#### Build Preview Tool:
```bash
python -m PyInstaller --onefile --console --name=EmailGeneratorPreview ^
    --add-data=config.ini;. ^
    --add-data=holidays.csv;. ^
    --add-data=Employees.csv;. ^
    --hidden-import=pandas ^
    --collect-all=pandas ^
    email_generator.py
```

#### Copy Data Files:
```bash
copy config.ini dist\
copy holidays.csv dist\
copy Employees.csv dist\
```

## Output

After building, the `dist` folder will contain:

- **HolidayReminderTool.exe** (~40-50 MB) - GUI application
- **EmailGeneratorPreview.exe** (~35-45 MB) - Preview tool
- **config.ini** - Configuration
- **holidays.csv** - Holiday data
- **Employees.csv** - Recipients
- **README_DISTRIBUTION.txt** - User guide

## Distribution

To distribute to end users:

1. **Zip the entire `dist` folder**
2. Share the zip file with users
3. Users extract and run `HolidayReminderTool.exe`
4. **No Python installation required** on user machines!

## Build Options Explained

- `--onefile`: Creates a single .exe file (vs. folder with dependencies)
- `--windowed`: No console window for GUI app (cleaner UX)
- `--console`: Shows console for preview tool (to see output)
- `--add-data`: Include data files in the executable
- `--hidden-import`: Ensure specific modules are included
- `--collect-all`: Include all submodules of a package

## Testing the Build

1. Navigate to dist folder: `cd dist`
2. Run: `.\HolidayReminderTool.exe`
3. Verify all features work:
   - GUI opens correctly
   - Buttons function properly
   - Email generation works
   - Outlook integration works

## Troubleshooting Build Issues

### Import Errors
- Add missing modules with `--hidden-import=module_name`
- Use `--collect-all=package_name` for complex packages

### Large File Size
- This is normal (40-50 MB) due to pandas and dependencies
- Can't be reduced significantly without compromising functionality

### Antivirus Warnings
- Normal for unsigned executables
- Users need to whitelist or click "Run anyway"

### Missing Data Files
- Ensure all `--add-data` paths are correct
- Use semicolon (;) on Windows, colon (:) on Linux/Mac

## Clean Build

To rebuild from scratch:
```bash
# Remove old build artifacts
rmdir /s /q build dist
del /q *.spec

# Run build again
python build_exe.py
```

## Advanced: Add Custom Icon

1. Create or download a `.ico` file
2. Replace `--icon=NONE` with `--icon=path\to\icon.ico`
3. Rebuild

## Version Management

Update version in README_DISTRIBUTION.txt after each build.

Current: Version 2.0 (Standalone)

---

**Built with PyInstaller - Making Python apps distributable since 2005!**
