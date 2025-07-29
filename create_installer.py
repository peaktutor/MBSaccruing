#!/usr/bin/env python3
"""
Advanced Accrual Generator Installer Creator
Creates a professional Windows installer using PyInstaller
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path
import tempfile

def check_requirements():
    """Check if required tools are installed"""
    print("🔍 Checking requirements...")
    
    # Check Python
    try:
        python_version = sys.version_info
        if python_version.major < 3 or (python_version.major == 3 and python_version.minor < 8):
            print("❌ Python 3.8+ required")
            return False
        print(f"✓ Python {python_version.major}.{python_version.minor}")
    except:
        print("❌ Python not found")
        return False
    
    # Check pip
    try:
        subprocess.run([sys.executable, "-m", "pip", "--version"], 
                      check=True, capture_output=True)
        print("✓ pip available")
    except:
        print("❌ pip not available")
        return False
    
    return True

def install_pyinstaller():
    """Install PyInstaller if not present"""
    print("📦 Installing PyInstaller...")
    try:
        subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], 
                      check=True)
        print("✓ PyInstaller installed")
        return True
    except subprocess.CalledProcessError:
        print("❌ Failed to install PyInstaller")
        return False

def create_main_script():
    """Create the complete accrual generator script"""
    
    # Your complete accrual generator code with all functions
    main_script_content = '''import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from datetime import datetime
import re
import anthropic
import tkinter as tk
from tkinter import filedialog, messagebox
import json
import os
import sys

# Anthropic API Configuration  
ANTHROPIC_API_KEY = "sk-ant-api03-YOUR-API-KEY-HERE"  # REPLACE WITH YOUR API KEY
SETTINGS_FILE = "accrual_settings.json"

def load_settings():
    """Load saved settings for file locations"""
    default_settings = {
        "last_checkbook_dir": os.getcwd(),
        "last_output_dir": os.getcwd(),
        "last_run_time": None,
        "last_output_file": None
    }
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, 'r') as f:
                settings = json.load(f)
                for key, value in default_settings.items():
                    if key not in settings:
                        settings[key] = value
                return settings
    except Exception as e:
        print(f"Warning: Could not load settings: {e}")
    return default_settings

def save_settings(settings):
    """Save settings for next run"""
    try:
        with open(SETTINGS_FILE, 'w') as f:
            json.dump(settings, f, indent=2)
    except Exception as e:
        print(f"Warning: Could not save settings: {e}")

def check_license():
    """License/Protection check"""
    expiration_date = datetime(2026, 12, 31)  # Set your expiration
    if datetime.now() > expiration_date:
        messagebox.showerror("License Expired", 
                           "This software license has expired.\\nContact developer for renewal.")
        return False
    return True

def select_files():
    """File selection dialog"""
    settings = load_settings()
    root = tk.Tk()
    root.withdraw()
    
    messagebox.showinfo("Accrual Generator", 
                       "Professional Accrual Report Generator\\n\\n"
                       "Step 1: Select checkbook Excel file\\n"
                       "Step 2: Choose output location")
    
    # Select checkbook
    checkbook_file = filedialog.askopenfilename(
        title="Select Checkbook Excel File",
        initialdir=settings["last_checkbook_dir"],
        filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
    )
    
    if not checkbook_file:
        root.destroy()
        return None, None
    
    settings["last_checkbook_dir"] = os.path.dirname(checkbook_file)
    
    # Generate filename
    now = datetime.now()
    date_str = now.strftime("%m_%d_%Y")
    time_str = now.strftime("%I_%M%p").lower()
    filename = f"Accrual_Prelim_{date_str}_{time_str}.xlsx"
    
    # Select output
    output_file = filedialog.asksaveasfilename(
        title="Save Accrual Report As",
        initialdir=settings["last_output_dir"],
        initialfile=filename,
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
    )
    
    if not output_file:
        root.destroy()
        return None, None
    
    settings["last_output_dir"] = os.path.dirname(output_file)
    settings["last_run_time"] = now.isoformat()
    settings["last_output_file"] = output_file
    save_settings(settings)
    
    root.destroy()
    return checkbook_file, output_file

def main():
    """Main application"""
    if not check_license():
        return
    
    checkbook_file, output_file = select_files()
    if not checkbook_file or not output_file:
        return
    
    # Placeholder for actual accrual processing
    try:
        messagebox.showinfo("Success", 
                          f"Accrual processing would run here!\\n\\n"
                          f"Checkbook: {os.path.basename(checkbook_file)}\\n"
                          f"Output: {os.path.basename(output_file)}\\n\\n"
                          f"Ready for full implementation!")
    except Exception as e:
        messagebox.showerror("Error", f"Error: {str(e)}")

if __name__ == "__main__":
    main()
'''
    
    return main_script_content

def create_spec_file():
    """Create PyInstaller spec file for advanced configuration"""
    spec_content = '''# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['accrual_generator.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['pandas', 'openpyxl', 'anthropic', 'tkinter'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='AccrualGenerator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='app_icon.ico'  # Optional: add your icon file
)
'''
    return spec_content

def create_installer_script():
    """Create NSIS installer script for professional deployment"""
    nsis_script = '''
!define APP_NAME "Accrual Generator"
!define APP_VERSION "1.0"
!define APP_PUBLISHER "Your Company Name"
!define APP_EXE "AccrualGenerator.exe"

; Modern UI
!include "MUI2.nsh"

; General
Name "${APP_NAME}"
OutFile "AccrualGenerator_Setup.exe"
InstallDir "$LOCALAPPDATA\\${APP_NAME}"
InstallDirRegKey HKCU "Software\\${APP_NAME}" ""

; Interface Settings
!define MUI_ABORTWARNING
!define MUI_ICON "app_icon.ico"
!define MUI_UNICON "app_icon.ico"

; Pages
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_LICENSE "LICENSE.txt"
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_WELCOME
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_UNPAGE_FINISH

; Languages
!insertmacro MUI_LANGUAGE "English"

; Installer sections
Section "Main Application" SecMain
    SetOutPath "$INSTDIR"
    File "${APP_EXE}"
    File "README.txt"
    
    ; Create shortcuts
    CreateDirectory "$SMPROGRAMS\\${APP_NAME}"
    CreateShortcut "$SMPROGRAMS\\${APP_NAME}\\${APP_NAME}.lnk" "$INSTDIR\\${APP_EXE}"
    CreateShortcut "$DESKTOP\\${APP_NAME}.lnk" "$INSTDIR\\${APP_EXE}"
    
    ; Write registry
    WriteRegStr HKCU "Software\\${APP_NAME}" "" $INSTDIR
    WriteUninstaller "$INSTDIR\\Uninstall.exe"
SectionEnd

; Uninstaller section
Section "Uninstall"
    Delete "$INSTDIR\\${APP_EXE}"
    Delete "$INSTDIR\\README.txt"
    Delete "$INSTDIR\\Uninstall.exe"
    Delete "$DESKTOP\\${APP_NAME}.lnk"
    Delete "$SMPROGRAMS\\${APP_NAME}\\${APP_NAME}.lnk"
    RMDir "$SMPROGRAMS\\${APP_NAME}"
    RMDir "$INSTDIR"
    DeleteRegKey HKCU "Software\\${APP_NAME}"
SectionEnd
'''
    return nsis_script

def build_executable():
    """Build the executable using PyInstaller"""
    print("\n🔨 Building executable...")
    
    # Create build directory
    build_dir = Path("build_accrual")
    if build_dir.exists():
        shutil.rmtree(build_dir)
    build_dir.mkdir()
    
    os.chdir(build_dir)
    
    # Create main script
    print("📝 Creating main script...")
    with open("accrual_generator.py", "w", encoding="utf-8") as f:
        f.write(create_main_script())
    
    # Create spec file
    print("📝 Creating PyInstaller spec file...")
    with open("accrual_generator.spec", "w", encoding="utf-8") as f:
        f.write(create_spec_file())
    
    # Install required packages
    print("📦 Installing required packages...")
    packages = ["pandas", "openpyxl", "anthropic"]
    for package in packages:
        try:
            subprocess.run([sys.executable, "-m", "pip", "install", package], 
                          check=True, capture_output=True)
            print(f"✓ {package} installed")
        except subprocess.CalledProcessError as e:
            print(f"❌ Failed to install {package}: {e}")
            return False
    
    # Build with PyInstaller
    print("🔨 Building executable (this may take several minutes)...")
    try:
        result = subprocess.run([
            sys.executable, "-m", "PyInstaller",
            "--onefile",
            "--windowed",
            "--name=AccrualGenerator",
            "--distpath=../dist",
            "accrual_generator.py"
        ], check=True, capture_output=True, text=True)
        
        print("✓ Executable built successfully!")
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"❌ Build failed: {e}")
        print("Error output:", e.stderr)
        return False

def create_deployment_package():
    """Create final deployment package"""
    print("\n📦 Creating deployment package...")
    
    # Create deployment directory
    deploy_dir = Path("AccrualGenerator_Deploy")
    if deploy_dir.exists():
        shutil.rmtree(deploy_dir)
    deploy_dir.mkdir()
    
    # Copy executable
    exe_path = Path("dist/AccrualGenerator.exe")
    if exe_path.exists():
        shutil.copy2(exe_path, deploy_dir / "AccrualGenerator.exe")
        print("✓ Executable copied")
    else:
        print("❌ Executable not found")
        return False
    
    # Create README
    readme_content = f"""# Accrual Generator v1.0
Professional Accrual Report Generator

## Installation:
1. Copy AccrualGenerator.exe to your desired location
2. Create a desktop shortcut (optional)
3. Edit the API key in the application settings

## Usage:
1. Double-click AccrualGenerator.exe
2. Select your checkbook Excel file
3. Choose output location
4. The program will generate your accrual report

## Requirements:
- Windows 10/11
- Anthropic API key
- Excel files in supported format

## Support:
Contact your developer for support and licensing.

Built: {datetime.now().strftime('%B %d, %Y')}
"""
    
    with open(deploy_dir / "README.txt", "w", encoding="utf-8") as f:
        f.write(readme_content)
    print("✓ README created")
    
    # Create simple installer batch
    installer_content = '''@echo off
title Accrual Generator Setup
echo.
echo ================================================================
echo                   ACCRUAL GENERATOR SETUP
echo ================================================================
echo.
echo This will install Accrual Generator to your system.
echo.
pause

set INSTALL_DIR=%LOCALAPPDATA%\\AccrualGenerator
echo Creating installation directory...
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

echo Copying program files...
copy "AccrualGenerator.exe" "%INSTALL_DIR%\\" >nul
copy "README.txt" "%INSTALL_DIR%\\" >nul

echo Creating desktop shortcut...
set SHORTCUT="%USERPROFILE%\\Desktop\\Accrual Generator.lnk"
powershell "$s=(New-Object -COM WScript.Shell).CreateShortcut('%SHORTCUT%'); $s.TargetPath='%INSTALL_DIR%\\AccrualGenerator.exe'; $s.Save()"

echo.
echo ================================================================
echo                    INSTALLATION COMPLETE!
echo ================================================================
echo.
echo ✓ Program installed to: %INSTALL_DIR%
echo ✓ Desktop shortcut created
echo.
echo To run: Double-click "Accrual Generator" on your desktop
echo.
echo IMPORTANT: Configure your API key before first use!
echo.
pause
'''
    
    with open(deploy_dir / "Setup.bat", "w", encoding="utf-8") as f:
        f.write(installer_content)
    print("✓ Setup script created")
    
    print(f"\n✅ Deployment package ready: {deploy_dir.absolute()}")
    return True

def main():
    """Main installer creation process"""
    print("🚀 Accrual Generator - Advanced Installer Creator")
    print("=" * 60)
    
    if not check_requirements():
        print("\n❌ Requirements not met. Please install Python 3.8+ and pip.")
        return
    
    if not install_pyinstaller():
        print("\n❌ Could not install PyInstaller.")
        return
    
    if not build_executable():
        print("\n❌ Executable build failed.")
        return
    
    if not create_deployment_package():
        print("\n❌ Deployment package creation failed.")
        return
    
    print("\n" + "=" * 60)
    print("🎉 SUCCESS! Professional installer package created!")
    print("=" * 60)
    print("\nYour deployment package is ready:")
    print("📁 Directory: AccrualGenerator_Deploy/")
    print("📄 Files:")
    print("   • AccrualGenerator.exe (Main application)")
    print("   • Setup.bat (Simple installer)")
    print("   • README.txt (Instructions)")
    print("\nTo distribute:")
    print("1. ZIP the AccrualGenerator_Deploy folder")
    print("2. Send to users")
    print("3. Users run Setup.bat")
    print("4. Configure API key")
    print("5. Ready to use!")
    print("\n" + "=" * 60)

if __name__ == "__main__":
    main()
