"""
Build script to create 32-bit and 64-bit executables.
Run this script to build the EML to PST Converter executables.

WINDOWS 7 COMPATIBILITY:
    - Use Python 3.8.x (last version supporting Windows 7)
    - Download from: https://www.python.org/downloads/release/python-3810/
    - For 32-bit Win7: Use "Windows x86 executable installer"
    - For 64-bit Win7: Use "Windows x86-64 executable installer"

Requirements:
    pip install pyinstaller pywin32

Usage:
    python build_exe.py
    python build_exe.py --win7    # Adds Win7 suffix to filename
"""

import subprocess
import sys
import os
import shutil
import struct
import argparse

# Minimum Python version info
MIN_PYTHON_WIN7 = (3, 8)  # Python 3.8 is the last to support Windows 7
MAX_PYTHON_WIN7 = (3, 9)  # Python 3.9+ dropped Windows 7 support


def get_python_arch():
    """Get the architecture of the current Python interpreter"""
    return struct.calcsize("P") * 8


def check_win7_compatibility():
    """Check if current Python version supports Windows 7"""
    version = sys.version_info[:2]
    if version >= MAX_PYTHON_WIN7:
        return False, f"Python {version[0]}.{version[1]} does NOT support Windows 7"
    return True, f"Python {version[0]}.{version[1]} supports Windows 7"


def check_pyinstaller():
    """Check if PyInstaller is installed, install if not"""
    try:
        import PyInstaller
        print(f"PyInstaller version: {PyInstaller.__version__}")
        return True
    except ImportError:
        print("PyInstaller not found. Installing...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
            print("PyInstaller installed successfully!")
            return True
        except subprocess.CalledProcessError as e:
            print(f"Failed to install PyInstaller: {e}")
            return False


def build_executable(win7_mode=False):
    """Build the executable for the current Python architecture"""
    arch = get_python_arch()
    win7_compat, win7_msg = check_win7_compatibility()
    
    print(f"\n{'='*60}")
    print(f"Building {arch}-bit executable...")
    print(f"Windows 7 compatible: {'YES' if win7_compat else 'NO'}")
    print(f"{'='*60}\n")
    
    if win7_mode and not win7_compat:
        print(f"WARNING: {win7_msg}")
        print("The executable will NOT run on Windows 7!")
        print("\nTo build a Windows 7 compatible executable:")
        print("  1. Install Python 3.8.x (32-bit or 64-bit)")
        print("  2. Run this script with that Python version")
        print("\nDownload Python 3.8.10 from:")
        print("  https://www.python.org/downloads/release/python-3810/")
        print("  - Windows x86 executable installer (32-bit)")
        print("  - Windows x86-64 executable installer (64-bit)")
        response = input("\nContinue anyway? (y/N): ")
        if response.lower() != 'y':
            return False
    
    # Get the directory of this script
    script_dir = os.path.dirname(os.path.abspath(__file__))
    main_script = os.path.join(script_dir, "eml_to_pst_converter.py")
    
    if not os.path.exists(main_script):
        print(f"Error: {main_script} not found!")
        return False
    
    # Output directory
    dist_dir = os.path.join(script_dir, "dist")
    build_dir = os.path.join(script_dir, "build")
    
    # Output name based on architecture and Win7 compatibility
    if win7_mode and win7_compat:
        output_name = f"EML_to_PST_Converter_x{arch}_Win7"
    else:
        output_name = f"EML_to_PST_Converter_x{arch}"
    
    # PyInstaller command
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",                          # Single executable
        "--windowed",                         # No console window (GUI app)
        f"--name={output_name}",              # Output name
        f"--distpath={dist_dir}",             # Output directory
        f"--workpath={build_dir}",            # Build directory
        "--clean",                            # Clean cache
        "--noconfirm",                        # Overwrite without asking
        # Icon (optional - uncomment and provide path if you have an icon)
        # "--icon=icon.ico",
        main_script
    ]
    
    print(f"Running: {' '.join(cmd)}\n")
    
    try:
        subprocess.check_call(cmd)
        exe_path = os.path.join(dist_dir, f"{output_name}.exe")
        if os.path.exists(exe_path):
            size_mb = os.path.getsize(exe_path) / (1024 * 1024)
            print(f"\n{'='*60}")
            print(f"SUCCESS! Executable created:")
            print(f"  {exe_path}")
            print(f"  Size: {size_mb:.2f} MB")
            print(f"  Windows 7 compatible: {'YES' if win7_compat else 'NO'}")
            print(f"{'='*60}\n")
            return True
        else:
            print("Error: Executable was not created!")
            return False
    except subprocess.CalledProcessError as e:
        print(f"Build failed: {e}")
        return False


def clean_build_artifacts():
    """Clean up build artifacts"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Remove build directory
    build_dir = os.path.join(script_dir, "build")
    if os.path.exists(build_dir):
        print(f"Removing {build_dir}...")
        shutil.rmtree(build_dir)
    
    # Remove spec file
    for f in os.listdir(script_dir):
        if f.endswith(".spec"):
            spec_file = os.path.join(script_dir, f)
            print(f"Removing {spec_file}...")
            os.remove(spec_file)
    
    print("Build artifacts cleaned.\n")


def main():
    parser = argparse.ArgumentParser(description="Build EML to PST Converter executable")
    parser.add_argument("--win7", action="store_true", 
                        help="Build for Windows 7 compatibility (adds Win7 suffix)")
    args = parser.parse_args()
    
    print("\n" + "="*60)
    print("   EML to PST Converter - Executable Builder")
    print("="*60 + "\n")
    
    arch = get_python_arch()
    win7_compat, win7_msg = check_win7_compatibility()
    
    print(f"Current Python: {sys.executable}")
    print(f"Python version: {sys.version}")
    print(f"Architecture: {arch}-bit")
    print(f"Windows 7 support: {win7_msg}\n")
    
    # Check/install PyInstaller
    if not check_pyinstaller():
        print("Cannot proceed without PyInstaller.")
        sys.exit(1)
    
    # Build executable
    success = build_executable(win7_mode=args.win7)
    
    # Clean up
    if success:
        clean_build_artifacts()
        
        print("\n" + "="*60)
        print("   BUILD INSTRUCTIONS")
        print("="*60)
        print("\nFor Windows 10/11 (any Python 3.7+):")
        print('  python build_exe.py')
        print("\nFor Windows 7 compatibility (REQUIRES Python 3.8.x):")
        print('  "C:\\Python38-32\\python.exe" build_exe.py --win7  (32-bit)')
        print('  "C:\\Python38\\python.exe" build_exe.py --win7     (64-bit)')
        print("\nDownload Python 3.8.10 for Windows 7:")
        print("  https://www.python.org/downloads/release/python-3810/")
        print("  - Windows x86 executable installer (32-bit)")
        print("  - Windows x86-64 executable installer (64-bit)")
    
    return 0 if success else 1


if __name__ == "__main__":
    sys.exit(main())
