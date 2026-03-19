# EML to PST Converter

A Python GUI application to convert EML/EMLX email files to Outlook PST format.

## Features

- **Folder Scanning**: Browse and select folders containing EML/EMLX files
- **Wildcard Pattern Support**: Use patterns like `*.eml`, `*.emlx`, or `*.eml;*.emlx` to filter files
- **Recursive Search**: Automatically finds files in subfolders
- **Duplicate Removal**: Option to skip duplicate emails based on content hash (SHA256)
- **Create or Append**: Create a new PST file or add to an existing one
- **Progress Tracking**: Real-time progress bar and status updates
- **File Management**: Right-click context menu to remove selected files or clear all
- **Secure**: Path traversal protection for attachments, secure temp file handling

## Download

Pre-built executables are available in the `dist` folder:
- `EML_to_PST_Converter_x64.exe` - 64-bit Windows
- `EML_to_PST_Converter_x32.exe` - 32-bit Windows (if built)

## Requirements

### For Running the Executable
- Windows 7/8/10/11
- Microsoft Outlook (32-bit Outlook for x32, 64-bit Outlook for x64)

### For Running from Source
- Python 3.7+
- Microsoft Outlook
- Windows OS

## Installation (from source)

1. Clone or download this repository
2. Install dependencies:

```bash
pip install -r requirements.txt
```

## Usage

### Using the Executable
Simply double-click `EML_to_PST_Converter_x64.exe` (or x32 version).

### Using Python
```bash
python eml_to_pst_converter.py
```

### Steps
1. Click **Add Files** to browse for a folder containing EML/EMLX files
2. Adjust the **File Pattern** if needed (supports wildcards):
   - `*.eml` - All EML files
   - `*.emlx` - All EMLX files  
   - `*.eml;*.emlx` - Both formats
3. Choose whether to **Create New PST File** or **Save to Existing PST File**
4. Optionally enable **Remove Duplicate Content** to skip duplicate emails
5. Click **Browse Destination** to select where to save the PST file
6. Click **Convert** to start the conversion

## Building Executables

### Build for Current Architecture (Windows 10/11)
```bash
python build_exe.py
```
Or double-click `build.bat`

### Build for Windows 7 Compatibility

**Important**: Windows 7 requires Python 3.8.x (the last version supporting Windows 7).

1. **Download Python 3.8.10** from:
   https://www.python.org/downloads/release/python-3810/
   - `Windows x86 executable installer` (32-bit)
   - `Windows x86-64 executable installer` (64-bit)

2. **Install to a known location**:
   - 32-bit: `C:\Python38-32\`
   - 64-bit: `C:\Python38\`

3. **Build Windows 7 compatible executables**:
```bash
# 32-bit Windows 7 compatible
"C:\Python38-32\python.exe" -m pip install pyinstaller pywin32
"C:\Python38-32\python.exe" build_exe.py --win7

# 64-bit Windows 7 compatible  
"C:\Python38\python.exe" -m pip install pyinstaller pywin32
"C:\Python38\python.exe" build_exe.py --win7
```

Or run `build_win7.bat` which does this automatically.

### Output Files

The executables will be created in the `dist` folder:

| File | OS Support | Outlook |
|------|------------|---------|
| `EML_to_PST_Converter_x64.exe` | Windows 10/11 | 64-bit |
| `EML_to_PST_Converter_x32.exe` | Windows 10/11 | 32-bit |
| `EML_to_PST_Converter_x64_Win7.exe` | Windows 7/8/10/11 | 64-bit |
| `EML_to_PST_Converter_x32_Win7.exe` | Windows 7/8/10/11 | 32-bit |

**Note**: Match your executable architecture with your Outlook installation:
- 64-bit Outlook → Use x64 executable
- 32-bit Outlook → Use x32 executable

## Wildcard Patterns

The file pattern field supports standard wildcard characters:
- `*` - Matches any number of characters
- `?` - Matches a single character

Examples:
- `*.eml` - All files ending with .eml
- `mail*.eml` - All .eml files starting with "mail"
- `*2024*.eml` - All .eml files containing "2024"

## Troubleshooting

### "Cannot connect to Outlook"
- Make sure Microsoft Outlook is installed
- Try running the application as Administrator
- Close Outlook completely and try again

### PST file is empty after conversion
- Check that you have write permissions to the destination folder
- Try creating the PST in a different location (e.g., Documents folder)
- Make sure Outlook is not running during conversion

### "pywin32 not available"
Install the required library:
```bash
pip install pywin32
```

## License

This project is open source and available for personal and commercial use.
