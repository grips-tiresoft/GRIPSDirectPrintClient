# GRIPS Direct Print Client - macOS Version

This is the macOS version of the GRIPS Direct Print Client that uses CUPS printing instead of SumatraPDF.

## Requirements

- macOS 10.14 or later
- Bash 4.0+ or Zsh
- `jq` - JSON processor for parsing configuration files
- CUPS printing system (built into macOS)

## Installation

1. Install `jq` using Homebrew:
   ```bash
   brew install jq
   ```

2. Make the script executable:
   ```bash
   chmod +x Print-GRDPFile.sh
   ```

3. Copy the configuration files:
   ```bash
   cp config-macos.json config.json
   ```

4. Update the `config.json` file with your settings.

## Usage

### Basic Usage

Print a PDF file:
```bash
./Print-GRDPFile.sh -i /path/to/file.pdf
```

Process a GRDP file:
```bash
./Print-GRDPFile.sh -i /path/to/file.grdp
```

### Advanced Usage

Specify custom config files:
```bash
./Print-GRDPFile.sh -i /path/to/file.pdf -c custom_config.json -u user_config.json
```

### Command Line Options

- `-i, --input`: Input file path (required)
- `-c, --config`: Path to config.json file (optional, defaults to script directory)
- `-u, --userconfig`: Path to userconfig.json file (optional, defaults to script directory)

## Configuration

The `config.json` file contains the following macOS-specific settings:

- **PDFPrinter**: Set to "CUPS" for macOS
- **PDFPrinter_command**: The CUPS command to use (default: "lp")
- **Sign_exe**: Path to signature application (adjust for your system)
- **TranscriptMaxAgeDays**: Number of days to keep transcript logs (default: 7)
- **ReleaseCheckDelay**: Seconds between update checks (default: 3600)

## Features

### CUPS Printing
- Uses native macOS CUPS printing system via the `lp` command
- Supports printer selection dialog if specified printer not found
- Handles output bin selection and additional print options
- Automatically prints to default printer for simple PDF files

### GRDP File Processing
- Extracts and processes `.grdp` (zip) files
- Reads `printsettings.json` for multiple print jobs
- Supports per-document printer and output bin settings
- Opens non-PDF files with default application

### Printer Selection
- Automatic printer detection using CUPS `lpstat` command
- Interactive printer selection dialog using AppleScript if printer not found
- Validates printer availability before submitting print jobs

### Auto-Update
- Checks for new releases from GitHub
- Downloads and installs updates automatically
- Supports both stable and prerelease versions
- Creates backup before updating

### Logging
- Creates timestamped transcript logs
- Automatically removes old logs based on configured age
- Logs all operations and errors for troubleshooting

### Multi-Language Support
- Uses the same `languages.json` file as Windows version
- Automatically detects OS language
- Falls back to English if language not found

## Differences from Windows Version

1. **Printing System**: Uses CUPS (`lp` command) instead of SumatraPDF
2. **Printer Selection**: Uses AppleScript dialogs instead of Windows Forms
3. **File Operations**: Uses bash/zsh instead of PowerShell
4. **Configuration**: Separate `config-macos.json` with macOS-specific paths
5. **Dependencies**: Requires `jq` for JSON parsing

## CUPS Print Options

The script supports standard CUPS print options:

- **Output Bin**: Specified via `-o outputbin=<bin>` option
- **Additional Options**: Any valid CUPS option can be passed through `AdditionalArgs`

### Common CUPS Options

```bash
-o media=Letter          # Paper size
-o sides=two-sided-long  # Duplex printing
-o outputbin=tray1       # Output tray
-o ColorModel=CMYK       # Color mode
-o Resolution=600dpi     # Print resolution
```

## Troubleshooting

### Script won't execute
Ensure the script has execute permissions:
```bash
chmod +x Print-GRDPFile.sh
```

### jq command not found
Install jq using Homebrew:
```bash
brew install jq
```

### Printer not found
The script will show a dialog to select an alternative printer. Verify printers are available:
```bash
lpstat -p
```

### Permission denied when printing
Check printer permissions and ensure you have access to the printer:
```bash
lpstat -p <printer-name>
```

## File Structure

```
GRIPSDirectPrintClient/
├── Print-GRDPFile.sh          # Main macOS script
├── config-macos.json          # macOS configuration template
├── config.json                # Active configuration
├── userconfig.json            # Optional user overrides
├── languages.json             # Multi-language strings
├── Transcripts/               # Log files directory
└── README-macOS.md            # This file
```

## Examples

### Example 1: Simple PDF Print
```bash
./Print-GRDPFile.sh -i ~/Documents/invoice.pdf
```

### Example 2: GRDP File with Multiple Documents
```bash
./Print-GRDPFile.sh -i ~/Downloads/print_job.grdp
```

### Example 3: Custom Configuration
```bash
./Print-GRDPFile.sh -i ~/Documents/report.pdf -c ~/custom/config.json
```

## Support

For issues or questions:
- Check the transcript logs in the `Transcripts/` directory
- Review printer status: `lpstat -p`
- Verify CUPS configuration: `cupsctl`

## License

Same as the Windows version - refer to main repository LICENSE file.
