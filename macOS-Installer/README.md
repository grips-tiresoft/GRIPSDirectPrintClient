# GRIPS Direct Print - macOS Installer

This directory contains the build scripts and resources needed to create a macOS installer (.pkg) for GRIPS Direct Print.

## Overview

The installer will:
- Create a macOS application bundle (GRIPSDirectPrint.app)
- Install it to `/Applications/` (available to all users)
- Register the .grdp file type association
- Automatically handle .grdp files when double-clicked

## Prerequisites

- macOS 10.13 or later
- Xcode Command Line Tools (for `pkgbuild` and `productbuild`)

The `jq` tool is bundled with the application, so you don't need to install it separately.

To install prerequisites:

```bash
# Install Xcode Command Line Tools
xcode-select --install
```

## Building the Installer

### Quick Build

To build the complete installer package:

```bash
cd macOS-Installer
chmod +x build-installer.sh
./build-installer.sh
```

This will create: `macOS-Installer/package/GRIPSDirectPrint-Installer.pkg`

### Step-by-Step Build

You can also build components separately:

```bash
# 1. Build just the application bundle
chmod +x build-app.sh
./build-app.sh

# This creates: macOS-Installer/build/GRIPSDirectPrint.app

# 2. Build the installer package
chmod +x build-installer.sh
./build-installer.sh
```

## Installation

### For End Users

1. Double-click `GRIPSDirectPrint-Installer.pkg`
2. Follow the installation wizard
3. Enter your password when prompted (required for system-wide installation)
4. The application will be installed to `/Applications/GRIPSDirectPrint.app`

### Testing Without Installing

You can test the application bundle without creating an installer:

```bash
# Build the app
./build-app.sh

# Test with a .grdp file
open -a "build/GRIPSDirectPrint.app" ../Test.grdp
```

## File Association

After installation:
- Double-clicking any `.grdp` file will automatically open it with GRIPS Direct Print
- The file will be processed according to its `printsettings.json`
- PDF files will be printed to the specified printers
- Other files (like .eml) will be opened with their default applications

## Structure

```
macOS-Installer/
├── app-template/           # Application bundle templates
│   ├── Info.plist         # Bundle metadata and file associations
│   └── GRIPSDirectPrint   # Launcher script (goes in MacOS/)
├── scripts/               # Installer scripts
│   └── postinstall       # Runs after installation
├── build-app.sh          # Builds the .app bundle
├── build-installer.sh    # Builds the .pkg installer
└── README.md             # This file

After building:
├── build/                # Built application bundle
│   └── GRIPSDirectPrint.app/
├── package/              # Installer package files
│   ├── GRIPSDirectPrint-Installer.pkg  # Final installer
│   └── ... (build artifacts)
```

## How It Works

### Application Bundle

The `.app` bundle structure:
```
GRIPSDirectPrint.app/
├── Contents/
│   ├── Info.plist                    # Metadata and file type registration
│   ├── MacOS/
│   │   └── GRIPSDirectPrint          # Launcher executable
│   └── Resources/
│       ├── Print-GRDPFile.sh         # Main processing script
│       ├── config-macos.json         # Configuration
│       ├── languages.json            # Localization strings
│       └── Transcripts/              # Log files directory
```

### Workflow

1. User double-clicks a `.grdp` file
2. macOS launches `GRIPSDirectPrint.app` with the file path as argument
3. The launcher script (`GRIPSDirectPrint`) receives the file path
4. It calls `Print-GRDPFile.sh` with the `-i <file>` argument
5. The main script processes the .grdp file:
   - Extracts the archive
   - Reads `printsettings.json`
   - Prints PDFs using CUPS
   - Opens other files with default applications

### File Type Registration

The `Info.plist` registers:
- UTI: `com.grips.directprint.grdp`
- File extension: `.grdp`
- Handler role: Owner (this app owns the file type)

The postinstall script calls `lsregister` to register the application with macOS Launch Services.

## Customization

### Changing the Installation Location

Edit `build-installer.sh` and modify:
```bash
--install-location "/"  # Change this to install elsewhere
```

For current user only, you could use:
```bash
--install-location "$HOME"
# And change the payload path to: Applications/ (relative to HOME)
```

### Modifying the Installer Appearance

Edit these files in `build-installer.sh`:
- `welcome.html` - Welcome screen
- `license.txt` - License agreement
- `conclusion.html` - Completion screen

### Adding an Icon

1. Create an `.icns` icon file
2. Copy it to `app-template/` as `AppIcon.icns`
3. Update `Info.plist` to reference it:
   ```xml
   <key>CFBundleIconFile</key>
   <string>AppIcon.icns</string>
   ```
4. Modify `build-app.sh` to copy the icon file

## Troubleshooting

### .grdp files not opening with the app

Try these steps:
1. Right-click a `.grdp` file → Get Info
2. Under "Open with:", select `GRIPSDirectPrint.app`
3. Click "Change All..."

Or rebuild the Launch Services database:
```bash
/System/Library/Frameworks/CoreServices.framework/Frameworks/LaunchServices.framework/Support/lsregister -kill -r -domain local -domain system -domain user
```

### Permission issues

If you get permission errors:
```bash
# Make sure scripts are executable
chmod +x macOS-Installer/build-app.sh
chmod +x macOS-Installer/build-installer.sh
chmod +x macOS-Installer/scripts/postinstall
chmod +x Print-GRDPFile.sh
```

### jq not found

Install jq:
```bash
brew install jq
```

### App won't open (macOS security)

First time opening the app:
1. Right-click the app → Open
2. Click "Open" in the security dialog

Or allow it in System Preferences:
1. System Preferences → Security & Privacy
2. Click "Open Anyway" for GRIPSDirectPrint

## Uninstallation

To uninstall:
```bash
sudo rm -rf /Applications/GRIPSDirectPrint.app
```

Then reset the file association:
1. Right-click a `.grdp` file → Get Info
2. Under "Open with:", select a different application
3. Click "Change All..."

## Distribution

The generated `.pkg` file can be:
- Distributed directly to users
- Hosted on a website for download
- Signed and notarized for wider distribution (requires Apple Developer account)

### Code Signing (Optional)

To sign the installer (requires Apple Developer certificate):

```bash
# Sign the app bundle first
codesign --deep --force --verify --verbose --sign "Developer ID Application: Your Name" build/GRIPSDirectPrint.app

# Then build the installer (it will include the signed app)
./build-installer.sh

# Sign the installer
productsign --sign "Developer ID Installer: Your Name" package/GRIPSDirectPrint-Installer.pkg package/GRIPSDirectPrint-Installer-Signed.pkg
```

## Support

For issues or questions:
- Check the Transcripts folder for logs: `/Applications/GRIPSDirectPrint.app/Contents/Resources/Transcripts/`
- Review the main script: [Print-GRDPFile.sh](../Print-GRDPFile.sh)

## Version History

- 1.0.0 - Initial release
