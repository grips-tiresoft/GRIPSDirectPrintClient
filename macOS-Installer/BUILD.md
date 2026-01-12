# Building the GRIPS Direct Print macOS Installer

This guide provides step-by-step instructions for building a new `.pkg` installer for GRIPS Direct Print on macOS.

## Prerequisites

Before you begin, ensure you have:

- macOS 10.13 or later
- Xcode Command Line Tools installed
- Access to the source repository
- (Optional) Apple Developer certificate for code signing

### Install Xcode Command Line Tools

If not already installed:

```bash
xcode-select --install
```

Verify installation:

```bash
pkgbuild --version
productbuild --version
```

## Quick Build Process

### 1. Navigate to the Installer Directory

```bash
cd /path/to/GRIPSDirectPrintClient/macOS-Installer
```

### 2. Make Build Scripts Executable

```bash
chmod +x build-app.sh build-installer.sh
```

### 3. Build the Installer

```bash
./build-installer.sh
```

This will:
- Clean any previous builds
- Create the application bundle from the template
- Copy all required resources (scripts, configs, jq binary)
- Build the `.pkg` installer
- Output: `macOS-Installer/package/GRIPSDirectPrint-Installer.pkg`

## Detailed Build Process

### Step 1: Build the Application Bundle

The application bundle can be built separately:

```bash
./build-app.sh
```

This creates `macOS-Installer/build/GRIPSDirectPrint.app/` with the following structure:

```
GRIPSDirectPrint.app/
└── Contents/
    ├── Info.plist
    ├── MacOS/
    │   └── GRIPSDirectPrint
    └── Resources/
        ├── Print-GRDPFile.sh
        ├── config-macos.json
        ├── languages.json
        ├── jq
        └── Transcripts/
```

### Step 2: Test the Application Bundle

Before creating the installer, test the app bundle:

```bash
# Test with a .grdp file
open -a "build/GRIPSDirectPrint.app" ../Test.grdp

# Or manually launch the app
./build/GRIPSDirectPrint.app/Contents/MacOS/GRIPSDirectPrint ../Test.grdp
```

### Step 3: Build the Installer Package

Once the app bundle is tested:

```bash
./build-installer.sh
```

The script will:
1. Run `build-app.sh` to create/update the app bundle
2. Create the package payload directory
3. Copy the app bundle to the payload
4. Build the component package using `pkgbuild`
5. Create the distribution package using `productbuild`
6. Output the final installer

## Versioning

### Update Version Before Building

Update the version in [config-macos.json](../config-macos.json):

```json
{
  "Version": "1.0.1",
  ...
}
```

The version should follow semantic versioning (e.g., `1.0.1`, `1.1.0`, `2.0.0`).

## Build Output

After a successful build, you'll have:

```
macOS-Installer/
├── build/
│   └── GRIPSDirectPrint.app/        # Application bundle
└── package/
    ├── GRIPSDirectPrint-Installer.pkg    # Final installer (ready for distribution)
    ├── payload/                          # Staging directory
    └── ... (other build artifacts)
```

## Distribution Package Components

The installer includes:

- **Application**: `GRIPSDirectPrint.app` installed to `/Applications/`
- **Scripts**: `postinstall` script that registers the app with macOS
- **Resources**:
  - `welcome.html` - Installation welcome screen
  - `license.txt` - License agreement
  - `conclusion.html` - Installation completion message
  - `distribution.xml` - Package distribution settings

## Code Signing (Optional but Recommended)

For distribution outside of testing, sign both the app and the installer.

### Prerequisites for Signing

- Apple Developer account
- Developer ID Application certificate (for app signing)
- Developer ID Installer certificate (for pkg signing)

### Sign the Application

```bash
# Sign the app bundle before building the installer
codesign --deep --force --verify --verbose \
  --sign "Developer ID Application: Your Name (TEAM_ID)" \
  --options runtime \
  build/GRIPSDirectPrint.app

# Verify signature
codesign --verify --verbose build/GRIPSDirectPrint.app
spctl --assess --verbose build/GRIPSDirectPrint.app
```

### Build and Sign the Installer

```bash
# Build the installer (with signed app)
./build-installer.sh

# Sign the installer
productsign --sign "Developer ID Installer: Your Name (TEAM_ID)" \
  package/GRIPSDirectPrint-Installer.pkg \
  package/GRIPSDirectPrint-Installer-Signed.pkg

# Verify installer signature
pkgutil --check-signature package/GRIPSDirectPrint-Installer-Signed.pkg
```

### Notarization (Required for macOS 10.15+)

After signing, notarize the installer for distribution:

```bash
# Upload for notarization
xcrun notarytool submit package/GRIPSDirectPrint-Installer-Signed.pkg \
  --apple-id "your-apple-id@example.com" \
  --team-id "TEAM_ID" \
  --password "app-specific-password" \
  --wait

# Staple the notarization ticket
xcrun stapler staple package/GRIPSDirectPrint-Installer-Signed.pkg

# Verify
xcrun stapler validate package/GRIPSDirectPrint-Installer-Signed.pkg
```

## Creating a GitHub Release with the Installer

### 1. Tag the Release

```bash
# Create and push a version tag
git tag v1.0.1
git push origin v1.0.1
```

### 2. Create GitHub Release

1. Go to your repository on GitHub
2. Click "Releases" → "Draft a new release"
3. Select your tag (e.g., `v1.0.1`)
4. Add release notes
5. **Attach the .pkg file**:
   - Drag and drop `GRIPSDirectPrint-Installer.pkg` (or the signed version)
   - File must be named with pattern: `GRIPSDirectPrint*.pkg`
6. Publish the release

### 3. Verify Auto-Update Detection

The auto-update mechanism in `Print-GRDPFile.sh` will:
- Check for new releases via GitHub API
- Look for assets matching `GRIPSDirectPrint*.pkg`
- Download and install automatically with user consent

## Troubleshooting

### Build Fails: "command not found: pkgbuild"

Install Xcode Command Line Tools:
```bash
xcode-select --install
```

### Build Fails: Permission Denied

Make scripts executable:
```bash
chmod +x macOS-Installer/build-app.sh
chmod +x macOS-Installer/build-installer.sh
chmod +x macOS-Installer/scripts/postinstall
chmod +x Print-GRDPFile.sh
chmod +x macOS-Installer/bin/jq
```

### Installer Builds but Won't Install

Check the installer structure:
```bash
pkgutil --payload-files package/GRIPSDirectPrint-Installer.pkg
```

Check for errors in the build output.

### Code Signing Fails

Verify your certificates:
```bash
security find-identity -v -p codesigning
```

List should include "Developer ID Application" and "Developer ID Installer" certificates.

## Clean Build

To start fresh:

```bash
# Remove all build artifacts
rm -rf build/ package/

# Rebuild
./build-installer.sh
```

## Testing the Installer

### Test Installation

```bash
# Install to test location (doesn't require sudo)
installer -pkg package/GRIPSDirectPrint-Installer.pkg \
  -target CurrentUserHomeDirectory \
  -verbose

# Or install system-wide (requires sudo)
sudo installer -pkg package/GRIPSDirectPrint-Installer.pkg \
  -target / \
  -verbose
```

### Verify Installation

```bash
# Check if app is installed
ls -la /Applications/GRIPSDirectPrint.app

# Check file association
mdls -name kMDItemContentType Test.grdp

# Test opening a .grdp file
open Test.grdp
```

### Check Logs

```bash
# View installation logs
cat /var/log/install.log | grep GRIPS

# View app logs
cat /Applications/GRIPSDirectPrint.app/Contents/Resources/Transcripts/*.txt
```

## Build Checklist

Before creating a production release:

- [ ] Update version in `config-macos.json`
- [ ] Test the main script (`Print-GRDPFile.sh`) with sample .grdp files
- [ ] Run `build-installer.sh` successfully
- [ ] Test the built app bundle before packaging
- [ ] Sign the app bundle (if distributing)
- [ ] Build and sign the installer package
- [ ] Test installation on a clean system
- [ ] Verify file type association works
- [ ] Test printing with real printers
- [ ] Notarize the installer (for macOS 10.15+)
- [ ] Create GitHub release with .pkg file attached
- [ ] Verify auto-update detection works

## Additional Resources

- [Apple Developer: Customizing the Installation Experience](https://developer.apple.com/library/archive/documentation/DeveloperTools/Reference/DistributionDefinitionRef/Chapters/Introduction.html)
- [Apple Developer: Notarizing macOS Software](https://developer.apple.com/documentation/security/notarizing_macos_software_before_distribution)
- [pkgbuild Manual Page](x-man-page://pkgbuild)
- [productbuild Manual Page](x-man-page://productbuild)

## Support

For issues or questions:
- Review [README.md](README.md) for general information
- Check build script output for specific errors
- Verify all prerequisites are installed
- Test on a clean macOS system
