#!/bin/zsh

# Build script for GRIPS Direct Print macOS Installer Package
# This script creates a .pkg installer that can be distributed

set -e  # Exit on any error

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"

APP_NAME="GRIPSDirectPrint.app"
BUILD_DIR="$SCRIPT_DIR/build"
PKG_DIR="$SCRIPT_DIR/package"
COMPONENT_PKG="$PKG_DIR/GRIPSDirectPrint-component.pkg"
FINAL_PKG="$PKG_DIR/GRIPSDirectPrint-Installer.pkg"

echo "=========================================="
echo "Building GRIPS Direct Print Installer"
echo "=========================================="

# Step 1: Build the app bundle first
echo ""
echo "Step 1: Building application bundle..."
"$SCRIPT_DIR/build-app.sh"

if [[ ! -d "$BUILD_DIR/$APP_NAME" ]]; then
    echo "ERROR: Application bundle was not created successfully."
    exit 1
fi

# Step 2: Create package directory structure
echo ""
echo "Step 2: Preparing package structure..."
rm -rf "$PKG_DIR"
mkdir -p "$PKG_DIR"
mkdir -p "$PKG_DIR/payload/Applications"

# Copy the app to the payload location
cp -R "$BUILD_DIR/$APP_NAME" "$PKG_DIR/payload/Applications/"

# Step 3: Make postinstall script executable
echo ""
echo "Step 3: Setting up scripts..."
chmod +x "$SCRIPT_DIR/scripts/postinstall"

# Step 4: Build component package
echo ""
echo "Step 4: Building component package..."
pkgbuild \
    --root "$PKG_DIR/payload" \
    --scripts "$SCRIPT_DIR/scripts" \
    --identifier "com.grips.directprint" \
    --version "1.0.0" \
    --install-location "/" \
    "$COMPONENT_PKG"

if [[ ! -f "$COMPONENT_PKG" ]]; then
    echo "ERROR: Component package creation failed."
    exit 1
fi

# Step 5: Create distribution XML
echo ""
echo "Step 5: Creating distribution definition..."
cat > "$PKG_DIR/distribution.xml" << 'EOF'
<?xml version="1.0" encoding="utf-8"?>
<installer-gui-script minSpecVersion="1">
    <title>GRIPS Direct Print</title>
    <organization>com.grips</organization>
    <domains enable_localSystem="true"/>
    <options customize="never" require-scripts="true" hostArchitectures="x86_64,arm64"/>
    
    <welcome file="welcome.html" mime-type="text/html"/>
    <license file="license.txt" mime-type="text/plain"/>
    <conclusion file="conclusion.html" mime-type="text/html"/>
    
    <pkg-ref id="com.grips.directprint" version="1.0.0">GRIPSDirectPrint-component.pkg</pkg-ref>
    
    <choices-outline>
        <line choice="default">
            <line choice="com.grips.directprint"/>
        </line>
    </choices-outline>
    
    <choice id="default"/>
    <choice id="com.grips.directprint" visible="false">
        <pkg-ref id="com.grips.directprint"/>
    </choice>
</installer-gui-script>
EOF

# Step 6: Create welcome text
cat > "$PKG_DIR/welcome.html" << 'EOF'
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif; }
    </style>
</head>
<body>
    <h1>Welcome to GRIPS Direct Print</h1>
    <p>This installer will install GRIPS Direct Print on your system.</p>
    <p>GRIPS Direct Print allows you to automatically print documents packaged in .grdp files.</p>
    <p>After installation, double-clicking any .grdp file will automatically process and print the documents it contains.</p>
</body>
</html>
EOF

# Step 7: Create license text
cat > "$PKG_DIR/license.txt" << 'EOF'
GRIPS Direct Print License

Copyright (c) 2026 GRIPS

Permission is hereby granted to use this software for printing documents.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND.
EOF

# Step 8: Create conclusion text
cat > "$PKG_DIR/conclusion.html" << 'EOF'
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif; }
    </style>
</head>
<body>
    <h1>Installation Complete</h1>
    <p>GRIPS Direct Print has been successfully installed.</p>
    <p>You can now open .grdp files, and they will be automatically processed.</p>
    <p>The application has been installed to: <strong>/Applications/GRIPSDirectPrint.app</strong></p>
    <h2>Next Steps:</h2>
    <ul>
        <li>Open a .grdp file to test the installation</li>
        <li>The first time you open a .grdp file, macOS may ask you to confirm</li>
        <li>Check the Downloads folder for any non-PDF files that are extracted</li>
    </ul>
</body>
</html>
EOF

# Step 9: Build the final product package
echo ""
echo "Step 6: Building final installer package..."
productbuild \
    --distribution "$PKG_DIR/distribution.xml" \
    --package-path "$PKG_DIR" \
    --resources "$PKG_DIR" \
    "$FINAL_PKG"

if [[ ! -f "$FINAL_PKG" ]]; then
    echo "ERROR: Final package creation failed."
    exit 1
fi

# Step 10: Display success message
echo ""
echo "=========================================="
echo "✓ Installation package created successfully!"
echo "=========================================="
echo ""
echo "Package location:"
echo "  $FINAL_PKG"
echo ""
echo "Package size: $(du -h "$FINAL_PKG" | cut -f1)"
echo ""
echo "To install, run:"
echo "  sudo installer -pkg \"$FINAL_PKG\" -target /"
echo ""
echo "Or double-click the .pkg file in Finder"
echo ""
