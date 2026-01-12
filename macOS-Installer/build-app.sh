#!/bin/zsh

# Build script for GRIPS Direct Print macOS Application Bundle
# This script creates the .app bundle structure and copies all necessary files

set -e  # Exit on any error

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"

APP_NAME="GRIPSDirectPrint.app"
BUILD_DIR="$SCRIPT_DIR/build"
APP_BUNDLE="$BUILD_DIR/$APP_NAME"

echo "Building GRIPS Direct Print application bundle..."
echo "Project root: $PROJECT_ROOT"
echo "Build directory: $BUILD_DIR"

# Clean previous build
if [[ -d "$BUILD_DIR" ]]; then
    echo "Cleaning previous build..."
    rm -rf "$BUILD_DIR"
fi

# Create build directory
mkdir -p "$BUILD_DIR"

# Build AppleScript app first
echo "Compiling AppleScript application..."
osacompile -o "$APP_BUNDLE" "$SCRIPT_DIR/app-template/launcher.applescript"

# Copy custom Info.plist to register file associations
echo "Copying custom Info.plist..."
cp "$SCRIPT_DIR/app-template/Info.plist" "$APP_BUNDLE/Contents/Info.plist"

# Copy all necessary files to Resources
echo "Copying application files..."

# List of files to copy
FILES_TO_COPY=(
    "Print-GRDPFile.sh"
    "config-macos.json"
    "languages.json"
)

for file in "${FILES_TO_COPY[@]}"; do
    if [[ -f "$PROJECT_ROOT/$file" ]]; then
        echo "  Copying $file..."
        cp "$PROJECT_ROOT/$file" "$APP_BUNDLE/Contents/Resources/"
    else
        echo "  WARNING: $file not found, skipping..."
    fi
done

# Copy bundled jq binary
echo "  Copying bundled jq binary..."
if [[ -f "$SCRIPT_DIR/bin/jq" ]]; then
    cp "$SCRIPT_DIR/bin/jq" "$APP_BUNDLE/Contents/Resources/"
    chmod +x "$APP_BUNDLE/Contents/Resources/jq"
else
    echo "  WARNING: jq binary not found at $SCRIPT_DIR/bin/jq"
fi

# Make the print script executable
chmod +x "$APP_BUNDLE/Contents/Resources/Print-GRDPFile.sh"

# Create Transcripts directory
mkdir -p "$APP_BUNDLE/Contents/Resources/Transcripts"

# Copy optional files if they exist
OPTIONAL_FILES=(
    "userconfig-macos.json"
    "last_update_check.txt"
)

for file in "${OPTIONAL_FILES[@]}"; do
    if [[ -f "$PROJECT_ROOT/$file" ]]; then
        echo "  Copying optional file $file..."
        cp "$PROJECT_ROOT/$file" "$APP_BUNDLE/Contents/Resources/"
    fi
done

# Sign the app bundle with adhoc signature
echo "Code signing the application..."
codesign --force --deep --sign - "$APP_BUNDLE"

echo ""
echo "✓ Application bundle created successfully at:"
echo "  $APP_BUNDLE"
echo ""
echo "You can test the app by running:"
echo "  open \"$APP_BUNDLE\""
echo ""
echo "Or test with a .grdp file:"
echo "  open -a \"$APP_BUNDLE\" path/to/file.grdp"
echo ""
