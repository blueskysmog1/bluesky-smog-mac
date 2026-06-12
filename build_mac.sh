#!/bin/bash
# Blue Sky Desktop Mac Build Script
# Usage: APP_SPECIFIC_PASSWORD="xxxx-xxxx-xxxx-xxxx" ./build_mac.sh v1.1.30

VERSION=${1:-v1.1.30}
APPLE_ID="michael@blueskysmog.net"
TEAM_ID="V5D7K5HQAY"
SIGN_ID="Developer ID Application: Michael Shaw (V5D7K5HQAY)"

if [ -z "$APP_SPECIFIC_PASSWORD" ]; then
  echo "ERROR: Set APP_SPECIFIC_PASSWORD env var first."
  exit 1
fi

echo "Building BlueSkyDesktop $VERSION..."

python3 -m PyInstaller --windowed --name "BlueSkyDesktop" --icon BlueSkyDesktop.icns -y \
  --hidden-import reportlab.graphics.barcode.code93 \
  --hidden-import reportlab.graphics.barcode.code39 \
  --hidden-import reportlab.graphics.barcode.common \
  --hidden-import reportlab.graphics.barcode.usps \
  --hidden-import reportlab.graphics.barcode.usps4s \
  --hidden-import reportlab.graphics.barcode.ecc200datamatrix \
  BlueSkyDesktopQt_mac.py

echo "Signing..."
codesign --deep --force --options runtime --sign "$SIGN_ID" dist/BlueSkyDesktop.app

echo "Creating DMG..."
create-dmg \
  --volname "Blue Sky Desktop" \
  --window-size 600 400 \
  --icon-size 128 \
  --icon "BlueSkyDesktop.app" 150 200 \
  --app-drop-link 450 200 \
  "dist/BlueSkyDesktop.dmg" \
  "dist/BlueSkyDesktop.app"

echo "Notarizing..."
xcrun notarytool submit dist/BlueSkyDesktop.dmg \
  --apple-id "$APPLE_ID" \
  --team-id "$TEAM_ID" \
  --password "$APP_SPECIFIC_PASSWORD" \
  --wait

echo "Stapling..."
xcrun stapler staple dist/BlueSkyDesktop.dmg

echo "Uploading $VERSION to GitHub..."
gh release create "$VERSION" dist/BlueSkyDesktop.dmg \
  --title "$VERSION" \
  --notes "Blue Sky Desktop $VERSION" 2>/dev/null || \
gh release upload "$VERSION" dist/BlueSkyDesktop.dmg --clobber

echo "Done! Released $VERSION"
