#!/bin/bash
# Fabrique « Adresses Outlook.app » sur un Mac (la construction automatique de
# GitHub le fait à chaque modification de mac/), et la fausse messagerie qui
# sert aux tests. Une seule app pour les Mac Intel et Apple Silicon.
#
#   bash mac/natif/construire.sh dist
#
# Résultat : dist/Adresses Outlook.app, dist/Adresses Outlook pour Mac.zip,
# dist/FauxOutlook.app (tests seulement).

set -euo pipefail
ICI="$(cd "$(dirname "$0")" && pwd)"
MAC="$(dirname "$ICI")"
OUT="${1:-$ICI/build}"
mkdir -p "$OUT"
OUT="$(cd "$OUT" && pwd)"

# Compile pour les deux familles de processeurs, réunies en un seul fichier.
universel() {
  swiftc -O -target arm64-apple-macos11 -o "$2.arm64" "$1"
  swiftc -O -target x86_64-apple-macos11 -o "$2.x86_64" "$1"
  lipo -create -output "$2" "$2.arm64" "$2.x86_64"
  rm -f "$2.arm64" "$2.x86_64"
}

# paquet <nom> <identifiant> <exécutable>
paquet() {
  local APP="$OUT/$1.app"
  rm -rf "$APP"
  mkdir -p "$APP/Contents/MacOS" "$APP/Contents/Resources"
  cp "$3" "$APP/Contents/MacOS/$1"
  printf 'APPL????' > "$APP/Contents/PkgInfo"
  cat > "$APP/Contents/Info.plist" <<PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>CFBundleExecutable</key><string>$1</string>
  <key>CFBundleIdentifier</key><string>$2</string>
  <key>CFBundleName</key><string>$1</string>
  <key>CFBundleDisplayName</key><string>$1</string>
  <key>CFBundlePackageType</key><string>APPL</string>
  <key>CFBundleShortVersionString</key><string>2.0</string>
  <key>CFBundleVersion</key><string>2</string>
  <key>CFBundleIconFile</key><string>AppIcon</string>
  <key>LSMinimumSystemVersion</key><string>11.0</string>
  <key>NSHighResolutionCapable</key><true/>
  <key>NSPrincipalClass</key><string>NSApplication</string>
</dict>
</plist>
PLIST
}

echo "== Compilation"
universel "$ICI/main.swift" "$OUT/adresses-outlook"
universel "$ICI/FauxOutlook.swift" "$OUT/faux-outlook"

echo "== Adresses Outlook.app"
paquet "Adresses Outlook" "com.aether.adresses-outlook" "$OUT/adresses-outlook"
APP="$OUT/Adresses Outlook.app"
cp "$ICI/AppIcon.icns" "$APP/Contents/Resources/AppIcon.icns"
mkdir -p "$APP/Contents/Resources/web"
cp "$MAC/interface.html" "$MAC/planif.js" "$MAC/moteur.js" "$APP/Contents/Resources/web/"
# Signature locale : macOS retient l'autorisation Accessibilité pour cette app.
codesign --force --deep --sign - "$APP"
codesign --verify --deep --strict "$APP"

echo "== FauxOutlook.app (tests)"
paquet "FauxOutlook" "com.aether.fauxoutlook" "$OUT/faux-outlook"
codesign --force --deep --sign - "$OUT/FauxOutlook.app"

rm -f "$OUT/adresses-outlook" "$OUT/faux-outlook"
rm -f "$OUT/Adresses Outlook pour Mac.zip"
ditto -c -k --keepParent "$APP" "$OUT/Adresses Outlook pour Mac.zip"
ls -la "$OUT"
