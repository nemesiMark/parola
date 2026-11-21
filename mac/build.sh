#!/bin/zsh
# Crea Parola.app per Mac con Apple Silicon (M1 e successivi) e la installa in Applicazioni.
# Uso:  ./mac/build.sh            crea e installa
#       ./mac/build.sh --no-install   crea soltanto (in mac/build/Parola.app)
set -euo pipefail

ROOT="${0:A:h:h}"
MAC="$ROOT/mac"
BUILD="$MAC/build"
APP="$BUILD/Parola.app"

echo "→ Preparo la cartella"
/bin/rm -rf -- "$APP"
mkdir -p "$APP/Contents/MacOS" "$APP/Contents/Resources/web/vendor" "$BUILD/tmp"

echo "→ Compilo l'app (arm64)"
swiftc -O -target arm64-apple-macos12 \
  -framework Cocoa -framework WebKit -framework UniformTypeIdentifiers \
  "$MAC/main.swift" -o "$APP/Contents/MacOS/Parola"

echo "→ Copio l'editor"
cp "$ROOT/index.html" "$ROOT/style.css" "$ROOT/app.js" "$ROOT/docx-export.js" "$APP/Contents/Resources/web/"
cp "$ROOT/vendor/"*.js "$APP/Contents/Resources/web/vendor/"
cp "$MAC/Info.plist" "$APP/Contents/Info.plist"

echo "→ Creo l'icona"
ICONSET="$BUILD/tmp/Parola.iconset"
/bin/rm -rf -- "$ICONSET"; mkdir -p "$ICONSET"
swiftc -O "$MAC/icon.swift" -o "$BUILD/tmp/icon"
"$BUILD/tmp/icon" "$BUILD/tmp/icon-1024.png"
for s in 16 32 128 256 512; do
  sips -z $s $s "$BUILD/tmp/icon-1024.png" --out "$ICONSET/icon_${s}x${s}.png" >/dev/null
  d=$((s * 2))
  sips -z $d $d "$BUILD/tmp/icon-1024.png" --out "$ICONSET/icon_${s}x${s}@2x.png" >/dev/null
done
iconutil -c icns "$ICONSET" -o "$APP/Contents/Resources/Parola.icns"

echo "→ Firmo l'app (firma locale)"
xattr -cr "$APP"
codesign --force --deep --sign - "$APP"

if [[ "${1:-}" != "--no-install" ]]; then
  echo "→ Installo in /Applications"
  /bin/rm -rf -- "/Applications/Parola.app"
  cp -R "$APP" "/Applications/Parola.app"
  /System/Library/Frameworks/CoreServices.framework/Frameworks/LaunchServices.framework/Support/lsregister -f "/Applications/Parola.app" || true
  echo "✓ Parola è in Applicazioni"
else
  echo "✓ Creata: $APP"
fi
