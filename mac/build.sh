#!/bin/zsh
# Crea Word.app per Mac con Apple Silicon (M1 e successivi) e la installa in Applicazioni.
# Uso:  ./mac/build.sh                 crea e installa (e toglie la vecchia Parola.app)
#       ./mac/build.sh --no-install    crea soltanto (in mac/build/Word.app)
set -euo pipefail

ROOT="${0:A:h:h}"
MAC="$ROOT/mac"
BUILD="$MAC/build"
APP="$BUILD/Word.app"
ID="it.nemesi.word"
OLD_ID="it.parola.editor"

echo "→ Preparo la cartella"
/bin/rm -rf -- "$APP" "$BUILD/Parola.app"
mkdir -p "$APP/Contents/MacOS" "$APP/Contents/Resources/web/vendor" "$BUILD/tmp"

echo "→ Compilo l'app (arm64)"
swiftc -O -target arm64-apple-macos12 \
  -framework Cocoa -framework WebKit -framework UniformTypeIdentifiers \
  "$MAC/main.swift" -o "$APP/Contents/MacOS/Word"

echo "→ Copio l'editor"
cp "$ROOT/index.html" "$ROOT/style.css" "$ROOT/app.js" "$ROOT/docx-export.js" "$ROOT/icon.svg" "$ROOT/installa.js" "$APP/Contents/Resources/web/"
cp "$ROOT/vendor/"*.js "$APP/Contents/Resources/web/vendor/"
cp "$MAC/Info.plist" "$APP/Contents/Info.plist"

echo "→ Creo l'icona (da icon.svg)"
ICONSET="$BUILD/tmp/Word.iconset"
/bin/rm -rf -- "$ICONSET"; mkdir -p "$ICONSET"
swiftc -O "$MAC/icon.swift" -o "$BUILD/tmp/icon"
"$BUILD/tmp/icon" "$ROOT/icon.svg" "$BUILD/tmp/icon-1024.png"
for s in 16 32 128 256 512; do
  sips -z $s $s "$BUILD/tmp/icon-1024.png" --out "$ICONSET/icon_${s}x${s}.png" >/dev/null
  d=$((s * 2))
  sips -z $d $d "$BUILD/tmp/icon-1024.png" --out "$ICONSET/icon_${s}x${s}@2x.png" >/dev/null
done
iconutil -c icns "$ICONSET" -o "$APP/Contents/Resources/Word.icns"

echo "→ Firmo l'app (firma locale)"
xattr -cr "$APP"
codesign --force --deep --sign - "$APP"

if [[ "${1:-}" != "--no-install" ]]; then
  LSREG=/System/Library/Frameworks/CoreServices.framework/Frameworks/LaunchServices.framework/Support/lsregister

  # Il documento salvato dentro la vecchia app (Parola) passa a Word, se Word non ne ha già uno
  if [[ -d "$HOME/Library/WebKit/$OLD_ID" && ! -d "$HOME/Library/WebKit/$ID" ]]; then
    echo "→ Porto in Word il documento salvato da Parola"
    osascript -e 'tell application id "'$OLD_ID'" to quit' >/dev/null 2>&1 || true
    sleep 1
    cp -R "$HOME/Library/WebKit/$OLD_ID" "$HOME/Library/WebKit/$ID"
  fi

  echo "→ Installo in /Applications"
  osascript -e 'tell application id "'$ID'" to quit' >/dev/null 2>&1 || true
  /bin/rm -rf -- "/Applications/Word.app"
  cp -R "$APP" "/Applications/Word.app"
  $LSREG -f "/Applications/Word.app" || true

  if [[ -d "/Applications/Parola.app" ]]; then
    echo "→ Elimino la vecchia Parola.app"
    osascript -e 'tell application id "'$OLD_ID'" to quit' >/dev/null 2>&1 || true
    $LSREG -u "/Applications/Parola.app" >/dev/null 2>&1 || true
    /bin/rm -rf -- "/Applications/Parola.app"
  fi
  echo "✓ Word è in Applicazioni"
else
  echo "✓ Creata: $APP"
fi
