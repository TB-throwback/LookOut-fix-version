#!/bin/bash

FILE="lookout_fix_version-beta-tb.xpi"

choice1() {
  echo "Warning! this will delete existing $FILE"
  read -rp "Continue [Y/N]? " c

  case "$c" in
    [Yy]) old_del ;;
    [Nn]) exit_script ;;
    *) choice1 ;;
  esac
}

old_del() {
  rm -f "$FILE"
  build
}

build() {
  # Try to find 7z (assumes it's installed and in PATH)
  if command -v 7z >/dev/null 2>&1; then
    ZIP_CMD="7z"
  elif command -v 7za >/dev/null 2>&1; then
    ZIP_CMD="7za"
  else
    echo "Error: 7z not found. Please install p7zip."
    exit 1
  fi

  "$ZIP_CMD" a -tzip "$FILE" \
    _locales api icons options scripts \
    background.html background.js message-content-script.js changes.txt LICENSE manifest.json
}

exit_script() {
  echo "Goodbye"
  read -rp "Press Enter to continue..."
  exit 0
}

# Start
choice1
