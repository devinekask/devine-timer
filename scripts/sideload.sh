#!/bin/sh
# Copies an add-in manifest into PowerPoint for Mac's sideload folder.
# Usage: scripts/sideload.sh <manifest.xml> <target-name.xml>
MANIFEST="$1"
TARGET_NAME="$2"
WEF="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"

if mkdir -p "$WEF" 2>/dev/null && cp "$MANIFEST" "$WEF/$TARGET_NAME" 2>/dev/null; then
  echo "Sideloaded $MANIFEST -> $WEF/$TARGET_NAME"
  echo "Restart PowerPoint, then: Insert > Add-ins > My Add-ins."
  exit 0
fi

cat <<MSG
macOS blocked access to PowerPoint's container folder.

Fix it one of two ways:
  A) System Settings > Privacy & Security > Full Disk Access > enable your terminal app,
     restart the terminal, and run this command again.
  B) Copy the manifest manually with Finder (opening it now):
       1. In Finder press Shift+Cmd+G and go to:
          ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/
       2. Create a folder named "wef" if it doesn't exist.
       3. Drag $MANIFEST into it, then restart PowerPoint.
MSG
open -R "$MANIFEST"
exit 1
