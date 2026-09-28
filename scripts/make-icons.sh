#!/bin/sh
# Renders the add-in icons from the SVG sources (Devine D with clock hands). Needs Inkscape.
# icon-small.svg drops the minute markers and uses thicker hands, so it stays readable at 32 px.
cd "$(dirname "$0")/../docs/assets" || exit 1
inkscape icon-small.svg --export-type=png --export-filename=icon-32.png -w 32 -h 32
inkscape icon.svg --export-type=png --export-filename=icon-64.png -w 64 -h 64
