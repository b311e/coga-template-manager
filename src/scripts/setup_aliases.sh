#!/usr/bin/env bash
# Setup script to add project scripts to PATH
# Run this with: source setup_aliases.sh

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SCRIPTS_PATH="$SCRIPT_DIR/src/scripts"

# Add to PATH if not already there
if [[ ":$PATH:" != *":$SCRIPTS_PATH:"* ]]; then
    export PATH="$SCRIPTS_PATH:$PATH"
    echo "Added $SCRIPTS_PATH to PATH"
    echo ""
    echo "You can now use these commands directly:"
    echo "  unpack <file>                    - Unpack OpenXML file to 'expanded' folder"
    echo "  pack <sourceDir> <outputFile>    - Pack directory to OpenXML file"
    echo "  style_list list <templateName>   - Generate style list for template"
    echo ""
    echo "Examples:"
    echo "  unpack templates/jbc/jbcBook/source/Book.xltx"
    echo "  style_list list jbcNormal"
else
    echo "Scripts directory already in PATH"
fi