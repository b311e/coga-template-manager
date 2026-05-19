#!/usr/bin/env bash
# Setup script to add project scripts to PATH
# Run this with: source setup_aliases.sh

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SCRIPTS_PATH="$SCRIPT_DIR/../.."
BIN_PATH="$SCRIPTS_PATH/bin"

# Add to PATH if not already there (prefer bin/ so top-level names map to wrapper scripts)
if [[ ":$PATH:" != *":$BIN_PATH:"* ]]; then
    export PATH="$BIN_PATH:$PATH"
    echo "Added $BIN_PATH to PATH"
    echo ""
    echo "You can now use these commands directly:"
    echo "  create <type> [output file name]    - Create new OpenXML templates/documents"
    echo "  unpack <input file name>            - Unpack OpenXML file to 'expanded' folder"
    echo "  pack <sourceDir> <outputFile>       - Pack directory to OpenXML file"
    echo "  style import <target> <source>      - Import styles from source to target Word doc"
    echo "  style import snippet <target> <snippet-file> [id] - Import styles from XML snippet"
    echo "  cleanup wordml <expanded-dir>       - Automated cleanup of WordprocessingML files"
    echo "  style list <templatePath>           - Generate style list for template"
    echo "  manifest <command> [options]        - Manage template manifests and registry"
    echo "  validate <file-path>                - Validate OOXML files and generate reports"
    echo ""
    echo "Examples:" 
    echo "  create excel-book-template Book"
    echo "  manifest generate      # regenerate master manifest"
    echo "  manifest validate      # validate master manifest"
    echo "  create word-doc-template jbcLetterhead"
    echo "  unpack templates/jbc/jbcBook/src/Book.xltx"
    echo "  style list builds/jbc/workspace/jbcNormal"
else
    echo "Scripts bin directory already in PATH"
fi
