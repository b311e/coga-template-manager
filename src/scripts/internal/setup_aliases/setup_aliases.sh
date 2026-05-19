#!/usr/bin/env bash
# Setup script to add the templx bin/ to PATH for the current shell session
# Run this with: source setup_aliases.sh

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
BIN_PATH="$(cd "$SCRIPT_DIR/../../bin" && pwd)"

if [[ ":$PATH:" != *":$BIN_PATH:"* ]]; then
    export PATH="$BIN_PATH:$PATH"
    echo "Added $BIN_PATH to PATH"
    echo ""
    echo "Top-level commands now available:"
    echo "  templx <command> [subcommand] [args...]  - Main dispatcher (preferred)"
    echo ""
    echo "Or call top-level commands directly:"
    echo "  pack, unpack, create, validate, xpathsel"
    echo "  style, inventory, cleanup, manifest"
    echo ""
    echo "Run 'templx help' for the full command list."
else
    echo "Scripts bin directory already in PATH"
fi
