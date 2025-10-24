#!/usr/bin/env bash
# COGA Template Manager Command Dispatcher

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

case "$1" in
    pack|unpack|create)
        exec "$SCRIPT_DIR/openxml" "$@"
        ;;
    inventory)
        shift
        exec "$SCRIPT_DIR/inventory" "$@"
        ;;
    help|--help|-h)
        echo "COGA Template Manager"
        echo ""
        echo "Usage: coga <command> [args...]"
        echo ""
        echo "Commands:"
        echo "  pack <file>         Package OpenXML document"
        echo "  unpack <file>       Unpack OpenXML document"
        echo "  create <type>       Create new template/document"
        echo "  inventory <cmd>     Generate template inventory"
        echo "  help                Show this help message"
        echo ""
        echo "See 'coga <command> help' for command-specific options"
        ;;
    *)
        echo "Unknown command: $1"
        echo "Run 'coga help' for usage information"
        exit 1
        ;;
esac