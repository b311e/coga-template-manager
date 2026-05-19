#!/usr/bin/env bash
# templx Command Dispatcher

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

case "$1" in
    pack)
        exec "$SCRIPT_DIR/../../commands/pack/pack" "$@"
        ;;
    unpack)
        exec "$SCRIPT_DIR/../../commands/unpack/unpack" "$@"
        ;;
    create)
        exec "$SCRIPT_DIR/../../commands/create/create" "$@"
        ;;
    validate)
        exec "$SCRIPT_DIR/../../commands/validate/validate" "$@"
        ;;
    style)
        exec "$SCRIPT_DIR/../../commands/style/style" "$@"
        ;;
    inventory)
        exec "$SCRIPT_DIR/../../commands/inventory/inventory" "$@"
        ;;
    cleanup)
        exec "$SCRIPT_DIR/../../commands/cleanup/cleanup" "$@"
        ;;
    manifest)
        exec "$SCRIPT_DIR/../../commands/manifest/manifest" "$@"
        ;;
    help|--help|-h)
        echo "templx"
        echo ""
        echo "Usage: templx <command> [subcommand] [args...]"
        echo ""
        echo "Commands:"
        echo "  pack <target-file>                         - Package OpenXML document from expanded directory"
        echo "  unpack <target-file>                       - Unpack OpenXML document to expanded directory"
        echo "  create <file-type> [name]                  - Create new template/document"
        echo "  validate <target-file>                     - Validate OpenXML document against schema"
        echo "  style list <template>                      - Extract and list styles from Word template"
        echo "  style import <target> <source>             - Import styles from source to target Word doc"
        echo "  style import snippet <target> <snippet>    - Import styles from XML snippet"
        echo "  inventory                                  - Generate template inventory"
        echo "  cleanup wordml <expanded-dir>              - Automated cleanup of WordprocessingML files"
        echo "  manifest generate                          - Regenerate master manifest (.templx/registry/manifest.json)"
        echo "  manifest validate                          - Validate master manifest"
        echo "  manifest list                              - List all templates"
        echo "  manifest update status <a> <t> <status>    - Update template status in master manifest"
        echo "  manifest add template <a> <t> <type>       - Guide for adding new template"
        echo "  help                                       - Show this help message"
        echo ""
        echo "Run '<command> help' for command-specific options"
        ;;
    *)
        echo "Unknown command: $1"
        echo "Run 'templx help' for usage information"
        exit 1
        ;;
esac