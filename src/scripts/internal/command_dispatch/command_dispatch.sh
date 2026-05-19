#!/usr/bin/env bash
# templx command dispatcher

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
COMMANDS_DIR="$SCRIPT_DIR/../../commands"

show_help() {
    cat <<EOF
templx — OpenXML template toolchain

Usage: templx <command> [subcommand] [args...]

Commands:
  pack <expanded-dir> [out]                  - Package OpenXML document from expanded directory
  unpack <file>                              - Unpack OpenXML document to expanded directory
  create <file-type> [name]                  - Create new template/document from scratch
  validate <file>                            - Validate OpenXML document against schema
  xpathsel <file> <xpath>                    - Run an XPath query against an XML file

  style list <expanded-dir>                  - List styles in an unpacked Word template
  style extract <expanded-dir> <name> <ids>  - Extract styles into a partial snippet
  style import map <target> [--mapping <f>]  - Map-driven import from partials manifest
  style import partial <target> <snippet>    - Import all styles from a single snippet
  style import doc <target> <source-doc>     - Copy styles from one Word document to another

  inventory generate                         - Generate docs/template-inventory.md
  cleanup wordml <expanded-dir>              - Full automated cleanup of WordprocessingML
  cleanup linkchar <styles.xml>              - Remove linked character styles
  cleanup noproof <xml-file>                 - Remove noProof elements
  cleanup tracking <xml-file>                - Remove RSID/tracking attributes

  manifest generate                          - Regenerate master manifest
  manifest validate                          - Validate master manifest
  manifest list                              - List all builds and templates
  manifest update status <a> <t> <status>    - Update template status
  manifest add template <a> <t> <type>       - Print instructions for new template

  help                                       - Show this help message

Run '<command> help' (or '<command> <subcommand> help') for command-specific options.
EOF
}

COMMAND="${1:-}"

case "$COMMAND" in
    pack|unpack|create|validate|xpathsel|style|inventory|cleanup|manifest)
        shift
        exec "$COMMANDS_DIR/$COMMAND/$COMMAND" "$@"
        ;;
    help|--help|-h|"")
        show_help
        ;;
    *)
        echo "Unknown command: $COMMAND" >&2
        echo "Run 'templx help' for usage information." >&2
        exit 1
        ;;
esac
