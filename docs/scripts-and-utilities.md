# Scripts and Utilities

Reference for all command-line tools in `src/scripts/`.

## Setup

Add `src/scripts/bin/` to your PATH by sourcing the setup script once per session:

**Bash (Git Bash):**
```bash
source src/scripts/internal/setup_aliases/setup_aliases.sh
```

**PowerShell:**
```powershell
. .\src\scripts\internal\setup_aliases\setup_aliases.ps1
```

## Directory Structure

```
src/scripts/
├── bin/        # Thin wrapper scripts — add this to PATH
├── commands/   # Actual implementations
└── internal/   # Infrastructure (command_dispatch, setup_aliases) — not user-facing
```

`bin/` wrappers follow this pattern:
```bash
#!/usr/bin/env bash
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
exec "$SCRIPT_DIR/../commands/COMMAND/COMMAND" "$@"
```

---

## Commands

### `pack` / `unpack`

Pack or unpack an OpenXML file to/from an expanded directory.

```bash
unpack builds/jbc/templates/jbcMemo/out/jbcMemo.dotm
pack   builds/jbc/templates/jbcMemo/out/jbcMemo_expanded jbcMemo.dotm
```

---

### `create`

Create a new template or document from a shell.

```bash
create <file-type> [name]
```

#### `create snippet_styles`

Create a new styles snippet file.

```bash
create_snippet_styles
```

---

### `validate`

Validate an Office Open XML file using the OpenXML SDK. Supports `.docx`, `.docm`, `.dotx`, `.dotm`, `.xlsx`, `.xlsm`, `.xltx`, `.xltm`, `.pptx`, `.pptm`, `.potx`, `.potm`.

```bash
validate builds/jbc/templates/jbcMemo/out/jbcMemo.dotm
```

Creates a timestamped report in a `reports/` folder next to the validated file:

```
out/
├── jbcMemo.dotm
└── reports/
    └── jbcMemo-validation-20251126-163258.txt
```

Report includes: file metadata, error descriptions, XPath locations, part URIs, and pass/fail summary.

Requires the `OpenXmlApp` to be built:
```powershell
cd src\OpenXmlApp
dotnet build --configuration Debug
```

---

### `manifest`

Manage the master manifest at `.templx/registry/manifest.json`. Never edit it by hand — use these commands.

```bash
manifest generate                              # Scan all builds/ and regenerate manifest
manifest validate                              # Validate manifest is present and valid JSON
manifest list                                  # List all templates across builds/
manifest update status <agency> <template> <status>   # Update a template's status
manifest add template <agency> <template> <type>      # Print guide for adding a new template
```

---

### `inventory`

Generate `docs/template-inventory.md` from the master manifest.

```bash
inventory
```

---

### `style`

Style import/export utilities.

```bash
style list <template-path>                     # Extract and list styles from a Word template
style import <target> <source>                 # Import styles from source into target
style import snippet <target> <snippet-file>   # Import styles from an XML snippet
```

---

### `cleanup`

Automated cleanup utilities for WordprocessingML files.

#### `cleanup wordml`

Full automated cleanup of an expanded WordprocessingML directory: removes RSID attributes, fixes Normal style, removes `noProof`/`iCs`/`szCs`/`bCs` elements, removes bidirectional language settings, and cleans empty `rPr`/`pPr` blocks.

```bash
cleanup wordml <expanded-template-dir>
```

#### `cleanup linkchar`

Remove linked character styles (`Heading1Char`–`Heading9Char`, `QuoteChar`, `SubtitleChar`, `HeaderChar`, `FooterChar`) and their `w:link` references from a `styles.xml`.

```bash
cleanup linkchar <styles.xml-path>
```

#### `cleanup noproof`

Remove `w:noProof` elements (and empty `w:rPr` blocks left behind) from a Word XML file.

```bash
cleanup noproof <xml-file-path>
```

#### `cleanup tracking`

Remove RSID attributes (`w:rsidR`, `w:rsidRPr`, etc.) and `w14:paraId`/`w14:textId` attributes from a Word XML file.

```bash
cleanup tracking <xml-file-path>
```

---

### `xpathsel`

Thin wrapper for `xmlstarlet` XPath queries.

```bash
xpathsel <file.xml> '/path/to/node'
```
