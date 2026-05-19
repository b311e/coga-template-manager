# templx

A template management system for Microsoft Office templates, including Word (.dotm) and Excel (.xltx) templates with standardized branding and themes.

## Setup on Windows

1. **Install Prerequisites**
   - [.NET 8.0 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)
   - [Git for Windows](https://git-scm.com/download/win) (includes Git Bash)
   - Python 3 (optional, only needed for manifest utilities)

2. **Clone Repository**
   ```bash
   git clone https://github.com/b311e/templx.git
   cd templx
   ```

3. **Build OpenXML Tools**
   ```bash
   dotnet build
   ```

4. **Set Up script shortcuts / PATH**
    The repository provides wrapper scripts in `src/scripts/bin` for convenience. Add that folder to your shell `PATH` so you can run commands like `pack`, `unpack`, and the `manifest` helpers directly.

    - Temporary (current Git Bash session):
       ```bash
       source src/scripts/internal/setup-aliases/setup_aliases.sh
       # or
       export PATH="$PWD/src/scripts/bin:$PATH"
       ```

    - PowerShell (current session):
       ```powershell
       $env:PATH = "$PWD\src\scripts\bin;" + $env:PATH
       ```

    - Persistent (Git Bash): add to `~/.bashrc`:
       ```bash
       export PATH="/c/code/templx/src/scripts/bin:$PATH"
       ```

    - Persistent (PowerShell, per-user):
       ```powershell
       [Environment]::SetEnvironmentVariable('PATH', $env:PATH + ';' + (Join-Path $PWD 'src\scripts\bin'), 'User')
       ```

    After adding the folder to your `PATH`, you can run hyphenated wrapper commands (examples below) or use the unified `manifest <action> <resource>` form.

5. **Start Working**
   - Templates are in `builds/<agency>/templates/`
   - Use `unpack` to extract .dotm files for editing
   - Use `pack` to rebuild templates after changes

## Quick Start

### Prerequisites
- [.NET 8.0 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)
- Windows environment (for deployment scripts)
- Access to CGA S Drive (for production deployment)

### Installation
1. Clone the repository:
   ```bash
   git clone https://github.com/b311e/templx.git
   cd templx
   ```

2. Build the application:
   ```bash
   dotnet build
   ```

3. Set up bash aliases (optional, but recommended):
   ```bash
   source src/scripts/internal/setup-aliases/setup_aliases.sh
   ```

## Usage

### Pack & Unpack

- Always work on templates under the `builds` directory.
- Unpacked working files should go under an `in` folder.
- Packed files should go under an `out` folder.

```bash
# Unpack: Expand spreadsheets, documents, or templates.
unpack <target-doc> <output-file-name>

# Pack: Zips the contents into a usable office file.
pack <expanded-file-in> <output-file-name-out>
```

### Create New Files

The `create` command writes the new template file to the current working directory unless you provide a path as the name.

```bash
# Create: Create a OpenXML file from scratch
create <create-type> <output-file-name>
```
**Create Types:**

| Create Type            | Ext.         | Description                 |
|-----------------|--------------|-----------------------------|
| `xl-book`         | xlsx         | Spreadsheet                 |
| `xl-mbook`        | xlsm         | Macro-enabled Spreadsheet   |
| `xl-template`     | xltx         | Spreadsheet Template                    |
| `xl-mtemplate`    | xltm         | Macro-enabled Spreadsheet Template      |
| `word-doc`        | docx         | Document                   |
| `word-mdoc`       | docm         | Macro-enabled Document      |
| `word-template`   | dotx         | Document Template                    |
| `word-mtemplate`  | dotm         | Macro-enabled Document Template      |

### Create New Snippets

Use the `create-snippet-styles` command to create new snippet files from a word doc or template.

```bash
create-snippet-styles <source-expanded-folder> <snippet-id> <style-id>,<style-id>
```

#### Snippet IDs

Snippet Ids should be unique, camel case, and follow this naming convention: 
`[agency][Template][StyleGroup][SnippetType]`.

   - `agency` = core or agency (jbc, olls, lcs, etc.)
   - `Template` = The name of the template the snippet comes from (e.g., Normal, AuditReport, etc.).
   - `StyleGroup` = The group the style belongs to. See [docs/snippets-order-of-operations.yml](docs/snippets-order-of-operations.yml)
   - `SnippetType` = The type of snippet it is (styles, numbering, etc.)

#### Style IDs

Style IDs is the unique name of a style within a word doc's `styles.xml`. 
- Style IDs never include spaces or special characters.
- Note: Although xml is case sensitive, the script is not. You do not need to match the case of a Style ID you enter in the command line with the case in the `styles.xml`.


### Style Utilities

#### List Styles

```bash
# Generate a style list for an unpacked template (saved to the template's docs folder)
style-generate-list <expandedPath>
```

#### Import & Merge Styles

There are three import helpers now:

- `style-import-map` (Python): read `docs/style-map.yml` (preferred) and import ordered styles from partial snippets into a target `styles.xml`. Auto-discovers a nearby `docs/style-map.yml` when run from a template folder and falls back to the legacy `docs/style-mapping.yml` for compatibility. Backups are opt-in with `--backup`.

   ```bash
   # Import styles from partials listed under style-map
   style-import-map <target-styles.xml>

   # Dry-run using auto-discovery of mapping
   style-import-map <target-styles.xml> --dry-run

   # Apply changes and create a .bak backup
   style-import-map <target-styles.xml> --backup
   ```

- `style-import-partial` (Python): replace styles in a target `styles.xml` from a single snippet XML file. Backups are opt-in with `--backup`.

   ```bash
   # Dry-run
   style-import-partial <target-styles.xml> <snippet-xml> --dry-run

   # Replace styles and create backup
   style-import-partial <target-styles.xml> <snippet-xml> --backup
   ```

- `style-import-doc` (.NET Open XML SDK): replace the entire `styles` part in a `.docx`/`.dotx` package with the `styles` part from another document. This uses the Open XML SDK for reliable package manipulation. Backups are opt-in with `--backup`.

   ```bash
   # Dry-run
   style-import-doc <target.docx> <source.docx> --dry-run

   # Apply and create backup
   style-import-doc <target.docx> <source.docx> --backup
   ```

### Clean Up

#### Remove Tracking (RSIDs, ParaId, and textId)

```bash
remove-tracking <target-file>
```

#### Remove No Proof

```bash
remove-noproof <target-file>
```

#### Remove Linked Character Styles

```bash
remove-styles-linkchar <path-to-styles.xml>
```

## Project Structure

```
templx/
├── .templx/                  # templx system metadata (hidden)
│   ├── registry/             # Agency registry and global configuration
│   └── schemas/              # Manifest validation schemas
│
├── builds/                   # Document / templates files under development
│   ├── core/                 # Custom core templates under development
│   └── <agency>/             # Agency templates and workspace files under development
│       ├── assets/
│       ├── system/           # Deployment and automation scripts
│       ├── templates/        # Document templates (Letterhead, Memo, etc.)
│       └── workspace/        # Office workspace templates (Book, Sheet, Normal, etc.)
│
├── dist/                     # Production-ready deployments
│   ├── <agency>/
│   │   ├── assets/
│   │   ├── system/
│   │   ├── templates/ 
│   │   ├── workspace/
│   └── scripts/              # Core deployment scripts (agency independent)
│
├── docs/                     # Additional project-specific documentation
│   ├── template-inventory.md    # Inventory of each agency's templates and their status
│   └── manifest-taxonomy.md  # Manifest system documentation
│
├── core/                  # Core templates and components
│   ├── macros/               # VBA macros and automation scripts
│   ├── manifests/            # Core manifest files
│   ├── partials/             # Reusable components
│   └── workspace/            # Core workspace templates
│
├── resources/                # General reference files and documentation (not project-specific)
│
├── src/
│   ├── OpenXmlApp/           # Core .NET application for template processing
│   ├── scripts/              # Bash utility scripts
│   └── shells/               # Blank shell files used as starting points for new build configs
│       ├── fields/           # Per-template-type field config shells (_committeeLetterhead.json, etc.)
│       ├── _config_build.json
│       └── _config_styles.json
```

## Workflows

### Add a New Template

1. **Create builds folder structure**:

   ```bash
   # Create src, out, in, docs, and config folders for the JBC Letterhead template
   mkdir -p builds/jbc/templates/jbcLetterhead/{src,out,in,docs,config}
   ```

   Copy the appropriate shell files from `src/shells/` into the new `config/` folder and fill them in:

   ```bash
   # Copy build and styles config shells
   cp src/shells/_config_build.json builds/jbc/templates/jbcLetterhead/config/config_build.json
   cp src/shells/_config_styles.json builds/jbc/templates/jbcLetterhead/config/_config_styles.json

   # Copy the field config shell that matches the template type
   cp src/shells/fields/_committeeLetterhead.json builds/jbc/templates/jbcLetterhead/config/config_fields.json
   ```

   Each `config_fields.json` includes a `"template"` key identifying which shell it came from. Add a new file to `src/shells/fields/` for each new template type that has unique fields (e.g., `_memo.json`, `_report.json`).

2. **Generate manifest** to auto-detect the new template:
   ```bash
   manifest generate
   ```

   This scans all `builds/` directories and regenerates `.templx/registry/manifest.json`.

### Development workflow

Work on templates within the appropriate `builds` directory.

1. Put the source file (original, packed file) in the `src` folder. Do not edit the file in this folder.

2. **Unpack** the file in the `src` file, then copy it into the template build folder and rename it to `in`.

   It should look like this:
   ```
   ├── builds/
   │   └── <agency>/
   │       └── templates/
   │           └── <templateName>/
   │               ├── docs/
   │               ├── in/
   │               ├── out/
   │               └── src/
   ```

3. **Make changes:** Make updates to expanded files within the `in` folder.

4. **Pack:** Once updates are complete and ready for testing, pack the files and store in the `out` folder.

5. **Validate:** Validate the packed file using one of the methods below.

   a. Validate the file using the `validate` command (recommended). This will generate a txt file within a `reports` folder in the same directory as the target file.
   
   b. Right click on the compiled file, select "Validate OOXML", then review the report. 

6. Download the file in the `out` folder and test on your local device in Microsoft Word, Excel, or PowerPoint.


## Template Registry & Manifests

### Manifest Structure

All template metadata is tracked in a single master manifest:

- **Master Manifest**: `.templx/registry/manifest.json` — all agencies and templates in one file, auto-generated by scanning `builds/`
- **Agency Registry**: `.templx/registry/agencies.json` — agency metadata (name, code, active flag) used by the inventory command
- **Schema & Taxonomy**: `.templx/schemas/manifest-schema.json`
  - **Schema documentation:** `docs/manifest-taxonomy.md`

Never edit `.templx/registry/manifest.json` by hand. Run `manifest generate` to regenerate it from the current state of `builds/`. Partials are automatically included — no separate partials manifest needed.

### Manifest Commands

```bash
# Regenerate master manifest (scans all builds/)
manifest generate

# Validate master manifest
manifest validate

# List all templates
manifest list

# Update a template's status
manifest update status jbc jbcMemo active

# Get guidance for adding a new template
manifest add template jbc jbcReport word-doc-template
```

## Deployment

### Deployment Scripts

- `dist/scripts/workspaceInstallPreProd.bat` - Deploy to PreProd environment
- `dist/scripts/workspaceDeploy.bat` - Deploy from PreProd to Production
- `dist/jbc/workspace/jbcWorkspaceInstall.bat` - JBC end-user installation script

## Template Inventory System

The template inventory can be found at [docs/template-inventory.md](docs/template-inventory.md).

### Generate Template Inventory
The inventory system provides a user-friendly overview of all templates across all agencies:

```bash
# Generate template inventory report
inventory generate

# Show help
inventory --help
```

### Inventory Output
The inventory generates `docs/template-inventory.md` with:
- **Current Status**: Template status pulled from manifests
- **Templates**: Complete listing of House, Senate, JBC, LCS, OLLS, OSA templates
- **Timestamp**: Generated date for tracking currency


**Note:** If you add, remove, or change templates or statuses in any agency manifest, rerun the inventory command to update the list.

### Getting Help
- Check the `docs/` folder for additional documentation
- Review template XML structure in `resources/defaults/`
- Contact the development team for support
