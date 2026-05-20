# Updates and Future Tasks

## 2026-05-19 - Refactor To Do's and Notes

- [X] Make src/scripts and commands use `-`, not `.`
- [X] Replace pack/unpack/test-clean with bash (no runtime needed, not replying on Open XML App)
- [X] Replace style import with Python (not replying on Open XML App)
- [X] Move outside schemas to resources/schemas
- [X] Move `C:\code\templx\resources\specs\ECMA 376 Office Open XML` schema files to resources/schemas/ooxml
- [X] Download referenced dublin core schemas and put them under resources/schemas as well.
- [X] Update validate command: Python + lxml — can validate against XSD schemas.
- [X] rename "base" folders to "default"
- [X] rename "workspace" folders to "foundation"
- [X] Update template folder names to use kabob-case (e.g., builds/base/templates/committeeLetterhead --> builds/base/templates/committee-letterhead)
- [X] get rid of .templx/registry and everything within it.
- [X] rename partials/document folder to partials/content
- [X] rename partials folder to components
- [X] strip agency prefix from template folder names
- [X] reanem the `build/templates/<agency>/<template-name>/config` folder to be `/data`.
- [X] make all bat scripts use - not _
- [X] move all deployment (dist) scripts to templx/dist folder
- [X] update builds folder structure (rename existing files)

    ```
    builds/
    ├── default/                      ← theme layer (was "base")
    │   ├── foundation/               ← shared-material location (was "workspace")
    │   │   ├── normal.dotx           ← Word foundation file (keep — mirrors Office)
    │   │   ├── assets
    │   │   ├── components/
    │   │   │   ├── content/
    │   │   │   ├── macros/
    │   │   │   ├── numbering/
    │   │   │   └── styles/
    │   │   └── data/                 ← system-default field values (the fallback layer)
    │   └── templates/
    └── <agency>/                     ← theme layer - Agency (column 1)
        ├── foundation/
        │   └── data/
        │       └── contact.json    
        └── templates/
            └── report/               ← Category (column 2) 
                ├── _base/            ← within-variant parent (the one true "base")
                ├── annual/           ← Template (Column 3): "Annual Report"
                │   └── data/
                │       └── fields.json
                └── quarterly/        ← Template (Column 3): "Quarterly Report"
    ```

- [] update pack script to pull from the ooxml indicator for the file type

- [ ] Update logic for style import. rename it to style sync

Every layer has its own components/styles/styles.json and components/styles/styles.xml. Listing a style in a layer's styles.json means "this layer overrides that style." Not listing it means "this layer is silent on it — keep looking up." Everything unstated is inherited. For any given style, the sync walks from most-specific layer to least and stops at the first styles.json that mentions it, then pulls from the styles.xml in the same directory.

    ```
    template/styles.json
    → category _base/styles.json
        → agency foundation/styles.json
        → default _base/styles.json
            → default foundation/styles.json   (the floor — must define everything)
    ```

Property-level (merge): annual's Heading2 lists only { "color": "red" }, and the resolver merges that over the inherited Heading2 (which itself may be agency-over-default). Change one property, write one property. Default's later spacing change does flow through, because annual only overrode color. 

Resolution becomes: for each style, for each property, walk the layers and take the first that sets it. `style sync` then shows provenance per property when it matters:

    ```
    Heading2:
    color    → annual          (template override)
    font     → acme/foundation  (agency override)
    size     → default          
    spacing  → default          
    ```

- [ ] Create script that creates `.templx/index.json. This will replace the manifest.json and agencies.json and the scripts that creates those. The command should be: 
builds reindex. OR: SQL database? Pros? Cons?
- [] index.json should be saved under .templx/builds
- [] each template gets its own folder here: .templx/builds/<agency>/templates/<template-name>

    ```json
    {
    "version": 1,
    "builds": [
        {
        "id": "",
        "created_at": "2026-05-19T14:32:01Z",
        "status": "success",
        "agency": "",
        "category": "",
        "template": "",
        "output": "builds/2026-05-19-abc123/output.dotx"
        }
    ]
    }
    ```

- [] create the following commands:
    - builds list -> reads index.json, prints the table
    - builds show <id> -> reads that build's metadata.json for full detail
    - builds open <id> -> opens the output file

- [] create very clear list of categories (Memo, Report, Agenda, Letterhead, Committee Letterhead, Analysis)

- [ ] Document style sync instructions: 

    1. If the template has a custom style, put the StyleId into the styles.json. 
    2. Run style sync

### Future

- [] create `.templx/config` (just create shell for now)
- [] create `.templx/logs` (just create shell for now, with note that it isnt working yet)


### Notes and New Requirements

We're building a design system.

Tokens
- colors, fonts, styles (the smallest reusable values)
- Tokens should be named semantically, not literally. Not blue-700 but color-primary. Not garamond-14 but text-body. Semantic names are what let the agency layer override color-heading-primary to their blue without every deliverable knowing the literal value.

Components: 
- content blocks, macros, styles, numbering, reusable components

library vs instances:
- base/agency workspaces are the library (definitions available to be used)
- deliverables are instances (specific compositions that consume the library)


CLI Requirements
- Follow CLI Best Practices
- git-like CLI design
    - Stateful working directory: Commands operate on context inferred from where you are



## 2025-01-27

### partials updates
- [x] change "snippet id" to "partial name" in partial.xml
- [x] change file naming for partials from kabob-case to snake_case
- [x] partial name should not include "styles" anymore. Need to remove this from the pattern entirely. The naming convention for partial_name is: [agency][Template][StyleGroup] 

### config (main build config file) updates
- [x] located in builds/agency/templateName/config
- [x] delete "id" field
- [x] move "extension" above "status"

### style-map, now config_styles updates
- [x] changed "style-map.yaml" to "config_styles.json"
- [x] config_styles.json will always be: templateName/config/config_styles.json"
- [x] the "partials" fields determine the source partial to import into the template. 
    - [x] "agency" targets the agency (folder name under builds/)
    - [x] "template" targets the template (folder name under builds/agency/ in pascalCase).
    - note: i am not sold on this naming convention.
- [x] config_styles schema is located in src/_config_styles.json

### schemas updates
- [x] moved .templx/schemas to src/schemas
- [x] renamed schemas/manifest-schema to schemas/manifest_schema.json
- [x] updated fields in manifest_schema.json
    - [x] template_entry is now build_entry
- [x] created build_types_manifest.json
    - [x] renamed "system" build_type to "dist"
- [x] created config_types_manifest.json
- [x] created style_groups_schema.json

### src/shells updates
- [x] created src/shells folder, moved all shells to this folder, and created new ones (_config_build.json, config_styles.json, sonfig_fields.json)
- [x] created partials/styles shell folder and front_matter.xml"

### To do:


---
- [ ] fix src/shells/config_fields.json to have template-specific fields for committee_templates (to reference collection field in committee templates?)
- [ ] create all other style partials shell files
- [ ] update other fields in schemas to flow with new system
- [ ] update manifest scripts to continue generate schemas in .templx/schemas, but reference the schemas under src/schemas
- [ ] make sure this schema is correct and follows best practices
- [ ] config_styles script must rely on the order in the json, rather than the old style-map-order-of-operations document.
- [ ] update scripts so config_styles.json is used to configure the source of style partials imported into a template.
- [ ] update scripts so the "config_type" field determines the type of config file (styles, numbering, macros, document, etc.) and therefore the target (referenced in the src/schemas).
- [ ] clean up the docs folders
- [ ] update the main README
- [ ] standardize naming conventions. i want to move to underscores (_), but need to make sure that makes sense everywhere
- [ ] figure out config_styles field names (style_group vs styles vs partials, etc.)


## DATE?????

src > scripts

- [ ] pull README files out of the individual commands folders, put everything directly under commands and undo nesting.
- [ ] fix bugs in commands (duplicate src, etc.) and add ability to use WIndows and Unix file paths for everything.
- [ ] clean up and remove obsolete code from scripts
- [ ] standardize naming conventions for scripts
- [ ] add tab complete for scripts
- [ ] update all help info to be accurate

partials

- [ ] rename snippet_id --> partial_id
- [ ] change everything to kabob_case

manifests & naming

- [ ] standardize manifests so there is just one master manifest
- [ ] update everything to use kabob_case
- [ ] update core --> base
- [ ] update manifest taxonomy to use kabob_case
- [ ] update taxonomy
      - remove agency from the top
      - remove manifest_version
      - templates --> builds (space > builds)
      - name --> build_name
      - type --> build_type
      - extention --> build_ext
      - remove src from manifest json. why is that even in there
      - status --> build_status
      - partials --> build_partial
      - add build_space as a field
        - base
        - jbc
        - lcs, etc... 
- [ ] move style-map.md from docs into a config folder
- [ ] change style-map.md --> style_map.json
- [ ] update all partial_id to be in kabob_case
- [ ] delete all old manifests

{
  "manifest_version": "1.0",
  "generated": "2026-01-09T...",
  "builds": {
    "core": {
      "workspace": {
        "normal": {
          "type": "word-doc-template",
          "extension": ".dotm",
          "directory": "builds/core/workspace/normal",
          "collection": "normal",
          "status": "active",
          "partials": {
            "body_styles": {
              "path": "builds/core/.../body-styles.xml",
              "partial_type": "styles"
            }
          }
        }
      },
      "templates": {
        "committee_letterhead": {
          "type": "word-doc-template",
          "extension": ".dotm",
          "directory": "builds/core/templates/committee_letterhead",
          "collection": "letterhead",
          "status": "in progress",
          "partials": { ... }
        }
      }
    }
  }
}

BAD IDEA:

builds: {
    core: {
        build_area: core
        build_types: {
            document_templates: {
                build_type: document_template
                committee_letterhead: {
                    build_id:
                    build_area: core
                    build_type: document_template
                    build_name: committee_letterhead
                    build_collection: letterheads
                    build_ext: dotm
                    build_status: active
                    build_path: builds/core/committee_letterhead
                    build_partials: {
                        partial_map: /config/partials_map.json"
                        styles: {
                            partial_type: style
                            styles_target: word/styles.xml
                            body: {
                                partial_id: ""
                                partial_location: <file path>
                                partial_style_group:body
                                partial_style_list: "Normal", "No Spacing", "Space Before"
                            }
                            callout: {}
                            caption: {}
                            character: {}
                            doc_default: {}
                            default: {}
                            foother: {}
                            header: {}
                            heading: {}
                            index: {}
                            toc: {}
                            list: {}
                            other: {}
                            toc: {}
                        numbering: {
                            numbering target: word/numbering.xml
                            numbering: {
                                partial_id: ""
                                partial_name: "committee_letterhead_numbering"
                                partial_type: numbering
                                partial_location: <file path>
                        macros: {
                            macros_target: ""
                            }
                                }
                            }
                        }
                    }
                }
        }
            }
            document_workspace: {
                normal: {
                    build_id: ""
                    build_name: core_normal
                    build_area: core
                    build_type: document_workspace
                    build_collection: normal
                    build_ext: dotm
                    build_status: active
                    build_path: builds/core/workspace/normal
                    build_partials: {
                        partial_map: ""
                }
                sheet : {}
                book: {}
                fonts: {}
            }
    }
}
