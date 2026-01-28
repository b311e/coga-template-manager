# Updates and Future Tasks

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
- [x] created partials/styles shell folder and front_matter.xml

### To do:

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
        - core
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
