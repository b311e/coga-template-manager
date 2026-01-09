# Updates and Future Tasks



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
      - 
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
