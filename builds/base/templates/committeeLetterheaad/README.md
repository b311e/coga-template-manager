# Committee Letterhead 2.0

## Rules

The config folder is where the json file containing the committee member names, etc. lives. 
- `_config.json` is the blank example file. `config.json` contains the data the template will use to auto-fill values.
- Do not change the name or location of `config.json`.

## Tasks / Future Updates

- [x] Create json file with test data (2026-01-08)
- [x] Create VBA macro that auto updates fields when you open the template  (2026-01-08)
- [x] Replace all the placeholder text with fields (still need to do the right side member names)
- [ ] Replace the core styles with those from the repo
- [ ] Use the file in the src folder to create the final, pulling in the core styles (bullets, etc.) but preserving the styling for Heading 1, 2, 3, etc. 
- [ ] Replace the text in the content controls with actual placeholder text (see the example in the body of the doc for code, the text read "test placeholder text")
- [ ] Reorder the styles that must stay in this template, including Colorado General Assembly, Committee Member, Committee Name, Address, Email, Phone, Date, Heading 2, Heading 3, Heading 4.
- [ ] Clean up VBA script so it isn't embarrassing to show the devs
- [ ] Fix the path for the config, because i renamed it config_fields.json

Can we connect this to the CLICS Endpoint API so the file values auto-update? Might be too complicated because the order of the names can change depending on the template.


See UPDATES file