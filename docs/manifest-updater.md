# Manifest Updaters

A concise reference for the two manifest updaters used in this project: the agency manifest helper and the partials manifest updater. Includes command examples and safety notes.

## Purpose

- Agency manifest helper: manage per-agency manifest files in `builds/{agency}/manifest.json`.
- Partials manifest updater: generate `foundation/partials/partials-manifest.json` from XML snippet files in `foundation/partials/`.

## Scripts & locations

- `src/scripts/manifest` — agency manifest helper (bash)
- `src/scripts/update_partials_manifest.py` — partials manifest updater (python)
- `src/scripts/update` — dispatcher: `update manifest partials` (also proxies other `manifest` commands)

## Commands (examples)

Agency manifest (run from project root):
```bash
# Generate/update agency manifest timestamp and scan for templates
./src/scripts/manifest generate jbc

# Validate manifests and registry
./src/scripts/manifest validate

# List templates across agencies
./src/scripts/manifest list

# Update a template status
./src/scripts/manifest update-status jbc jbcNormal active
```

Partials manifest updater:
```bash
# Dry-run (prints JSON to stdout)
python3 src/scripts/update_partials_manifest.py --dry-run

# Update manifest file (creates backup)
python3 src/scripts/update_partials_manifest.py
```

Dispatcher (short form):
```bash
# preferred: uses dispatcher to run the partials updater
./src/scripts/update manifest partials


## Common commands (quick reference)

Agency manifest
```bash
# Generate/update agency manifest timestamp and scan for templates
manifest generate jbc    # or ./src/scripts/manifest generate jbc

# Validate manifests and registry
manifest validate         # or ./src/scripts/manifest validate

# List templates across agencies
manifest list             # or ./src/scripts/manifest list

# Update a template status
manifest update-status jbc jbcNormal active    # or ./src/scripts/manifest update-status jbc jbcNormal active
```

Partials manifest updater
```bash
# Dry-run (prints JSON to stdout)
update manifest partials --dry-run   # or python3 src/scripts/update_partials_manifest.py --dry-run

# Update manifest file (creates backup)
update manifest partials             # or python3 src/scripts/update_partials_manifest.py

# proxy to agency manifest commands via dispatcher
./src/scripts/update manifest generate jbc
```

If you add `src/scripts` to your PATH (see next section) you can call `update manifest partials` directly.

## Safety & backups

- Before writing, the partials updater saves a backup: `foundation/partials/partials-manifest.json.bak`.
- Use `--dry-run` to preview changes without writing.
- To restore the previous manifest:
```bash
mv foundation/partials/partials-manifest.json.bak foundation/partials/partials-manifest.json
```

## Integration (PATH / aliases)

To run the short commands without `./src/scripts/` prefix, add `src/scripts` to your PATH. You can source the provided helper once per session:
```bash
source src/scripts/setup_aliases.sh
# or add 'export PATH="$PWD/src/scripts:$PATH"' to your ~/.bashrc
```

## Troubleshooting

- No entry for a snippet: ensure the file is valid XML and ends with `.xml`.
- Duplicate ids: rename the snippet `id` or filename to avoid collisions.
