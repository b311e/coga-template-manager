#!/usr/bin/env python3
"""
templx style import doc — copy word/styles.xml from one Word document to another.

Usage: style_import_doc.py <target-docx> <source-docx> [--dry-run] [--backup]
"""

import argparse
import os
import shutil
import sys
import zipfile
import tempfile


STYLES_PART = "word/styles.xml"


def normalize_path(path_str):
    """Normalize cross-platform paths (Windows, Git Bash, Unix)."""
    if not path_str:
        return path_str
    normalized = path_str.replace("\\", "/")
    if os.name == "nt":
        import re
        match = re.match(r"^/([a-zA-Z])(/.*)", normalized)
        if match:
            normalized = f"{match.group(1).upper()}:{match.group(2)}"
    return os.path.normpath(normalized)


def copy_styles(target_path, source_path, dry_run=False, backup=False):
    # Read styles.xml from source
    with zipfile.ZipFile(source_path, "r") as src_zip:
        if STYLES_PART not in src_zip.namelist():
            print(f"Error: source document has no {STYLES_PART}: {source_path}", file=sys.stderr)
            return 4
        styles_xml = src_zip.read(STYLES_PART)

    if dry_run:
        print(f"Dry run: would replace {STYLES_PART} in '{target_path}' from '{source_path}'")
        return 0

    if backup:
        backup_path = target_path + ".bak"
        shutil.copy2(target_path, backup_path)

    # Rewrite target zip, replacing styles.xml
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".tmp", dir=os.path.dirname(target_path))
    os.close(tmp_fd)
    try:
        with zipfile.ZipFile(target_path, "r") as tgt_zip, \
             zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as out_zip:
            for item in tgt_zip.infolist():
                if item.filename == STYLES_PART:
                    out_zip.writestr(item, styles_xml)
                else:
                    out_zip.writestr(item, tgt_zip.read(item.filename))

        shutil.move(tmp_path, target_path)
    except Exception:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise

    if backup:
        print(f"Wrote {target_path} (backup at {target_path}.bak)")
    else:
        print(f"Wrote {target_path}")
    return 0


def main():
    parser = argparse.ArgumentParser(prog="style_import_doc")
    parser.add_argument("target", help="target .docx/.dotx to modify")
    parser.add_argument("source", help="source .docx/.dotx to copy styles from")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--backup", action="store_true", help="create a .bak backup before writing")
    args = parser.parse_args()

    target = normalize_path(args.target)
    source = normalize_path(args.source)

    if not os.path.exists(source):
        print(f"Error: source document not found: {source}", file=sys.stderr)
        sys.exit(3)
    if not os.path.exists(target):
        print(f"Error: target document not found: {target}", file=sys.stderr)
        sys.exit(3)

    sys.exit(copy_styles(target, source, dry_run=args.dry_run, backup=args.backup))


if __name__ == "__main__":
    main()
