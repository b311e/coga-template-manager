#!/usr/bin/env python3
"""
OOXML Validator — validates Office Open XML documents against ECMA-376 XSD schemas.
Usage: validate_ooxml.py <file-path>
"""

import sys
import zipfile
import re
from pathlib import Path
from datetime import datetime

try:
    from lxml import etree
except ImportError:
    print("Error: lxml is required. Install with: pip install lxml", file=sys.stderr)
    sys.exit(1)

# Paths relative to this script: src/scripts/internal/ -> project root is 3 levels up
_SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = _SCRIPT_DIR.parent.parent.parent
SCHEMA_DIR    = PROJECT_ROOT / "resources" / "schemas"
OOXML_DIR     = SCHEMA_DIR / "ooxml"
TRANSITIONAL  = OOXML_DIR / "transitional"
OPC_DIR       = OOXML_DIR / "opc"
DC_DIR        = SCHEMA_DIR / "dublin-core"

# Parts to validate and which schema applies to each.
# Patterns are matched in order; first match wins.
PART_SCHEMA_MAP = [
    (re.compile(r"^\[Content_Types\]\.xml$"),           OPC_DIR / "opc-contentTypes.xsd"),
    (re.compile(r"^(_rels|.*/_rels)/.*\.rels$"),        OPC_DIR / "opc-relationships.xsd"),
    (re.compile(r"^docProps/core\.xml$"),               OPC_DIR / "opc-coreProperties.xsd"),
    (re.compile(r"^word/document\.xml$"),               TRANSITIONAL / "wml.xsd"),
    (re.compile(r"^xl/workbook\.xml$"),                 TRANSITIONAL / "sml.xsd"),
    (re.compile(r"^ppt/presentation\.xml$"),            TRANSITIONAL / "pml.xsd"),
]

# Maps external schema URLs to local filenames within the same directory.
# Used by LocalSchemaResolver to avoid network access at validation time.
_EXTERNAL_SCHEMA_MAP = {
    "http://dublincore.org/schemas/xmls/qdc/2003/04/02/dc.xsd":       DC_DIR / "dc.xsd",
    "http://dublincore.org/schemas/xmls/qdc/2003/04/02/dcterms.xsd":  DC_DIR / "dcterms.xsd",
    "http://dublincore.org/schemas/xmls/qdc/2003/04/02/dcmitype.xsd": DC_DIR / "dcmitype.xsd",
    "http://www.w3.org/2001/03/xml.xsd":                              DC_DIR / "xml.xsd",
}


class LocalSchemaResolver(etree.Resolver):
    """Redirects external schema URL imports to local files."""

    def resolve(self, url, id, context):
        local = _EXTERNAL_SCHEMA_MAP.get(url)
        if local and local.exists():
            return self.resolve_filename(str(local), context)
        return None


def load_schema(xsd_path: Path) -> "etree.XMLSchema | None":
    """Load an XSD schema, routing any external imports to local copies."""
    parser = etree.XMLParser()
    parser.resolvers.add(LocalSchemaResolver())
    try:
        schema_doc = etree.parse(str(xsd_path), parser)
        return etree.XMLSchema(schema_doc)
    except etree.XMLSchemaParseError as e:
        return None


def validate_file(file_path: Path) -> tuple:
    """
    Validate each known part of an OOXML file against its schema.
    Returns (error_count, message_lines).
    """
    messages = []
    schema_cache: dict = {}

    with zipfile.ZipFile(file_path) as zf:
        for name in zf.namelist():
            xsd_path = None
            for pattern, xsd in PART_SCHEMA_MAP:
                if pattern.match(name):
                    xsd_path = xsd
                    break

            if xsd_path is None:
                continue

            if xsd_path not in schema_cache:
                schema_cache[xsd_path] = load_schema(xsd_path)

            schema = schema_cache[xsd_path]
            if schema is None:
                messages.append(f"  WARN  Could not load schema for part: {name}")
                continue

            try:
                doc = etree.fromstring(zf.read(name))
                schema.assertValid(doc)
            except etree.DocumentInvalid as e:
                for err in e.error_log:
                    messages.append(f"  ERROR [{name}] line {err.line}: {err.message}")
            except etree.XMLSyntaxError as e:
                messages.append(f"  ERROR [{name}] XML syntax error: {e}")

    error_count = sum(1 for m in messages if m.startswith("  ERROR"))
    return error_count, messages


def main():
    if len(sys.argv) < 2:
        print("Usage: validate_ooxml.py <file-path>", file=sys.stderr)
        sys.exit(1)

    file_path = Path(sys.argv[1]).resolve()

    if not file_path.exists():
        print(f"Error: File not found: {file_path}", file=sys.stderr)
        sys.exit(1)

    report_dir = file_path.parent / "reports"
    report_dir.mkdir(exist_ok=True)
    report_path = report_dir / f"{file_path.stem}-report.txt"

    print(f"Validating: {file_path.name}")
    error_count, messages = validate_file(file_path)

    with open(report_path, "w") as f:
        f.write("OOXML Validation Report\n")
        f.write("=======================\n")
        f.write(f"File:      {file_path.name}\n")
        f.write(f"Full Path: {file_path}\n")
        f.write(f"Timestamp: {datetime.now():%Y-%m-%d %H:%M:%S}\n")
        f.write(f"File Size: {file_path.stat().st_size:,} bytes\n\n")

        if messages:
            f.write("\n".join(messages) + "\n\n")
        else:
            f.write("No validation errors found.\n\n")

        f.write("=======================\n")
        f.write(f"Validation {'PASSED' if error_count == 0 else 'FAILED'}\n")
        f.write(f"Total Errors: {error_count}\n")

    result = "PASSED" if error_count == 0 else "FAILED"
    print(f"Validation {result} — {error_count} error(s) found")
    print(f"Report saved to: {report_path}")

    sys.exit(0 if error_count == 0 else 1)


if __name__ == "__main__":
    main()
