#!/usr/bin/env python3
"""
Remove tracking metadata attributes from Word XML files.

These attributes are used by Word for change tracking, versioning, and collaborative
editing but are not needed for clean template files.

Tracking attributes removed:
- w:rsidR - Revision ID for when the element was inserted/created
- w:rsidRPr - Revision ID for when run properties were last modified
- w:rsidRDefault - Revision ID for default/base content of a paragraph
- w:rsidP - Revision ID for when paragraph properties were last modified
- w:rsidSect - Revision ID for when section properties were last modified
- w:rsidTr - Revision ID for when a table row was inserted/created
- w:rsidDel - Revision ID for when content was deleted
- w14:paraId - Paragraph identifier for collaborative editing tracking
- w14:textId - Text content identifier for collaborative editing tracking
"""

import sys
import re
from pathlib import Path

def clean_tracking_attributes(file_path):
    """Remove all tracking metadata attributes from an XML file."""
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Pattern 1: RSID attributes - space + w:rsid + any letters + ="8 hex chars"
    # Handles: rsidR, rsidRPr, rsidRDefault, rsidP, rsidSect, rsidTr, rsidDel, etc.
    rsid_pattern = r' w:rsid[A-Za-z]*="[0-9A-Fa-f]{8}"'
    
    # Pattern 2: Paragraph ID - space + w14:paraId + ="8 hex chars"
    paraid_pattern = r' w14:paraId="[0-9A-Fa-f]{8}"'
    
    # Pattern 3: Text ID - space + w14:textId + ="8 hex chars"
    textid_pattern = r' w14:textId="[0-9A-Fa-f]{8}"'
    
    # Count occurrences of each type
    rsid_count = len(re.findall(rsid_pattern, content))
    paraid_count = len(re.findall(paraid_pattern, content))
    textid_count = len(re.findall(textid_pattern, content))
    total_count = rsid_count + paraid_count + textid_count
    
    if total_count == 0:
        print(f"No tracking attributes found in {file_path.name}")
        return 0
    
    # Remove all matches
    cleaned_content = content
    cleaned_content = re.sub(rsid_pattern, '', cleaned_content)
    cleaned_content = re.sub(paraid_pattern, '', cleaned_content)
    cleaned_content = re.sub(textid_pattern, '', cleaned_content)
    
    # Write back
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(cleaned_content)
    
    # Return breakdown for reporting
    return (rsid_count, paraid_count, textid_count)

def main():
    if len(sys.argv) < 2:
        print("Usage: cleanup tracking <xml-file-path>")
        print("  xml-file-path - Path to Word XML file (document.xml, styles.xml, etc.)")
        print("")
        print("Removes tracking metadata attributes from Word XML files.")
        print("These attributes are used by Word for change tracking and versioning")
        print("but aren't needed in clean template files.")
        print("")
        print("Removes:")
        print("  - All RSID types: rsidR, rsidRPr, rsidRDefault, rsidP, rsidSect, rsidTr, rsidDel")
        print("  - Paragraph IDs: w14:paraId")
        print("  - Text IDs: w14:textId")
        sys.exit(1)
    
    file_path = Path(sys.argv[1])
    
    if not file_path.exists():
        print(f"Error: File not found: {file_path}")
        sys.exit(1)
    
    if not file_path.is_file():
        print(f"Error: Not a file: {file_path}")
        sys.exit(1)
    
    print(f"Cleaning tracking attributes from: {file_path.name}")
    
    rsid_count, paraid_count, textid_count = clean_tracking_attributes(file_path)
    total_count = rsid_count + paraid_count + textid_count
    
    if total_count > 0:
        print(f"✓ Removed {total_count} tracking attribute(s):")
        if rsid_count > 0:
            print(f"  - {rsid_count} RSID attribute(s)")
        if paraid_count > 0:
            print(f"  - {paraid_count} paraId attribute(s)")
        if textid_count > 0:
            print(f"  - {textid_count} textId attribute(s)")
    else:
        print("No tracking attributes found")

if __name__ == "__main__":
    main()
