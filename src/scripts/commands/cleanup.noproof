#!/usr/bin/env python3
"""
Remove noProof elements from Word XML files.

The noProof element tells Word not to check spelling/grammar for specific text.
This script handles both minified and pretty-printed XML formats.

Removes two patterns:
1. Complete <w:rPr><w:noProof/></w:rPr> blocks (empty run properties containing only noProof)
2. Standalone <w:noProof/> within properties that have other content
"""

import sys
import re
from pathlib import Path

def clean_noproof(file_path):
    """Remove noProof elements from an XML file."""
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Pretty-printed XML only
    # Pattern 1: Complete <w:rPr> blocks containing ONLY <w:noProof/>
    # Remove the entire block including the lines
    pattern_rpr_block = re.compile(r'^[ \t]*<w:rPr>\n[ \t]*<w:noProof/>\n[ \t]*</w:rPr>\n', re.MULTILINE)
    
    # Pattern 2: Standalone <w:noProof/> lines (when rPr has other content)
    # Remove just the noProof line (with its indentation and newline)
    pattern_standalone = re.compile(r'^[ \t]*<w:noProof/>\n', re.MULTILINE)
    
    # Count and remove - order matters! Remove complete blocks first, then standalone
    rpr_blocks = len(pattern_rpr_block.findall(content))
    content = pattern_rpr_block.sub('', content)
    
    standalone = len(pattern_standalone.findall(content))
    content = pattern_standalone.sub('', content)
    
    total_count = rpr_blocks + standalone
    
    if total_count == 0:
        return (0, 0)
    
    # Write back
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(content)
    
    return (rpr_blocks, standalone)

def main():
    if len(sys.argv) < 2:
        print("Usage: cleanup noproof <xml-file-path>")
        print("  xml-file-path - Path to Word XML file (document.xml, styles.xml, etc.)")
        print("")
        print("Removes noProof elements from Word XML files.")
        print("The noProof element disables spell/grammar checking for specific text.")
        print("")
        print("Removes two patterns:")
        print("  1. Complete <w:rPr><w:noProof/></w:rPr> blocks")
        print("  2. Standalone <w:noProof/> elements within properties")
        sys.exit(1)
    
    file_path = Path(sys.argv[1])
    
    if not file_path.exists():
        print(f"Error: File not found: {file_path}")
        sys.exit(1)
    
    if not file_path.is_file():
        print(f"Error: Not a file: {file_path}")
        sys.exit(1)
    
    print(f"Cleaning noProof elements from: {file_path.name}")
    
    rpr_count, standalone_count = clean_noproof(file_path)
    total_count = rpr_count + standalone_count
    
    if total_count > 0:
        print(f"✓ Removed {total_count} noProof element(s):")
        if rpr_count > 0:
            print(f"  - {rpr_count} complete <w:rPr><w:noProof/></w:rPr> block(s)")
        if standalone_count > 0:
            print(f"  - {standalone_count} standalone <w:noProof/> element(s)")
    else:
        print("No noProof elements found")

if __name__ == "__main__":
    main()
