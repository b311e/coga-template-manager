#!/usr/bin/env python3
"""
Remove linked character styles from Word styles.xml file.
Removes HeadingXChar (1-9), QuoteChar, SubtitleChar, HeaderChar, FooterChar styles
and their link references from the parent paragraph styles.
"""

import sys
import re
from pathlib import Path

# Character styles to remove
CHAR_STYLES_TO_REMOVE = [
    'Heading1Char', 'Heading2Char', 'Heading3Char', 'Heading4Char', 'Heading5Char',
    'Heading6Char', 'Heading7Char', 'Heading8Char', 'Heading9Char',
    'QuoteChar', 'SubtitleChar', 'HeaderChar', 'FooterChar'
]

def remove_char_styles(styles_xml_path):
    """Remove character styles and their link references from styles.xml."""
    
    if not styles_xml_path.exists():
        print(f"Error: File not found: {styles_xml_path}")
        sys.exit(1)
    
    if not styles_xml_path.is_file():
        print(f"Error: Not a file: {styles_xml_path}")
        sys.exit(1)
    
    # Read the file
    with open(styles_xml_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    original_content = content
    styles_removed = 0
    links_removed = 0
    
    # Remove entire w:style elements for character styles
    for char_style in CHAR_STYLES_TO_REMOVE:
        # Simple pattern - just remove the style element
        style_pattern = rf'<w:style\s[^>]*w:styleId="{char_style}"[^>]*>.*?</w:style>'
        
        matches = re.findall(style_pattern, content, re.DOTALL)
        if matches:
            content = re.sub(style_pattern, '', content, flags=re.DOTALL)
            styles_removed += len(matches)
    
    # Remove all consecutive blank lines (multiple newlines with only whitespace between)
    content = re.sub(r'\n\s*\n+', '\n', content)
    
    # Remove w:link references to character styles from paragraph styles
    # Pattern matches entire line: optional indent + <w:link w:val="XxxChar" /> + newline
    # This preserves formatting by removing the entire line
    for char_style in CHAR_STYLES_TO_REMOVE:
        # Match: optional leading whitespace, the link element, optional trailing whitespace, newline
        link_pattern = rf'^[ \t]*<w:link w:val="{char_style}"\s*/>\s*\n'
        
        matches = re.findall(link_pattern, content, re.MULTILINE)
        if matches:
            content = re.sub(link_pattern, '', content, flags=re.MULTILINE)
            links_removed += len(matches)
    
    # Check if any changes were made
    if content == original_content:
        print("No character styles or link references found")
        return
    
    # Write back to file
    with open(styles_xml_path, 'w', encoding='utf-8') as f:
        f.write(content)
    
    print(f"✓ Removed {styles_removed} character style(s)")
    print(f"✓ Removed {links_removed} link reference(s)")
    print(f"Total changes: {styles_removed + links_removed}")

def main():
    if len(sys.argv) < 2:
        print("Usage: cleanup linkchar <styles.xml-path>")
        print("  styles.xml-path - Path to styles.xml file")
        print("")
        print("Removes linked character styles and their references from styles.xml:")
        print("  - Heading1Char through Heading9Char")
        print("  - QuoteChar, SubtitleChar, HeaderChar, FooterChar")
        sys.exit(1)
    
    styles_xml_path = Path(sys.argv[1])
    
    print(f"Removing character styles from: {styles_xml_path.name}")
    remove_char_styles(styles_xml_path)

if __name__ == "__main__":
    main()
