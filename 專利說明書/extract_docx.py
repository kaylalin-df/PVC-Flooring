#!/usr/bin/env python3
"""
Extract text from a .docx file using zipfile + xml.etree.ElementTree.
No external dependencies required - uses only Python standard library.
"""

import zipfile
import xml.etree.ElementTree as ET
import os
import sys

def extract_text_from_docx(docx_path):
    """Extract all paragraph text from a .docx file."""
    # Word namespace
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

    paragraphs = []

    with zipfile.ZipFile(docx_path, 'r') as z:
        # Read the main document XML
        with z.open('word/document.xml') as f:
            tree = ET.parse(f)
            root = tree.getroot()

            # Find all paragraphs
            for para in root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'):
                texts = []
                for run in para.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
                    if run.text:
                        texts.append(run.text)
                if texts:
                    paragraphs.append(''.join(texts))
                else:
                    paragraphs.append('')  # Empty paragraph (blank line)

    return '\n'.join(paragraphs)


if __name__ == '__main__':
    script_dir = os.path.dirname(os.path.abspath(__file__))

    docx_file = os.path.join(script_dir, 'CP38568-回收晶舟盒製成的生態箱-說明書(新型)_v0_20250122.docx')
    output_file = os.path.join(script_dir, 'extracted_patent.txt')

    if not os.path.exists(docx_file):
        print(f"Error: File not found: {docx_file}", file=sys.stderr)
        sys.exit(1)

    text = extract_text_from_docx(docx_file)

    # Print to stdout
    print(text)

    # Save to txt file
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(text)

    print(f"\n--- Saved to: {output_file} ---", file=sys.stderr)
