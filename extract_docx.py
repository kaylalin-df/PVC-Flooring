import zipfile
import xml.etree.ElementTree as ET
import sys

docx_path = r"c:\Users\kayla.lin\ai project\PVC Flooring\廠商提供-20251001合作意向書-應昌 (1).docx"

try:
    with zipfile.ZipFile(docx_path, 'r') as z:
        with z.open('word/document.xml') as f:
            tree = ET.parse(f)
            root = tree.getroot()

            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

            paragraphs = []
            for para in root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'):
                texts = []
                for node in para.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
                    if node.text:
                        texts.append(node.text)
                paragraphs.append(''.join(texts))

            print('\n'.join(paragraphs))
except Exception as e:
    print(f"Error: {e}", file=sys.stderr)
    sys.exit(1)
