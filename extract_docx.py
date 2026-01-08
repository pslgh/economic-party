import zipfile
import xml.etree.ElementTree as ET
import sys
import os

def get_text(docx_path):
    if not os.path.exists(docx_path):
        return f"Error: File not found at {docx_path}"
        
    try:
        with zipfile.ZipFile(docx_path) as document:
            xml_content = document.read('word/document.xml')
        
        tree = ET.fromstring(xml_content)
        text = []
        
        # Determine namespace
        # Usually it's http://schemas.openxmlformats.org/wordprocessingml/2006/main
        # accessible via the 'w' prefix in tag names if we parse intelligently, 
        # but simpler to just look for 'p' and 't' tags ending with the right suffix
        
        for element in tree.iter():
            if element.tag.endswith('}p'):
                text.append('\n')
            elif element.tag.endswith('}t'):
                if element.text:
                    text.append(element.text)
            elif element.tag.endswith('}tab'):
                text.append('\t')
                
        return ''.join(text)
    except Exception as e:
        return f"Error reading docx: {str(e)}"

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python extract_docx.py <path_to_docx>")
    else:
        print(get_text(sys.argv[1]))
