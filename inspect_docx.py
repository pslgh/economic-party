import zipfile
import sys
import os

def inspect_docx(docx_path):
    if not os.path.exists(docx_path):
        print(f"Error: File not found at {docx_path}")
        return

    try:
        with zipfile.ZipFile(docx_path) as z:
            print("Files in docx archive:")
            for info in z.infolist():
                print(f"{info.filename} ({info.file_size} bytes)")
            
            # Extract document.xml
            with open('debug_doc.xml', 'wb') as f:
                f.write(z.read('word/document.xml'))
            print("\nExtracted word/document.xml to debug_doc.xml")
            
            # Check for media
            media = [n for n in z.namelist() if n.startswith('word/media/')]
            if media:
                print(f"\nFound {len(media)} media files.")
                # Extract the first image to see
                first_image = media[0]
                ext = os.path.splitext(first_image)[1]
                target = f"extracted_image{ext}"
                with open(target, 'wb') as f:
                    f.write(z.read(first_image))
                print(f"Extracted first image to {target}")

    except Exception as e:
        print(f"Error inspecting docx: {str(e)}")

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python inspect_docx.py <path_to_docx>")
    else:
        inspect_docx(sys.argv[1])
