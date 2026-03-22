import zipfile
import xml.etree.ElementTree as ET
import sys

def read_docx(path):
    document_path = 'word/document.xml'
    word_namespace = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    para = word_namespace + 'p'
    text = word_namespace + 't'

    try:
        docx = zipfile.ZipFile(path)
        xml_content = docx.read(document_path)
        tree = ET.XML(xml_content)
        
        paragraphs = []
        for paragraph in tree.iter(para):
            texts = [node.text
                     for node in paragraph.iter(text)
                     if node.text]
            if texts:
                paragraphs.append(''.join(texts))
                
        return '\n\n'.join(paragraphs)
    except Exception as e:
        return f"Error: {e}"

if __name__ == '__main__':
    print(read_docx(sys.argv[1]))
