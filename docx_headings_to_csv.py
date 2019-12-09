from pathlib import Path
from docx import Document
import pandas as pd

def main():
    p = r'test_data/test.docx'
    PATH = Path(p)
    get_headings_to_csv(PATH, csv_name="headings.csv")

def get_heading_tuple(paragraph):
    if paragraph.style.name.startswith('Heading'):
        return (paragraph.style.name, paragraph.text)

def _get_headings(doc):
    headings = []
    for para in doc.paragraphs:
        heading = get_heading_tuple(para)
        if heading is not None:
            headings.append(heading)
    return headings

def get_headings(doc):
    "Return dict of (heading_num, heading_text). NOTE: Support 3 levels only"
    headings = _get_headings(doc)
    h1_c, h2_c, h3_c = 0, 0, 0
    numbered_headings = []
    for heading in headings:
        if heading[0] == 'Heading 1':
            h1_c += 1
            h2_c = 0
            h3_c = 0
            numbered_headings.append((f"{h1_c}", heading[1]))
        if heading[0] == 'Heading 2':
            if h1_c < 1: h1_c = 1
            h2_c += 1
            h3_c = 0
            numbered_headings.append((f"{h1_c}.{h2_c}", heading[1]))
        if heading[0] == 'Heading 3':
            if h2_c < 1: h2_c = 1
            h3_c += 1
            numbered_headings.append((f"{h1_c}.{h2_c}.{h3_c}", heading[1]))
    return dict(numbered_headings)
        
def get_headings_to_csv(docx_p, csv_name="headings.csv"):
    "Parse docx file and save headings to csv. NOTE: Support 3 levels only"
    doc = Document(docx_p)
    headings = get_headings(doc)
    headings = pd.DataFrame.from_dict(headings, orient='index')
    headings.to_csv(csv_name, header=False)


if __name__ == "__main__":
    main()
