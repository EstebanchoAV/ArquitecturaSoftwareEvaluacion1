import docx
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

doc = docx.Document('To-be_HaciendaSoft.docx')

for element in doc.element.body:
    tag = element.tag.split('}')[-1]
    if tag == 'p':
        # Get paragraph text - simple way
        text = element.text if hasattr(element, 'text') else ''
        # Actually, let's get all text from runs
        from docx.oxml.ns import qn
        runs = element.findall('.//'+qn('w:t'))
        text = ''.join(r.text for r in runs if r.text)
        if text.strip():
            print(text.strip())
    elif tag == 'tbl':
        from docx.table import Table
        tbl = Table(element, doc)
        for row in tbl.rows:
            cells = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
            print(' | '.join(cells))
        print()
