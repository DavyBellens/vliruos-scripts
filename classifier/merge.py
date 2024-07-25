import os
import PyPDF2
from pathlib import Path

cd = Path(r'C:\Users\DavyBellens\OneDrive - VLIR-UOS\Documenten\vliruos-scripts\classifier')
merged = cd / 'merged'
extracted = cd / 'extracted'

if not os.path.exists(merged):
    os.makedirs(merged)

levels = [x for x in os.listdir(extracted)]

for level in levels:
    institutions = [x for x in os.listdir(extracted / level) if x != 'failed.txt']

    for institution in institutions:
        if not os.path.exists(merged / institution):
            os.makedirs(merged / institution)

        file = PyPDF2.PdfWriter()
        files = os.listdir(extracted / level / institution)

        for f in files:
            pdf = PyPDF2.PdfReader(extracted / level / institution / f)

            for page in pdf.pages:
                file.add_page(page)

        file.write(f"{merged}/{institution}/{level}.pdf")