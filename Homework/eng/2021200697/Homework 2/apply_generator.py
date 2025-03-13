from docxtpl import DocxTemplate
from docx2pdf import convert
import pandas as pd
import os

df = pd.read_excel("programs.xlsx")
template = DocxTemplate("statement_template.docx")
output_dir = "Applications"

for idx, row in df.iterrows():
    university = row["University Names"]
    for col in ["Major1", "Major2", "Major3"]:
        program = row[col]

        context = {
            "university_name": university,
            "program_name": program
        }

        filename = f"{university.replace(' ', '_').replace('-', '_')}_{program.replace(' ', '_')}.docx"
        output_path = os.path.join(output_dir, filename)

        template.render(context)
        template.save(output_path)
        convert("Applications")
