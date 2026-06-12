from docxtpl import DocxTemplate
import pandas as pd

doc = DocxTemplate("letter_template.docx")
df = pd.read_excel("client_data.xlsx")

records = df.to_dict(orient="records")

for record in records:
    doc.render(record)
    file_name = f"Correspondence_{record['client_name']}.docx"
    doc.save(file_name)
