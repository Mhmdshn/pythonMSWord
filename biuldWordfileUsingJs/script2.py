import pandas as pd
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches

# Load Excel sheet into DataFrame
df = pd.read_excel("people.xlsx")

# Load Word template
doc = DocxTemplate("template.docx")

# Convert rows into a list of dictionaries with InlineImage
people = []
for _, row in df.iterrows():
    person = {
        "name": row["name"],
        "image": InlineImage(doc, row["image"], width=Inches(2))
    }
    people.append(person)

# Render the template
doc.render({"people": people})
doc.save("output.docx")