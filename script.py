from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches

# Load template
doc = DocxTemplate("template.docx")

# List of people
people = [
    {"name": "Alice", "image": "images/1.jpg"},
    {"name": "Bob", "image": "images/2.jpg"},
    {"name": "Charlie", "image": "images/3.jpg"},
]

# Convert images to InlineImage objects
for person in people:
    person["image"] = InlineImage(doc, person["image"], width=Inches(2))

# Render template
doc.render({"people": people})

# Save result
doc.save("output.docx")