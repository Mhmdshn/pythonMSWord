import csv
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches

# Load template
doc = DocxTemplate("template.docx")

# Read orders from CSV file
orders = []
with open("orders.csv", newline='', encoding='utf-8') as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        if "image" not in row or not row["image"].strip():
            raise ValueError(f"Missing 'image' field in row: {row}")
        orders.append(row)

# Convert image file paths to InlineImage objects
for order in orders:
    order["image"] = InlineImage(doc, order["image"], width=Inches(2))

# Render template with orders
doc.render({"orders": orders})

# Save result
doc.save("output.docx")
