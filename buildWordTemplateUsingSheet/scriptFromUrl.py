import requests
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
from datetime import datetime

# URL providing JSON data
url = "https://script.google.com/macros/s/AKfycbz-pdvGbjiIC_xs_NlcdrzSNaw9a6yhLeB3OilrWo7cJdhmLnYefEjO5jqJPs0Y3iIe/exec"

# Fetch data from the URL
response = requests.get(url)
response.raise_for_status()
orders = response.json()

# Load template
doc = DocxTemplate("template.docx")

# Convert images to InlineImage objects
for order in orders:
    image_path = order.get(["live"])
    if image_path:
        order["live"] = InlineImage(doc, image_path, width=Inches(2))

for order in orders:
    image_path = order.get(["b"])
    if image_path:
        order["b"] = InlineImage(doc, image_path, width=Inches(2))

# Render template
doc.render({"orders": orders})

# Generate timestamped filename
timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
output_filename = f"output_{timestamp}.docx"

# Save result
doc.save(output_filename)
