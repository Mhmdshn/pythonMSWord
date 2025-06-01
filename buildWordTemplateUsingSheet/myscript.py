import os
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
from datetime import datetime

# Determine the script's directory
script_dir = os.path.dirname(os.path.abspath(__file__))

# Path to the images directory
image_dir = os.path.join(script_dir, 'images')

# Load the Word template
template_path = os.path.join(script_dir, 'template.docx')
doc = DocxTemplate(template_path)

# URL providing JSON data
url = "https://script.google.com/macros/s/AKfycbz-pdvGbjiIC_xs_NlcdrzSNaw9a6yhLeB3OilrWo7cJdhmLnYefEjO5jqJPs0Y3iIe/exec"

# Fetch data from the URL
import requests
response = requests.get(url)
response.raise_for_status()
orders = response.json()

# Process each order
for idx, order in enumerate(orders):
    for key, value in order.items():
        if isinstance(value, str) and value.lower().endswith(('.png', '.jpg', '.jpeg')):
            # Construct the full path to the image file
            image_path = os.path.join(image_dir, os.path.basename(value))
            if os.path.exists(image_path):
                # Create InlineImage
                order[key] = InlineImage(doc, image_path, width=Inches(2))
            else:
                print(f"Image file not found: {image_path}")
                order[key] = None  # Handle missing images gracefully

# Render the template with the orders data
doc.render({"orders": orders})

# Generate timestamped filename for the output document
timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
output_filename = os.path.join(script_dir, f"output_{timestamp}.docx")

# Save the generated document
doc.save(output_filename)
