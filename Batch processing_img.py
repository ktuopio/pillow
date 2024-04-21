from PIL import Image, ImageDraw, ImageFont
import openpyxl
import os

# Paths to files and folders
base_image_path = "path/to/base/image.png"
photo_folder = "path/to/photo/folder"
excel_file = "path/to/excel/file.xlsx"
output_folder = "path/to/output/folder"

# Create output folder if it doesn't exist
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Text parameters
font_regular = ImageFont.truetype("path/to/regular/font.ttf", 52)
font_bold1 = ImageFont.truetype("path/to/bold/font.ttf", 73)
font_bold2 = ImageFont.truetype("path/to/bold/font.ttf", 52)

# Load base image
base_image = Image.open(base_image_path)

# Load Excel file
workbook = openpyxl.load_workbook(excel_file)
sheet = workbook.active

# Function to add text to image
def add_text(draw, text, font, color, x, y):
    draw.text((x, y), text, fill=color, font=font)

# Read data from Excel and add text and photo to image
for row in sheet.iter_rows(min_row=2, values_only=True):
    # Create a copy of the base image
    image_with_text = base_image.copy()
    draw = ImageDraw.Draw(image_with_text)

    file_name = row[0]
    text1, text2, text3 = row[1], row[2], row[3]

    photo_path = os.path.join(photo_folder, file_name)
    if os.path.exists(photo_path):
        # Load photo
        photo = Image.open(photo_path)
        # Resize photo to 600x600
        photo.thumbnail((600, 600))
        # Paste photo onto the image
        image_with_text.paste(photo, (500, 500))

        # Add text
        add_text(draw, text1, font_regular, "#191717", 40, 205)
        add_text(draw, text2, font_bold1, "#d42e12", 40, 375)
        add_text(draw, text3, font_bold2, "#ffffff", 45, 490)

        # Save image with text and photo
        output_path = os.path.join(output_folder, f"{os.path.splitext(file_name)[0]}_with_text.png")
        image_with_text.save(output_path)
        print(f"Image with text and photo saved: {output_path}")

print("Done!")
