# Manipulating images
This script automates the process of adding text and photos to images based on data stored in an Excel spreadsheet. Here's a breakdown of what it does:

* Load the base image and Excel data.
* Iterate through each row of the Excel data.
* For each row:
  * Load the corresponding photo.
  * Resize and paste the photo onto the base image.
  * Add specified text to the image.
  * Save the modified image.
* Print a completion message when all images are processed.
