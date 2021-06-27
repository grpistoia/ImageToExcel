###############################################################################
# Gustavo Pistoia - First Python program .. see what the fuss is about..
# Load an image and generate an Excel file with cells coloured as the image
###############################################################################

# The library to write Excel files
import xlsxwriter
	
# Pillow library for images processing
from PIL import Image

###############################################################################
# GLOBALS YOU CAN CHANGE
###############################################################################

# Source image file
IMAGE_FILE_NAME  = 'input_image.jpg'
# Target Excel file
EXCEL_FILE_NAME  = 'output_file.xls'
# Width of the image to generate
MAX_IMAGE_WIDTH  = 1024.0
# Height of the image to generate
MAX_IMAGE_HEIGHT = 768.0
# Max number of colors to use
MAX_USED_COLORS  = 255

###############################################################################
# Excel limits .. but XLSX may vary
###############################################################################

# max number of rows and columns in XLS
LIMIT_XLS_ROWCOL = 65535 
# max number of colors in XLS
LIMIT_XLS_COLOR  = 255
# This is how big the square is.. the biggest the less I can zoom
XLS_CELL_SIZE = max(min(1024, 768) / max(MAX_IMAGE_WIDTH, MAX_IMAGE_HEIGHT), 5)

###############################################################################
# Conversion functions
###############################################################################

# Convert an array pixel color [R][B][G] to string Hex format
def pixel_rgb_to_hex(myPixel):
    r, g, b = myPixel
    return '#{:02X}{:02X}{:02X}'.format(r, g, b)
    
# Produce a ratio for witdh or height to be applied later as float
def calculate_ratio(user_input, image_value, file_limit):
    return min(user_input, image_value, file_limit) / float(image_value)

###############################################################################
# Open the image and applies some restrictions
###############################################################################

print("Opening file '" + IMAGE_FILE_NAME + "'..")

# Load the image
input_image = Image.open(IMAGE_FILE_NAME)

# This is the original size
image_width, image_height = input_image.size

print("Adjusting colors..")

# reduce colors to fix into Excel's limitation
input_image = input_image.convert('P', 
                                  palette = Image.ADAPTIVE,
                                  colors = min(MAX_USED_COLORS, 
                                               LIMIT_XLS_COLOR))

# Convert to RGB .. because transparent is not needed
input_image = input_image.convert('RGB') # must be after color reduction

print("Original image size is " + str((image_width, image_height))+ "..")

# Calculate each ratio and see which is used as the driving one
ratio_x = calculate_ratio(MAX_IMAGE_WIDTH,  image_width,  LIMIT_XLS_ROWCOL);
ratio_y = calculate_ratio(MAX_IMAGE_HEIGHT, image_height, LIMIT_XLS_ROWCOL);
ratio_f = ratio_x if ratio_x < ratio_y else ratio_y;

# Apply the final ratio to the image width and height
image_width  = int(image_width  * ratio_f)
image_height = int(image_height * ratio_f)

print("Adjusting size to " + str((image_width, image_height)) + "..")

# Now resize the image as expected
input_image = input_image.resize( (image_width, image_height) )

# Turn image into an array of pixels of RBG (access is slow)
image_pixels = input_image.load()

###############################################################################
# Create the excel file
###############################################################################

print("Creating file..")

# Create the output file
workbook = xlsxwriter.Workbook(EXCEL_FILE_NAME)
# With a worksheet called as the image...
worksheet = workbook.add_worksheet(IMAGE_FILE_NAME)

###############################################################################
# Create the formats based on the conversion dictionary
# Notice is destructive, and replaces the value with the new formatter
###############################################################################

print("Creating color styles..")

# Create a dictionary key = pixel color, value = format
color_to_style = {  }

# I tested a cache but makes no difference
# Notice 2d array creation: [[0]*width]*height... is buggy..
#cached_pixels = [[0 for c in range(image_width)] for r in range(image_height)]

# Go through rows first..
for row in range(0, image_height):
    # Then columns.. although PIL does it column-row
    for col in range(0, image_width):        
        # Get the pixel array color (this is slow)
        pixel_color = image_pixels[col, row] # notice (column & row) API
        # Without effort the dictionary remove duplicates
        color_to_style[pixel_color] = '?' # 'value' not relevant yet

# Now that the dictionary contains only few elements, apply the format
for pixel_color in color_to_style.keys():
    # key is the pixel color (i.e: 255,0,128) .. which I turn to hex..
    hex_color = pixel_rgb_to_hex(pixel_color)
    # now store format in the value in dictionary
    color_to_style[pixel_color] = workbook.add_format({'bg_color': hex_color})

###############################################################################
# Finally it can now write the content of the file
###############################################################################

print("Saving content..")

# Set the width of a range of columns
worksheet.set_column_pixels(0, image_width, XLS_CELL_SIZE)

# Again iterate by row first..
for row in range(0, image_height):
    # Set the row height (set_row_pixel does not work as per column)
    worksheet.set_row(row, XLS_CELL_SIZE)
    # Now iterate on columns..
    for col in range(0, image_width):
        # get the pixel
        pixel_color = image_pixels[col, row] # notice (column & row) API
        # get the associated format 
        loaded_format = color_to_style[pixel_color]
        # write the cell..        
        worksheet.write(row, col, '', loaded_format) # notice row & column

# ready to finish now..
workbook.close()

print("File '" + workbook.filename + "' saved.")

###############################################################################
