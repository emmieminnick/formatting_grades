# Team 10 Section 003
# Make a program that will automatically format and summarize the important 
# information about each of the classes they teach from the imported Excel files.

# Import libraries
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font

# Load the Excel workbook
myWorkbook = openpyxl.load_workbook('Poorly_Organized_Data_1.xlsx')

# Get the active sheet
currSheet = myWorkbook.active

# Create a new workbook for the formatted data
formattedWorkbook = Workbook()

# Remove the default sheet created in the workbook
formattedWorkbook.remove(formattedWorkbook['Sheet'])

# Identify different classes and create worksheets for classes
class_names = set()

# Find unique class names
for row in currSheet.iter_rows(min_row = 2, values_only=True) :
    # Class name is in first column
    class_name = row[0]
    class_names.add(class_name)

# Create a worksheet for each unique class
for class_name in class_names :
    formattedWorkbook.create_sheet(title=class_name)

# Save the workbook and close it
formattedWorkbook.save(filename = 'formatted_grades.xlsx')
formattedWorkbook.close()