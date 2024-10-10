import os
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib import colors

def convert_lowercase_to_uppercase(input_string):
    result = ''
    for char in input_string:
        # Check if the character is lowercase
        if 'a' <= char <= 'z':
            # Convert lowercase to uppercase
            char = chr(ord(char) - 32)  # Convert ASCII value
        result += char
    return result

df = os.path.join(os.path.expanduser("~"),"Downloads","class_list.xlsx")
df_students = pd.read_excel(df)

df2 = os.path.join(os.path.expanduser("~"), "Downloads", "subject_list.xlsx")
df_course_info = pd.read_excel(df2)


# Step 1: Read the first Excel file with student data
def print_subject_list():
   
    # Step 2: Extract the subject code provided by the user
    subject_codes = input("Enter subject code: ")  # Replace this with the user input
    subject_code = convert_lowercase_to_uppercase(subject_codes)

    # Extract relevant columns
   # Filter students who have 'TRUE' for the provided subject code
   
    df_students[subject_code] = df_students[subject_code].astype(str)

    # Now apply the string operations
    df_students[subject_code] = df_students[subject_code].str.upper().str.strip()

    # Filter the DataFrame for rows where the subject code column is 'TRUE'
    df_filtereds = df_students[['FULL NAME', 'MATRIC NO', subject_code]][df_students[subject_code] == 'TRUE']
    df_filtered = df_filtereds.copy()  # Make a copy to avoid modifying the original DataFrame
    df_filtered.insert(0, 'Numbering', range(1, len(df_filtered) + 1))

# Display the modified DataFrame



    # Check if any students are enrolled in the subject
    if df_filtereds.empty:
        print(f"No students found for subject {subject_code}")
        return
    # Step 3: Read the second Excel sheet with the subject title information
    

    # Extract course title from cell B1 of the second sheet
    course_row = df_course_info[df_course_info['SUBJECT CODE'] == subject_code]
    if not course_row.empty:
        course_title = course_row.iloc[0, 1]  # Assuming title is in the second column
    else:
        course_title = "Title not found" 
    file_name = (f"{subject_code} class_list.pdf")
    downloads_path = os.path.join(os.path.expanduser('~'), 'Downloads')
    pdf_filename = os.path.join(downloads_path, file_name)
    # Step 4: Create the PDF

    c = canvas.Canvas(pdf_filename, pagesize=letter)

    # Create the heading with the updated course title
    heading = f"UNIVERSITY OF NIGERIA, NSUKKA\nDEPARTMENT: STATISTICS\nCOURSE: {subject_code} - {course_title}"
    c.setFont("Helvetica-Bold", 12)
    c.setFillColor(colors.black)

    # Draw the heading with line breaks
    heading_textobject = c.beginText(100, 750)
    heading_textobject.setFont("Helvetica-Bold", 12)
    heading_textobject.setFillColor(colors.black)
    heading_textobject.textLines(heading)
    c.drawText(heading_textobject)

    # Define colors for rows
    odd_color = colors.lightgrey
    even_color = colors.white

    font_size = 9

    # Add table headers with different colors
    header_y = 700
    c.setFillColor(colors.green)  # Set fill color for the top row
    c.setStrokeColor(colors.black)  # Set the border color
    c.setLineWidth(1)  # Set the border line width
    c.setFont("Helvetica-Bold", font_size)  # Set bold font for headers

    # Draw green rectangle for the top row
    c.rect(90, header_y - 15, 500, 20, fill=True)

    # Set text color to white for the top row
    c.setFillColor(colors.white)
    c.drawString(100, header_y - 10, "NO")
    c.drawString(150, header_y - 10, "Full Name")
    c.drawString(350, header_y - 10, "Matric No")
    c.drawString(500, header_y - 10, "Signature")

    # Add extracted data to the PDF
    y_coordinate = 680  # Start after header
    line_height = 20  # Height between each row
    row_count = 0

    for index, row in df_filtered.iterrows():
        numbering = str(row['Numbering'])
        full_name = row['FULL NAME']
        matric_no = row['MATRIC NO']
        
        offering_subject = row[subject_code]
        if offering_subject == 'TRUE':
            signature = ' '
        else:
            signature = ' '
        
        # Alternate row colors
        if row_count % 2 == 0:
            c.setFillColor(odd_color)
        else:
            c.setFillColor(even_color)
        
        # Draw rectangles for row background
        c.rect(90, y_coordinate - 15, 500, line_height, fill=True)
        
        # Place Full Name, Matric No, and Signature on separate rows with colors and adjusted font size
        c.setFillColor(colors.black)
        c.setFont("Helvetica", font_size)
        c.drawString(100, y_coordinate - 10, numbering)
        c.drawString(150, y_coordinate - 10, full_name)
        c.drawString(350, y_coordinate - 10, matric_no)
        c.drawString(500, y_coordinate - 10, signature)
        
        # Move to the next row
        y_coordinate -= line_height
        row_count += 1

    if os.path.exists(pdf_filename):
        os.remove(pdf_filename)
        print(f"Existing file {subject_code}_class_list.pdf deleted.")
    # Save and close the PDF
    c.save()
    print(f"Class list for subject {subject_code} created and saved to {pdf_filename}")

def list_subject_code():
    # Read the second Excel sheet
    df_course_info = pd.read_excel(df2)

    # Extract the 'Subject Code' column
    subject_codes = df_course_info['SUBJECT CODE']

    # Display or use the extracted subject codes
    print(subject_codes)

try:
    print_subject_list()
except Exception as e:
    print (f"invalid Subject Code {e}")
    print("these are the subject code:")
    list_subject_code()
