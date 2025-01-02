import os
import pandas as pd
from docxtpl import DocxTemplate

# Define the folder where the reports will be saved
output_folder = "generated_reports"

# Check if the folder exists, and create it if not
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Load the template
template = DocxTemplate("reportsheet.docx")

# Read data from the Excel file, skipping the first 5 rows to account for the header starting at row 6
data = pd.read_excel("results.xlsx", header=5)

# Clean up column names by stripping any extra spaces (if needed)
data.columns = data.columns.str.strip()

# Check the column names to ensure correct matching (optional, for debugging)
print(data.columns)

# Iterate over each row in the Excel file
for index, row in data.iterrows():
    # Convert the row to a dictionary for rendering
    student_data = row.to_dict()

    # Convert numerical scores to integers (if they are floats)
    for key, value in student_data.items():
        if isinstance(value, float) and value.is_integer():  # Check if it's a float with .0
            student_data[key] = int(value)  # Convert to an integer

    # Use the 'name' column to generate the filename (assuming 'name' exists)
    student_name = student_data.get('name', 'Unknown Name')  # Default to 'Unknown Name' if missing
    
    # Ensure the student name is a valid string before replacing spaces
    if isinstance(student_name, str):
        filename = f"{student_name.replace(' ', '_')}_Report.docx"
    else:
        filename = f"Unknown_Student_{index}_Report.docx"  # Handle invalid names

    # Render the template with student data
    template.render(student_data)

    # Save the document in the specified folder with the student's name
    output_path = os.path.join(output_folder, filename)
    template.save(output_path)

    print(f"Generated report for {student_name}: {output_path}")
