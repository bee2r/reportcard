import pandas as pd

# Load the data from the Excel file
def load_data(file_name):
    # Read the Excel file into a DataFrame
    return pd.read_excel(file_name)
    print("Columns in the Excel file:", list(data.columns))
    return data

# Search for student by name or admission number
def find_student(data, key):
    # Check if the key matches 'Name of Students' or 'Admission No'
    student_data = data[(data['Name of Students'] == key) | (data['Admission No'] == key)]
    return student_data

# Generate the report card
def generate_report_card(student_data):
    if student_data.empty:
        print("Student not found.")
        return
    else:
        # Extract student details
        student = student_data.iloc[0]
        print("\n=== Report Card ===")
        print(f"Name: {student['Name of Students']}")
        print(f"Admission No: {student['Admission No']}")
        print("\nSubjects and Marks:")
        
        # Print each subject and mark
        for column in student.index[3:-5]:  # Subjects columns
            print(f"{column}: {student[column]}")
        
        # Summary
        print("\n=== Summary ===")
        print(f"Total Subjects: {student['Total Number of subjects']}")
        print(f"Total Marks Obtained: {student['Total marks']}")
        print(f"Average: {student['Average']}")
        print(f"Position: {student['Position']}")
        print("=================\n")

# Main function to interact with the user
def main():
    file_name = "results.xlsx"  # Replace with the correct path if needed
    try:
        data = load_data(file_name)
        print("Data loaded successfully!")
        while True:
            key = input("Enter student's Name or Admission No (or 'exit' to quit): ").strip()
            if key.lower() == 'exit':
                break
            student_data = find_student(data, key)
            generate_report_card(student_data)
    except FileNotFoundError:
        print("The file 'results.xlsx' was not found. Please ensure it's in the same directory.")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
