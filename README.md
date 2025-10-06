![image alt](https://github.com/MkrManishk/Marksheet-Generator/blob/e0d626152736fc21f6dc97583d076a0d28f030a8/Screenshot%20(90).png)
# Marksheet-Generator

MARKSHEET GENERATOR - Setup and Usage Instructions

===== INSTALLATION =====

1. Ensure Python 3.x is installed on your system

2. Install required dependencies:
   pip install fpdf openpyxl matplotlib

   OR use the requirements.txt file:
   pip install -r requirements.txt

===== RUNNING THE APPLICATION =====

Run the application with:
   python marksheet_generator.py

===== FEATURES =====

1. STUDENT INFORMATION
   - Enter student name, roll number, and class/section
   - All fields are validated before calculation

2. SUBJECT MARKS
   - 5 subjects with customizable names
   - Enter marks obtained and total marks for each subject
   - Default total marks set to 100 (can be changed)
   - Validation ensures marks don't exceed total marks

3. CALCULATE RESULT
   - Computes total marks, percentage, and grade
   - Determines Pass/Fail status (33% passing criteria)
   - Displays a bar chart showing performance across subjects

4. CLEAR
   - Resets all input fields and results
   - Clears the performance chart

5. GENERATE PDF
   - Creates a professional PDF marksheet
   - Includes all student details and subject-wise marks
   - Shows total, percentage, grade, and result
   - Saves with timestamp in filename

6. GENERATE EXCEL
   - Exports marksheet to Excel format
   - Formatted with colors and proper alignment
   - Includes subject-wise percentage calculation
   - Saves with timestamp in filename

===== GRADING SYSTEM =====

A+  : 90% and above
A   : 80% - 89%
B+  : 70% - 79%
B   : 60% - 69%
C   : 50% - 59%
D   : 40% - 49%
E   : 33% - 39%
F   : Below 33% (Fail)

===== USER INTERFACE =====

- Modern, clean design with light grey/white background
- Soft blue buttons with hover effects
- Responsive layout that adapts to window resizing
- Clear sections for input, results, and visualization
- Real-time performance chart using matplotlib

===== FILE OUTPUTS =====

Generated files are saved in the current directory with format:
- PDF: marksheet_[RollNumber]_[Timestamp].pdf
- Excel: marksheet_[RollNumber]_[Timestamp].xlsx

===== NOTES =====

- All input fields are validated before processing
- Marks cannot be negative or exceed total marks
- Results must be calculated before generating PDF/Excel
- The application uses Tkinter (included with Python)
- Performance chart automatically updates on calculation
