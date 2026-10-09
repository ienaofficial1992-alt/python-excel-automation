from openpyxl import load_workbook

# Grade বের করার function
def get_grade(marks):
    if marks >= 80:
        return "A"
    elif marks >= 60:
        return "B"
    elif marks >= 50:
        return "C"
    else:
        return "F"


# Excel file খুলছি
workbook = load_workbook("students.xlsx")
sheet = workbook.active

# Heading
sheet["C1"] = "Result"
sheet["D1"] = "Grade"

# Counter
a_count = 0
b_count = 0
c_count = 0
f_count = 0

# Total marks
total_marks = 0

# প্রতিটি student-এর data process
for row in range(2, sheet.max_row + 1):

    marks = sheet[f"B{row}"].value
    total_marks += marks

    # Result
    if marks >= 80:
        result = "Nice"
    elif marks >= 50:
        result = "Need improvement"
    else:
        result = "Fail"

    # Grade
    grade = get_grade(marks)

    # Excel-এ লেখা
    sheet[f"C{row}"] = result
    sheet[f"D{row}"] = grade

    # Grade count
    if grade == "A":
        a_count += 1
    elif grade == "B":
        b_count += 1
    elif grade == "C":
        c_count += 1
    elif grade == "F":
        f_count += 1


# মোট student
total_students = sheet.max_row - 1

# Average
average_marks = total_marks / total_students


# Summary
print("----- STUDENT REPORT -----")
print("Total students:", total_students)
print("Total marks:", total_marks)
print("Average marks:", round(average_marks, 2))
print("A:", a_count)
print("B:", b_count)
print("C:", c_count)
print("F:", f_count)


# নতুন Excel file
summery=workbook.create_sheet("Summery")
summery["A1"] = "student report summery"
summery["A3"] = "total student"
summery["B3"] = total_students
summery["A4"] = "total marks"
summery["B4"] = total_marks
summery["A5"] = "average marks"
summery["B5"] = average_marks
summery["A7"] = "grade"
summery["B7"] = ("number of students")
summery["A8"] = "A"
summery["B8"] = a_count
summery["A9"] = "B"
summery["B9"] = b_count
summery["A10"] = "C"
summery["B10"] = c_count
summery["A11"] = "F"
summery["B11"] = f_count
from openpyxl.styles import Font, PatternFill, Alignment

# Main sheet-এর heading format
for cell in sheet[1]:
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill(
        fill_type="solid",
        fgColor="1565C0"
    )
    cell.alignment = Alignment(horizontal="center")

# Column width ঠিক করা
sheet.column_dimensions["A"].width = 18
sheet.column_dimensions["B"].width = 12
sheet.column_dimensions["C"].width = 22
sheet.column_dimensions["D"].width = 12

# Summary sheet-এর heading format
for cell in summery[1]:
    cell.font = Font(bold=True, size=14)
    cell.alignment = Alignment(horizontal="center")

summery.column_dimensions["A"].width = 25
summery.column_dimensions["B"].width = 22
from openpyxl.styles import Font, PatternFill, Alignment

# Summary title
summery["A1"].font = Font(
    bold=True,
    size=16,
    color="FFFFFF"
)
summery["A1"].fill = PatternFill(
    fill_type="solid",
    fgColor="1565C0"
)

# Average marks: 2 decimal places
summery["B5"].number_format = "0.00"

# Center alignment
for row in summery.iter_rows():
    for cell in row:
        cell.alignment = Alignment(vertical="center")

# Column width
summery.column_dimensions["A"].width = 25
summery.column_dimensions["B"].width = 24
# Heading merge করা
summery.merge_cells("A1:B1")

# Heading center করা
summery["A1"].alignment = Alignment(
    horizontal="center",
    vertical="center"
)
summery.row_dimensions[1].height = 30
from openpyxl.styles import PatternFill, Font

# Grade অনুযায়ী রং
grade_colors = {
    "A": "008000",
    "B": "0000FF",
    "C": "FFA500",
    "F": "FF0000"
}

for row in range(8, 12):
    grade = summery[f"A{row}"].value
    color = grade_colors[grade]

    summery[f"A{row}"].fill = PatternFill(
        fill_type="solid",
        fgColor=color
    )

    summery[f"A{row}"].font = Font(
        bold=True,
        color="FFFFFF"
    )
workbook.save("students_final_report.xlsx")

print("Report created successfully!")