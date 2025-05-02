from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.title = "Student_scores"

ws.append(["Name", "Grade", "Remarks"])

wb.save("Student_scores.xlsx")
print("Excel file created successfully.")