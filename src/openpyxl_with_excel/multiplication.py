import openpyxl

wb = openpyxl.load_workbook("C:\\Users\\PC\\OneDrive\\Documents\\Multiplication_Table.xlsx")
sheet = wb.active

num = int(input("Enter a number to generate its multiplication table: "))
for row_pos in range(1, num):
    for col_pos in range(1,num):
        sheet.cell(row = row_pos, column = col_pos, value = row_pos*col_pos)

wb.save("C:\\Users\\PC\\OneDrive\\Documents\\Multiplication_Table.xlsx")
print("Multiplication table generated.")