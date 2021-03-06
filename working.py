from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font

<<<<<<< HEAD
data = {
	"Joe": {
		"math": 65,
		"science": 78,
		"english": 98,
		"gym": 89
	},
	"Bill": {
		"math": 55,
		"science": 72,
		"english": 87,
		"gym": 95
	},
	"Tim": {
		"math": 100,
		"science": 45,
		"english": 75,
		"gym": 92
	},
	"Sally": {
		"math": 30,
		"science": 25,
		"english": 45,
		"gym": 100
	},
	"Jane": {
		"math": 100,
		"science": 100,
		"english": 100,
		"gym": 60
	}
}

wb = Workbook()
ws = wb.active
ws.title = "Grades"

headings = ['Name'] +  list(data['Joe'].keys())
ws.append(headings)

for person in data:
    grades = list(data[person].values())
    ws.append([person] + grades)

for col in range(2, len(data['Joe']) + 2):
    char = get_column_letter(col)
    ws[char + "7"] = f"=SUM({char + '2'}:{char + '6'})/{len(data)}"
    
for col in range(1,6):
    ws[get_column_letter(col) + '1'].font = Font(bold=True, color="00ff0000")

wb.save("Grades.xlsx")
=======
wb = load_workbook('GTN.xlsx')
ws = wb.active
print(ws['A2'].value)
ws['A2'].value = "Test"
print(ws['A2'].value)

wb.save('GTN.xlsx')
>>>>>>> 3d4ae598f8517ab8a93af3b7c1e9a918e23cf3ab
