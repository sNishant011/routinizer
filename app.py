# program for reading excel file named all_routine containing routine of all sections to make 
# other individual routines and then writing them to a new excel file
import os
from openpyxl import load_workbook, Workbook

all_routine = load_workbook('test.xlsx')
# getting active worksheet (first worksheet)
active_sheet = all_routine.active
all_groups = []
# creating a list of all groups
for i in range(1, 12):
	all_groups.append(f"L5CG{i}")

# making folder for storing individual routine
try:
	os.mkdir("individual_routines")
	print("New directory called individual_routines created.")
except:
	print("Directory already exist continuing further.")

for group in all_groups:
	# creating a new workbook for each group
	wb = Workbook()
	ws = wb.active
	# title of worksheet
	ws.title = group + "_routine_3rdSem"

	routine_heading = []

	for cell in active_sheet[2]:

		# getting value of each cell of second row ( 'Day', 'Time'..)
		routine_heading.append(cell.value)
	
	# appending routine heading to the first row of the sheet
	ws.append(routine_heading)

	# first two rows of the original routine contains unneccessary data
	for row in active_sheet.iter_rows(3, active_sheet.max_row):

		# to mitigate error bug to matching of group name while using in operator
		# L5CG1 in L5CG10 and L5CG11 will
		if group == 'L5CG1':
			class_grouped = group+'+'
		else:
			class_grouped = group
		if (group == row[7].value) or (class_grouped in row[7].value):

			# row that will contain all the information about the class
			new_row = []
			for i in range(active_sheet.max_column):

				# appending cell value to new row
				new_row.append(row[i].value)

			# appending each day routine to the new excel file
			ws.append(new_row)

			# saving file inside individual_routines folder
			wb.save(f"individual_routines/{group}_routine_3rdSem.xlsx")
	wb.close()
all_routine.close()
