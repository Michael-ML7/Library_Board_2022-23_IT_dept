# language used: Python 3
# local python version: 3.10.5 (though the code will work with all Python 3 versions)

from openpyxl import Workbook, load_workbook # this library can be used to deal with the Excel files

wb = load_workbook('22-23 Library Aptitude Test - Coding Option 2 - Data.xlsx') # load the workbook downloaded from Google Sheets
file = wb.active # set to the active sheet in the workbook
# print(file) # debug line

i = 1 # iterator for each student
mx_books = 0 # marking the maximum books one has borrowed
cnt_ppl = 0 # counting the number of students in the sheet
while file["A%s" %i].value != None: # while we've not reached the end of the log
	cnt_ppl += 1
	j = i # this person's log history starts here
	while file["A%s" %(j + 1)].value != None and file["A%s" %(j + 1)].value[0] == '2': # loop through all books this person has borrrowed
		j += 1 # increment the pointer j as there are still borrow history made by that person
	mx_books = max(mx_books, (j - i + 1)) # record the maximum books a person has borrowed

	for k in range(i + 1, j + 1): # loop through all the books this person has borrowed
		file.move_range("A%s:A%s" %(k, k), rows=(i-k), cols=(k-i)) # move them to a single line as described in the task description
	i = j + 1 # change pointer i to j + 1 as we finished deadling with this person's borrowing history

lst_line = 0 # marking the last line of the log book
while file["A%s" &lst_line].value != None:
	lst_line += 1 # this line is not empty go on with the next line
lst_line -= 1 # as the row[lst_line] is now empty, to get the last row with content, we minus one from the number

for k in range(lst_line, 0, -1): # loop through all the lines of student logs
	if file["A%s" %k].value == None: # if this cell is empty -> this entire row is empty
		file.delete_rows(idx=k, amount=1) # delete empty rows

file.insert_rows(idx=1, amount=2) # insert two rows at the top of the document for 1. Headers 2. Librarian example
file.insert_cols(idx=2, amount=2) # insert two columns at the left of the document for class and class numberS
file.insert_cols(idx=3, amount=2) # insert two columns at the right of column 3 (column 3 should contain the Chinese name and the English name of the student) for storing their Chinese name and English name
file["A1"].value = "Class" # Preset value given in the task description
file["B1"].value = "Class No" # Preset value given in the task description
file["C1"].value = "English & Chinese Name" # Preset value given in the task description
file["D1"].value = "English Name" # Preset value given in the task description
file["E1"].value = "Chinese Name" # Preset value given in the task description
file["A2"].value = "7H" # Preset value given in the task description
file["B2"].value = 1 # Preset value given in the task description
file["C2"].value = "Librarian 圖書館管理員" # Preset value given in the task description
file["D2"].value = "LIBRARIAN" # Preset value given in the task description
file["E2"].value = "圖書館管理員" # Preset value given in the task description
file["F2"].value = "19950605 CBK 855 書一本" # Preset value given in the task description

file.move_range("A3:A%s" %(3 + cnt_ppl), rows=0, cols=2) # move the all student names on column A to column C

for k in range(3, 3 + cnt_ppl): # iterating the students
	file["A%s" %k].value = "%s%s" %((k - 3) % 6 + 1, chr(((k - 3) % 7) + ord('A'))) # generating classes from the cycle of 1A, 2B, 3C, 4D, 5E, 6F, as classes are not given in the Google Sheet
	file["B%s" %k].value = k # generating class numbers, as class numbers are not given in the Google Sheet, x is generated as the class number for the x-th student given in the Google Sheet

for k in range(3, 3 + cnt_ppl): # iterating the students
	full_name = file["C%s" %k].value # copy and store the students full name in variable fulL_name
	eng_name = "" # create varible to store the students English name
	for x in full_name: # for every character in the student's full name
		if ('A' <= x and x <= 'Z') or ('a' <= x and x <= 'z') or x == ' ': # if the character matches the format of an English name
			eng_name = eng_name + x # append the character to the student's English name
	chi_name = "" # same for the Chinese name
	for x in full_name: # same as above
		if not(('A' <= x and x <= 'Z') or ('a' <= x and x <= 'z') or x == ' '): # if the character does NOT match the format of an English name -> must be a Chinese character -> a part of the student's Chinese name
			chi_name = chi_name + x # same as above
	
	eng_name = eng_name.upper() # change the format of the student's English name to all uppercase, to match the requirements in the task description
	file["D%s" %k].value = eng_name # store the student's English name is its corresponding position
	file["E%s" %k].value = chi_name # store the student's Chinese name is its corresponding position

for k in range(mx_books): # loop through the maximum number of books a student has borrowed
	# print(chr(ord('F') + k)) # debug line
	file["%s1" % (chr(ord('F') + k))].value = "Book %s" %(k + 1) # set the header cell to the format given in the task description "Book x"

wb.save('22-23 Library Aptitude Test - Coding Option 2 - Data.xlsx') # finally save the formatted excel document

# I dont know not how to colour the cells so I did them manually after running this code