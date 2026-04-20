from openpyxl import load_workbook


file_path = input("Enter path name")
workbook = load_workbook(file_path)
sheet = workbook.active


room_list = []
for row in sheet.iter_rows(values_only=True):
    room_list.append(list(row))

unmatched = []
matched = True

for line in room_list:
    if line[0] != line[2]:

        matched = False
        unmatched.append(tuple((line[0], line[2])))

if unmatched:
    print("Error: There is a mismatch")
    print(unmatched)
else:
    print("Success: All the rooms matched!")
