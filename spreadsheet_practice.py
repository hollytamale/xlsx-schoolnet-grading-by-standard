import openpyxl as xl
from openpyxl import Workbook
from schoolnetFunctions import pull_test_id, pull_names_from_schoolnet, pull_point_columns

"""NOTES TO THE NOODLE:
- 10/14, added standards questions
- Next steps, 
    - standardizing Standards inputs
    - Consolidating student info into one tag? Good idea, I think?
    - averaging grades by standard
    - Consolidate info by average or standard on first sheet, sheet 1.
    - lastly, major clean up and reorganizing. Hm. 
"""

# wb = Workbook()
wb = xl.load_workbook('2023-24Grades.xlsx')
sheet = wb.active

file_to_import = 'TestResults_5174646.xlsx' # Replace additional file name here
testid = pull_test_id(file_to_import)

if wb.sheetnames.count(testid) == True:
    sheet_new = wb[testid]
else:
    sheet_new = wb.create_sheet(testid, 1)

# Unpack array of redacted names and IDs to a new sheet from pull_names function
names_array = pull_names_from_schoolnet(file_to_import)
for item in range(0, len(names_array)):
    names_row = names_array[item]
    for point in range(len(names_row)):
        sheet_new.cell(item + 1, point + 1).value = str(names_array[item][point])

# Unpack points array to same test sheet from pull_points function
points_array = pull_point_columns(file_to_import)
for item in range(0, len(points_array)):
    points_row = points_array[item]
    for point in range(0, len(points_row)):
        sheet_new.cell(item + 1, point + len(names_row) + 1).value = points_array[item][point]

print(bool(sheet_new.cell(1,1).value))

# Need to standardize how standards are input, or they won't categorize together
if sheet_new.cell(1, 1).value:
    resubmit = input("Do you wish to resubmit standards for this quiz? (Yes/No): ").lower()
    if resubmit == "yes":
        # This junk is redundant. I need to avoid repeating myself.. laterz.
        sheet_new.insert_rows(1)
        print(f"For test ID #{testid}, identify standard for: ")
        for question in range(len(points_row)):
            q_standard = input(f"Q{question + 1}: ")
            ques_num = "Q" + str(question + 1)  # Formats question number
            ques_column_pos = question + len(names_row) + 1
            sheet_new.cell(1, ques_column_pos).value = f'{ques_num}:{q_standard}'
    else:
        print("Okay, no changes will be made.")
# Never goes to this loop, either. How to fix, and redundancy, too?
elif sheet_new.cell(1, 1).value == False:
    sheet_new.insert_rows(1)
    print(f"For test ID #{testid}, identify standard for: ")
    for question in range(len(points_row)):
        q_standard = input(f"Q{question + 1}: ")
        ques_num = "Q" + str(question + 1) # Formats question number
        ques_column_pos = question + len(names_row) + 1
        sheet_new.cell(1, ques_column_pos).value = f'{ques_num}:{q_standard}'

# Also, slice string for what comes before and after the colon to find Q# and Standard easily?

wb.save('2023-24Grades.xlsx')
