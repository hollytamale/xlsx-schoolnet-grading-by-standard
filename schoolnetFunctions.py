import openpyxl as xl
#
# def replacevalue(columnnumber, char):
#     for row in range(1, sheet.max_row):
#         old_value = sheet.cell(row + 1, columnnumber).value
#         new_value = old_value[0:char]
#         sheet.cell(row + 1, columnnumber).value = new_value

def pull_point_columns(new_schoolnet_file):
    wb = xl.load_workbook(new_schoolnet_file)
    sheet = wb.active
    # Generate list of N + # (less than number of columns)
    n_column = ["P" + str(num) for num in range(sheet.max_column + 1)]
    point_column_list = []
    # Find start and end of P columns for range, rather than set values
    for column in range(1, sheet.max_column + 1):
        col_header = sheet.cell(1, column)
        for value in n_column:
            if value == col_header.value:
                point_column_list.append(column)
                point_column_min = point_column_list[0]
                point_column_max = point_column_list[-1]
    # Found columns for points -- example, min is column 23 and max is column 30.
    # Now, find and store values for all rows from those columns.
    points_array = []
    for row in range(2, sheet.max_row + 1):
        points_list = []
        for i in range((point_column_max + 1) - point_column_min):
            point_column = point_column_min + i
            point_thing = sheet.cell(row, point_column).value
            points_list += [point_thing]
        points_array.append(points_list)
    return points_array

# # Add up points in P columns per student, returns sum and average
# for name in range(2, sheet.max_row + 1):
#     last_name = sheet.cell(name, 1).value
#     first_name = sheet.cell(name, 2).value
#     point_sum = 0
#     for points in range(point_column_min, point_column_max + 1):
#         point_per_question = sheet.cell(name, points)
#         # print(point_per_question)
#         # print(point_per_question.value)
#         point_sum += int(point_per_question.value)
#         point_average = point_sum / len(point_column_list)


# def pull_point_columns(new_schoolnet_file):
#     wb = xl.load_workbook(new_schoolnet_file)
#     sheet = wb['Sheet1']
#     # Generate list of N + value less than number of columns
#     n_column = ["P" + str(num) for num in range(sheet.max_column + 1)]
#     point_column_list = []
#     # Find start and end of P columns for range, rather than set values
#     for column in range(1, sheet.max_column + 1):
#         col_header = sheet.cell(1, column)
#         for value in n_column:
#             if value == col_header.value:
#                 point_column_list.append(column)
#                 point_column_min = point_column_list[0]
#                 point_column_max = point_column_list[-1]
#     # Add up points in P columns per student, returns sum and average
#     for name in range(2, sheet.max_row + 1):
#         last_name = sheet.cell(name, 1).value
#         first_name = sheet.cell(name, 2).value
#         point_sum = 0
#         for points in range(point_column_min, point_column_max + 1):
#             point_per_question = sheet.cell(name, points)
#             # print(point_per_question)
#             # print(point_per_question.value)
#             point_sum += int(point_per_question.value)
#             point_average = point_sum / len(point_column_list)


def pull_test_id(new_schoolnet_file):
    wb = xl.load_workbook(new_schoolnet_file)
    sheet = wb.active
    testid = str(sheet.cell(2, 8).value)
    return testid


def pull_names_from_schoolnet(new_schoolnet_file):
    wb = xl.load_workbook(new_schoolnet_file)
    sheet = wb.active
    # Store in array/list
    data_array = []
    for row in range(1, sheet.max_row):
        data_list = []
        for cell in range(1, 2):
            last_name = sheet.cell(row + 1, 1).value
            las_nam = last_name[0:2]
            first_name = sheet.cell(row + 1, 2).value
            fir_nam = first_name[0:2]
            id_number = str(sheet.cell(row + 1, 3).value)
            id_num = id_number[0:2]
            data_list = [las_nam, fir_nam, id_num]
            data_array.append(data_list)
            # pull_point_columns(new_schoolnet_file)
                # 'Name': las_nam + fir_nam,
                # 'ID': id_num}
    return data_array

    # wb = xl.load_workbook(schoolnet_file)
    # sheet = wb['Sheet1']
    # testid = str(sheet.cell(2, 8).value)
    # this_sheet = wb.create_sheet(testid)

        # for row in range(1, sheet.max_row):
        #     teacher_name = sheet.cell(row + 1, 6).value
        #     split_name = teacher_name.split()
        #     teacher_last_name = str(split_name[1:])
        #     teacher_last_name = teacher_last_name[2:4]
        #     this_sheet.cell(row + 1, 2).value = str(teacher_last_name)
        # for row in range(1, sheet.max_row):
        #     test_data = sheet.cell(row + 1, 8).value
        #     this_sheet.cell(row + 1, 3).value = test_data


# for row in range(1, sheet.max_row):
#     last_name = sheet.cell(row + 1, 1).value
#     las_nam = last_name[0:1]
#     sheet.cell(row + 1, 1).value = las_nam
#
# for row in range(1, sheet.max_row):
#     first_name = sheet.cell(row + 1, 2).value
#     fir_nam = first_name[0:1]
#     sheet.cell(row + 1, 2).value = fir_nam
#
# for row in range(1, sheet.max_row):
#     id_number = str(sheet.cell(row + 1, 3).value)
#     id_num = id_number[0:1]
#     sheet.cell(row + 1, 3).value = id_num