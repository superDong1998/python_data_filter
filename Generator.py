import xlrd
import sys

# 院系 团学联 班级 个人公众号
departments = []
stu_organizations = []
personal_accounts = []
class_accounts = []

# rows in source_1 file
rows = []

# data come from weekly file
data = []

# result data arrays
personal_result = []
stu_dept_result = []
class_result = []


def read_source_xls(filename):
    workbook = xlrd.open_workbook(filename)
    table = workbook.sheet_by_index(0)

    for i in range(table.nrows):
        row_value = table.row_values(i)
        rows.append(row_value)

    split_oranization(rows)


def read_weekly_xlsx(filename):
    workbook1 = xlrd.open_workbook(filename)
    table1 = workbook1.sheet_by_index(0)
    for i in range(table1.nrows):
        row_value = table1.row_values(i)
        if i == 0 or row_value[0] == '':
            continue
        data.append(row_value[0])


def find_data():
    for i in range(len(data)):
        if data[i] in departments or data[i] in stu_organizations:
            stu_dept_result.append(data[i])
        if data[i] in personal_accounts:
            personal_result.append(data[i])
        if data[i] in class_accounts:
            class_result.append(data[i])
        if len(stu_dept_result) > 4 and len(personal_result) > 2 and len(class_result) > 4:
            break


def main(argv):
    if (len(argv) < 3):
        usage()
        exit(0)
    print("working...")
    read_source_xls(argv[1])
    read_weekly_xlsx(argv[2])
    find_data()
    print("result:\n")
    print("个人榜：\n")
    for j in range(len(personal_result)):
        if j <= 2:
            print(personal_result[j])
        else:
            break
    print("\n院系、团学联榜：\n")
    for i in range(len(stu_dept_result)):
        if i <= 4:
            print(stu_dept_result[i])
        else:
            break
    print("\n班级榜：\n")
    for k in range(len(class_result)):
        if k <= 4:
            print(class_result[k])
        else:
            break
    print("\ndone.")


def usage():
    print("please usage input file: ")
    print("example:    ")
    print(" ")
    print("        python Generator.py sourcefile.xls 20190519-20190525.xlsx")
    print(" ")
    print("thanks for using it!")


def split_oranization(rows):
    depart_flag = False
    stu_flag = False
    personal_flag = False
    class_flag = False
    for i in range(len(rows)):
        if (rows[i][0] == '院系'):
            depart_flag = True
            stu_flag = False
            personal_flag = False
            class_flag = False
        elif (rows[i][0] == '团学联'):
            depart_flag = False
            stu_flag = True
            personal_flag = False
            class_flag = False
        elif (rows[i][0] == '班级'):
            depart_flag = False
            stu_flag = False
            personal_flag = False
            class_flag = True
        elif (rows[i][0] == '个人'):
            depart_flag = False
            stu_flag = False
            personal_flag = True
            class_flag = False
        elif isinstance(rows[i][0], str):
            depart_flag = False
            stu_flag = False
            personal_flag = False
            class_flag = False
        if depart_flag and rows[i][1] != '':
            departments.append(rows[i][1])
            continue
        if stu_flag and rows[i][1] != '':
            stu_organizations.append(rows[i][1])
            continue
        if personal_flag and rows[i][1] != '':
            personal_accounts.append(rows[i][1])
            continue
        if class_flag and rows[i][1] != '':
            class_accounts.append(rows[i][1])
            continue


main(sys.argv)
