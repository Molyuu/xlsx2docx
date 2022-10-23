#!/usr/bin/python


from docxtpl import DocxTemplate
from openpyxl import load_workbook, Workbook
from shutil import move, copy
from os import remove
from sys import argv

xlsx = ""
docx = ""


def column_to_name(colnum):
    str = ""
    while not (colnum // 26 == 0 and colnum % 26 == 0):
        temp = 25
        if colnum % 26 == 0:
            str += chr(temp + 65)
        else:
            str += chr(colnum % 26 - 1 + 65)
        colnum //= 26
    return str


def format_excel(excel):
    copy(excel, "./%s_gork.xlsx" % excel)
    a_excel = "%s_gork.xlsx" % excel
    wb = load_workbook(a_excel, data_only=True)
    sheet = wb.active

    for i in range(sheet.max_row):
        for j in range(sheet.max_column):
            """
            print(i + 1, column_to_name(j + 1))
            """
            cell_type = sheet["%s%d" % (column_to_name(j + 1), i + 1)].number_format
            fin_type = ['_ * #,##0.00_ ;_ * \\-#,##0.00_ ;_ * "-"??_ ;_ @_ ']
            # '_ * #,##0.00_ ;_ * \\-#,##0.00_ ;_ * "-"??_ ;_ @_ '
            # '_ * #,##0.00_ ;_ * \\-#,##0.00_ ;_ * "-"??_ ;_ @_'
            if str(cell_type) == fin_type[0]:
                if type(sheet["%s%d" % (column_to_name(j + 1), i + 1)].value) == str:
                    break
                else:
                    float_value = float(
                        sheet["%s%d" % (column_to_name(j + 1), i + 1)].value
                    )
                sheet["%s%d" % (column_to_name(j + 1), i + 1)] = f"{float_value:,.2f}"
    wb.save("%s_gork.xlsx" % excel)


if len(argv) == 1:
    print("======xlsx批量填入docx======")
    print("请仔细阅读本目录下的使用说明.pdf！！！！")
    print("请仔细阅读本目录下的使用说明.pdf！！！！")
    print("请仔细阅读本目录下的使用说明.pdf！！！！")
    print("========================")
    a = input("你是否已经阅读了 使用说明.pdf 并完全理解了如何操作？(请输入yes或no)")
    if a == "yes":
        print("好的！")
        xlsx = "SRC.xlsx"
        docx = "SRC.docx"
    else:
        print("好的！")
        b = input("请按下回车键以退出程序！")
        exit(0)
elif len(argv) == 3:
    xlsx = argv[1]
    docx = argv[2]

format_excel(xlsx)
doc = DocxTemplate(docx)
wb = load_workbook("%s_gork.xlsx" % xlsx, data_only=True, keep_vba=True)
sheet = wb.active
for row in sheet:
    rowMap = map(lambda x: x.value, row)
    if row[0].row == 1:
        title = list(rowMap)
    else:
        context = dict(zip(title, rowMap))
        filename = context["filename"]
        print("处理", filename, "中", end="\r")
        doc.render(context)
        doc.save("%s.docx" % filename)
        move("%s.docx" % filename, "./OUT")
        print("处理", filename, "完毕！.")
remove("SRC.xlsx_gork.xlsx")

c = input("处理完成！按下回车键来退出")

