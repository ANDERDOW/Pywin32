from win32com.client import constants
import Contents, Create_table
import Get_excel


def font_display(location, font, size, bold):
    location.Range.Font.Name = font
    location.Range.Font.Size = size
    location.Range.Font.Bold = bold


def fulfill_table1(document,location):
    table1 = Create_table.create_table(document,location, 10, 4)
    table1_row_num = table1.Columns(2).Cells.Count
    # table1_column_num = table1.Rows(2).Cells.Count

    # set the font location
    for i in range(table1_row_num):
        table1.Rows(i + 1).Cells.VerticalAlignment = constants.wdCellAlignVerticalBottom
    # set the font in row1~3
    for i in range(3):
        table1.Rows(i + 1).Cells.VerticalAlignment = constants.wdAlignVerticalTop

    table1.Cell(4, 1).Merge(table1.Cell(10, 1))
    # new_table1_row_num = table1.Rows(1).Cells.Count
    # 这里是个疑问点，因为会抛出error:Cannot access individual rows in this collection because the table has vertically merged cells
    # 需要找可以确认合并后行的数目，此处先用数字代替
    # 目前通过content中的数目来进行数字的代替，不再使用行号
    font_display(location,"Times New Roman",10,False)

    first_column_content = Contents.Table1().column1_content
    second_column_content = Contents.Table1().column2_content
    third_column_content = Contents.Table1().column3_content
    fourth_column_content = Contents.Table1().column4_content

    # fulfill the first column
    for column in range(len(first_column_content)):
        table1.Columns(1).Cells(column + 1).Range.Text = first_column_content[column]
    # fulfill the second column
    for column in range(len(second_column_content)):
        table1.Columns(2).Cells(column + 1).Range.Text = second_column_content[column]
    # fulfill the third column
    for column in range(len(third_column_content)):
        table1.Columns(3).Cells(column + 1).Range.Text = third_column_content[column]
    for column in range(len(fourth_column_content)):
        table1.Columns(4).Cells(column + 1).Range.Text = fourth_column_content[column]


def fulfill_table2(document,location):

    table2_row_num = len(Get_excel.Main.summary_list)
    table2 = Create_table.create_table(document,location, table2_row_num, 4)

    # set the font location

    font_display(document.Paragraphs(3),"Times New Roman",10,False)

    first_column_content = Get_excel.Main.ID_list
    second_column_content = Get_excel.Main.summary_list

    # fulfill the first column
    for column in range(len(first_column_content)-1):
        # 此时有减去一行，否则会报错
        table2.Columns(1).Cells(column + 2).Range.Text = column+1
    table2.Columns(1).Cells(1).Range.Text = "Index"

    # fulfill the second column
    for column in range(len(first_column_content)):
        table2.Columns(2).Cells(column + 1).Range.Text = first_column_content[column]
    # fulfill the third column
    for column in range(len(second_column_content)):
        table2.Columns(3).Cells(column + 1).Range.Text = second_column_content[column]
    # fulfill the fourth column
    for column in range(len(second_column_content)-1):
        # 此时有减去一行，否则会报错
        table2.Columns(4).Cells(column + 2).Range.Text = "优化项"
    table2.Columns(4).Cells(1).Range.Text = "Change Reason"

    # for i in range(table2_row_num):
    #     table2.Rows(i + 1).Cells.AutoFit()
    #     table2.Rows(i + 1).Cells.VerticalAlignment = constants.wdCellAlignVerticalBottom
    table2.Rows(1).Cells.Shading.Texture = 60
    for i in range(4):
        table2.Columns(i+1).Cells.AutoFit()


def f_table1(document,table):

    first_column_content = Contents.Table1().column1_content
    second_column_content = Contents.Table1().column2_content
    third_column_content = Contents.Table1().column3_content
    fourth_column_content = Contents.Table1().column4_content

    # fulfill the first column
    for column in range(len(first_column_content)):
        table.Columns(1).Cells(column + 1).Range.Text = first_column_content[column]
    # fulfill the second column
    for column in range(len(second_column_content)):
        table.Columns(2).Cells(column + 1).Range.Text = second_column_content[column]
    # fulfill the third column
    for column in range(len(third_column_content)):
        table.Columns(3).Cells(column + 1).Range.Text = third_column_content[column]
    for column in range(len(fourth_column_content)):
        table.Columns(4).Cells(column + 1).Range.Text = fourth_column_content[column]


def f_table2(document,table):

    table2_row_num = len(Get_excel.Main.summary_list)

    first_column_content = Get_excel.Main.ID_list
    second_column_content = Get_excel.Main.summary_list

    # fulfill the second column
    for column in range(len(first_column_content)):
        table.Columns(2).Cells(column + 1).Range.Text = first_column_content[column]
    # fulfill the third column
    for column in range(len(second_column_content)):
        table.Columns(3).Cells(column + 1).Range.Text = second_column_content[column]
    # fulfill the fourth column
    for column in range(len(second_column_content)-1):
        # 此时有减去一行，否则会报错
        table.Columns(4).Cells(column + 2).Range.Text = "优化项"
    table.Columns(4).Cells(1).Range.Text = "Change Reason"

    # for i in range(table2_row_num):
    #     table2.Rows(i + 1).Cells.AutoFit()
    #     table2.Rows(i + 1).Cells.VerticalAlignment = constants.wdCellAlignVerticalBottom
    for i in range(4):
        table.Columns(i+1).Cells.AutoFit()
