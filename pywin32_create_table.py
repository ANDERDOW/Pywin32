from win32com.client import Dispatch, constants
import title_table, doc_contents
import win32com
# very important!! If you use dispatch to open word, the constants attributes can't be found.
word = win32com.client.gencache.EnsureDispatch('Word.Application')
word.Visible = -1

path = r"C:\Users\mshaffu\Desktop\untitled\Generate Change Orders Documents\1.docx"
doc = word.Documents.Add()

# define document
document1 = title_table.Create_Table1(doc)
# define the init location
location_paragraph = document1.location_paragraph(1)
# create paragraphs
document1.create_para(location_paragraph, 2)
title = doc.Range(0, 0).InsertBefore("Design")
table1 = document1.create_table(document1.location_paragraph(3),10,4)
table1_row_num = table1.Columns(2).Cells.Count
table1_column_num = table1.Rows(2).Cells.Count

# set the font location
for i in range(table1_row_num):
    table1.Rows(i + 1).Cells.AutoFit()
    table1.Rows(i+1).Cells.VerticalAlignment = constants.wdAlignVerticalCenter
# set the font in row1~3
for i in range(3):
    table1.Rows(i+1).Cells.VerticalAlignment = constants.wdAlignVerticalTop


table1.Cell(4, 1).Merge(table1.Cell(10, 1))
# new_table1_row_num = table1.Rows(1).Cells.Count
# 这里是个疑问点，因为会抛出error:Cannot access individual rows in this collection because the table has vertically merged cells
# 需要找可以确认合并后行的数目，此处先用数字代替
document1.roman_font_display(document1.location_paragraph(1))

first_column_content = doc_contents.table1_content().column1_content
second_column_content = doc_contents.table1_content().column2_content

# fulfill the first column
for column in range(4):
    table1.Columns(1).Cells(column+1).Range.Text = first_column_content[column]
# fulfill the second column
for column in range(table1_row_num):
    table1.Columns(2).Cells(column+1).Range.Text = second_column_content[column]
