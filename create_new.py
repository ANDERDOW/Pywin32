from win32com.client import Dispatch
from win32com.client import constants
import Create_table, Create_paragraph
import Contents
import Fulfill_table
import sys, os, datetime
import win32com


def font_display(location, font, size, bold):
    location.Range.Font.Name = font
    location.Range.Font.Size = size
    location.Range.Font.Bold = bold


def get_log_file_name():
    now = datetime.datetime.now()
    file_name = "{0}-{1}-{2}-{3}-{4}-{5}".format(now.year, now.month, now.day, now.hour, now.minute, now.second)
    return "log_" + file_name + ".log"


def create_log_folder():
    log_folder_path = os.path.join(os.getcwd(), "log")
    if os.path.exists(log_folder_path):
        return
    os.mkdir(log_folder_path)


def main():
    # try:
    # very important!! If you use dispatch to open word, the constants attributes can't be found.
    word = win32com.client.gencache.EnsureDispatch('Word.Application')
    word.Visible = -1

    doc = word.Documents.Add()
    # 定位初始位置
    first_location = doc.Paragraphs(1)
    # 创建段落
    Create_paragraph.create_para(first_location, 10)
    # 再次分配位置
    head_location = []
    for para in range(8):
        head_location.append(doc.Paragraphs(para+1))

    Fulfill_table.fulfill_table1(doc,doc.Paragraphs(1))
    doc.Sections(1).Headers(constants.wdHeaderFooterPrimary).Range.Text = Contents.Header.header
    head_location[1].Range.InsertBefore(Contents.Title.title1)
    head_location[1].Range.InsertAfter(Contents.Title.title1_1)
    for i in range(len(head_location)):
        font_display(head_location[i],"Times New Roman",11,False)
    table2 = Fulfill_table.fulfill_table2(doc,head_location[3])

    pwd = sys.path[0]
    # 创建并填充表格
    # Fulfill_table.fulfill_table2(doc)
    doc.SaveAs(pwd + "\1.docx")
    # doc.Close()
    # word.Quit()


    # except Exception:
    #     errstr = str(EOFError)
    #     with open(os.path.join(os.getcwd(), get_log_file_name()), "w") as file:
    #         file.write(errstr)


if __name__ == "__main__":
    main()
