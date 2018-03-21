from win32com.client import constants
import win32com

class Create_Table1():
    # 定义函数，用于创建段落或定义格式等
    def __init__(self, file):
        self.file = file

    def location_paragraph(self, num):
        return self.file.Paragraphs(num)

    def create_para(self, paragraph, num):
        for i in range(num):
            self.file.Range(0, 0).Paragraphs.Add()

    def roman_font_display(self, location):
        # 自己使用指代遍历一遍后发现location即为self.file.Paragraphs(num),
        # 因此无需再加self对象，否则会报错
        location.Range.Font.Name = "Times New Roman"
        location.Range.Font.Size = 10
        location.Range.Font.Bold = True

    def create_table(self, location, row, column):
        self.table = self.file.Tables.Add(location.Range, row, column)
        self.table.Borders.InsideLineStyle = constants.wdLineStyleSingle
        self.table.Borders.OutsideLineStyle = constants.wdLineStyleSingle
        self.table.PreferredWidth = (100 / 0.88) * column  # unit: points
        # 很重要的一步，需要返回table给对象，否则后续以None type类型无法进行其余操作
        return self.table
