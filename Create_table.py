from win32com.client import constants


def create_table(document, location, row, column):
    table = document.Tables.Add(location.Range, row, column)
    table.Borders.InsideLineStyle = constants.wdLineStyleSingle
    table.Borders.OutsideLineStyle = constants.wdLineStyleSingle
    table.PreferredWidth = (100 / 0.88) * column  # unit: points
    # 很重要的一步，需要返回table给对象，否则后续以None type类型无法进行其余操作
    return table
