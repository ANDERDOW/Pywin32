import datetime, sys

now = datetime.datetime.now()
# This file is used to list the contents which would be used in document.
pwd = sys.path[0]
path = pwd + "\Graseby.xml"
sheet_name = "SW Issue"
requestor = ["FF"]
reviewer = ["CD", "CD", "CD", "CD", "CD", "CD"]
function_name =["RND", "RND", "RND", "RND", "DAS", "RA"]
CR_No =["CR-2018-004"]
project_name = ["GrasebyC9"]
version = ["V001"]
project_phase = ["SVA"]
review_date =[str(now.year)+"-"+str(now.month)+"-"+str(now.day)]


class Header:
    header = "Design Change Request\n设计更改申请表"


class Title:
    title1 = "1. Change Description and Change reason 变更描述及变更原因:"
    title1_1 = "A. 根据Mantis问题讨论结果，执行以下优化项，升级软件并验证:"


class Table1:
    column1_content = ["Change Request No.\n变更申请号", "Project\n项目",
                       "Change Requestor\n变更申请者",
                       "Attendees and Signatures\n评审者及签名"]
    column2_content = CR_No + project_name + requestor + ["Function部门"] + function_name
    column3_content = ["Version\n版本", "Project Phase\n项目阶段", "Review Date\n评审日期", "Name\n姓名"] + reviewer
    column4_content = version + project_phase + review_date + ["Signature/Date\n签字/日期"]
