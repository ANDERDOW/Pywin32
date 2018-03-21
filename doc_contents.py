from win32com.client import constants

# This file is used to list the contents which would be used in document.

class table1_content():
    column1_content = ["Change Request No\n变更申请号", "Project\n项目",
                            "Change Requestor\n变更申请者",
                            "Attendees and Signatures\n评审者及签名"]
    column2_content = ["CR-2018-004","GrasebyC9","傅伟","Function部门"]+\
                      ["RND"]*4+["DAS","RA"]
