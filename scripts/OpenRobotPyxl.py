import openpyxl
from robot.api.deco import keyword, library


class OpenRobotPyxl:

    def __init__(self):
        self.active_sheet = None
        self.active_book = None
        pass

    @keyword("Create New Workbook")
    def create_new_workbook(self, path, book_name, sheet_name, postion=0):
        wb = openpyxl.Workbook()
        book_name = book_name + ".xlsx"
        ws = wb.create_sheet(sheet_name, postion)
        wb.save(path + "/" + book_name)
        self.active_book, self.active_sheet = wb, ws
        return self.active_book

    @keyword('Get Active Sheet')
    def get_active_sheet(self):
        if self.active_book and self.active_sheet:
            return self.active_sheet

    @keyword('Active Sheet Name')
    def get_active_sheet_name(self):
        if self.active_book and self.active_sheet:
            return self.active_sheet.name





