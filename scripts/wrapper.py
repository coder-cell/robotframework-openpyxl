import openpyxl


class OpenPyxl:

    def __init__(self):
        self.active_sheet = None
        self.active_book = None
        pass

    def create_new_workbook(self, book_name, sheet_name, postion=0):
        wb = openpyxl.Workbook()
        ws = wb.create_sheet(sheet_name, postion)
        wb.save(book_name)
        self.active_book, self.active_sheet = wb, ws
        return self.active_book

    def get_active_sheet(self):
        if self.active_book and self.active_sheet:
            return self.active_sheet

    def get_active_sheet_name(self):
        if self.active_book and self.active_sheet:
            return self.active_sheet.name





