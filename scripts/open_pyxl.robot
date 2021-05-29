*** Settings ***
Documentation   Open PyXL Keywords
Library     scripts.OpenRobotPyxl    WITH NAME   pyxlobj

*** Keywords ***
Create Workbook
    [Documentation]     The keyword to create new workbook with input as path, file name,
    ...                 sheet name and position of the sheet in the workbook

    [Arguments]    ${PATH}    ${BOOK_NAME}    ${SHEET_NAME}    ${POSTION}
    pyxlobj.Create New Workbook   ${PATH}   ${BOOK_NAME}    ${SHEET_NAME}    ${POSTION}

Close Workbook
    [Documentation]     The keywords to close the created workbook.
    pyxlobj.Close Workbook

Get Active Sheet
    [Documentation]      Get Active Sheet
    ${value}    pyxlobj.Get Active Sheet
    [return]    ${value}

Get Active Sheet Name
    [Documentation]     Get Active Sheet Name in the current workbook
    ${value}    pyxlobj.Active Sheet Name
    [return]    ${value}