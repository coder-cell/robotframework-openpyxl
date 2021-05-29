*** Settings ***
Documentation   Open PyXL Keywords
Library     scripts.OpenRobotPyxl    WITH NAME   pyxl

*** Keywords ***
Create Workbook
    [Documentation]     The keyword to create new workbook with input as path, file name,
    ...                 sheet name and position of the sheet in the workbook.

    [Arguments]    ${PATH}    ${BOOK_NAME}    ${SHEET_NAME}    ${POSTION}
    pyxl.Create New Workbook   ${PATH}   ${BOOK_NAME}    ${SHEET_NAME}    ${POSTION}

Close Workbook
    [Documentation]     The keywords to close the created workbook.
    pyxl.Close Workbook

Get Active Sheet
    [Documentation]      Get Active Sheet in the current workbook.
    ${value}    pyxl.Get Active Sheet
    [return]    ${value}

Get Active Sheet Name
    [Documentation]     Get Active Sheet Name in the current workbook.
    ${value}    pyxl.Active Sheet Name
    [return]    ${value}

Set Sheet Name
    [Documentation]
    [Arguments]     ${Title}
    ${sheet}    pyxl.Get Active Sheet
    Log     ${Title}
    ${sheet.title}  Set Variable    ${Title}

Open Workbook
    [Documentation]
    [Arguments]     ${Path}     ${Book Name}
    ${value}    pyxl.Load Workbook      ${Path}     ${Book Name}
    [return]    ${value}
