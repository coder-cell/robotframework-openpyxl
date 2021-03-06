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

Add Sheet
    [Documentation]
    [Arguments]     ${Sheet Name}   ${Position}=${0}
    pyxl.Add Sheet  ${Sheet Name}   ${Position}

Open Workbook
    [Documentation]
    [Arguments]     ${Path}     ${Book Name}
    ${value}    pyxl.Load Workbook      ${Path}     ${Book Name}
    [return]    ${value}

Set Cell Value
    [Documentation]
    [Arguments]     ${row}  ${col}  ${value}
    pyxl.Set Cell Value  ${row}  ${col}  ${value}

Get Cell Value
    [Documentation]
    [Arguments]     ${row}  ${col}
    ${value}    pyxl.Get Cell Value     ${row}  ${col}
    [return]    ${value}

Insert Row
    [Documentation]
    [Arguments]     ${row number}
    ${value}    pyxl.Insert Row    ${row number}
    [return]    ${value}


Insert Column
    [Documentation]
    [Arguments]     ${column number}
    ${value}    pyxl.Insert Column    ${column number}
    [return]    ${value}


Delete Row
    [Documentation]
    [Arguments]     ${row number}
    ${value}    pyxl.Delete Row    ${row number}
    [return]    ${value}


Delete Column
    [Documentation]
    [Arguments]     ${column number}
    ${value}    pyxl.Delete Column    ${column number}
    [return]    ${value}

Insert List To Row
    [Documentation]
    [Arguments]     ${Row Number}   ${Column Number}    ${List Data}
    ${value}    pyxl.Convert List to Row    ${Row Number}   ${Column Number}    ${List Data}
    [return]    ${value}


Insert List To Column
    [Documentation]
    [Arguments]     ${Row Number}   ${Column Number}    ${List Data}
    ${value}    pyxl.Convert List to Column    ${Row Number}   ${Column Number}    ${List Data}
    [return]    ${value}