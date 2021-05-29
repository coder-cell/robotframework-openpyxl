*** Settings ***
Documentation   Open PyXL Keywords
Library     scripts/OpenRobotPyxl.py

*** Keywords ***
Create Workbook
    [Arguments]    ${PATH}    ${BOOK_NAME}    ${SHEET_NAME}    ${POSTION}
    Create New Workbook   ${PATH}   ${BOOK_NAME}    ${SHEET_NAME}    ${POSTION}

Active Sheet
    [return]    Get Active Sheet