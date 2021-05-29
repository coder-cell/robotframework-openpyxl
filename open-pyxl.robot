*** Settings ***
Documentation   Open PyXL Keywords
Library    scripts/wrapper.py

*** Keywords ***
Create Workbook
    [Arguments]    ${BOOK_NAME}    ${SHEET_NAME}    ${POSTION}
       Create New Workbook    ${BOOK_NAME}    ${SHEET_NAME}    ${POSTION}

