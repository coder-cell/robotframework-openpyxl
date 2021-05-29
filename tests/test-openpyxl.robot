*** Settings ***
Documentation    Test Open Pyxl Robot file Wrapper
Resource    scripts/open_pyxl.robot

*** Test Cases ***
Test Create Workbook
    Create Workbook   output    TestBook  TestSheet   ${0}
    Get Active Sheet
    Get Active Sheet Name
    Set Sheet Name   Introduction
    ${sheetname}    Get Active Sheet Name
    Close Workbook

Test Load Workbook
    Open Workbook   output  TestBook.xlsx
    Get Active Sheet Name
    Set Sheet Name   Introduction_Modified
    ${sheetname}    Get Active Sheet Name
    Close Workbook

