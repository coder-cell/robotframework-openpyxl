*** Settings ***
Documentation    Test Open Pyxl Robot file Wrapper
Resource    scripts/open_pyxl.robot

*** Test Cases ***
Test Create Workbook
    Create Workbook   output     TestBook  TestSheet   ${0}
    ${Sheet Name}   Active Sheet