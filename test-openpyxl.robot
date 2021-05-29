*** Settings ***
Documentation    Test Open Pyxl Robot file Wrapper
Library    open-pyxl.robot

*** Test Cases ***
Create Workbook
    Create Workbook   TestBook  TestSheet   ${0}
