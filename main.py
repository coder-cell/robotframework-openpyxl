import robot

status = robot.run("tests/test-openpyxl.robot", outputdir="output", exclude=['skip'])
exit(status)
