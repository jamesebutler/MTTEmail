rem Deploy application changes to Production server
robocopy .\bin\Release\ C:\PublishPrograms\MTTEmail\ *MTTEmail.exe
robocopy .\bin\Release\ C:\PublishPrograms\MTTEmail\ *.dll
robocopy .\bin\Release\ C:\PublishPrograms\MTTEmail\ *.pdb
robocopy .\bin\Release\ C:\PublishPrograms\MTTEmail\ *.config
pause