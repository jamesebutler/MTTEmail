rem Deploy application changes to Production server
robocopy .\bin\Release\ \\GPIAZMESWEBP01\C$\International Paper\MTT\MTTEmail\ *MTTEmail.exe
robocopy .\bin\Release\ \\GPIAZMESWEBP01\C$\International Paper\MTT\MTTEmail\ *.dll
robocopy .\bin\Release\ \\GPIAZMESWEBP01\C$\International Paper\MTT\MTTEmail\ *.pdb
robocopy .\bin\Release\ \\GPIAZMESWEBP01\C$\International Paper\MTT\MTTEmail\ *.config
pause
