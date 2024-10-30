Set objAccess = CreateObject("Access.Application")
objAccess.OpenCurrentDatabase "C:\Users\rajas\Desktop\Database\filecompare.accdb"
objAccess.Run "ImportExcelFiles", "C:\Users\rajas\Desktop\Excelcompare\Lastweekfile.xlsx", "C:\Users\rajas\Desktop\Excelcompare\Thisweekfile.xlsx"
objAccess.Quit
Set objAccess = Nothing