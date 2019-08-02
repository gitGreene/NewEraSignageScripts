dim fso, theFolder, folderName, oldFileName, newFileName, pathName

set fso = CreateObject("Scripting.FileSystemObject")
'ENTER FOLDER NAME HERE:
set theFolder = fso.GetFolder("")
folderName = theFolder.name
slideNumber = 1

inputPrefix = InputBox("Enter Prefix:", "Input prefix")

inputStartDate = InputBox("Enter date in MM/DD/YYYY Format", "Input start date")
startDay = WeekDayName(DatePart("w", inputStartDate), true)
startMonth = MonthName(DatePart("m", inputStartDate), true)
startDate = DatePart("d", inputStartDate)

inputEndDate = InputBox("Enter date in MM/DD/YYYY Format", "Input end date")

endDay = WeekDayName(DatePart("w", inputEndDate), true)
endMonth = MonthName(DatePart("m", inputEndDate), true)
endDate = DatePart("d", inputEndDate)
 
if IsDate(inputStartDate)=true AND IsDate(inputEndDate)=true then 
  inputStartDate = CDate(inputStartDate)
  inputEndDate = CDate(inputEndDate)
  renameFiles()
else
  msgbox("Incorrect Date Format")
End If



Function renameFiles()

For Each File in theFolder.Files

if StrComp(File.Name, "MiscRenameScript.vbs")<>0 Then
newFileName = inputPrefix & " " & startDay & " " & startMonth & " " & startDate & " - " & endDay & " " & endMonth & " " & endDate & " Slide " & slideNumber
File.Name = Replace(File.Name, File.Name, newFileName)
slideNumber = slideNumber + 1
End If

Next
End Function