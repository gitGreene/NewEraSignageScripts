dim fso, folder, folderName, dateInput, day, month, date, oldFileName, newFileName, pathName

set fso = CreateObject("Scripting.FileSystemObject")
set folder = fso.GetFolder("C:\Users\MITCH.GREENE\OneDrive - New Era Technology\Desktop\ScriptingTests\MenuRename")
slideNumber = 1

inputPrefix = InputBox("Enter Prefix:", "Input prefix")
inputDate = InputBox("Enter start date in MM/DD/YYYY Format", "User date")
 
if IsDate(inputDate)=true then 
  inputDate = CDate(inputDate)
  renameFiles()
else
  msgbox("Incorrect Date Format")
End If


Function renameFiles()

For Each File in Folder.Files

oldFileName = File.Name

day = WeekDayName(DatePart("w", inputDate), true)
month = MonthName(DatePart("m", inputDate), true)
date = DatePart("d", inputDate)

newFileName = inputPrefix & " " & day & " " & month & " " & date & " Slide " & slideNumber

Select case File.Name
  Case "Slide1.JPG"
    File.Name = Replace(File.Name, "Slide1", newFileName)
    addDay()
  Case "Slide2.JPG"
    File.Name = Replace(File.Name, "Slide2", newFileName)
    addDay()
  Case "Slide3.JPG"
    File.Name = Replace(File.Name, "Slide3", newFileName)
    addDay()
  Case "Slide4.JPG"
    File.Name = Replace(File.Name, "Slide4", newFileName)
    addDay()
  Case "Slide5.JPG"
    File.Name = Replace(File.Name, "Slide5", newFileName)
    addDay()
End Select

Next
End Function


Function addDay()

inputDate = DateAdd("d", 1, inputDate)
slideNumber = slideNumber + 1

End Function