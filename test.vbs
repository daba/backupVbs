
Option Explicit

Dim strToday
Dim saveFile 
Dim saveDir
Dim objFso

saveFile = "C:\Documents and Settings\Administrator\work\kanri-table.xls"
saveDir = "C:\Documents and Settings\Administrator\work\backup"
strToday = Year(Now()) & Right("0"&Month(Now),2)  & Right("0"&Day(Now),2)

MsgBox(strToday)
MsgBox(saveDir & "kanri-table" & strToday & ".xls")

Set objFso = CreateObject("Scripting.FileSystemObject")

if

then

else
objFso.CopyFile saveFile, saveDir & "\kanri-table" & strToday & ".xls"


Set objFso  = Nothing