dim ObjFSO , Log, objLog
Set objFSO = CreateObject("Scripting.FileSystemObject")
Const ForReading = 1, ForWriting = 2, ForAppending = 8


if objFSO.FileExists("d:\my.log") then
  Set objLog = objFSO.OpenTextFile("d:\my.log", ForAppending )
else
 Set objLog = objFSO.CreateTextFile("d:\my.log")  
end if
objLog.WriteLine "Whatever output you want.2"
'objLog.WriteLine "You can even add " & strVariables & " if you want"


'Set Log = objFSO.OpenTextFile("d:\my.log", For_Writing, True)
'Log.WriteLine "jhg"

objLog.Close