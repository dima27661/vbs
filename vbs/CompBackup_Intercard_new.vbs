Option Explicit
Dim fso, folder, subFil, fld, objRegExp, bpath, apath, zpath, cpath, i, t 
Dim sd, ed, bb, nn, Max_, Min, tz, service, Process, s, wsh, cc(), dl(), dn, sReturn
Dim expired_day_count, objLog, objBatchFile , log_path, batch_path
Const ForReading = 1, ForWriting = 2, ForAppending = 8

Min = "20180101"
Max_ = "20180607"



bpath = "D:\Backup\Intercard\DB\"
apath = "D:\Backup\Intercard\DB_Arch\"
cpath = "D:\Temp\DB_Arch\"
zpath = "C:\Progra~1\7-Zip\7z.exe"
'zpath = "C:\Program Files (x86)\WinRAR\WinRAR.exe"
log_path = "d:\logFile.log"
batch_path = "d:\bat_del.bat"

expired_day_count = 3

Function Fdaymonth(dm)
    If(Len(dm)=1) Then
        Fdaymonth="0"&dm
    Else
        Fdaymonth=dm
    End If
End Function

Function ClearArhive(day_count , FolderPath )
Dim fso_del, folder_del, subFil_del, fld_del
Dim FileDate 

Set fso_del = CreateObject("Scripting.FileSystemObject")
Set folder_del = fso_del.GetFolder(FolderPath)
Set subFil_del = folder_del.Files

   objLog.WriteLine date & " " & time() & " ClearArhive begin scan " & FolderPath 
              

    For Each fld_del In subFil_del
        FileDate = fld_del.DateLastModified
     objLog.WriteLine date & " " & time() & " check file date " & fso_del.GetAbsolutePathName(fld_del) & ", filedate = " & FileDate   
'          fso_del.DeleteFile (fso_del.GetAbsolutePathName(fld_del))
        If (FileDate < (Date - day_count)) Then
     objLog.WriteLine date & " " & time() & " delete " & fso_del.GetAbsolutePathName(fld_del)
     objBatchFile.WriteLine "del " & chr(34) & fso_del.GetAbsolutePathName(fld_del) & chr(34)
          fso_del.DeleteFile (fso_del.GetAbsolutePathName(fld_del))
        End If
    Next
      Set subFil_del = Nothing: Set folder_del = Nothing: Set fso_del = Nothing

End Function

nn = 0
ReDim cc(0)
Set fso = CreateObject("Scripting.FileSystemObject")

if fso.FileExists(batch_path) then
  Set objBatchFile = fso.OpenTextFile(batch_path, ForWriting)
else
  Set objBatchFile = fso.CreateTextFile(batch_path)  
end if

if fso.FileExists(log_path) then
  Set objLog = fso.OpenTextFile(log_path, ForAppending )
else
 Set objLog = fso.CreateTextFile(log_path)  
end if




objLog.WriteLine  date & " " & time() & " Script run"



Set folder = fso.GetFolder(bpath)
Set subFil = folder.files
	For Each fld in subFil
		bb=Instr(1,fld.name,201)
		If (bb <> 0) Then
			ReDim Preserve cc(nn)
			cc(nn) = (Mid(fld.name,bb,4) & Mid(fld.name,(bb+5),2) & Mid(fld.name,bb+8,2))
			If Max_ < cc(nn) Then Max_ = cc(nn)
			If Min > cc(nn) Then Min = cc(nn)
			nn=nn+1
		End If
	Next
sd = CDate(Mid(Min,1,4) & "-" & Mid(Min,5,2) & "-" & Mid(Min,7,2))
ed = CDate(Mid(Max_,1,4) & "-" & Mid(Max_,5,2) & "-" & Mid(Max_,7,2))
Do
	ReDim dl(1)
	dn = 0
	tz = True
	set service = GetObject ("winmgmts:")
	for each Process in Service.InstancesOf ("Win32_Process")
		If Process.Name = "7z.exe" then
			WScript.Sleep 1000
			tz = False
		End If
	Next
	If tz Then
		Set objRegExp = CreateObject("VBScript.RegExp")
		t = Year(sd) & "_" & Fdaymonth(Month(sd)) & "_" & Fdaymonth(Day(sd))
		objRegExp.Pattern = t
		For Each fld in subFil
			ReDim Preserve dl(dn)
			if objRegExp.Test(fld.name) Then
				s = s & Chr(34) & bpath & fld.name & Chr(34) & Chr(32)
				dl(dn) = bpath & fld.name

                objLog.WriteLine date & " " & time() & " база для архива " & dl(dn)

				dn = dn + 1
			End If
		Next	
		If ( s <> "" ) Then
			Set wsh = WScript.CreateObject("WScript.Shell")
			sReturn = wsh.Run(zpath & " a -t7z -ssw -mx9 " & apath & "Intercard_" & t & ".7z " & s, 3, TRUE)
            objLog.WriteLine date & " " & time() & " " & zpath & " a -t7z -ssw -mx9 " & apath & "Intercard_" & t & ".7z " & s  
			Set wsh = Nothing
			If (sReturn = 0) Then
'                On Error GoTo 0
				fso.CopyFile apath & "Intercard_" & t & ".7z", cpath, TRUE
                 objLog.WriteLine date & " " & time() & " CopyFile " & apath & "Intercard_" & t & ".7z" & cpath
				For i=0 to (dn-1)
                    objLog.WriteLine date & " " & time() & " DeleteFile " & dl(i)
					fso.DeleteFile dl(i)
				Next
                 objLog.WriteLine date & " " & time() & " ClearArhive expired_day_count = " & expired_day_count & ", cpath = " & cpath
                ClearArhive expired_day_count, cpath
'            0:
'              On Error Resume Next
			End If
		End If
		s=""
		sd=sd+1
	End If
Loop Until (sd = (ed+1))  
objBatchFile.Close 
objLog.Close

