Option Explicit
Dim fso, folder, subFil, fld, objRegExp, bpath, apath, zpath, cpath, i, t
Dim sd, ed, bb, nn, Max, Min, tz, service, Process, s, wsh, cc(), dl(), dn, fldel, fsd, fsc
Dim expired_day_count
Min = "20180101"
Max = "20170101"

'bpath = "\\192.168.102.108\Data\"
bpath = "D:\temp\from\"
'apath = "D:\BackUp_Gourmet\Data_Arh\"
apath = "D:\temp\to\"
cpath = "\\192.168.137.254\BackUp\BackUp_Gourmet\"
zpath = "C:\Progra~1\7-Zip\7z.exe"

expired_day_count = 2

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
    For Each fld_del In subFil_del
        FileDate = fld_del.DateLastModified
'          fso_del.DeleteFile (fso_del.GetAbsolutePathName(fld_del))
        If (FileDate < (Date - day_count)) Then
          fso_del.DeleteFile (fso_del.GetAbsolutePathName(fld_del))
        End If
    Next
      Set subFil_del = Nothing: Set folder_del = Nothing: Set fso_del = Nothing

End Function

nn = 0
ReDim cc(0)
Set fso = CreateObject("Scripting.FileSystemObject")
Set folder = fso.GetFolder(bpath)
Set subFil = folder.files
	For Each fld in subFil
		bb=Instr(1,fld.name,201)
		If (bb <> 0) Then
			ReDim Preserve cc(nn)
			cc(nn) = (Mid(fld.name,bb,4) & Mid(fld.name,(bb+5),2) & Mid(fld.name,bb+8,2))
			If Max < cc(nn) Then Max = cc(nn)
			If Min > cc(nn) Then Min = cc(nn)
			nn=nn+1
		End If
	Next
sd = CDate(Mid(Min,1,4) & "-" & Mid(Min,5,2) & "-" & Mid(Min,7,2))
ed = CDate(Mid(Max,1,4) & "-" & Mid(Max,5,2) & "-" & Mid(Max,7,2))
Do
	ReDim dl(0)
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
'			if objRegExp.Test(fld.name) Then
				s = s & Chr(34) & bpath & fld.name & Chr(34) & Chr(32)
				dl(dn) = bpath & fld.name
				dn = dn + 1
'			End If
		Next
		If ( s <> "" ) Then
			Set wsh = WScript.CreateObject("WScript.Shell")
			Dim sReturn
			sReturn = wsh.Run(zpath & " a -t7z -ssw -mx9 " & apath & "Gourmet_" & t & ".7z " & s, 3, TRUE)
			Set wsh = Nothing
			If (sReturn = 0) Then
				fso.CopyFile apath & "Gourmet_" & t & ".7z", cpath, TRUE
				For i=0 to (dn-1)
					fso.DeleteFile(dl(i))
				Next
	ClearArhive expired_day_count, apath                
	ClearArhive expired_day_count, cpath
'            0:
'              On Error Resume Next
			End If
		End If
		s=""
		sd=sd+1
	End If
Loop Until (sd = (ed+1))   
