Option Explicit
Dim fso, folder, subFil, fld, objRegExp, bpath, apath, zpath, cpath, i, t
Dim sd, ed, bb, nn, Max, Min, tz, service, Process, s, wsh, cc(), dl(), dn, fldel, fsd, fsc
Dim expired_day_count As Integer


Min = "20150101"
Max = "20130101"

expired_day_count = 3

bpath = "D:\1C_Backup\DB\"
apath = "D:\1C_Backup\DB_Arch\"
cpath = "E:\Backup_1C\DB_Arch\"
zpath = "C:\Progra~1\7-Zip\7z.exe"



Function Fdaymonth(dm)
    If (Len(dm) = 1) Then
        Fdaymonth = "0" & dm
    Else
        Fdaymonth = dm
    End If
End Function


Function ClearArhive(day_count As Integer, FolderPath As String)
Dim fso_del, folder_del, subFil_del, fld_del
Dim FileDate As Date

Set fso_del = CreateObject("Scripting.FileSystemObject")
Set folder_del = fso_del.GetFolder(FolderPath)
Set subFil_del = folder_del.Files
    For Each fld_del In subFil_del
        FileDate = fld_del.DateLastModified
          fso_del.DeleteFile (fso_del.GetAbsolutePathName(fld_del))
        If (FileDate < (Date - day_count)) Then
          fso_del.DeleteFile (fso_del.GetAbsolutePathName(fld_del))
        End If
    Next
      Set subFil_del = Nothing: Set folder_del = Nothing: Set fso_del = Nothing

End Function







ReDim cc(0)
Set fso = CreateObject("Scripting.FileSystemObject")
Set folder = fso.GetFolder(bpath)
Set subFil = folder.Files
    For Each fld In subFil
        bb = InStr(1, fld.Name, 201)
        If (bb <> 0) Then
            ReDim Preserve cc(nn)
            cc(nn) = (Mid(fld.Name, bb, 4) & Mid(fld.Name, (bb + 5), 2) & Mid(fld.Name, bb + 8, 2))
            If Max < cc(nn) Then Max = cc(nn)
            If Min > cc(nn) Then Min = cc(nn)
            nn = nn + 1
        End If
    Next
sd = CDate(Mid(Min, 1, 4) & "-" & Mid(Min, 5, 2) & "-" & Mid(Min, 7, 2))
ed = CDate(Mid(Max, 1, 4) & "-" & Mid(Max, 5, 2) & "-" & Mid(Max, 7, 2))
Do
    ReDim dl(0)
    dn = 0
    tz = True
    Set service = GetObject("winmgmts:")
    For Each Process In service.InstancesOf("Win32_Process")
        If Process.Name = "7z.exe" Then
          WScript.Sleep 1000

            tz = False
        End If
    Next
    
    
    If tz Then
        Set objRegExp = CreateObject("VBScript.RegExp")
        t = Year(sd) & "_" & Fdaymonth(Month(sd)) & "_" & Fdaymonth(Day(sd))
        objRegExp.Pattern = t
        For Each fld In subFil
            ReDim Preserve dl(dn)
            If objRegExp.Test(fld.Name) Then
                s = s & Chr(34) & bpath & fld.Name & Chr(34) & Chr(32)
                dl(dn) = bpath & fld.Name
                dn = dn + 1
            End If
        Next
        If (s <> "") Then
            Set wsh = WScript.CreateObject("WScript.Shell")
            Dim sReturn
            sReturn = wsh.Run(zpath & " a -t7z -ssw -mx9 " & apath & "1C_" & t & ".7z " & s, 3, True)

            
            Set wsh = Nothing
            If (sReturn = 0) Then
                On Error GoTo 0
                fso.CopyFile apath & "1C_" & t & ".7z", cpath, True
                For i = 0 To (dn - 1)
                    fso.DeleteFile (dl(i))
                Next
                ClearArhive expired_day_count, cpath
0:
On Error Resume Next
            End If
        End If
        s = ""
        sd = sd + 1
    End If
Loop Until (sd = (ed + 1))


