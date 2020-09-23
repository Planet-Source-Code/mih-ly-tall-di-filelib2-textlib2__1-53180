Attribute VB_Name = "Globals"
Public Const BUFF_LEN = 50000
Public Entrys() As String
Public CurrEntrys() As Long
Public WordWarp As Byte
Public TextPass As String
Public NamePass As String
Public StatePass As String

Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long

Public Function DelFile(FileName As String) As Boolean
SetFileAttributes FileName, 0
If DeleteFile(FileName) = 0 Then DelFile = False Else DelFile = True
End Function

Public Function IsValidFileName(ByVal FNam As String, Optional DevChk As Boolean = False) As Boolean
Dim fnCount As Byte, DevNames(0 To 22) As String
IsValidFileName = True

If (Len(FNam) > 255) Or (Len(FNam) = 0) Or (FNam = Null) Then IsValidFileName = False: Exit Function
For fnCount = 0 To 31
   If InStr(1, FNam, Chr(fnCount)) <> 0 Then IsValidFileName = False: Exit Function
Next

If InStr(1, FNam, "\") <> 0 Then IsValidFileName = False: Exit Function
If InStr(1, FNam, "/") <> 0 Then IsValidFileName = False: Exit Function
'< > : " / \ |
If InStr(1, FNam, "<") <> 0 Then IsValidFileName = False: Exit Function
If InStr(1, FNam, ">") <> 0 Then IsValidFileName = False: Exit Function
If InStr(1, FNam, ":") <> 0 Then IsValidFileName = False: Exit Function
If InStr(1, FNam, "|") <> 0 Then IsValidFileName = False: Exit Function
If InStr(1, FNam, Chr(34)) <> 0 Then IsValidFileName = False: Exit Function

If DevChk = False Then Exit Function
DevNames(0) = "CON"
DevNames(1) = "PRN"
DevNames(2) = "AUX"
DevNames(3) = "CLOCK$"
DevNames(4) = "NUL"
For fnCount = 1 To 9
   DevNames(4 + fnCount) = "COM" & fnCount
Next
For fnCount = 1 To 9
   DevNames(13 + fnCount) = "LPT" & fnCount
Next

For fnCount = 0 To 22
   If InStr(1, FNam, DevNames(fnCount), vbTextCompare) <> 0 Then IsValidFileName = False: Exit Function
   If InStr(1, FNam, DevNames(fnCount) & ".", vbTextCompare) <> 0 Then IsValidFileName = False: Exit Function
Next

End Function
