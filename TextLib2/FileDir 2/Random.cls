VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Random"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public Enum W32F_Errors
   W32F_UNKNOWN_ERROR = 45600
   W32F_FILE_ALREADY_OPEN
   W32F_PROBLEM_OPENING_FILE
   W32F_FILE_ALREADY_CLOSED
   W32F_Problem_seeking
End Enum

Private Const W32F_SOURCE = "Win32File Object"

Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_READ = &H80000000
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_FLAG_RANDOM_ACCESS = &H10000000
Private Const CREATE_ALWAYS = 2
Private Const OPEN_ALWAYS = 4
Private Const INVALID_HANDLE_VALUE = -1

Private Const FILE_BEGIN = 0
Private Const FILE_CURRENT = 1
Private Const FILE_END = 2

Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Any) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32.dll" (ByVal hFile As Long) As Long

Private hFile As Long, sFName As String, fAutoFlush As Boolean

Public Property Get FileHandle() As Long
If hFile = INVALID_HANDLE_VALUE Then
   RaiseError W32F_FILE_ALREADY_CLOSED
End If

FileHandle = hFile
End Property

Public Property Get FileName() As String
If hFile = INVALID_HANDLE_VALUE Then
   RaiseError W32F_FILE_ALREADY_CLOSED
End If

FileName = sFName
End Property

Public Property Get IsOpen() As Boolean
IsOpen = hFile <> INVALID_HANDLE_VALUE
End Property

Public Property Get AutoFlush() As Boolean
If hFile = INVALID_HANDLE_VALUE Then
   RaiseError W32F_FILE_ALREADY_CLOSED
End If

AutoFlush = fAutoFlush
End Property

Public Property Let AutoFlush(ByVal NewVal As Boolean)
If hFile = INVALID_HANDLE_VALUE Then
   RaiseError W32F_FILE_ALREADY_CLOSED
End If

fAutoFlush = NewVal
End Property

Public Sub OpenFile(ByVal sFileName As String)
If hFile <> INVALID_HANDLE_VALUE Then
   RaiseError W32F_FILE_ALREADY_OPEN, sFName
End If

hFile = CreateFile(sFileName, GENERIC_WRITE Or GENERIC_READ, 0, _
  0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL Or FILE_FLAG_RANDOM_ACCESS, 0)

If hFile = INVALID_HANDLE_VALUE Then
   RaiseError W32F_PROBLEM_OPENING_FILE, sFileName
End If
sFName = sFileName
End Sub

Public Sub CloseFile()
If hFile = INVALID_HANDLE_VALUE Then
   RaiseError W32F_FILE_ALREADY_CLOSED
End If

CloseHandle hFile
sFName = ""
fAutoFlush = False
hFile = INVALID_HANDLE_VALUE
End Sub

Public Function ReadBytes(ByVal ByteCount As Long) As Variant
Dim BytesRead As Long, Bytes() As Byte
If hFile = INVALID_HANDLE_VALUE Then
   RaiseError W32F_FILE_ALREADY_CLOSED
End If

ReDim Bytes(0 To ByteCount - 1) As Byte
ReadFile hFile, Bytes(0), ByteCount, BytesRead, 0
ReadBytes = Bytes
End Function

Public Sub WriteBytes(DataBytes() As Byte)
Dim fSuccess As Long, BytesToWrite As Long, BytesWritten As Long
If hFile = INVALID_HANDLE_VALUE Then
   RaiseError W32F_FILE_ALREADY_CLOSED
End If

BytesToWrite = UBound(DataBytes) - LBound(DataBytes) + 1
fSuccess = WriteFile(hFile, DataBytes(LBound(DataBytes)), _
BytesToWrite, BytesWritten, 0)
If fAutoFlush Then Flush
End Sub

Public Sub Flush()
If hFile = INVALID_HANDLE_VALUE Then
   RaiseError W32F_FILE_ALREADY_CLOSED
End If

FlushFileBuffers hFile
End Sub

Public Sub SeekAbsoluteOld(ByVal HighPos As Long, ByVal LowPos As Long)
If hFile = INVALID_HANDLE_VALUE Then
   RaiseError W32F_FILE_ALREADY_CLOSED
End If

LowPos = SetFilePointer(hFile, LowPos, HighPos, FILE_BEGIN)
End Sub

Public Sub SeekAbsolute(NewPos As Double)
Dim FilH As Long, FilL As Long

If hFile = INVALID_HANDLE_VALUE Then
   RaiseError W32F_FILE_ALREADY_CLOSED
End If

FilH = GetHighPart(NewPos)
FilL = GetLowPart(NewPos)
SetFilePointer hFile, FilL, FilH, FILE_BEGIN
End Sub

Public Sub SeekRelative(ByVal Offset As Long)
Dim TempLow As Long, TempErr As Long
If hFile = INVALID_HANDLE_VALUE Then
   RaiseError W32F_FILE_ALREADY_CLOSED
End If

TempLow = SetFilePointer(hFile, Offset, ByVal 0&, FILE_CURRENT)
If TempLow = -1 Then
   TempErr = Err.LastDllError
   If TempErr Then
      RaiseError W32F_Problem_seeking, "Error " & TempErr & "." & _
      vbCrLf & CStr(TempErr)
   End If
End If

End Sub

Private Sub Class_Initialize()
hFile = INVALID_HANDLE_VALUE
End Sub

Private Sub Class_Terminate()
If hFile <> INVALID_HANDLE_VALUE Then CloseHandle hFile
End Sub

Private Sub RaiseError(ByVal ErrorCode As W32F_Errors, Optional sExtra)
Dim Win32Err As Long, Win32Text As String
Win32Err = Err.LastDllError
If Win32Err Then
   Win32Text = vbCrLf & "Error " & Win32Err & vbCrLf & _
   DecodeAPIErrors(Win32Err)
End If

Select Case ErrorCode
Case W32F_FILE_ALREADY_OPEN
   Err.Raise W32F_FILE_ALREADY_OPEN, W32F_SOURCE, _
   "The file '" & sExtra & "' is already open." & Win32Text
Case W32F_PROBLEM_OPENING_FILE
   Err.Raise W32F_PROBLEM_OPENING_FILE, W32F_SOURCE, _
   "Error opening '" & sExtra & "'." & Win32Text
Case W32F_FILE_ALREADY_CLOSED
   Err.Raise W32F_FILE_ALREADY_CLOSED, W32F_SOURCE, _
   "There is no open file."
Case W32F_Problem_seeking
   Err.Raise W32F_Problem_seeking, W32F_SOURCE, _
   "Seek Error." & vbCrLf & sExtra
Case Else
   Err.Raise W32F_UNKNOWN_ERROR, W32F_SOURCE, _
   "Unknown error." & Win32Text
End Select

End Sub

Private Function DecodeAPIErrors(ByVal ErrorCode As Long) As String
Dim sMessage As String, MessageLength As Long
sMessage = Space$(256)
MessageLength = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0&, _
ErrorCode, 0&, sMessage, 256&, 0&)
If MessageLength > 0 Then
   DecodeAPIErrors = Left(sMessage, MessageLength)
Else
   DecodeAPIErrors = "Unknown Error."
End If

End Function

Public Function GetCurrentPos() As Double
Dim FilH As Long, FilL As Long

If hFile = INVALID_HANDLE_VALUE Then
   RaiseError W32F_FILE_ALREADY_CLOSED
End If

FilH = 0
FilL = SetFilePointer(hFile, 0, FilH, 1)
GetCurrentPos = DoubleFromLInt(FilH, FilL)
End Function

Public Function GetHighPart(Inp As Double) As Long
GetHighPart = Int(Inp / 2 ^ 32)
End Function

Public Function GetLowPart(Inp As Double) As Long
GetLowPart = Inp - Int(Inp / 2 ^ 32) * 2 ^ 32
End Function

Public Function DoubleFromLInt(HighPart As Long, LowPart As Long) As Double
DoubleFromLInt = HighPart * 2 ^ 32 + LowPart
End Function

Public Function GetFileLen() As Double
Dim FilH As Long, FilL As Long

If hFile = INVALID_HANDLE_VALUE Then
   RaiseError W32F_FILE_ALREADY_CLOSED
End If

FilL = GetFileSize(hFile, FilH)
GetFileLen = DoubleFromLInt(FilH, FilL)

End Function

Public Sub SetEOFHere()
If hFile = INVALID_HANDLE_VALUE Then
   RaiseError W32F_FILE_ALREADY_CLOSED
End If

SetEndOfFile hFile
End Sub

Public Sub SetEOFToPos(NewEOF As Double)
Dim SavedPos As Double
If hFile = INVALID_HANDLE_VALUE Then
   RaiseError W32F_FILE_ALREADY_CLOSED
End If

SavedPos = GetCurrentPos
SeekAbsolute NewEOF
SetEOFHere
SeekAbsolute SavedPos

End Sub
