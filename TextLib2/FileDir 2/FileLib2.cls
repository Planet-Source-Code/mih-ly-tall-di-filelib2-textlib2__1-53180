VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileLib2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarFileHeader As String
Private mvarLibName As String
Private mvarBufferLength As Long

Private Type Entry
   EntryName As String
   EntryPos As Double
   EntryLen As Double
End Type

Public Enum FlagsTypes
   User0 = 1
   User1 = 2
   User2 = 4
   User3 = 8
   User4 = 16
   User5 = 32
   User6 = 64
   User7 = 128
End Enum

Private Buffer As String
Private Entrys() As Entry, TableStart As Double, IsAny As Byte
Private Flags As Byte

Public Property Let BufferLength(ByVal vData As Long)
Attribute BufferLength.VB_Description = "Returns/sets the length of the read-write buffer."
    mvarBufferLength = vData
End Property

Public Property Get BufferLength() As Long
    BufferLength = mvarBufferLength
End Property

Public Property Get LibName() As String
Attribute LibName.VB_Description = "Returns the opened library's name."
    LibName = mvarLibName
End Property

Public Property Let FileHeader(ByVal vData As String)
Attribute FileHeader.VB_Description = "A 3-byte length file identifier. This only applies on opening the library. If this differs from the library's header, the library will not be loaded. If you create a library when opening, it will have this header."
    mvarFileHeader = Left(vData, 3)
End Property

Public Property Get FileHeader() As String
    FileHeader = mvarFileHeader
End Property

Public Property Get NumEntrys() As Long
Attribute NumEntrys.VB_Description = "Returns the number of entrys contained in the library."
If IsAny = 0 Then
   NumEntrys = 0
Else
   NumEntrys = UBound(Entrys) + 1
End If
End Property


Public Function GetFlag(What As FlagsTypes) As Boolean
Attribute GetFlag.VB_Description = "Returns one of the library's flags."
If Flags And What Then GetFlag = True
End Function

Public Sub SetFlag(What As FlagsTypes, State As Boolean)
Attribute SetFlag.VB_Description = "Sets one of the flags of the library.(You can use these to anything, like a read-only flag, type flags, etc.)"
If mvarLibName = "" Then Exit Sub

Dim Lib As Random
Set Lib = New Random

If State = True Then
   Flags = Flags Or What
Else
   Flags = Flags And Not What
End If

Lib.OpenFile mvarLibName
Lib.SeekAbsolute 4
Lib.WriteBytes StrConv(DecTo256(Flags), vbFromUnicode)
Lib.CloseFile
Set Lib = Nothing

End Sub

Public Sub RenameEntry(ByVal EntryNum As Long, ByVal NewName As String)
Attribute RenameEntry.VB_Description = "Sets the name of the current entry, end updates the library with the new name."
If (EntryNum < 0) Or (EntryNum > UBound(Entrys)) Or (IsAny = 0) Then Exit Sub
If NewName = "" Then Exit Sub

Entrys(EntryNum).EntryName = NewName
WriteTable
End Sub

Public Function GetEntryName(ByVal EntryNum As Long) As String
Attribute GetEntryName.VB_Description = "Retrives the name associated with the given entry."
If (EntryNum < 0) Or (EntryNum > UBound(Entrys)) Or (IsAny = 0) Then Exit Function
GetEntryName = Entrys(EntryNum).EntryName
End Function

Public Function GetEntryLength(ByVal EntryNum As Long) As Double
Attribute GetEntryLength.VB_Description = "Retrives the give entry's size, in bytes."
If (EntryNum < 0) Or (EntryNum > UBound(Entrys)) Or (IsAny = 0) Then Exit Function
GetEntryLength = Entrys(EntryNum).EntryLen
End Function

Public Function ExtractEntryString(ByVal EntryNum As Long, Optional ByVal FromPos As Double = -1, Optional ByVal ToPos As Double = -1) As String
Attribute ExtractEntryString.VB_Description = "You can read an entry directly with this function. If FromPos is -1 or omitted, it will start from the entry's start. if ToPos is -1 or omitted, it will stop at the entry's end. If ToPos>Entry's length, it will stop at the entry's end."
If (EntryNum < 0) Or (EntryNum > UBound(Entrys)) Or (IsAny = 0) Then Exit Function

Dim Lib As Random
Set Lib = New Random
Lib.OpenFile mvarLibName

If FromPos = -1 Then FromPos = 0
If ToPos = -1 Then ToPos = Entrys(EntryNum).EntryLen
If ToPos > Entrys(EntryNum).EntryLen Then ToPos = Entrys(EntryNum).EntryLen
ToPos = ToPos - FromPos

Lib.SeekAbsolute Entrys(EntryNum).EntryPos + FromPos

ExtractEntryString = ""
Do
   If Lib.GetCurrentPos + mvarBufferLength < Entrys(EntryNum).EntryPos + ToPos + FromPos Then
      Buffer = StrConv(Lib.ReadBytes(mvarBufferLength), vbUnicode)
      ExtractEntryString = ExtractEntryString & Buffer
   Else
      Buffer = StrConv(Lib.ReadBytes((Entrys(EntryNum).EntryPos + ToPos + FromPos) - Lib.GetCurrentPos), vbUnicode)
      ExtractEntryString = ExtractEntryString & Buffer
      Exit Do
   End If
Loop
Lib.CloseFile
Set Lib = Nothing
Buffer = ""

End Function

Public Sub ExtractEntry(ByVal EntryNum As Long, ByVal OriginalName As Boolean, Optional ExtractPath As String = "", Optional ByVal TargetFileName As String = "")
Attribute ExtractEntry.VB_Description = "Extracts the given entry to ExtractPath (if OriginalName=True; useful when extracting entrys with names that contain directory structures), or to TargetFileName (if OriginalName=False)."
If (EntryNum < 0) Or (EntryNum > UBound(Entrys)) Or (IsAny = 0) Then Exit Sub

Dim Lib As Random, Target As Random
Set Lib = New Random: Set Target = New Random
Lib.OpenFile mvarLibName

If OriginalName = True Then
   If ExtractPath = "" Then
      Set Lib = Nothing
      Set Target = Nothing
      Exit Sub
   End If
   If Right(ExtractPath, 1) <> "\" Then ExtractPath = ExtractPath & "\"
   CreatePath ExtractPath & Mid(Entrys(EntryNum).EntryName, 1, InStrRev(Entrys(EntryNum).EntryName, "\"))
   Target.OpenFile ExtractPath & Entrys(EntryNum).EntryName
Else
   If TargetFileName = "" Then
      Set Lib = Nothing
      Set Target = Nothing
      Exit Sub
   End If
   Target.OpenFile TargetFileName
End If

Target.SeekAbsolute 0
Lib.SeekAbsolute Entrys(EntryNum).EntryPos
Do
   If Lib.GetCurrentPos + mvarBufferLength < Entrys(EntryNum).EntryPos + Entrys(EntryNum).EntryLen Then
      Buffer = StrConv(Lib.ReadBytes(mvarBufferLength), vbUnicode)
      Target.WriteBytes StrConv(Buffer, vbFromUnicode)
   Else
      Buffer = StrConv(Lib.ReadBytes((Entrys(EntryNum).EntryPos + Entrys(EntryNum).EntryLen) - Lib.GetCurrentPos), vbUnicode)
      Target.WriteBytes StrConv(Buffer, vbFromUnicode)
      Exit Do
   End If
Loop
Target.CloseFile
Lib.CloseFile
Set Lib = Nothing
Set Target = Nothing
Buffer = ""

End Sub

Public Sub AddStringEntry(ByVal EntryData As String, ByVal EntryName As String)
Attribute AddStringEntry.VB_Description = "Adds the EntryData string as an entry. It is wise to name the entry using the standard file naming rules."
Dim Lib As Random, StrPos As Long
Set Lib = New Random
Lib.OpenFile mvarLibName

If IsAny = 0 Then
   Entrys(0).EntryName = EntryName
   Entrys(0).EntryPos = TableStart
   Entrys(0).EntryLen = Len(EntryData)
   IsAny = 1
Else
   ReDim Preserve Entrys(UBound(Entrys) + 1)
   Entrys(UBound(Entrys)).EntryName = EntryName
   Entrys(UBound(Entrys)).EntryPos = TableStart
   Entrys(UBound(Entrys)).EntryLen = Len(EntryData)
End If
Lib.SeekAbsolute TableStart
Lib.SetEOFHere

StrPos = 1
Do
   If StrPos + mvarBufferLength < Len(EntryData) Then
      Buffer = Mid(EntryData, StrPos, mvarBufferLength)
      StrPos = StrPos + mvarBufferLength
      Lib.WriteBytes StrConv(Buffer, vbFromUnicode)
   Else
      Buffer = Mid(EntryData, StrPos, Len(EntryData) - StrPos + 1)
      Lib.WriteBytes StrConv(Buffer, vbFromUnicode)
      Exit Do
   End If
Loop
TableStart = TableStart + Len(EntryData)

Lib.CloseFile
Set Lib = Nothing
Buffer = ""

WriteTable
End Sub

Public Sub AddEntry(ByVal FileName As String, ByVal UseOrigFileName As Boolean, Optional ByVal RelativeRoot As String = "", Optional ByVal AlternativeName As String = "")
Attribute AddEntry.VB_Description = "If UseOrigFileName, you must give a RelativeRoot, else you must provide an AlternativeName."
If (FileName = "") Then Exit Sub
If (UseOrigFileName = False) And (AlternativeName = "") Then Exit Sub

Dim Lib As Random, Source As Random
Set Lib = New Random: Set Source = New Random
Lib.OpenFile mvarLibName
Source.OpenFile FileName
Source.SeekAbsolute 0

If UseOrigFileName = True Then
   If Right(RelativeRoot, 1) <> "\" Then RelativeRoot = RelativeRoot & "\"
   FileName = Mid(FileName, Len(RelativeRoot) + 1)
Else
   FileName = AlternativeName
End If

If IsAny = 0 Then
   Entrys(0).EntryName = FileName
   Entrys(0).EntryPos = TableStart
   Entrys(0).EntryLen = Source.GetFileLen
   IsAny = 1
Else
   ReDim Preserve Entrys(UBound(Entrys) + 1)
   Entrys(UBound(Entrys)).EntryName = FileName
   Entrys(UBound(Entrys)).EntryPos = TableStart
   Entrys(UBound(Entrys)).EntryLen = Source.GetFileLen
End If
Lib.SeekAbsolute TableStart
Lib.SetEOFHere

Do
   If Source.GetCurrentPos + mvarBufferLength < Source.GetFileLen Then
      Buffer = StrConv(Source.ReadBytes(mvarBufferLength), vbUnicode)
      Lib.WriteBytes StrConv(Buffer, vbFromUnicode)
   Else
      Buffer = StrConv(Source.ReadBytes(Source.GetFileLen - Source.GetCurrentPos), vbUnicode)
      Lib.WriteBytes StrConv(Buffer, vbFromUnicode)
      Exit Do
   End If
Loop
TableStart = TableStart + Source.GetFileLen

Source.CloseFile
Lib.CloseFile
Set Lib = Nothing
Set Source = Nothing
Buffer = ""

WriteTable
End Sub

Public Sub AddDirectory(Path As String, RelativeRoot As String, Optional FileSpec As String = "*.*", Optional Recurse As Boolean = True)
Attribute AddDirectory.VB_Description = "Adds the contents of a whole directory to the library.  Relative root:The directory that acts as a root for the files in the library. If empty, Path will be used for RelativeRoot. If RealtiveRoot contains only a Space character, then no realtive root will be used."
If Path = "" Then Exit Sub

If RelativeRoot = "" Then RelativeRoot = Path
If RelativeRoot = " " Then RelativeRoot = ""

Dim Lib As Random, Source As Random, Files As Collection, FileName As Variant
Set Lib = New Random: Set Source = New Random
Lib.OpenFile mvarLibName

Lib.SeekAbsolute TableStart
Lib.SetEOFHere

Set Files = GetAllFiles(Path, FileSpec, Recurse)

For Each FileName In Files
   Source.OpenFile FileName
   Source.SeekAbsolute 0
   
   If Right(RelativeRoot, 1) <> "\" Then RelativeRoot = RelativeRoot & "\"
   FileName = Mid(FileName, Len(RelativeRoot) + 1)
   
   If IsAny = 0 Then
      Entrys(0).EntryName = FileName
      Entrys(0).EntryPos = TableStart
      Entrys(0).EntryLen = Source.GetFileLen
      IsAny = 1
   Else
      ReDim Preserve Entrys(UBound(Entrys) + 1)
      Entrys(UBound(Entrys)).EntryName = FileName
      Entrys(UBound(Entrys)).EntryPos = TableStart
      Entrys(UBound(Entrys)).EntryLen = Source.GetFileLen
   End If
   
   Do
      If Source.GetCurrentPos + mvarBufferLength < Source.GetFileLen Then
         Buffer = StrConv(Source.ReadBytes(mvarBufferLength), vbUnicode)
         Lib.WriteBytes StrConv(Buffer, vbFromUnicode)
      Else
         Buffer = StrConv(Source.ReadBytes(Source.GetFileLen - Source.GetCurrentPos), vbUnicode)
         Lib.WriteBytes StrConv(Buffer, vbFromUnicode)
         Exit Do
      End If
   Loop
   
   TableStart = TableStart + Source.GetFileLen
   Source.CloseFile
   Buffer = ""
Next FileName

Set Source = Nothing
Lib.CloseFile
Set Lib = Nothing

WriteTable

End Sub

Public Sub DeleteEntrys(ByVal StartEntryNum As Long, ByVal EndEntryNum As Long, TempFileName As String)
Attribute DeleteEntrys.VB_Description = "Deletes entrys from StartEntryNum to EndEntryNum. Examples: Start=1, End=1 ->delete entry 1. Start=1, End=3 ->delete entrys 1,2,3. TempFileName: name of a temporary file to store library data during delete."
If (IsAny = 0) Or (StartEntryNum > UBound(Entrys)) Or (EndEntryNum > UBound(Entrys)) Or (StartEntryNum < 0) Or (EndEntryNum < 0) Or (StartEntryNum > EndEntryNum) Then Exit Sub

Dim Lib As Random, TempFile As Random
Set Lib = New Random: Set TempFile = New Random

Lib.OpenFile mvarLibName

If (EndEntryNum = UBound(Entrys)) Then
'Best case, we don't need the Tempfile
   Lib.SetEOFToPos Entrys(StartEntryNum).EntryPos
   
   If StartEntryNum > 0 Then
      ReDim Preserve Entrys(StartEntryNum - 1)
   Else
      ReDim Preserve Entrys(0)
      IsAny = 0
   End If
Else
   TempFile.OpenFile TempFileName
   TempFile.SeekAbsolute 0
   Lib.SeekAbsolute Entrys(EndEntryNum + 1).EntryPos
   Do
      If Lib.GetCurrentPos + mvarBufferLength < TableStart Then
         Buffer = StrConv(Lib.ReadBytes(mvarBufferLength), vbUnicode)
         TempFile.WriteBytes StrConv(Buffer, vbFromUnicode)
      Else
         Buffer = StrConv(Lib.ReadBytes(TableStart - Lib.GetCurrentPos), vbUnicode)
         TempFile.WriteBytes StrConv(Buffer, vbFromUnicode)
         Exit Do
      End If
   Loop
   Buffer = ""
   Lib.SeekAbsolute Entrys(StartEntryNum).EntryPos
   Lib.SetEOFHere
   
   TempFile.SeekAbsolute 0
   
   Do
      If TempFile.GetCurrentPos + mvarBufferLength < TempFile.GetFileLen Then
         Buffer = StrConv(TempFile.ReadBytes(mvarBufferLength), vbUnicode)
         Lib.WriteBytes StrConv(Buffer, vbFromUnicode)
      Else
         Buffer = StrConv(TempFile.ReadBytes(TempFile.GetFileLen - TempFile.GetCurrentPos), vbUnicode)
         Lib.WriteBytes StrConv(Buffer, vbFromUnicode)
         Exit Do
      End If
   Loop
   TempFile.CloseFile
   SetFileAttributes TempFileName, 0
   DeleteFile TempFileName
   
   Set TempFile = Nothing
   Buffer = ""
   
   For x = EndEntryNum + 1 To UBound(Entrys)
      Entrys(x - (EndEntryNum - StartEntryNum + 1)) = Entrys(x)
   Next x
   ReDim Preserve Entrys(UBound(Entrys) - (EndEntryNum - StartEntryNum + 1))
   
   Entrys(0).EntryPos = 8
   For x = 1 To UBound(Entrys)
      Entrys(x).EntryPos = Entrys(x - 1).EntryPos + Entrys(x - 1).EntryLen
   Next x
End If

Lib.CloseFile
Set Lib = Nothing

TableStart = CalcTablePos
WriteTable
End Sub

Public Function CloseLib()
Attribute CloseLib.VB_Description = "Closes the library and resets all internal variables."
ReDim Entrys(0)
IsAny = 0

Flags = 0
TableStart = 0
mvarFileHeader = "FL2"
mvarLibName = ""
mvarBufferLength = 50000
Buffer = ""

End Function

Public Function OpenLib(ByVal LibName As String, Optional CreateFlags As Byte = 0) As Byte
Attribute OpenLib.VB_Description = "Opens an FL2 file. Returns 1 if succeeded, 0 if not. If the library not exists, it will be created. If creation is successfull, the function returns 2."
Dim Lib As Random, Temp As String

Set Lib = New Random

If Dir(LibName) <> "" Then
'Library exists
   Lib.OpenFile LibName
   Temp = StrConv(Lib.ReadBytes(8), vbUnicode)
   
   If Left(Temp, 3) <> mvarFileHeader Then
      OpenLib = 0
      ReDim Entrys(0)
      IsAny = 0
      Set Lib = Nothing
      Exit Function
   End If
   
   Flags = BackToDec(Mid(Temp, 4, 1))
   
   mvarLibName = LibName
   TableStart = BackToDec(Right(Temp, 4))
   Lib.CloseFile
   Set Lib = Nothing
   ParseTable
   OpenLib = 1
Else
'Creating library...
   Lib.OpenFile LibName
   Flags = CreateFlags
   Temp = mvarFileHeader & DecTo256(Flags) & CompleteString(DecTo256(8), 4) & String(4, Chr(0))
   Lib.WriteBytes StrConv(Temp, vbFromUnicode)
   Lib.CloseFile
   Set Lib = Nothing
   mvarLibName = LibName
   TableStart = 8
   ParseTable
   OpenLib = 2
End If

End Function

Private Sub Class_Initialize()
Me.BufferLength = 50000
Me.FileHeader = "FL2"
Flags = 0
End Sub

Private Function CalcTablePos() As Double
CalcTablePos = 9
If IsAny = 0 Then
   Exit Function
End If

CalcTablePos = Entrys(UBound(Entrys)).EntryPos + Entrys(UBound(Entrys)).EntryLen
End Function

Private Function CalcTableStartDumb() As Double
CalcTableStart = 9
If IsAny = 0 Then
   Exit Function
End If

For x = 0 To UBound(Entrys)
   CalcTableStart = CalcTableStart + Entrys(x).EntryLen
Next x

End Function

'Ez elõtt a TableStart-ot be KELL állítani!
Private Sub WriteTable()
Dim Lib As Random, Temp As String
Set Lib = New Random
Lib.OpenFile mvarLibName
Lib.SeekAbsolute 4
Lib.WriteBytes StrConv(CompleteString(DecTo256(TableStart), 4), vbFromUnicode)
Lib.SeekAbsolute TableStart
Lib.SetEOFHere

Temp = CompleteString(DecTo256(UBound(Entrys) + 1), 4)
For x = 0 To UBound(Entrys)
   Temp = Temp & CTableEntryString(Entrys(x))
Next x
Lib.WriteBytes StrConv(Temp, vbFromUnicode)
Temp = ""
Lib.CloseFile
Set Lib = Nothing

End Sub

Private Sub ParseTable()
Dim Lib As Random, Temp As String
Set Lib = New Random
Lib.OpenFile mvarLibName
Lib.SeekAbsolute TableStart

Temp = StrConv(Lib.ReadBytes(4), vbUnicode)
If BackToDec(Temp) = 0 Then
   ReDim Entrys(0)
   IsAny = 0
   Exit Sub
End If
IsAny = 1

ReDim Entrys(BackToDec(Temp) - 1)

Buffer = ""
Do
   If Lib.GetCurrentPos + mvarBufferLength < Lib.GetFileLen Then
      Buffer = Buffer & StrConv(Lib.ReadBytes(mvarBufferLength), vbUnicode)
   Else
      Buffer = Buffer & StrConv(Lib.ReadBytes(Lib.GetFileLen - Lib.GetCurrentPos), vbUnicode)
      Exit Do
   End If
Loop
Lib.CloseFile
Set Lib = Nothing

For x = 0 To UBound(Entrys)
   Temp = Left(Buffer, 2)
   Temp = Left(Buffer, 2 + BackToDec(Temp) + 8)
   Entrys(x) = CTableEntryType(Temp)
   Buffer = Mid(Buffer, Len(Temp) + 1)
Next x
Buffer = ""

End Sub

Private Function CTableEntryType(TCEntry As String) As Entry
CTableEntryType.EntryName = Mid(TCEntry, 3, BackToDec(Left(TCEntry, 2)))
CTableEntryType.EntryPos = BackToDec(Mid(TCEntry, Len(CTableEntryType.EntryName) + 3, 4))
CTableEntryType.EntryLen = BackToDec(Mid(TCEntry, Len(CTableEntryType.EntryName) + 7, 4))
End Function

Private Function CTableEntryString(TCEntry As Entry) As String
CTableEntryString = CompleteString(DecTo256(Len(TCEntry.EntryName)), 2) & TCEntry.EntryName
CTableEntryString = CTableEntryString & CompleteString(DecTo256(TCEntry.EntryPos), 4)
CTableEntryString = CTableEntryString & CompleteString(DecTo256(TCEntry.EntryLen), 4)
End Function

Private Sub Class_Terminate()
CloseLib
End Sub
