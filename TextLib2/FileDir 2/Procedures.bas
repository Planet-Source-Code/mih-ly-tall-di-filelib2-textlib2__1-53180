Attribute VB_Name = "Procedures"
Public Declare Function SetFileAttributes Lib "kernel32.dll" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Public Declare Function DeleteFile Lib "kernel32.dll" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Public Function GetAllFiles(ByVal Path As String, ByVal FileSpec As String, Optional RecurseDirs As Boolean) As Collection
    Dim spec As Variant
    Dim file As Variant
    Dim subdir As Variant
    Dim subdirs As New Collection
    Dim specs() As String
    
    ' initialize the result
    Set GetAllFiles = New Collection
    
    ' ensure that path has a trailing backslash
    If Right$(Path, 1) <> "\" Then Path = Path & "\"
    
    ' get the list of provided file specifications
    specs() = Split(FileSpec, ";")
    
    ' this is necessary to ignore duplicates in result
    ' caused by overlapping file specifications
    On Error Resume Next
                
    ' at each iteration search for a different filespec
    For Each spec In specs
        ' start the search
        file = Dir$(Path & spec)
        Do While Len(file)
            ' we've found a new file
            file = Path & file
            GetAllFiles.Add file, file
            ' get ready for the next iteration
            file = Dir$
        Loop
    Next
    
    ' first, build the list of subdirectories to be searched
    If RecurseDirs Then
        ' get the collection of subdirectories
        ' start the search
        file = Dir$(Path & "*.*", vbDirectory)
        Do While Len(file)
            ' we've found a new directory
            If file = "." Or file = ".." Then
                ' exclude the "." and ".." entries
            ElseIf (GetAttr(Path & file) And vbDirectory) = 0 Then
                ' ignore regular files
            Else
                ' this is a directory, include the path in the collection
                file = Path & file
                subdirs.Add file, file
            End If
            ' get next directory
            file = Dir$
        Loop
        
        ' parse each subdirectory
        For Each subdir In subdirs
            ' use GetAllFiles recursively
            For Each file In GetAllFiles(subdir, FileSpec, True)
                GetAllFiles.Add file, file
            Next
        Next
    End If
    
End Function

Public Function DecTo256(ByVal Dec As Variant) As String
Dim ConTmp As String

Do
ConTmp = Chr$(Dec Mod 256) & ConTmp
If Dec <= 255 Then GoTo Done
Dec = Int(Dec / 256)
Loop

Done:
DecTo256 = ConTmp
End Function

Public Function BackToDec(ByVal Back As String) As Variant
Dim z As Integer, BackConTmp As Long

For z = Len(Back) - 1 To 0 Step -1
    BackConTmp = BackConTmp + Asc(Mid$(Back, Len(Back) - z, 1)) * (256 ^ z)
Next z

BackToDec = BackConTmp
End Function

Public Function CompleteString(InNum As String, ByVal OutLen As Byte) As String
OutLen = OutLen - Len(InNum)
If OutLen <= 0 Then
   CompleteString = InNum
   Exit Function
End If
CompleteString = String(OutLen, Chr(0)) & InNum
End Function

Public Sub CreatePath(ByVal sPath As String)
Dim PathNodes() As String, PathCounter As Long
ReDim PathNodes(0 To PathCounter)
'Build Pathnodes
Do
    PathNodes(PathCounter) = Mid$(sPath, 1, InStr(1, sPath, "\", vbTextCompare))
    sPath = Right(sPath, Len(sPath) - Len(PathNodes(PathCounter)))
    If InStr(1, sPath, "\", vbTextCompare) <= 0 Then
        PathCounter = PathCounter + 1
        ReDim Preserve PathNodes(0 To PathCounter)
        PathNodes(PathCounter) = sPath
        Exit Do
    End If
    PathCounter = PathCounter + 1
    ReDim Preserve PathNodes(0 To PathCounter)
Loop

ChDrive PathNodes(0)
ChDir "\"
For PathCounter = 1 To UBound(PathNodes)
    If PathNodes(PathCounter) = "" Then GoTo SkipThis
    If Dir(PathNodes(PathCounter), vbDirectory) <> "" Then
        ChDir PathNodes(PathCounter)
    Else
        MkDir PathNodes(PathCounter)
        ChDir PathNodes(PathCounter)
    End If
SkipThis:
Next PathCounter

End Sub
