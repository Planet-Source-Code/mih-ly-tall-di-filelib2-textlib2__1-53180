VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FileLib 
   Appearance      =   0  'Flat
   Caption         =   "Text Library 2"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10230
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton SaveButt 
      Caption         =   "&Save"
      Height          =   255
      Left            =   9480
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin RichTextLib.RichTextBox EntryBox 
      Height          =   5775
      Index           =   1
      Left            =   2640
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   10186
      _Version        =   393217
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"Main.frx":0442
   End
   Begin RichTextLib.RichTextBox EntryBox 
      Height          =   5775
      Index           =   0
      Left            =   2640
      TabIndex        =   7
      Top             =   240
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   10186
      _Version        =   393217
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"Main.frx":052C
   End
   Begin VB.Frame Frame1 
      Height          =   375
      Left            =   -120
      TabIndex        =   4
      Top             =   6000
      Width           =   10455
      Begin VB.Label LibStatLab 
         Alignment       =   1  'Right Justify
         Caption         =   "Lib stats: Library not opened"
         Height          =   255
         Left            =   5160
         TabIndex        =   6
         Top             =   120
         Width           =   5175
      End
      Begin VB.Label StatLab 
         Caption         =   "library not opened."
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   120
         Width           =   4215
      End
      Begin VB.Label SLab 
         Caption         =   "Status:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   615
      End
   End
   Begin MSComDlg.CommonDialog CmDlg 
      Left            =   1320
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Text Libary (*.TxL)|*.TxL"
   End
   Begin VB.ListBox EntryList 
      Appearance      =   0  'Flat
      Height          =   5295
      ItemData        =   "Main.frx":0616
      Left            =   0
      List            =   "Main.frx":0618
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   2535
   End
   Begin VB.ComboBox CatList 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.PictureBox DragLab 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   5655
      Left            =   2520
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5655
      ScaleWidth      =   135
      TabIndex        =   10
      Top             =   240
      Width           =   135
   End
   Begin VB.Label ENameLab 
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   0
      Width           =   7575
   End
   Begin VB.Label InfoLab 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Categories && Entries:"
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1455
   End
   Begin VB.Menu FileMnu 
      Caption         =   "&File"
      Begin VB.Menu NewMnu 
         Caption         =   "&New..."
      End
      Begin VB.Menu OpenMnu 
         Caption         =   "&Open..."
      End
      Begin VB.Menu Div4Mnu 
         Caption         =   "-"
      End
      Begin VB.Menu MergeMnu 
         Caption         =   "Merge..."
      End
      Begin VB.Menu Div3Mnu 
         Caption         =   "-"
      End
      Begin VB.Menu ExitMnu 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu EditMnu 
      Caption         =   "&Edit"
      Enabled         =   0   'False
      Begin VB.Menu CopyMnu 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu SelAllMnu 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu CopyAllMnu 
         Caption         =   "Copy All"
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu ViewMnu 
      Caption         =   "&View"
      Begin VB.Menu LineBrkMnu 
         Caption         =   "Break lines"
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu EntryMnu 
      Caption         =   "&Entry"
      Enabled         =   0   'False
      Begin VB.Menu AddMnu 
         Caption         =   "&Add..."
      End
      Begin VB.Menu AddDirMnu 
         Caption         =   "Add directory..."
      End
      Begin VB.Menu DeleteMnu 
         Caption         =   "&Delete"
      End
      Begin VB.Menu DelCatMnu 
         Caption         =   "De&lete category"
      End
      Begin VB.Menu Div2Mnu 
         Caption         =   "-"
      End
      Begin VB.Menu RenEntryMnu 
         Caption         =   "Re&name entry..."
      End
      Begin VB.Menu RenCatMnu 
         Caption         =   "Rena&me category..."
      End
      Begin VB.Menu Div6Mnu 
         Caption         =   "-"
      End
      Begin VB.Menu ExtractMnu 
         Caption         =   "Ex&tract to..."
      End
      Begin VB.Menu ExtractCatMnu 
         Caption         =   "Extract &category to..."
      End
      Begin VB.Menu ExtractAllMnu 
         Caption         =   "Ext&ract all to..."
      End
   End
   Begin VB.Menu HelpMnu 
      Caption         =   "&Help"
      Begin VB.Menu AssocMnu 
         Caption         =   "Associate with "".TxL"" files"
      End
      Begin VB.Menu Div7Mnu 
         Caption         =   "-"
      End
      Begin VB.Menu AboutMnu 
         Caption         =   "A&bout..."
      End
   End
End
Attribute VB_Name = "FileLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Private FL As New FileLib2, oldX As Integer, ListPercent As Integer

Private Sub AboutMnu_Click()
AboutForm.Show 1, Me
End Sub

Private Sub AddDirMnu_Click()
Dim SrcDir As String
SrcDir = CommonTree("Add a whole directory", Me.hwnd)
If SrcDir = "" Then Exit Sub
FL.AddDirectory SrcDir, SrcDir
ParseLib
End Sub

Private Sub AddMnu_Click()
Dim SaveCat As Integer
StatePass = CatList.ListIndex

AddEntryForm.Show 1, Me
If StatePass = -3 Then Exit Sub
FL.AddStringEntry TextPass, NamePass

SaveCat = CatList.ListIndex
ParseLib
StatLab = "entry added."
CatList.ListIndex = SaveCat
End Sub

Private Sub AssocMnu_Click()
If MsgBox("You are going to associate .TxL files with Text Library." & vbNewLine & "Please confirm.", vbExclamation + vbOKCancel, "Confirm") = vbOK Then
   Dim FilTyp As filetype
   FilTyp.Extension = "TxL"
   FilTyp.ContentType = "Text Database"
   FilTyp.FullName = "Text Libary File"
   FilTyp.ProperName = "TxL"
   FilTyp.Commands.Captions.Add "open"
   If Right(App.Path, 1) <> "\" Then
      FilTyp.Commands.Commands.Add Chr(34) & App.Path & "\" & App.EXEName & Chr(34) & " %1"
      FilTyp.IconPath = App.Path & "\" & App.EXEName & ".exe"
   Else
      FilTyp.Commands.Commands.Add App.Path & App.EXEName & " %1"
      FilTyp.IconPath = App.Path & App.EXEName & ".exe"
   End If
   FilTyp.IconIndex = 0
   CreateExtension FilTyp
End If

End Sub

Private Sub CatList_Click()
If FL.NumEntrys = 0 Then Exit Sub

ReDim CurrEntrys(0)
EntryList.Clear
EntryBox(WordWarp) = ""
SaveButt.Visible = False
ENameLab = ""

For X = 0 To UBound(Entrys)
   If Mid(Entrys(X), 1, InStr(1, Entrys(X), "\") - 1) = CatList.Text Then
      CurrEntrys(UBound(CurrEntrys)) = X
      ReDim Preserve CurrEntrys(UBound(CurrEntrys) + 1)
   End If
Next X
ReDim Preserve CurrEntrys(UBound(CurrEntrys) - 1)

For X = 0 To UBound(CurrEntrys)
   EntryList.AddItem Mid$(Entrys(CurrEntrys(X)), InStr(1, Entrys(CurrEntrys(X)), "\") + 1)
   EntryList.ItemData(EntryList.NewIndex) = X
Next X

LibStatLab = "Lib stats: Entries: " & UBound(Entrys) + 1 & " ...in current category: " & EntryList.ListCount & " categories: " & CatList.ListCount & " "
End Sub

Private Sub CopyAllMnu_Click()
Clipboard.SetText EntryBox(WordWarp).Text
End Sub

Private Sub CopyMnu_Click()
Clipboard.SetText EntryBox(WordWarp).SelText
End Sub

Private Sub DelCatMnu_Click()
If CatList.ListIndex = -1 Then Exit Sub
If MsgBox("Are you sure you want to delete this category?", vbQuestion + vbYesNo) = vbYes Then
'   For X = 0 To EntryList.ListCount - 1
'      FL.DeleteEntrys CurrEntrys(EntryList.ItemData(X)), CurrEntrys(EntryList.ItemData(X)), FL.LibName & ".TMP"
'   Next X
   For X = 0 To UBound(CurrEntrys)
      FL.DeleteEntrys CurrEntrys(X), CurrEntrys(X), FL.LibName & ".TMP"
      For Y = X To UBound(CurrEntrys)
         CurrEntrys(Y) = CurrEntrys(Y) - 1
      Next Y
   Next X
End If

ParseLib
StatLab = "category deleted."
End Sub

Private Sub DeleteMnu_Click()
Dim SaveCat As String
If EntryList.ListIndex = -1 Then Exit Sub
If MsgBox("Are you sure you want to delete this entry?", vbQuestion + vbYesNo) = vbYes Then
   FL.DeleteEntrys CurrEntrys(EntryList.ItemData(EntryList.ListIndex)), CurrEntrys(EntryList.ItemData(EntryList.ListIndex)), FL.LibName & ".TMP"
End If

If UBound(CurrEntrys) > 0 Then
   SaveCat = CatList.Text
Else
   SaveCat = ""
End If

ParseLib

StatLab = "entry deleted"

For X = 0 To CatList.ListCount - 1
   If SaveCat = CatList.List(X) Then
      CatList.ListIndex = X
      Exit Sub
   End If
Next X

End Sub

Private Sub DragLab_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
oldX = X
End Sub

Private Sub DragLab_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
   If EntryList.Width + (X - oldX) < InfoLab.Width Then Exit Sub
   If Me.ScaleWidth - 135 - (EntryList.Width + (X - oldX)) < 120 Then Exit Sub
   DragLab.Cls
   
   EntryList.Width = EntryList.Width + (X - oldX)
   
   CatList.Width = EntryList.Width
   EntryBox(0).Width = Me.ScaleWidth - 135 - EntryList.Width
   EntryBox(0).Left = EntryList.Width + 135
   EntryBox(1).Width = Me.ScaleWidth - 135 - EntryList.Width
   EntryBox(1).Left = EntryList.Width + 135
   ENameLab.Left = EntryBox(0).Left
   EntryList.Refresh
   'Form_Resize
   oldX = X
End If
End Sub

Private Sub DragLab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
DragLab.Left = EntryList.Width

DragLab.ForeColor = vbButtonShadow
DragLab.DrawWidth = 2
DragLab.Line (DragLab.ScaleWidth / 2, Screen.TwipsPerPixelY * 3)-(DragLab.Width / 2, DragLab.ScaleHeight - Screen.TwipsPerPixelY * 3)
DragLab.ForeColor = vb3DHighlight
DragLab.DrawWidth = 1
DragLab.Line (DragLab.ScaleWidth / 2, Screen.TwipsPerPixelY * 3)-(DragLab.Width / 2, DragLab.ScaleHeight - Screen.TwipsPerPixelY * 3)

End Sub

Private Sub EntryBox_Change(Index As Integer)
If EntryBox(worwrap).Text = "" Then EditMnu.Enabled = False Else EditMnu.Enabled = True
End Sub

Private Sub EntryBox_KeyPress(Index As Integer, KeyAscii As Integer)
If EntryList.ListIndex > -1 Then
   SaveButt.Visible = True
Else
    KeyAscii = 0
End If
End Sub

Private Sub EntryList_Click()
Dim WorkEntry As Long
WorkEntry = CurrEntrys(EntryList.ItemData(EntryList.ListIndex))
ENameLab = Replace(EntryList.Text, "&", "&&")
EntryBox(WordWarp).Text = FL.ExtractEntryString(WorkEntry)
SaveButt.Visible = False

End Sub

Private Sub ExitMnu_Click()
FL.CloseLib
Set FL = Nothing
Unload Me
End
End Sub

Private Sub ExtractAllMnu_Click()
Dim ExtractPath As String
If CatList.ListCount = 0 Then Exit Sub
ExtractPath = CommonTree("Select dir to extract to", Me.hwnd)

If ExtractPath = "" Then Exit Sub
   
For X = 0 To UBound(Entrys)
   FL.ExtractEntry X, True, ExtractPath
   StatLab = "extracting " & X + 1 & " of " & UBound(Entrys) & "..."
   DoEvents
Next X

StatLab = "category extracted."
End Sub

Private Sub ExtractCatMnu_Click()
Dim ExtractPath As String
If CatList.ListIndex = -1 Then Exit Sub
ExtractPath = CommonTree("Select dir to extract to", Me.hwnd)

If ExtractPath = "" Then Exit Sub
If Right(ExtractPath, 1) <> "\" Then ExtractPath = ExtractPath & "\"
   
For X = 0 To EntryList.ListCount - 1
   FL.ExtractEntry CurrEntrys(EntryList.ItemData(X)), False, "", ExtractPath & EntryList.List(X)
   StatLab = "extracting " & X + 1 & " of " & EntryList.ListCount & "..."
   DoEvents
Next X

StatLab = "category extracted."
End Sub

Private Sub ExtractMnu_Click()
Dim ExtractPath As String
If EntryList.ListIndex = -1 Then Exit Sub
ExtractPath = CommonTree("Select dir to extract to", Me.hwnd)

If ExtractPath = "" Then Exit Sub
If Right(ExtractPath, 1) <> "\" Then ExtractPath = ExtractPath & "\"

FL.ExtractEntry CurrEntrys(EntryList.ItemData(EntryList.ListIndex)), False, "", ExtractPath & EntryList.Text

StatLab = "entry extracted."
End Sub

Private Sub Form_Load()
FL.FileHeader = "TxL"
FL.BufferLength = BUFF_LEN
If Command <> "" Then
'Loading library...
   FL.CloseLib
   
   FL.FileHeader = "TxL"
   FL.BufferLength = BUFF_LEN
   
   If FL.OpenLib(Command) = 0 Then
      MsgBox "The library cannot be opened.", vbCritical
      StatLab = "library cannot be opened."
      FL.CloseLib
   Else
      StatLab = "library opened successfully."
      Me.Caption = "Text Library 2 - " & StrConv(Mid(FL.LibName, InStrRev(FL.LibName, "\") + 1, InStrRev(FL.LibName, ".") - 1), vbProperCase)
   End If
   
   ParseLib
   EntryMnu.Enabled = True
End If
DragLab.ForeColor = vbButtonShadow
DragLab.DrawWidth = 2
DragLab.Line (DragLab.ScaleWidth / 2, Screen.TwipsPerPixelY * 3)-(DragLab.Width / 2, DragLab.ScaleHeight - Screen.TwipsPerPixelY * 3)
DragLab.ForeColor = vb3DHighlight
DragLab.DrawWidth = 1
DragLab.Line (DragLab.ScaleWidth / 2, Screen.TwipsPerPixelY * 3)-(DragLab.Width / 2, DragLab.ScaleHeight - Screen.TwipsPerPixelY * 3)

End Sub

'Private Sub Form_Resize()
'If (Me.ScaleHeight = 0) And (Me.ScaleWidth = 0) Then Exit Sub
'
'If Me.ScaleHeight < 1170 Then Me.Height = 1170 + 690
'If Me.ScaleWidth < 3090 Then Me.Width = 3090 + 120
'Frame1.Top = Me.ScaleHeight - 345
'Frame1.Width = Me.ScaleWidth + 225
'EntryBox(0).Height = Me.ScaleHeight - 570
'EntryBox(0).Width = Me.ScaleWidth - 135 - EntryList.Width
'EntryBox(0).RightMargin = EntryBox(0).Width + 10000000
'EntryBox(0).Left = EntryList.Width + 135
'
'EntryBox(1).Height = Me.ScaleHeight - 570
'EntryBox(0).Width = Me.ScaleWidth - 135 - EntryList.Width
'EntryBox(1).Left = EntryList.Width + 135
'
'ENameLab.Left = EntryBox(0).Left
'DragLab.Height = Me.ScaleHeight
'
'EntryList.Height = Me.ScaleHeight - 1050
'StatLab.Left = 720
'StatLab.Width = Frame1.Width / 2 - 120 * 2
'LibStatLab.Left = StatLab.Left + StatLab.Width
'LibStatLab.Width = StatLab.Width + 120 * 2 - 660
'End Sub

Private Sub Form_Resize()
If (Me.ScaleHeight = 0) And (Me.ScaleWidth = 0) Then Exit Sub

If Me.ScaleHeight < 1170 Then Me.Height = 1170 + 690
If Me.ScaleWidth < 3090 Then Me.Width = 3090 + 120
Frame1.Top = Me.ScaleHeight - 345
Frame1.Width = Me.ScaleWidth + 225
EntryBox(0).Height = Me.ScaleHeight - 570
EntryBox(0).Width = Me.ScaleWidth - 2655 - Screen.TwipsPerPixelX * 5
EntryBox(0).RightMargin = EntryBox(0).Width + 10000000
EntryBox(0).Left = 2535 + 135
EntryBox(1).Height = Me.ScaleHeight - 570
EntryBox(1).Width = Me.ScaleWidth - 2655 - Screen.TwipsPerPixelX * 5
EntryBox(1).Left = 2535 + 135
SaveButt.Left = EntryBox(1).Left + EntryBox(1).Width - SaveButt.Width
ENameLab.Left = EntryBox(1).Left

EntryList.Height = Me.ScaleHeight - 1050
EntryList.Width = 2535
CatList.Width = 2535
DragLab.Left = EntryList.Width
DragLab.Height = Me.ScaleHeight - (Me.ScaleHeight - Frame1.Top)

DragLab.ForeColor = vbButtonShadow
DragLab.DrawWidth = 2
DragLab.Line (DragLab.ScaleWidth / 2, Screen.TwipsPerPixelY * 3)-(DragLab.Width / 2, DragLab.ScaleHeight - Screen.TwipsPerPixelY * 3)
DragLab.ForeColor = vb3DHighlight
DragLab.DrawWidth = 1
DragLab.Line (DragLab.ScaleWidth / 2, Screen.TwipsPerPixelY * 3)-(DragLab.Width / 2, DragLab.ScaleHeight - Screen.TwipsPerPixelY * 3)

StatLab.Left = 720
StatLab.Width = Frame1.Width / 2 - 120 * 2
LibStatLab.Left = StatLab.Left + StatLab.Width
LibStatLab.Width = StatLab.Width + 120 * 2 - 660
End Sub


Private Sub LineBrkMnu_Click()

If WordWarp = 0 Then
   LineBrkMnu.Checked = True
   WordWarp = 1
   EntryBox(1) = EntryBox(0)
   EntryBox(0) = ""
   If EntryBox(1).Text <> "" Then EditMnu.Enabled = True Else EditMnu.Enabled = False
   EntryBox(0).Visible = False
   EntryBox(1).Visible = True
Else
   LineBrkMnu.Checked = False
   WordWarp = 0
   EntryBox(0) = EntryBox(1)
   EntryBox(1) = ""
   If EntryBox(0).Text <> "" Then EditMnu.Enabled = True Else EditMnu.Enabled = False
   EntryBox(1).Visible = False
   EntryBox(0).Visible = True
End If

End Sub

Private Sub MergeMnu_Click()
MergeForm.Show 1, Me
End Sub

Private Sub NewMnu_Click()
On Error GoTo EEE
FL.CloseLib
cmdlg.Flags = cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNOverwritePrompt
cmdlg.ShowSave
If Dir(cmdlg.FileName) <> "" Then
   If DelFile(cmdlg.FileName) = False Then
      MsgBox "The library cannot be created.", vbCritical
      StatLab = "library cannot be created."
   
      EntryBox(WordWarp) = ""
      EntryList.Clear
      CatList.Clear
      ENameLab = ""
      ReDim Entrys(0)
      
      GoTo EEE
   End If
End If

FL.CloseLib

FL.FileHeader = "TxL"
FL.BufferLength = BUFF_LEN

If FL.OpenLib(cmdlg.FileName) = 0 Then
   MsgBox "The library cannot be created.", vbCritical
   StatLab = "library cannot be created."
   FL.CloseLib
   
   EntryBox(WordWarp) = ""
   EntryList.Clear
   CatList.Clear
   ENameLab = ""
   ReDim Entrys(0)

   GoTo EEE
End If

StatLab = "library created successfully."
Me.Caption = "Text Library 2 - " & Mid(FL.LibName, InStrRev(FL.LibName, "\") + 1)

ParseLib
EntryMnu.Enabled = True
EEE:
End Sub

Private Sub ParseLib()
EntryBox(WordWarp) = ""
SaveButt.Visible = False
EntryList.Clear
CatList.Clear
ENameLab = ""

ReDim Entrys(0)
If FL.NumEntrys = 0 Then
   'LibStatLab = "Lib stats: Entries: 0 ...in current category: 0 categories: 0 "
   LibStatLab = "Lib stats: libary is empty "
   Exit Sub
End If
ReDim Entrys(FL.NumEntrys - 1)

For X = 0 To UBound(Entrys)
   Entrys(X) = FL.GetEntryName(X)
   
   For Y = 0 To CatList.ListCount - 1
      If CatList.List(Y) = Mid$(Entrys(X), 1, InStr(1, Entrys(X), "\") - 1) Then
         GoTo CategoryAdded
      End If
   Next Y
   CatList.AddItem Mid$(Entrys(X), 1, InStr(1, Entrys(X), "\") - 1)

CategoryAdded:
Next X

LibStatLab = "Lib stats: Entries: " & UBound(Entrys) + 1 & " ...in current category: " & EntryList.ListCount & " categories: " & CatList.ListCount ' & " "
End Sub

Private Sub OpenMnu_Click()
On Error GoTo EEE
cmdlg.Flags = cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNFileMustExist
cmdlg.ShowOpen

FL.CloseLib

FL.FileHeader = "TxL"
FL.BufferLength = BUFF_LEN

If FL.OpenLib(cmdlg.FileName) = 0 Then
   MsgBox "The library cannot be opened.", vbCritical
   StatLab = "library cannot be opened."
   FL.CloseLib
   
   EntryBox(WordWarp) = ""
   EntryList.Clear
   CatList.Clear
   ENameLab = ""
   ReDim Entrys(0)

   GoTo EEE
End If

StatLab = "library opened successfully."
Me.Caption = "Text Library 2 - " & Mid(FL.LibName, InStrRev(FL.LibName, "\") + 1)

ParseLib
EntryMnu.Enabled = True
EEE:
End Sub

Private Sub RenCatMnu_Click()
Dim CatSave As Integer, EntrySave As Integer
If CatList.ListIndex = -1 Then Exit Sub

NamePass = CatList.Text
RenameForm.Show 1, Me

If NamePass = "" Then Exit Sub
For X = 0 To UBound(CurrEntrys)
   FL.RenameEntry CurrEntrys(EntryList.ItemData(X)), NamePass & "\" & EntryList.List(X)
Next X

CatSave = CatList.ListIndex
EntrySave = EntryList.ListIndex
ParseLib

CatList.ListIndex = CatSave
EntryList.ListIndex = EntrySave

End Sub

Private Sub RenEntryMnu_Click()
Dim CatSave As Integer, EntrySave As Integer
If EntryList.ListIndex = -1 Then Exit Sub

NamePass = EntryList.Text
RenameForm.Show 1, Me

If NamePass = "" Then Exit Sub
FL.RenameEntry CurrEntrys(EntryList.ItemData(EntryList.ListIndex)), CatList.Text & "\" & NamePass

CatSave = CatList.ListIndex
EntrySave = EntryList.ListIndex
ParseLib

CatList.ListIndex = CatSave
EntryList.ListIndex = EntrySave

End Sub

Private Sub SaveButt_Click()
Dim ModEntryName As String, SaveCat As Integer, SaveEntry As Integer

ModEntryName = FL.GetEntryName(CurrEntrys(EntryList.ItemData(EntryList.ListIndex)))
SaveCat = CatList.ListIndex
SaveEntry = EntryList.ListIndex

FL.DeleteEntrys CurrEntrys(EntryList.ItemData(EntryList.ListIndex)), CurrEntrys(EntryList.ItemData(EntryList.ListIndex)), FL.LibName & ".TMP"
FL.AddStringEntry EntryBox(WordWarp).Text, ModEntryName
SaveButt.Visible = False
ParseLib

CatList.ListIndex = SaveCat
EntryList.ListIndex = SaveEntry
End Sub

Private Sub SelAllMnu_Click()
EntryBox(WordWarp).SelStart = 0
EntryBox(WordWarp).SelLength = Len(EntryBox(WordWarp).Text)
End Sub
