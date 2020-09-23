VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form AddEntryForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Entry"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7785
   Icon            =   "AddEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox CatList 
      Height          =   315
      ItemData        =   "AddEntry.frx":000C
      Left            =   1080
      List            =   "AddEntry.frx":000E
      TabIndex        =   9
      Top             =   120
      Width           =   2655
   End
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   5880
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "All Files (*.*)|*.*"
   End
   Begin VB.CommandButton CancelButt 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3945
      TabIndex        =   7
      Top             =   7080
      Width           =   1695
   End
   Begin VB.CommandButton OkButt 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2145
      TabIndex        =   6
      Top             =   7080
      Width           =   1695
   End
   Begin VB.TextBox EntryBox 
      Appearance      =   0  'Flat
      Height          =   5775
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   1080
      Width           =   7575
   End
   Begin VB.TextBox EntryNameBox 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   480
      Width           =   2655
   End
   Begin VB.CommandButton LoadButt 
      Caption         =   "Load file..."
      Height          =   375
      Left            =   6360
      TabIndex        =   5
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   $"AddEntry.frx":0010
      Height          =   615
      Left            =   3840
      TabIndex        =   8
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label3 
      Caption         =   "Entry text:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Entry name:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Category:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "AddEntryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelButt_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = FileLib.Icon

If Entrys(0) = "" Then GoTo EE
For X = 0 To UBound(Entrys)
   For Y = 0 To CatList.ListCount - 1
      If CatList.List(Y) = Mid$(Entrys(X), 1, InStr(1, Entrys(X), "\") - 1) Then
         GoTo CategoryAdded
      End If
   Next Y
   CatList.AddItem Mid$(Entrys(X), 1, InStr(1, Entrys(X), "\") - 1)

CategoryAdded:
Next X
EE:
CatList.ListIndex = StatePass
StatePass = -3
End Sub

Private Sub LoadButt_Click()
On Error GoTo EEE
cmdlg.Flags = cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNFileMustExist
cmdlg.ShowOpen

Open cmdlg.FileName For Input As #1
EntryBox = Input(LOF(1), #1)
Close #1

If InStrRev(cmdlg.FileName, ".") <> 0 Then
   EntryNameBox = Mid(cmdlg.FileName, InStrRev(cmdlg.FileName, "\") + 1, InStrRev(cmdlg.FileName, ".") - (InStrRev(cmdlg.FileName, "\") + 1))
Else
   EntryNameBox = Mid(cmdlg.FileName, InStrRev(cmdlg.FileName, "\") + 1)
End If
EEE:
End Sub

Private Sub OkButt_Click()
If (CatList.Text = "") Or (EntryBox = "") Or (EntryNameBox = "") Then Exit Sub
If Not ((IsValidFileName(CatList.Text, True)) Or (IsValidFileName(EntryNameBox, True))) Then
   MsgBox "The entry and category names must follow standard file naming conventions!", vbCritical
   Exit Sub
End If

NamePass = CatList.Text & "\" & EntryNameBox

For X = 0 To UBound(Entrys)
   If Entrys(X) = NamePass Then
      MsgBox "The new entry must have a unique name!", vbExclamation
      Exit Sub
   End If
Next X

TextPass = EntryBox
StatePass = -2
Unload Me
End Sub
