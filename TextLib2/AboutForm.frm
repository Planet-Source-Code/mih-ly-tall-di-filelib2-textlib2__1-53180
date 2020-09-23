VERSION 5.00
Begin VB.Form AboutForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Text Library by Msi"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   Icon            =   "AboutForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label MailLab 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "msist@freemail.hu"
      Height          =   195
      Left            =   2100
      TabIndex        =   2
      Top             =   1680
      Width           =   1305
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "If you found a bug, have an idea to improve my program, or you have any other comments, send your mail to:"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   5295
   End
   Begin VB.Label Label1 
      Caption         =   $"AboutForm.frx":000C
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "AboutForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = FileLib.Icon
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
UnLineMail
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
UnLineMail
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
UnLineMail
End Sub

Private Sub MailLab_Click()
ShellExecute Me.hwnd, "open", "mailto:msist@freemail.hu", "", "", 0
End Sub

Private Sub MailLab_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
LineMail
End Sub

Private Sub UnLineMail()
MailLab.ForeColor = vbWindowText
MailLab.FontUnderline = False
End Sub

Private Sub LineMail()
MailLab.ForeColor = vbHighlight
MailLab.FontUnderline = True
End Sub
