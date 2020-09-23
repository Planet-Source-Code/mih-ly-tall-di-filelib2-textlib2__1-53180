VERSION 5.00
Begin VB.Form RenameForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rename"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   Icon            =   "RenameForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CancelButt 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2648
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton OkButt 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   848
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox NewNameBox 
      Height          =   285
      Left            =   188
      TabIndex        =   1
      Text            =   "NewNames"
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Enter new name:"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "RenameForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelButt_Click()
Unload Me
End Sub

Private Sub Form_Load()
NewNameBox = NamePass
NewNameBox.SelStart = 0
NewNameBox.SelLength = Len(NewNameBox)
NamePass = ""
End Sub

Private Sub OkButt_Click()
NamePass = NewNameBox
Unload Me
End Sub
