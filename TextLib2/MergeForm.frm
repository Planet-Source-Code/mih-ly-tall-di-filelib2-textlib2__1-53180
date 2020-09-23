VERSION 5.00
Begin VB.Form MergeForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Merge information"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   Icon            =   "MergeForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label6 
      Caption         =   $"MergeForm.frx":000C
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   6135
   End
   Begin VB.Label Label5 
      Caption         =   "4. Use the ""Entry/Add directory..."" menu, then select your temorary directory."
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   6135
   End
   Begin VB.Label Label4 
      Caption         =   "3. Load the target library."
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   6135
   End
   Begin VB.Label Label3 
      Caption         =   "2. Extract all of the files to a temporary directory."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   6135
   End
   Begin VB.Label Label2 
      Caption         =   "1. Load the library you want to merge."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   6135
   End
   Begin VB.Label Label1 
      Caption         =   "If you want to merge two librarys, there is a way, but it's not easy."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "MergeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = FileLib.Icon
End Sub
