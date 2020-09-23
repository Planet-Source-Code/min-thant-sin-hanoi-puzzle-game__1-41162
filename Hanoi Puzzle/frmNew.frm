VERSION 5.00
Begin VB.Form frmNew 
   Caption         =   "New Game"
   ClientHeight    =   1095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4155
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   1095
   ScaleWidth      =   4155
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Okey-dokey!"
      Default         =   -1  'True
      Height          =   390
      Left            =   2850
      TabIndex        =   1
      Top             =   75
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   2850
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.ComboBox cboNumDisks 
      Height          =   315
      Left            =   1125
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   525
      Width           =   1590
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   75
      Picture         =   "frmNew.frx":0000
      Top             =   75
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Choose a puzzle size:"
      Height          =   195
      Left            =   1125
      TabIndex        =   3
      Top             =   225
      Width           =   1530
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
      Unload Me
End Sub

Private Sub cmdOK_Click()
      Dim x As Byte
      
      'Extract number of disks.
      NumDisks = Val(VBA.Left(cboNumDisks.Text, 1)) - 1
      PegHeight = (NumDisks + 1) * DiskHeight + 1
      ReDim Disks(NumDisks) As DISKINFO
      ReDim Priority(NumDisks) As Integer
      
      Call DrawPegs
      Call frmHanoi.PreparePuzzle
      Call PrintText
      
      boolNewGame = True
      cboNumDisks.SetFocus
      Unload Me
End Sub

Private Sub Form_Load()
      Dim x As Byte
      For x = MinimumDisks To MaximumDisks
            cboNumDisks.AddItem CStr(x) & "   disks"
      Next x
            cboNumDisks.ListIndex = (NumDisks - MinimumDisks) + 1
End Sub
