VERSION 5.00
Begin VB.Form frmWin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Achievement of Homo sapiens..."
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&Okay"
      Default         =   -1  'True
      Height          =   390
      Left            =   1425
      TabIndex        =   1
      Top             =   1350
      Width           =   1515
   End
   Begin VB.Image imgNormalSmile 
      Height          =   480
      Left            =   150
      Picture         =   "frmWin.frx":0000
      Top             =   1350
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWideSmile 
      Height          =   480
      Left            =   150
      Picture         =   "frmWin.frx":0442
      Top             =   825
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgSmile 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   150
      Top             =   225
      Width           =   480
   End
   Begin VB.Label lblComment 
      Caption         =   "Congratulations!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Left            =   825
      TabIndex        =   0
      Top             =   225
      Width           =   3240
   End
End
Attribute VB_Name = "frmWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
      Unload Me
End Sub

Private Sub Form_Load()
      Dim msg As String
      
      If MoveCount <= TypicalMoves + 1 Then
            msg = "Excellent! Incredible! I admire you." & vbNewLine
            
            imgSmile.Picture = imgWideSmile.Picture
      Else
            msg = "Congratulations!" & vbNewLine
            msg = msg & "But you need to improve your skill." & vbNewLine
            
            imgSmile.Picture = imgNormalSmile.Picture
      End If
      msg = msg & "You made " & MoveCount & " moves." & vbNewLine
      msg = msg & "The typical moves is " & TypicalMoves & "."
      
      lblComment.Caption = msg
End Sub
