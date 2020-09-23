VERSION 5.00
Begin VB.Form frmOption 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hanoi options"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2265
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   2265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&Okey-dokey"
      Default         =   -1  'True
      Height          =   390
      Left            =   75
      TabIndex        =   6
      Top             =   2250
      Width           =   2115
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sound"
      Height          =   990
      Left            =   75
      TabIndex        =   3
      Top             =   1125
      Width           =   2115
      Begin VB.OptionButton optSound 
         Caption         =   "O&ff"
         Height          =   240
         Index           =   1
         Left            =   225
         TabIndex        =   5
         Tag             =   "0"
         Top             =   600
         Width           =   840
      End
      Begin VB.OptionButton optSound 
         Caption         =   "&On"
         Height          =   240
         Index           =   0
         Left            =   225
         TabIndex        =   4
         Tag             =   "-1"
         Top             =   300
         Value           =   -1  'True
         Width           =   915
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Move disk by :"
      Height          =   990
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   2115
      Begin VB.OptionButton optMoveType 
         Caption         =   "&Dragging"
         Height          =   240
         Index           =   1
         Left            =   225
         TabIndex        =   2
         Tag             =   "1"
         Top             =   300
         Value           =   -1  'True
         Width           =   1365
      End
      Begin VB.OptionButton optMoveType 
         Caption         =   "&Clicking"
         Height          =   240
         Index           =   0
         Left            =   225
         TabIndex        =   1
         Tag             =   "0"
         Top             =   600
         Width           =   1290
      End
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
      Me.Hide
End Sub

Private Sub optMoveType_Click(Index As Integer)
      MoveType = Val(optMoveType(Index).Tag)
End Sub

Private Sub optSound_Click(Index As Integer)
      boolSoundOn = Val(optSound(Index).Tag)
End Sub
