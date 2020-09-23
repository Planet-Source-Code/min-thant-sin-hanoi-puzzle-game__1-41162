Attribute VB_Name = "Declarations"
Option Explicit

Public Enum enumMoveType
      CLICK_MOVE = 0
      DRAG_MOVE = 1
End Enum

Public MoveType As enumMoveType

'enumPegType is not necessary
Public Enum enumPegType
      SOURCE_PEG = 0
      SPARE_PEG = 1
      DESTINATION_PEG = 2
End Enum

Public Type POINTAPI
      x As Long
      y As Long
End Type

Public Type PEGINFO
      x As Integer
      y As Integer
End Type

Public Type DISKINFO
      Index As Integer
      CurrentPeg As enumPegType
      x As Integer
      y As Integer
      HomeX As Integer
      HomeY As Integer
      Width As Integer
      Height As Integer
End Type

'Graphics
Public Const SRCINVERTANDDEST = &H220B24
Public Const SRCCOPY = &HCC0020

'Sounds
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2

'Regions
Public Const RGN_AND = 1
Public Const RGN_DIFF = 4
Public Const NULLREGION = 1

Public Const MinimumDisks = 3
Public Const MaximumDisks = 7

Public Const PegWidth = 25
Public PegHeight

Public Const MaxPegHeight = 197
Public Const NumPegs = 2 'Total 3

Public Const DiskHeight = 24
Public Const SmallestDiskWidth = 60
Public Const DiskWidthIncrement = 20
Public Const MaxDiskWidth = SmallestDiskWidth + (MaximumDisks * DiskWidthIncrement)

Public NumDisks As Integer

Public MoveCount As Long
Public TypicalMoves As Integer

Public MovingPiece   As Integer
Public OldMouseX    As Integer
Public OldMouseY    As Integer
Public OldDiskX As Integer
Public OldDiskY As Integer

Public rgnLast As Long
Public rgnOther As Long
Public rgnMoving As Long
Public rgnCombine As Long

Public rctPeg(NumPegs) As RECT
Public Peg As PEGINFO
Public Pegs(NumPegs) As PEGINFO

Public Disk As DISKINFO
Public Disks() As DISKINFO

Public Priority() As Integer

Public boolSoundOn As Boolean
Public boolNewGame As Boolean
Public boolExcellent As Boolean

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetRectRgn Lib "gdi32" (ByVal hRgn As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Public Declare Function ClipCursorBynum Lib "user32" Alias "ClipCursor" (ByVal lpRect As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
