VERSION 5.00
Begin VB.Form frmHanoi 
   Caption         =   "Hanoi Puzzle"
   ClientHeight    =   7815
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11400
   Icon            =   "Hanoi Puzzle.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   521
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   760
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picFinalDisk 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   2925
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   179
      TabIndex        =   9
      Top             =   825
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.PictureBox picDiskImage 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   2925
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   179
      TabIndex        =   8
      Top             =   450
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.PictureBox picDiskMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   2925
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   179
      TabIndex        =   7
      Top             =   75
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.PictureBox picPegMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2925
      Left            =   2925
      Picture         =   "Hanoi Puzzle.frx":164A
      ScaleHeight     =   195
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   152
      TabIndex        =   6
      Top             =   3900
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.PictureBox picPegBitmap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2925
      Left            =   75
      Picture         =   "Hanoi Puzzle.frx":171E4
      ScaleHeight     =   195
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   152
      TabIndex        =   5
      Top             =   3900
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.PictureBox picWholeMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2520
      Left            =   2925
      Picture         =   "Hanoi Puzzle.frx":2CD7E
      ScaleHeight     =   168
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   4
      Top             =   1275
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.PictureBox picBoard 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1110
      Left            =   75
      ScaleHeight     =   74
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   55
      TabIndex        =   0
      Top             =   75
      Width           =   825
   End
   Begin VB.PictureBox picBackground 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1110
      Left            =   975
      Picture         =   "Hanoi Puzzle.frx":43020
      ScaleHeight     =   1110
      ScaleWidth      =   900
      TabIndex        =   3
      Top             =   75
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox picWholeBitmap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2520
      Left            =   75
      Picture         =   "Hanoi Puzzle.frx":4E96B
      ScaleHeight     =   168
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   2
      Top             =   1275
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.PictureBox picTmpBoard 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      ForeColor       =   &H0000FF00&
      Height          =   1110
      Left            =   1950
      ScaleHeight     =   74
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   55
      TabIndex        =   1
      Top             =   75
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuNew 
         Caption         =   "&New...               "
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options..."
         Shortcut        =   {F3}
      End
      Begin VB.Menu sepExit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHowToPlay 
         Caption         =   "&How to play..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu sepAbout 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmHanoi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'A lot of thanks to a developer who wrote
'Jigsaw (puzzle) in VB.
'I used some of his code and idea.

Dim boolDragging As Boolean

Private Sub Form_Load()
      boolDragging = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
      Call ReleaseMouse
      Call CleanUp
      End
End Sub

Private Sub mnuAbout_Click()
      MsgBox "Developed by Min Thant Sin in 2002.", vbInformation, "Sneaky, sneaky, sneaky..."
End Sub

Private Sub mnuExit_Click()
      Unload Me
End Sub
      
Private Sub GetMaskAndImage(PieceIndex)
      Disk = Disks(Priority(PieceIndex))
      Call BitBlt(picDiskMask.hdc, 0, 0, Disk.Width, Disk.Height, picWholeMask.hdc, Disk.HomeX, Disk.HomeY, vbSrcCopy)
      Call BitBlt(picDiskImage.hdc, 0, 0, Disk.Width, Disk.Height, picWholeBitmap.hdc, Disk.HomeX, Disk.HomeY, vbSrcCopy)
      Call BitBlt(picDiskImage.hdc, 0, 0, Disk.Width, Disk.Height, picDiskMask.hdc, 0, 0, vbSrcInvert And vbNotSrcCopy)
End Sub

Private Sub DisplayAPiece(destHdc&, PieceIndex%)
      Disk = Disks(Priority(PieceIndex))
      
      Call BitBlt(picFinalDisk.hdc, 0, 0, Disk.Width, Disk.Height, picTmpBoard.hdc, Disk.x, Disk.y, vbSrcCopy)
      Call BitBlt(picFinalDisk.hdc, 0, 0, Disk.Width, Disk.Height, picDiskMask.hdc, 0, 0, vbSrcAnd)
      Call BitBlt(picFinalDisk.hdc, 0, 0, Disk.Width, Disk.Height, picDiskImage.hdc, 0, 0, vbSrcInvert)
      Call BitBlt(destHdc, Disk.x, Disk.y, Disk.Width, Disk.Height, picFinalDisk.hdc, 0, 0, vbSrcCopy)
End Sub

Sub DisplayAPeg(destHdc&, PegIndex%)
      Dim StartX, StartY
      
      StartX = (MaximumDisks - NumDisks) * PegWidth
      StartY = (MaximumDisks - NumDisks) * DiskHeight
      PegHeight = (NumDisks + 1) * DiskHeight + 1
      
      Peg = Pegs(PegIndex)
      
      Call BitBlt(picPegBitmap.hdc, StartX, StartY, PegWidth, PegHeight, picPegMask.hdc, StartX, StartY, SRCINVERTANDDEST)    'vbSrcInvert And vbNotSrcCopy
      Call BitBlt(destHdc, Peg.x, Peg.y, PegWidth, PegHeight, picPegMask.hdc, StartX, StartY, vbSrcAnd)
      Call BitBlt(destHdc, Peg.x, Peg.y, PegWidth, PegHeight, picPegBitmap.hdc, StartX, StartY, vbSrcInvert)
End Sub

Private Sub PrepareToMovePiece()
      Dim I%
      
      Disk = Disks(Priority(0))
      
      Call GetMaskAndImage(0)
      'Bring the selected item to top.
      Call DisplayAPiece(picBoard.hdc, 0)
                
      Call SetRectRgn(rgnMoving, Disk.x, Disk.y, Disk.x + Disk.Width, Disk.y + Disk.Height)
      Call SelectClipRgn(picTmpBoard.hdc, rgnMoving)
      Call BitBlt(picTmpBoard.hdc, Disk.x, Disk.y, Disk.Width, Disk.Height, picBackground.hdc, Disk.x, Disk.y, vbSrcCopy)
      
      'You have to repaint the pegs.
      For I = 0 To NumPegs
            Peg = Pegs(I)
            Call SetRectRgn(rgnOther, Peg.x, Peg.y, Peg.x + PegWidth, Peg.y + PegHeight)
            If CombineRgn(rgnCombine, rgnMoving, rgnOther, RGN_AND) <> NULLREGION Then
                  DisplayAPeg picTmpBoard.hdc, I
            End If
      Next I
      
      'Paint the disks.
      For I = NumDisks To 1 Step -1
            Disk = Disks(Priority(I))
            Call SetRectRgn(rgnOther, Disk.x, Disk.y, Disk.x + Disk.Width, Disk.y + Disk.Height)
            If CombineRgn(rgnCombine, rgnMoving, rgnOther, RGN_AND) <> NULLREGION Then
                  GetMaskAndImage I
                  DisplayAPiece (picTmpBoard.hdc), I
            End If
      Next I
      
      Call SelectClipRgn(picTmpBoard.hdc, 0)
      GetMaskAndImage 0
      
      MovingPiece = True
End Sub

Private Sub MovePiece(x As Single, y As Single)
      Dim LastDiskX%, LastDiskY%
      Dim TopPiece%
      Dim tmpDiskX%, tmpDiskY%
      
      TopPiece = Priority(0)
      
      With Disks(TopPiece)
            LastDiskX = .x
            LastDiskY = .y
            tmpDiskX = .x + (x - OldMouseX)
            tmpDiskY = .y + (y - OldMouseY)
            If tmpDiskX <= 0 Then tmpDiskX = 0
            If tmpDiskX >= (frmHanoi.ScaleWidth - .Width) Then tmpDiskX = (frmHanoi.ScaleWidth - .Width)
            If tmpDiskY <= 0 Then tmpDiskY = 0
            If tmpDiskY >= (frmHanoi.ScaleHeight - .Height) Then tmpDiskY = (frmHanoi.ScaleHeight - .Height)
            
            .x = tmpDiskX
            .y = tmpDiskY
      End With
      
      DisplayAPiece (picBoard.hdc), 0
      
      Call SetRectRgn(rgnLast, LastDiskX, LastDiskY, LastDiskX + Disk.Width, LastDiskY + Disk.Height)
      Call SetRectRgn(rgnMoving, Disks(TopPiece).x, Disks(TopPiece).y, Disks(TopPiece).x + Disk.Width, Disks(TopPiece).y + Disk.Height)
      Call CombineRgn(rgnCombine, rgnLast, rgnMoving, RGN_DIFF)
      Call SelectClipRgn(picBoard.hdc, rgnCombine)
      Call BitBlt(picBoard.hdc, LastDiskX, LastDiskY, Disk.Width, Disk.Height, picTmpBoard.hdc, LastDiskX, LastDiskY, vbSrcCopy)
      Call SelectClipRgn(picBoard.hdc, 0)
      
      OldMouseX = x
      OldMouseY = y
End Sub

Sub DetectMouse(x As Single, y As Single)
      Dim I As Integer
      
      For I = 0 To NumDisks
            Disk = Disks(Priority(I))
            If (x > Disk.x) And (x < Disk.x + Disk.Width) Then
                  If (y > Disk.y) And (y < Disk.y + Disk.Height) Then
                        'Move the disk only if the mouse is over the disk.
                        If GetPixel(picWholeMask.hdc, Disk.HomeX + (x - Disk.x), Disk.HomeY + (y - Disk.y)) = vbRed Then
                              If Disk.Index = GetTopDisk(Disk.CurrentPeg) Then
                                    OldDiskX = Disk.x
                                    OldDiskY = Disk.y
                                    OldMouseX = x
                                    OldMouseY = y
                                    Call SetPiecePriority(I)
                                    Call PrepareToMovePiece
                                    Call LockMouseToForm(Me)
                                    Exit For
                              End If
                        End If
                  End If
            End If
      Next I
End Sub

Private Sub mnuHowToPlay_Click()
      Call DisplayHowToPlay
End Sub

Private Sub mnuNew_Click()
      frmNew.Show vbModal
End Sub

Private Sub mnuOptions_Click()
      frmOption.Show vbModal
End Sub

Private Sub picBoard_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
      If Not boolNewGame Then Exit Sub
      If MoveType = DRAG_MOVE Then
            Call DetectMouse(x, y)
      End If
End Sub

Private Sub picBoard_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
      If MovingPiece Then MovePiece x, y
End Sub

Private Sub picBoard_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
      If Not boolNewGame Then Exit Sub
      Dim I As Integer       'Counter variable
            
      'Check the type of moving.
      If MoveType = CLICK_MOVE Then
            If Not MovingPiece Then
                  Call DetectMouse(x, y)
                  Exit Sub
            End If
      End If
      
      If Not MovingPiece Then Exit Sub
      
      'The rest are executed when MovingPiece is True
      Dim DestPeg As Integer  'Destination peg
      Dim RetVal As Integer     'The return value of <GetTopDisk> function
      Dim LastDiskX As Single  'The disk's last X position when
                                                'the mouse button is released
      Dim LastDiskY As Single  'The disk's last Y position
                                                'when the mouse button is released
      Dim boolOverPeg As Boolean  'Flag set to True to indicate that
                                                      'the disk is over one of the pegs
      Dim rctDisk As RECT   'The rectangle structure of the disk.
                                            'Used to determine if the disk is over
                                            'one of the pegs' vicinity.
      
      LastDiskX = Disks(Priority(0)).x
      LastDiskY = Disks(Priority(0)).y
      
      With rctDisk
            .Left = Disks(Priority(0)).x
            .Top = Disks(Priority(0)).y
            .Right = .Left + Disks(Priority(0)).Width
            .Bottom = .Top + Disks(Priority(0)).Height
      End With
      
      'Check if the disk is over one of the pegs.
      boolOverPeg = False
      For I = 0 To NumPegs
            If OverLap(rctPeg(I), rctDisk) Then
                  DestPeg = I
                  boolOverPeg = True
                  Exit For
            End If
      Next I
      
      Select Case boolOverPeg
      'It is over one of the pegs. Now, we must determine
      'whether to put this disk in this peg or not.
      Case Is = True
            'Get the top disk of the destination peg.
            RetVal = GetTopDisk(DestPeg)
            
            '-1 means there is no disk in this peg. So, we put it
            'at the bottom of the destination peg.
            If RetVal = -1 Then
                  Disks(Priority(0)).x = Pegs(DestPeg).x + (PegWidth - Disks(Priority(0)).Width) / 2
                  Disks(Priority(0)).y = (Pegs(DestPeg).y + PegHeight) - DiskHeight
                  Disks(Priority(0)).CurrentPeg = DestPeg
                  
                  MoveCount = MoveCount + 1
                  If boolSoundOn Then PlayWave "move.wav"
                  
            'One or more disks are in the destination peg.
            'We need to compare the disk with the top disk of
            'the destination peg.
            Else
                  
                  'If the disk is smaller than the top disk of
                  'destination peg, put it on this top disk.
                  If Disks(Priority(0)).Index < GetTopDisk(DestPeg) Then
                        Disks(Priority(0)).x = Pegs(DestPeg).x + (PegWidth - Disks(Priority(0)).Width) / 2
                        Disks(Priority(0)).y = (Disks(RetVal).y - DiskHeight)
                        Disks(Priority(0)).CurrentPeg = I
                        
                        MoveCount = MoveCount + 1
                        If boolSoundOn Then PlayWave "move.wav"
                                          
                  'The disk may be larger than the top disk of
                  'destination peg, or it may be the disk itself
                  'that the user is moving to the old peg.
                  Else
                        
                        'The X and Y distances the disk has been moved.
                        Dim DiffX%, DiffY%
                        
                        'Back into the old box.
                        If Disks(Priority(0)).Index = GetTopDisk(DestPeg) Then
                              Disks(Priority(0)).x = OldDiskX
                              Disks(Priority(0)).y = OldDiskY
                              
                              DiffX = Abs(LastDiskX - Disks(Priority(0)).x)
                              DiffY = Abs(LastDiskY - Disks(Priority(0)).y)
                              
                              'The user is just clicking in dragging mode. Squirrely, restless!!
                              If (DiffX = 0 And DiffY = 0) Then
                                   Call ReleaseMouse
                                    MovingPiece = False
                                    Exit Sub
                              End If
                              
                              'Only play the sound if the disk has been moved
                              'slightly above.
                              If DiffY > 2 Then
                                    If boolSoundOn Then PlayWave "move.wav"
                              End If
                              
                        'The disk is larger than the top disk of dest peg.
                        Else
                              Disks(Priority(0)).x = OldDiskX
                              Disks(Priority(0)).y = OldDiskY
                        End If
                  End If
            End If
            
      'It is not over any of the pegs.
      'Just put it back to its old place.
      Case Is = False
            Disks(Priority(0)).x = OldDiskX
            Disks(Priority(0)).y = OldDiskY
      End Select
      
      'Paint the disk.
      Call BitBlt(picBoard.hdc, LastDiskX, LastDiskY, Disks(Priority(0)).Width, Disks(Priority(0)).Height, picTmpBoard.hdc, LastDiskX, LastDiskY, vbSrcCopy)
      Call DisplayAPiece(picBoard.hdc, 0)
      Call DisplayAPiece(picTmpBoard.hdc, 0)
      Call ReleaseMouse
      
      picTmpBoard.Refresh
      MovingPiece = False
      
      If CheckWinGame Then
            TypicalMoves = (2 ^ (NumDisks + 1)) - 1
            frmWin.Show vbModal
            MoveCount = 0
            boolNewGame = False
      End If
End Sub

Private Sub picBoard_Paint()
      Call BitBlt(picBoard.hdc, 0, 0, picTmpBoard.Width, picTmpBoard.Height, picTmpBoard.hdc, 0, 0, vbSrcCopy)
End Sub

Public Sub PreparePuzzle()
      picTmpBoard.Cls
      
      Dim I%
      
      'Blit the three Pegs.
      For I = 0 To NumPegs
            Call DisplayAPeg(picTmpBoard.hdc, I)
      Next I
      
      PegHeight = ((NumDisks + 1) * DiskHeight) + 1
      
      For I = 0 To NumDisks
            With Disks(I)
                  .Index = I
                  .CurrentPeg = SOURCE_PEG
                  .Height = DiskHeight
                  .Width = SmallestDiskWidth + (DiskWidthIncrement * I)
                  .x = Pegs(0).x + (PegWidth - .Width) / 2
                  .y = (Pegs(0).y + PegHeight) - ((NumDisks + 1) - I) * DiskHeight
                  .HomeX = 0
                  .HomeY = DiskHeight * I
            End With
            
            Priority(I) = I
      Next I
      
      For I = NumDisks To 0 Step -1
            Call GetMaskAndImage(I)
            Call DisplayAPiece(picTmpBoard.hdc, I)
      Next I
      
      Call DrawBottomLine
      picBoard.Refresh
End Sub

Function CheckWinGame() As Boolean
      Dim I As Integer
      CheckWinGame = True
      For I = 0 To NumDisks
            If Disks(I).CurrentPeg <> DESTINATION_PEG Then
                  CheckWinGame = False
                  Exit Function
            End If
      Next I
End Function
