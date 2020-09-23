Attribute VB_Name = "MainSub"
Option Explicit

Public Sub Main()
      TypicalMoves = 0
      MoveCount = 0
      
      'We always start the game with 3 disks.
      'For clarity, I do not attempt to write to registry!
      NumDisks = 2
            
      PegHeight = (NumDisks + 1) * DiskHeight + 1
      
      ReDim Disks(NumDisks) As DISKINFO
      ReDim Priority(NumDisks) As Integer
      
      'Always start the game with sound on, and
      'move the disk by clicking.
      boolSoundOn = True
      boolNewGame = True
      MoveType = DRAG_MOVE
      
      Dim nWidth%, nHeight%
      nWidth = Screen.Width / Screen.TwipsPerPixelX
      nHeight = Screen.Height / Screen.TwipsPerPixelY
      
      With frmHanoi
            .picTmpBoard.Picture = .picBackground.Picture
            
            .picBoard.Move 0, 0, nWidth, nHeight
            .picTmpBoard.Move 0, 0, nWidth, nHeight
            .picBackground.Move 0, 0, nWidth, nHeight
            
            .picFinalDisk.Move 0, 0, MaxDiskWidth, DiskHeight
            .picDiskMask.Move 0, 0, MaxDiskWidth, DiskHeight
            .picDiskImage.Move 0, 0, MaxDiskWidth, DiskHeight
      End With
                        
      'Create the necessary regions.
      'Remember you must delete them when your program ends.
      rgnLast = CreateRectRgn(0, 0, 0, 0)
      rgnOther = CreateRectRgn(0, 0, 0, 0)
      rgnMoving = CreateRectRgn(0, 0, 0, 0)
      rgnCombine = CreateRectRgn(0, 0, 0, 0)
                              
      Call DrawPegs
      Call frmHanoi.PreparePuzzle
      Call PrintText
      
      frmHanoi.Show
End Sub

'This sub draws the three pegs.
Public Sub DrawPegs()
      Dim I%
      Dim OffSet As Single
      Dim ScreenWidth As Integer
      
      ScreenWidth = (Screen.Width / Screen.TwipsPerPixelX)
      OffSet = ScreenWidth / 6
      
      For I = 0 To NumPegs
            Pegs(I).x = OffSet + (ScreenWidth / 3 * I)
            Pegs(I).y = MaxPegHeight
            
            With rctPeg(I)
                  .Left = Pegs(I).x - (PegWidth * 4)
                  .Top = 0
                  .Right = Pegs(I).x + PegWidth + (PegWidth * 4)
                  .Bottom = Pegs(I).y + PegHeight
            End With
      Next I
End Sub

'This sub prints the peg numbers under the three pegs.
Public Sub PrintText()
      Dim CaptionWidth%
      Dim I%, nX%, nY%
      Dim PegName$
      
      frmHanoi.picTmpBoard.ForeColor = vbWhite
      frmHanoi.picTmpBoard.FontSize = 14
      
      For I = 0 To NumPegs
            PegName = "Peg " & CStr(I + 1)
            CaptionWidth = frmHanoi.picTmpBoard.TextWidth(PegName)
            nX = Pegs(I).x + (PegWidth - CaptionWidth) / 2
            TextOut frmHanoi.picTmpBoard.hdc, nX, Pegs(I).y + PegHeight, PegName, Len(PegName)
      Next I
      
      'Const strInfo = "Move all the disks from Peg (1) to Peg(3)"
      'TextOut frmHanoi.picTmpBoard.hdc, 0, 0, strInfo, Len(strInfo)
End Sub
