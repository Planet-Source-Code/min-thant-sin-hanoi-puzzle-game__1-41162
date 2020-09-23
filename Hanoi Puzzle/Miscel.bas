Attribute VB_Name = "Miscel"
Option Explicit

Sub DisplayHowToPlay()
      Dim msg As String
      msg = "Your goal is to move all the bamboos from Peg (1) to Peg (3)." & vbCrLf & _
             "There are certain rules :" & vbCrLf & _
             "" & vbCrLf & _
             "(1) When a bamboo is moved, it must be placed on one of the three pegs (Peg 1 or Peg 2 or Peg 3)." & vbCrLf & _
             "" & vbCrLf & _
             "(2) Only one disk may be moved at a time, and it must be the top disk on one of the pegs." & vbCrLf & _
             "" & vbCrLf & _
             "(3) A larger disk may never be placed on top of a smaller one."
      MsgBox msg, vbInformation, "How to play this lousy game..."
End Sub

Sub DrawBottomLine()
      Dim LineY%, LineWidth%
      frmHanoi.picTmpBoard.ForeColor = vbWhite
      frmHanoi.picTmpBoard.DrawWidth = 1
      
      LineY = (Pegs(0).Y + PegHeight) + 1
      LineWidth = Screen.Width / Screen.TwipsPerPixelX
      frmHanoi.picTmpBoard.Line (0, LineY)-(LineWidth, LineY)
      
      frmHanoi.picTmpBoard.ForeColor = vbBlack
      frmHanoi.picTmpBoard.Line (0, LineY + 1)-(LineWidth, LineY + 1)
End Sub

Public Sub PlayWave(FileName As String)
      On Error Resume Next
      If FileLen(App.Path & "\" & FileName) > 4200 Then Exit Sub
      sndPlaySound App.Path & "\" & FileName, SND_ASYNC
      If Err Then
            Err = 0
            sndPlaySound "", SND_NODEFAULT
      End If
End Sub
