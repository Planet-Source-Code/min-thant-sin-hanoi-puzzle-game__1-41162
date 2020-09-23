Attribute VB_Name = "APIFunctions"
Option Explicit

Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long

Function OverLap(ObjectRectA As RECT, ObjectRectB As RECT) As Boolean
      'This sub takes two rectangles and determines if they overlap each other
      Dim ReturnRect As RECT  'rectangle structure
      
      'If the two rectangles overlap in any way
      If IntersectRect(ReturnRect, ObjectRectA, ObjectRectB) Then
            OverLap = True
      End If
End Function

Public Function GetTopDisk(ByVal PegIndex As Integer) As Integer
      Dim I As Integer
      Dim TopDisk As Integer  'The disk with minimum index value.
      Dim boolGotADisk As Boolean
      
      'Set to -1 to indicate that there are no disks in the current peg.
      TopDisk = -1
      boolGotADisk = False
      
      'Get a disk.
      For I = 0 To NumDisks
            If Not boolGotADisk Then
                  If Disks(I).CurrentPeg = PegIndex Then
                        TopDisk = Disks(I).Index
                        boolGotADisk = True
                  End If
            Else
                  If Disks(I).Index < TopDisk Then
                        TopDisk = Disks(I).Index
                  End If
            End If
      Next I
      
      'Return minimum disk index
      GetTopDisk = TopDisk
End Function

Public Sub SetPiecePriority(DiskIndex As Integer)
      Dim Temp%, I%
      Temp = Priority(DiskIndex)
      For I = DiskIndex To 1 Step -1
            Priority(I) = Priority(I - 1)
      Next
      Priority(0) = Temp
End Sub

Public Sub CleanUp()
      'Clean up the mess I've done.
      Call DeleteObject(rgnMoving)
      Call DeleteObject(rgnLast)
      Call DeleteObject(rgnOther)
      Call DeleteObject(rgnCombine)
      Set frmHanoi.picTmpBoard.Picture = Nothing
End Sub

Sub LockMouseToForm(frmDest As Form)
      Dim ClipRect As RECT, ClipPt As POINTAPI
      Dim dummy&
      
      'Clip the cursor to the form.
      'Get the SCREEN coordinates of the form's origin
      ClipPt.x = 0
      ClipPt.y = 0
      dummy = ClientToScreen(frmDest.hwnd, ClipPt)
      
      'Set the clip rectangle
      With ClipRect
            .Top = ClipPt.y     'Top of form
            .Left = ClipPt.x   'Left of form, etc.
            .Right = frmDest.ScaleWidth + ClipPt.x
            .Bottom = frmDest.ScaleHeight + ClipPt.y
      End With
      
      dummy = ClipCursor(ClipRect) 'Clip the cursor to the ClipRect rectangle
      'Call ShowCursor(False)
End Sub

Sub ReleaseMouse()
      'Unclip the cursor
      Call ClipCursorBynum(0&)
      'Call ShowCursor(True)
End Sub
