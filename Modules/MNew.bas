Attribute VB_Name = "MNew"
Option Explicit

Function Screenshot(aPB As PictureBox, SrcRect As WinAPIRect) As Screenshot
    Set Screenshot = New Screenshot: Screenshot.New_ aPB, SrcRect
End Function

Public Function WinAPIRect(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long) As WinAPIRect
    With WinAPIRect: .Left = x: .Top = y: .Right = x + w: .Bottom = y + h: End With
End Function

