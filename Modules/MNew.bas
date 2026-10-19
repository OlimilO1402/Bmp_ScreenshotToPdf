Attribute VB_Name = "MNew"
Option Explicit

Public Function SScreen(DstPB As PictureBox, SrcRect As WndRect) As SScreen
    Set SScreen = New SScreen: SScreen.New_ DstPB, SrcRect
End Function

Public Function Screenshot(PicBmp As StdPicBmp, SrcRect As WndRect) As Screenshot
    Set Screenshot = New Screenshot: Screenshot.New_ PicBmp, SrcRect
End Function

Public Function WinAPIRect(ByVal X As Long, ByVal y As Long, ByVal W As Long, ByVal H As Long) As WinAPIRect
    With WinAPIRect: .Left = X: .Top = y: .Right = X + W: .Bottom = y + H: End With
End Function
'
'Public Function WinAPIRect(ByVal L As Long, ByVal T As Long, ByVal R As Long, ByVal B As Long) As WinAPIRect
'    With WinAPIRect: .Left = L: .Top = T: .Right = R: .Bottom = B: End With
'End Function

Public Function WndRect(r As WinAPIRect) As WndRect
    Set WndRect = New WndRect: WndRect.New_ r
End Function
Public Function WndRectHWnd(ByVal hwnd As LongPtr) As WndRect
    Set WndRectHWnd = New WndRect: WndRectHWnd.NewFromHWnd hwnd
End Function
Public Function WndRectFromMousePoint(p As WinAPIPoint) As WndRect
    Set WndRectFromMousePoint = New WndRect: WndRectFromMousePoint.NewFromMousePoint p
End Function

Public Function WinAPIPoint(ByVal X As Long, ByVal y As Long) As WinAPIPoint
    With WinAPIPoint: .X = X: .y = y: End With
End Function

Public Function WinAPISize(ByVal W As Long, ByVal H As Long) As WinAPISize
    With WinAPISize: .Width = W: .Height = H: End With
End Function

Public Function StdPicBmp(StdPic As StdPicture) As StdPicBmp
    Set StdPicBmp = New StdPicBmp: StdPicBmp.New_ StdPic
End Function

Public Function FocusRect(ByVal ahDC As LongPtr) As FocusRect
    Set FocusRect = New FocusRect: FocusRect.New_ ahDC
End Function

Public Function List(Of_T As EDataType, _
                     Optional ArrColStrTypList, _
                     Optional ByVal IsHashed As Boolean = False, _
                     Optional ByVal Capacity As Long = 32, _
                     Optional ByVal GrowRate As Single = 2, _
                     Optional ByVal GrowSize As Long = 0) As List
    Set List = New List: List.New_ Of_T, ArrColStrTypList, IsHashed, Capacity, GrowRate, GrowSize
End Function

Public Function PathFileName(ByVal aPathFileName As String, _
                     Optional ByVal aFileName As String, _
                     Optional ByVal aExt As String) As PathFileName
    Set PathFileName = New PathFileName: PathFileName.New_ aPathFileName, aFileName, aExt
End Function


