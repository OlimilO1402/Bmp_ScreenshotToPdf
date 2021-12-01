Attribute VB_Name = "MNew"
Option Explicit

Function Screenshot(aPB As PictureBox, SrcRect As WinAPIRect) As Screenshot
    Set Screenshot = New Screenshot: Screenshot.New_ aPB, SrcRect
End Function

Public Function WinAPIRect(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long) As WinAPIRect
    With WinAPIRect: .Left = x: .Top = y: .Right = x + w: .Bottom = y + h: End With
End Function

Public Function SScreen(aDstPB As PictureBox, SrcRect As WinAPIRect) As SScreen
    Set SScreen = New SScreen: SScreen.New_ aDstPB, SrcRect
End Function

Public Function StdPicBmp(aStdPic As StdPicture) As StdPicBmp
    Set StdPicBmp = New StdPicBmp: StdPicBmp.New_ aStdPic
End Function

Public Function List(Of_T As EDataType, _
                     Optional ArrColStrTypList, _
                     Optional ByVal IsHashed As Boolean = False, _
                     Optional ByVal Capacity As Long = 32, _
                     Optional ByVal GrowRate As Single = 2, _
                     Optional ByVal GrowSize As Long = 0) As List
    Set List = New List: List.New_ Of_T, ArrColStrTypList, IsHashed, Capacity, GrowRate, GrowSize
End Function

