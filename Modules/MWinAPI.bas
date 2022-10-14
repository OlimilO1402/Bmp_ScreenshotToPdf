Attribute VB_Name = "MwinAPI"
Option Explicit

'#If VBA7 = 0 Then
'    Public Enum LongPtr
'        [_]
'    End Enum
'#End If

Public Type WinAPIRect
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Public Type WinAPIPoint
    X As Long
    Y As Long
End Type

Public Type WinAPISize
    Width  As Long
    Height As Long
End Type

'Public Declare Function DrawFocusRect Lib "user32" (ByVal hhdc As LongPtr, lpRect As WinAPIRect) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr

'Public Function HimetricToPixel(Himetric As Single) As Long
Public Function HimToPix(ByVal Himetric As Single) As Long
    Dim dpi    As Long:    dpi = 96   'dots per inch
    Dim mmpi   As Long:   mmpi = 2540 'mm per inch * 100
    'Dim HiDpi  As Long:  HiDpi = 1440
    HimToPix = (Himetric / mmpi) * dpi 'HiDpi
End Function

