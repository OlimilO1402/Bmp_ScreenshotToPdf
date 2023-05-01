Attribute VB_Name = "MwinAPI"
Option Explicit

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

        
Const RGN_OR   As Long = 2
Const RGN_DIFF As Long = 4

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As LongPtr
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As LongPtr, ByVal hSrcRgn1 As LongPtr, ByVal hSrcRgn2 As LongPtr, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As LongPtr, ByVal hRgn As LongPtr, ByVal bRedraw As Long) As Long

Public Declare Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr

'Public Function HimetricToPixel(Himetric As Single) As Long
Public Function HimToPix(ByVal Himetric As Single) As Long
    Dim dpi    As Long:    dpi = 96   'dots per inch
    Dim mmpi   As Long:   mmpi = 2540 'mm per inch * 100
    'Dim HiDpi  As Long:  HiDpi = 1440
    HimToPix = (Himetric / mmpi) * dpi 'HiDpi
End Function

Public Function Max(V1, V2)
    If V1 > V2 Then Max = V1 Else Max = V2
End Function
Public Function Min(V1, V2)
    If V1 < V2 Then Min = V1 Else Min = V2
End Function

Public Sub GlassForm(aForm As Form)
    
    If aForm.WindowState = vbMinimized Then Exit Sub
    
    'Ganzen Form als Region festlegen:
    Dim FWidth  As Long:   FWidth = aForm.ScaleX(aForm.Width, vbTwips, vbPixels)
    Dim FHeight As Long:  FHeight = aForm.ScaleY(aForm.Height, vbTwips, vbPixels)
    Dim ROuter As LongPtr: ROuter = CreateRectRgn(0, 0, FWidth, FHeight)
    
    'Ränder und Titel abzziehen & als Region festlegen
    aForm.ScaleMode = vbPixels
    Dim FBorder As Long:  FBorder = (FWidth - aForm.ScaleWidth) / 2
    Dim FTitle  As Long:   FTitle = FHeight - FBorder - aForm.ScaleHeight
    Dim RInner As LongPtr: RInner = CreateRectRgn(FBorder, FTitle, FWidth - FBorder, FHeight - FBorder)
    
    'Innere von der äußeren Region abziehen
    Dim RCombined As LongPtr: RCombined = CreateRectRgn(0, 0, 0, 0)
    CombineRgn RCombined, ROuter, RInner, RGN_DIFF
    
    Call SetWindowRgn(aForm.hWnd, RCombined, True)
End Sub
