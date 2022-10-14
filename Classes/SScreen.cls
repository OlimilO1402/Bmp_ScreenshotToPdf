VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetDesktopWindow Lib "user32" () As LongPtr

'Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As WinAPIRect) As Long

Private Declare Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr

'https://docs.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-stretchblt
'BOOL StretchBlt(
'  [in] HDC   hdcDst,
'  [in] int   xDst,
'  [in] int   yDst,
'  [in] int   wDst,
'  [in] int   hDst,
'  [in] HDC   hdcSrc,
'  [in] int   xSrc,
'  [in] int   ySrc,
'  [in] int   wSrc,
'  [in] int   hSrc,
'  [in] DWORD rop
');
Private Declare Function StretchBlt Lib "gdi32" ( _
    ByVal Dst_hDC As LongPtr, ByVal xDst As Long, ByVal yDst As Long, ByVal wDst As Long, ByVal hDst As Long, _
    ByVal Src_hDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal wSrc As Long, ByVal hSrc As Long, ByVal dwRop As Long) As Long
'follows: dst = src (Destination=Source)

Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long 'Ptr

Private Const SRCCOPY As Long = &HCC0020


'Screen aka Desktop is Source (Src)
Private m_SrchWnd As LongPtr
Private m_SrchDC  As LongPtr
Private m_SrcRect As WndRect 'WinAPIRect
Private m_MaxRect As WndRect 'WinAPIRect

Private m_DstPB   As PictureBox
Private m_DstHWnd As LongPtr

'm_SrcHDC = GetDC(m_SrcHWnd)
'm_SrcRect = m_MaxRect

Private Sub Class_Initialize()
    m_SrchWnd = GetDesktopWindow
    'Dim hr As Long: hr = GetWindowRect(m_SrchWnd, m_MaxRect)
    Set m_MaxRect = MNew.WndRectHWnd(m_SrchWnd)
End Sub

Private Sub Class_Terminate()
    Dim hr As Long: hr = ReleaseDC(m_SrchWnd, m_SrchDC)
End Sub

Friend Sub New_(aDstPB As PictureBox, SrcRect As WndRect) ' WinAPIRect)
    Set m_DstPB = aDstPB: m_DstHWnd = m_DstPB.hWnd
    Set m_SrcRect = SrcRect
    'SetSrcRect SrcRect
End Sub

'Friend Function SScreen(aDstPB As PictureBox, SrcRect As WinAPIRect) As SScreen
'    Set SScreen = New SScreen: SScreen.New_ aDstPB, SrcRect
'End Function

Private Sub SetSrcRect(SrcRect As WinAPIRect)
    With m_SrcRect
        .Left = Min(m_MaxRect.Right, SrcRect.Right)
        .Top = Min(m_MaxRect.Top, SrcRect.Top)
        .Right = Min(m_MaxRect.Right, SrcRect.Right)
        .Bottom = Min(m_MaxRect.Bottom, SrcRect.Bottom)
    End With
End Sub

Private Function Max(V1, V2)
    If V1 > V2 Then Max = V1 Else Max = V2
End Function
Private Function Min(V1, V2)
    If V1 < V2 Then Min = V1 Else Min = V2
End Function

'Public Property Get Count() As Long
'    'returns how many Screens there are in the System
'End Property
'
'Public Property Get hWnd() As LongPtr
'    hWnd = m_hWnd
'End Property
'Public Property Get hDC() As LongPtr
'    hDC = GetDC(m_hWnd)
'End Property
'
'Public Property Get Index() As Long '1, 2, 3 or 4 or more
'    'returns the Index for this screen
'End Property
'
Public Property Get Width() As Long 'Width in Pixels
    'returns the width in pixels for this screen
    Width = m_MaxRect.Right - m_MaxRect.Left 'Left normally should be 0
End Property

Public Property Get Height() As Long 'Height in Pixels
    'returns the height in pixels for this screen
    Height = m_MaxRect.Bottom - m_MaxRect.Top 'Top normally should be 0
End Property

Public Function Shot() As StdPicBmp
    'm_DstPB.AutoRedraw = False
    
    'Dim DstRect As WinAPIRect: DstRect = MNew.WinAPIRect(x, y, w, h)
    Dim DstHDC As Long: DstHDC = m_DstPB.hDC
    Dim SrchDC As Long: SrchDC = m_SrchDC
    
'    Dim hr As Long: hr = StretchBlt(DstHDC, 0, 0, w, h, _
'                                    SrchDC, x, y, w, h, SRCCOPY)
    Set Shot = m_DstPB.Image
'    hr = ReleaseDC(m_SrchWnd, m_SrchDC)
    
    'm_DstPB.Refresh
    'm_DstPB.AutoRedraw = True
End Function
