VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Screenshot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'http://www.activevb.de/tipps/vb6tipps/tipp0078.html
'VB 5/6-Tipp 0078: Screenshot, Fotografie des Bildschirms erstellen


'http://www.activevb.de/tipps/vb6tipps/tipp0311.html
'VB 5/6-Tipp 0311: Gezielten Snapshot erstellen

#If VBA7 = 0 Then
    Public Enum LongPtr
        [_]
    End Enum
#End If

Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As LongPtr

'Private Declare Function ChildWindowFromPoint Lib "user32" (ByVal hWndParent As LongPtr, ByVal pt As POINTAPI) As LongPtr
Private Declare Function ChildWindowFromPoint Lib "user32" (ByVal hWndParent As LongPtr, ByVal xPoint As Long, ByVal yPoint As Long) As LongPtr

Private Declare Function GetDesktopWindow Lib "user32" () As LongPtr

Private Declare Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As WinAPIRect) As Long
        
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
'follows: dst = src

'https://docs.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-bitblt
'BOOL BitBlt(
'  [in] HDC   hdcDst,
'  [in] int   x,
'  [in] int   y,
'  [in] int   cx,
'  [in] int   cy,
'  [in] HDC   hdcSrc,
'  [in] int   x1,
'  [in] int   y1,
'  [in] DWORD rop
');

Private Declare Function BitBlt Lib "gdi32" ( _
    ByVal Dst_hDC As LongPtr, ByVal xDst As Long, ByVal yDst As Long, ByVal wDst As Long, ByVal hDst As Long, _
    ByVal Src_hDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'follows: dst = src

Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long 'Ptr

Private Const SRCCOPY As Long = &HCC0020

Private m_SrcHWnd As LongPtr
Private m_SrcHDC  As LongPtr
Private m_MaxRect As WinAPIRect
Private m_SrcRect As WinAPIRect

Private m_DstPB   As PictureBox
Private m_DstHWnd As LongPtr
Private m_DstHDC  As LongPtr

Private m_DstRect As WinAPIRect

Private Sub Class_Initialize()
    m_SrcHWnd = GetDesktopWindow
    'm_SrcHDC = GetDC(m_SrcHWnd)
    Dim hr As Long: hr = GetWindowRect(m_SrcHWnd, m_MaxRect)
    m_SrcRect = m_MaxRect
    m_DstRect = m_SrcRect
End Sub

Private Sub Class_Terminate()
    Dim hr As Long: hr = ReleaseDC(m_SrcHWnd, m_SrcHDC)
End Sub

Friend Sub New_(aPB As PictureBox, SrcRect As WinAPIRect)
    Set m_DstPB = aPB
    m_DstHWnd = m_DstPB.hWnd
    m_SrcRect = SrcRect
End Sub

Private Function Screenshot(aPB As PictureBox, SrcRect As WinAPIRect) As Screenshot
    Set Screenshot = New Screenshot: Screenshot.New_ aPB, SrcRect
End Function

Public Property Get HandleWnd() As LongPtr
    HandleWnd = m_SrcHWnd
End Property
Public Property Let HandleWnd(ByVal Value As LongPtr)
    m_SrcHWnd = Value
End Property

Public Function Shot() As StdPicture
    
    Dim hr As Long
    m_DstPB.AutoRedraw = True
    
    'm_DstHDC = GetDC(m_DstPB.hwnd)
    'm_DstHDC = m_DstPB.hdc
    
    'hr = GetWindowRect(m_SrcHWnd, m_SrcRect)
    Dim SrcRectWidth  As Long:  SrcRectWidth = m_SrcRect.Right - m_SrcRect.Left
    Dim SrcRectHeight As Long: SrcRectHeight = m_SrcRect.Bottom - m_SrcRect.Top
    
    m_DstPB.Width = SrcRectWidth * Screen.TwipsPerPixelX
    m_DstPB.Height = SrcRectHeight * Screen.TwipsPerPixelY
    
    m_DstHDC = m_DstPB.hdc
    
    'hr = GetWindowRect(m_DstHWnd, m_DstRect)
    
    'm_DstRect = m_SrcRect
    'm_DstRect.Left = 0: m_DstRect.Top = 0
    
    'hr = StretchBlt(m_DstHDC, m_DstRect.Left, m_DstRect.Top, m_SrcRect.Width, m_SrcRect.Height, _
    '                m_SrcHDC, m_SrcRect.Left, m_SrcRect.Top, m_SrcRect.Width, m_SrcRect.Height, SRCCOPY)
    m_SrcHDC = GetDC(m_SrcHWnd)
    
    hr = StretchBlt(m_DstHDC, 0, 0, SrcRectWidth, SrcRectHeight, _
                    m_SrcHDC, m_SrcRect.Left, m_SrcRect.Top, SrcRectWidth, SrcRectHeight, SRCCOPY)
    
    'hr = BitBlt(m_DstHDC, 0, 0, m_SrcRect.Width, m_SrcRect.Height, _
    '            m_SrcHDC, m_SrcRect.Left, m_SrcRect.Top, m_SrcRect.Width, m_SrcRect.Height, SRCCOPY)
    
    
    Set Shot = m_DstPB.Image
    hr = ReleaseDC(m_SrcHWnd, m_SrcHDC)
    'hr = ReleaseDC(m_DstHWnd, m_DstHDC)
    m_DstPB.Refresh
    m_DstPB.AutoRedraw = False
End Function

Private Function Max(V1, V2)
    If V1 > V2 Then Max = V1 Else Max = V2
End Function
Private Function Min(V1, V2)
    If V1 < V2 Then Min = V1 Else Min = V2
End Function

Friend Function RectFromPoint(aPt As WinAPIPoint) As WinAPIRect
    Dim hWndParent As LongPtr: hWndParent = WindowFromPoint(aPt.x, aPt.y)
    Dim hWndChild  As LongPtr: hWndChild = ChildWindowFromPoint(hWndParent, aPt.x, aPt.y)
    Dim hWnd As LongPtr: If hWndChild = 0 Then hWnd = hWndParent Else hWnd = hWndChild
    If hWnd = 0 Then Exit Function
    Dim hr As Long: hr = GetWindowRect(hWnd, RectFromPoint)
End Function

Friend Property Get SrcRect() As WinAPIRect
    SrcRect = m_SrcRect
End Property
Friend Property Let SrcRect(Value As WinAPIRect)
    m_SrcRect = Value
End Property

Public Property Get SrcLeft() As Long
    SrcLeft = m_SrcRect.Left
End Property
Public Property Let SrcLeft(ByVal Value As Long)
    m_SrcRect.Left = Value
End Property

Public Property Get SrcTop() As Long
    SrcTop = m_SrcRect.Top
End Property
Public Property Let SrcTop(ByVal Value As Long)
    m_SrcRect.Top = Value
End Property

Public Property Get SrcWidth() As Long
    SrcWidth = m_SrcRect.Right - m_SrcRect.Left
End Property
Public Property Let SrcWidth(ByVal Value As Long)
    m_SrcRect.Right = m_SrcRect.Left + Value
End Property

Public Property Get SrcHeight() As Long
    SrcHeight = m_SrcRect.Bottom - m_SrcRect.Top
End Property
Public Property Let SrcHeight(ByVal Value As Long)
    m_SrcRect.Bottom = m_SrcRect.Top + Value
End Property


Friend Property Get DstRect() As WinAPIRect
    DstRect = m_DstRect
End Property
Friend Property Let DstRect(Value As WinAPIRect)
    m_DstRect = Value
End Property

Public Property Get DstLeft() As Long
    DstLeft = m_DstRect.Left
End Property
Public Property Let DstLeft(ByVal Value As Long)
    m_DstRect.Left = Value
End Property

Public Property Get DstTop() As Long
    DstTop = m_DstRect.Top
End Property
Public Property Let DstTop(ByVal Value As Long)
    m_DstRect.Top = Value
End Property

Public Property Get DstWidth() As Long
    DstWidth = m_DstRect.Right - m_DstRect.Left
End Property
Public Property Let DstWidth(ByVal Value As Long)
    m_DstRect.Right = m_DstRect.Left + Value
End Property

Public Property Get DstHeight() As Long
    DstHeight = m_DstRect.Bottom - m_DstRect.Top
End Property
Public Property Let DstHeight(ByVal Value As Long)
    m_DstRect.Bottom = m_DstRect.Top + Value
End Property


