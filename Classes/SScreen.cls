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
'http://www.activevb.de/tipps/vb6tipps/tipp0078.html
'VB 5/6-Tipp 0078: Screenshot, Fotografie des Bildschirms erstellen

'http://www.activevb.de/tipps/vb6tipps/tipp0311.html
'VB 5/6-Tipp 0311: Gezielten Snapshot erstellen

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getdesktopwindow
Private Declare Function GetDesktopWindow Lib "user32" () As LongPtr
'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getdc
Private Declare Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr

'https://docs.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-stretchblt
Private Declare Function StretchBlt Lib "gdi32" ( _
    ByVal Dst_hDC As LongPtr, ByVal xDst As Long, ByVal yDst As Long, ByVal wDst As Long, ByVal hDst As Long, _
    ByVal Src_hDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal wSrc As Long, ByVal hSrc As Long, ByVal dwRop As Long) As Long
'follows: dst = src (Destination=Source)
'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-releasedc
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long 'Ptr

Private Const SRCCOPY As Long = &HCC0020

'Screen aka Desktop is Source (Src)
Private m_MaxRect As WndRect
Private m_DstPB   As PictureBox

Private m_SrchWnd As LongPtr
Private m_SrchDC  As LongPtr
Private m_SrcRect As WndRect

Private m_DsthWnd As LongPtr
Private m_DsthDC  As LongPtr
Private m_DstRect As WndRect

Private Sub Class_Initialize()
    m_SrchWnd = GetDesktopWindow
    Set m_MaxRect = MNew.WndRectHWnd(m_SrchWnd)
    Set m_SrcRect = m_MaxRect.Clone
    Set m_DstRect = m_SrcRect.Clone
End Sub

Private Sub Class_Terminate()
    Dim hr As Long: hr = ReleaseDC(m_SrchWnd, m_SrchDC)
End Sub

Friend Sub New_(DstPB As PictureBox, SrcRect As WndRect)
    Set m_DstPB = DstPB: m_DsthWnd = m_DstPB.hWnd: Set m_SrcRect = SrcRect
End Sub
'Friend Function SScreen(DstPB As PictureBox, SrcRect As WinAPIRect) As SScreen
'    Set SScreen = New SScreen: SScreen.New_ DstPB, SrcRect
'End Function

Public Property Get DesktophDC() As LongPtr
    Dim Desktp_hWnd As LongPtr: Desktp_hWnd = m_SrchWnd
    DesktophDC = GetDC(Desktp_hWnd)
End Property

Public Function Shot() As Screenshot 'StdPicBmp
    Dim hr As Long
    Dim SrcRectWidth  As Long:  SrcRectWidth = m_SrcRect.Right - m_SrcRect.Left
    Dim SrcRectHeight As Long: SrcRectHeight = m_SrcRect.Bottom - m_SrcRect.Top
    m_DstPB.Width = SrcRectWidth   '* Screen.TwipsPerPixelX
    m_DstPB.Height = SrcRectHeight '* Screen.TwipsPerPixelY
    m_DsthDC = m_DstPB.hDC
    m_SrchDC = DesktophDC ' GetDC(m_SrchWnd)
    hr = StretchBlt(m_DsthDC, 0, 0, SrcRectWidth, SrcRectHeight, _
                    m_SrchDC, m_SrcRect.Left, m_SrcRect.Top, SrcRectWidth, SrcRectHeight, SRCCOPY)
    Dim pic As StdPicture: Set pic = m_DstPB.Image
    Set Shot = MNew.Screenshot(MNew.StdPicBmp(pic), m_SrcRect.Clone)
    hr = ReleaseDC(m_SrchWnd, m_SrchDC)
End Function

Friend Property Get SrcRect() As WndRect
    Set SrcRect = m_SrcRect
End Property

Friend Property Get DstRect() As WndRect
    Set DstRect = m_DstRect
End Property

Public Sub ClearScreen()
    MwinAPI.UpdateWindow m_SrchWnd
End Sub
