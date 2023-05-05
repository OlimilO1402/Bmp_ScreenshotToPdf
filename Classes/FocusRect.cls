VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FocusRect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'the deviceContext where to draw
Private m_hDC     As LongPtr
Private m_Rect    As WndRect
Private m_Visible As Boolean
'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-drawfocusrect
Private Declare Function DrawFocusRect Lib "user32" (ByVal hhdc As LongPtr, lpRect As Any) As Long 'WinAPIRect) As Long
'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-releasedc
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long 'Ptr

Private Sub Class_Initialize()
    Set m_Rect = New WndRect
End Sub

Friend Sub New_(ByVal ahDC As LongPtr)
    m_hDC = ahDC
End Sub
'Public Function FocusRect(ByVal ahDC As LongPtr) As FocusRect
'    Set FocusRect = New FocusRect: FocusRect.New_ ahDC
'End Function

Private Sub Class_Terminate()
    'Me.Delete 'nope leads to ghost-FocusRects
End Sub

Public Property Get Visible() As Boolean
    Visible = m_Visible
End Property
Public Sub Delete()
    If Not m_Visible Then Exit Sub
    DrawFocusRect m_hDC, ByVal m_Rect.Ptr
    m_Visible = False
End Sub

Public Property Get WndRect() As WndRect
    Set WndRect = m_Rect
End Property

Friend Sub Draw(NewRect As WndRect)
    'Maybe better we make it 3-Pixel thick? no
    If m_Rect.Equals(NewRect) Then Exit Sub
    If m_Visible Then
        'First delete the old Rect
        'DrawFocusRect m_hDC, ByVal m_Rect.Ptr
        Delete
    End If
    'Then we draw the new Rect
    m_Rect.NewC NewRect
    DrawFocusRect m_hDC, ByVal m_Rect.Ptr
    m_Visible = True ' Not m_Visible
End Sub
