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
'Private Type WinAPIRect
'    Left   As Long
'    Top    As Long
'    Right  As Long
'    Bottom As Long
'End Type
'Private Type RECT
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type
'Public Type WinAPIPoint
'    X As Long
'    y As Long
'End Type

'Private m_Rect As WinAPIRect


'the deviceContext where to draw
Private m_hDC     As LongPtr
Private m_Rect    As WndRect
Private m_Visible As Boolean
Private Declare Function DrawFocusRect Lib "user32" (ByVal hhdc As LongPtr, lpRect As Any) As Long 'WinAPIRect) As Long

Private Sub Class_Initialize()
    Set m_Rect = New WndRect
End Sub

Friend Sub New_(ByVal ahDC As LongPtr)
    m_hDC = ahDC
End Sub

Private Sub Class_Terminate()
    Me.Delete
End Sub

Public Property Get Visible() As Boolean
    Visible = m_Visible
End Property
Public Sub Delete()
    If Not m_Visible Then Exit Sub
    DrawFocusRect m_hDC, m_Rect
End Sub

Public Property Get WndRect() As WndRect
    Set WndRect = m_Rect
End Property

'Public Function FocusRect(ByVal ahDC As LongPtr) As FocusRect
'    Set FocusRect = New FocusRect: FocusRect.New_ ahDC
'End Function
'
'Friend Function Equals(other As WinAPIRect) As Boolean
'    With m_Rect
'        If .Left <> other.Left Then Exit Function
'        If .Top <> other.Top Then Exit Function
'        If .Right <> other.Right Then Exit Function
'        If .Bottom <> other.Bottom Then Exit Function
'    End With
'    Equals = True
'End Function

Friend Sub Draw(NewRect As WndRect) 'WinAPIRect)

'Besser das FocusRect 3-Pixel breit machen
    If m_Rect.Equals(NewRect) Then Exit Sub
    'If Me.Equals(NewRect) Then Exit Sub
    'Zuerst das alte Rect löschen
    DrawFocusRect m_hDC, ByVal m_Rect.Ptr
    'Dann  das neue Rect zeichnen
    m_Rect.NewC NewRect
    DrawFocusRect m_hDC, ByVal m_Rect.Ptr
    m_Visible = Not m_Visible
End Sub
