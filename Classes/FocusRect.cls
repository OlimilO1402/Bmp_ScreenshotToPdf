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

Private m_Rect As WinAPIRect
Private Declare Function DrawFocusRect Lib "user32" (ByVal hhdc As LongPtr, lpRect As WinAPIRect) As Long
'the deviceContext where to draw
Private m_hDC As LongPtr
Private m_Visible As Boolean

Friend Sub New_(ByVal ahDC As LongPtr)
    m_hDC = ahDC
End Sub

Public Property Get Visible() As Boolean
    Visible = m_Visible
End Property
'Public Function FocusRect(ByVal ahDC As LongPtr) As FocusRect
'    Set FocusRect = New FocusRect: FocusRect.New_ ahDC
'End Function

Friend Function Equals(other As WinAPIRect) As Boolean
    With m_Rect
        If .Left <> other.Left Then Exit Function
        If .Top <> other.Top Then Exit Function
        If .Right <> other.Right Then Exit Function
        If .Bottom <> other.Bottom Then Exit Function
    End With
    Equals = True
End Function

Friend Sub Draw(NewRect As WinAPIRect)

'Besser das FocusRect 3-Pixel breit machen

    If Me.Equals(NewRect) Then Exit Sub
    'Zuerst das alte Rect löschen
    DrawFocusRect m_hDC, m_Rect
    'Dann  das neue Rect zeichnen
    m_Rect = NewRect
    DrawFocusRect m_hDC, m_Rect
    m_Visible = Not m_Visible
End Sub
