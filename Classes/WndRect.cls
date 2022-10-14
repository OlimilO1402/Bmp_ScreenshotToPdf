VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WndRect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Rect As MwinAPI.WinAPIRect

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getwindowrect
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As WinAPIRect) As Long

'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-windowfrompoint
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As LongPtr

'Private Declare Function ChildWindowFromPoint Lib "user32" (ByVal hWndParent As LongPtr, ByVal pt As POINTAPI) As LongPtr
Private Declare Function ChildWindowFromPoint Lib "user32" (ByVal hWndParent As LongPtr, ByVal xPoint As Long, ByVal yPoint As Long) As LongPtr

Friend Sub New_(R As WinAPIRect)
    m_Rect = R
End Sub
Friend Sub NewC(other As WndRect)
    m_Rect.Left = other.Left
    m_Rect.Top = other.Top
    m_Rect.Right = other.Right
    m_Rect.Bottom = other.Bottom
End Sub

Friend Sub NewFromMousePoint(mousepoint As WinAPIPoint)
    Dim hWndParent As LongPtr: hWndParent = WindowFromPoint(mousepoint.X, mousepoint.Y)
    Dim hWndChild  As LongPtr: hWndChild = ChildWindowFromPoint(hWndParent, mousepoint.X, mousepoint.Y)
    'Dim hWnd       As LongPtr: If hWndChild = 0 Then hWnd = hWndParent Else hWnd = hWndChild
    'If hWnd = 0 Then Exit Sub
    'NewFromHWnd hWnd
    NewFromHWnd IIf(hWndChild, hWndChild, hWndParent)
End Sub

Public Sub NewFromHWnd(hWnd As LongPtr)
    Dim b As Boolean: b = GetWindowRect(hWnd, m_Rect)
    If Not b Then ErrHandler "FromHWnd", , b
End Sub

Public Function Clone() As WndRect
    Set Clone = New WndRect: Clone.NewC Me
End Function

Public Function Equals(other As WndRect) As Boolean
    If m_Rect.Left <> other.Left Then Exit Function
    If m_Rect.Top <> other.Top Then Exit Function
    If m_Rect.Right <> other.Right Then Exit Function
    If m_Rect.Bottom <> other.Bottom Then Exit Function
    Equals = True
End Function

Public Function IsSame(other As WndRect) As Boolean
    'works here an for many other classes with the property Ptr
    'IsSame = Me.Ptr = other.Ptr
    'works for every class:
    IsSame = ObjPtr(Me) = ObjPtr(other)
End Function

Friend Property Get WinAPIRect() As WinAPIRect
    WinAPIRect = m_Rect
End Property

Friend Property Get Position() As WinAPIPoint
    Position = MNew.WinAPIPoint(m_Rect.Left, m_Rect.Top)
End Property

Friend Property Get Size() As WinAPISize
    Size = MNew.WinAPISize(Me.Width, Me.Height)
End Property

Public Property Get Ptr() As LongPtr
    Ptr = VarPtr(m_Rect)
End Property

Public Property Get Left() As Long
    Left = m_Rect.Left
End Property
Public Property Let Left(ByVal Value As Long)
    m_Rect.Left = Value
End Property

Public Property Get Top() As Long
    Top = m_Rect.Top
End Property
Public Property Let Top(ByVal Value As Long)
    m_Rect.Top = Value
End Property

Public Property Get Right() As Long
    Right = m_Rect.Right
End Property
Public Property Let Right(ByVal Value As Long)
    m_Rect.Right = Value
End Property

Public Property Get Bottom() As Long
    Bottom = m_Rect.Bottom
End Property
Public Property Let Bottom(ByVal Value As Long)
    m_Rect.Bottom = Value
End Property

Public Property Get Width() As Long
    'Width = VBA.Math.Abs(m_Rect.Right - m_Rect.Left)
    Width = m_Rect.Right - m_Rect.Left
End Property
Public Property Let Width(ByVal Value As Long)
    m_Rect.Right = m_Rect.Left + Value
End Property

Public Property Get Height() As Long
    'Height = VBA.Math.Abs(m_Rect.Bottom - m_Rect.Top)
    Height = m_Rect.Bottom - m_Rect.Top
End Property
Public Property Let Height(ByVal Value As Long)
    m_Rect.Bottom = m_Rect.Top + Value
End Property

Friend Function ToStr() As String
    With m_Rect
        ToStr = TypeName(Me) & " = {L: " & .Left & "; T: " & Me.Top & "; R: " & Me.Right & "; B: " & Me.Bottom & "};"
    End With
End Function

Public Function Move(Optional L, Optional T, Optional W, Optional H) As WndRect
    If Not IsMissing(L) Then m_Rect.Left = L
    If Not IsMissing(T) Then m_Rect.Top = T
    If Not IsMissing(W) Then Me.Width = W
    If Not IsMissing(H) Then Me.Height = H
    Set Move = Me
End Function

Public Function Shrink(Optional ByVal Value As Long = 1) As WndRect
    If Me.Width < Value And Me.Height < Value Then Exit Function
    With m_Rect
        .Left = .Left + Value
        .Top = .Top + Value
        .Right = .Right - Value
        .Bottom = .Bottom - Value
    End With
    Set Shrink = Me
End Function
Public Function ShrinkWOTop(Optional ByVal Value As Long = 1) As WndRect
    If Me.Width < Value And Me.Height < Value Then Exit Function
    With m_Rect
        .Left = .Left + Value
        '.Top = .Top + Value
        .Right = .Right - Value
        .Bottom = .Bottom - Value
    End With
    Set ShrinkWOTop = Me
End Function

Public Function Widen(Optional ByVal Value As Long = 1) As WndRect
    If Me.Width < Value And Me.Height < Value Then Exit Function
    With m_Rect
        .Left = .Left - Value
        .Top = .Top - Value
        .Right = .Right + Value
        .Bottom = .Bottom + Value
    End With
    Set Widen = Me
End Function

Public Function WidenWH(Optional ByVal Value As Long = 1) As WndRect
    If Me.Width < Value And Me.Height < Value Then Exit Function
    With m_Rect
        '.Left = .Left - Value
        '.Top = .Top - Value
        .Right = .Right + Value
        .Bottom = .Bottom + Value
    End With
    Set WidenWH = Me
End Function

''copy this same function to every class, form or module
''the name of the class or form will be added automatically
''in standard-modules the function "TypeName(Me)" will not work, so simply replace it with the name of the Module
'' v ############################## v '   Local ErrHandler   ' v ############################## v '
Private Function ErrHandler(ByVal FuncName As String, _
                            Optional ByVal AddInfo As String, _
                            Optional WinApiError, _
                            Optional bLoud As Boolean = True, _
                            Optional bErrLog As Boolean = True, _
                            Optional vbDecor As VbMsgBoxStyle = vbOKCancel, _
                            Optional bRetry As Boolean) As VbMsgBoxResult

    If bRetry Then

        ErrHandler = MessErrorRetry(TypeName(Me), FuncName, AddInfo, WinApiError, bErrLog)

    Else

        ErrHandler = MessError(TypeName(Me), FuncName, AddInfo, WinApiError, bLoud, bErrLog, vbDecor)

    End If

End Function

