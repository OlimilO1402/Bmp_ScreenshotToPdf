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
'class Screenshot, it holds the Rect, the StdPic and a Name
Private m_Name    As String
Private m_PicBmp  As StdPicBmp
Private m_SrcRect As WndRect

Friend Sub New_(PicBmp As StdPicBmp, SrcRect As WndRect)
    Set m_PicBmp = PicBmp: Set m_SrcRect = SrcRect
End Sub
'Public Function Screenshot(PicBmp As StdPicBmp, SrcRect As WndRect) As Screenshot
'    Set Screenshot = New Screenshot: Screenshot.New_ PicBmp, SrcRect
'End Function

Public Property Get Name() As String
    Name = m_Name
End Property
Public Property Let Name(ByVal Value As String)
    m_Name = Value
End Property

Public Property Get Picture() As StdPicBmp
    Set Picture = m_PicBmp
End Property

Friend Property Get SrcRect() As WndRect
    Set SrcRect = m_SrcRect
End Property
