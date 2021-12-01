VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StdPicBmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type WinAPIBitmap
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    bmBits       As Long
End Type

Private Declare Function GetObjectW Lib "gdi32" (ByVal hObject As LongPtr, ByVal nCount As Long, ByRef lpObject As Any) As Long

Private m_StdPic As StdPicture
Private m_bmp    As WinAPIBitmap

Friend Sub New_(aStdPic As StdPicture)
    Set m_StdPic = aStdPic
    Dim hr As Long: hr = GetObjectW(m_StdPic.Handle, LenB(m_bmp), m_bmp)
End Sub

Public Property Get Width() As Long
    Width = m_bmp.bmWidth
End Property

Public Property Get Height() As Long
    Height = m_bmp.bmHeight
End Property

Public Property Get Handle() As LongPtr
    Handle = m_StdPic.Handle
End Property

'Public Property Get hDC() As LongPtr
'    '
'End Property
