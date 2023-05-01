VERSION 5.00
Begin VB.Form FScreenshotBorders 
   Appearance      =   0  '2D
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Änderbares Werkzeugfenster
   Caption         =   "Drag window-borders over desired screenshot-range and close the window!"
   ClientHeight    =   8340
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
End
Attribute VB_Name = "FScreenshotBorders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Result As VbMsgBoxResult
Private m_MyRect As WndRect
Private spx As Single

Private Sub Form_Load()
    spx = Screen.TwipsPerPixelX
    m_Result = 0
End Sub

Friend Function ShowDialog(Fowner As Form, Rect_out As WndRect) As VbMsgBoxResult
    Set m_MyRect = Rect_out.Clone
    spx = Screen.TwipsPerPixelX
    With m_MyRect
        Me.Move .Left * spx, .Top * spx, .Width * spx, .Height * spx
    End With
    Me.Show vbModal, Fowner
    ShowDialog = m_Result
    If ShowDialog = vbOK Then
        Rect_out.NewC m_MyRect
    End If
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        m_Result = vbCancel
        Unload Me
    End If
End Sub

Private Sub Form_Resize()
    GlassForm Me
    m_MyRect.Left = Me.Left / spx
    m_MyRect.Top = Me.Top / spx
    m_MyRect.Width = Me.Width / spx
    m_MyRect.Height = Me.Height / spx
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If m_Result <> vbCancel Then
        m_MyRect.Left = Me.Left / spx
        m_MyRect.Top = Me.Top / spx
        m_MyRect.Width = Me.Width / spx
        m_MyRect.Height = Me.Height / spx
        m_Result = vbOK
    End If
End Sub

