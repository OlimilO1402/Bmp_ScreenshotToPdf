VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Form1"
   ClientHeight    =   7455
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   11055
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7455
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnInfo 
      Caption         =   "Info"
      Height          =   375
      Left            =   9720
      TabIndex        =   14
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton BtnClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   8400
      TabIndex        =   12
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton BtnSavePictures 
      Caption         =   "SavePictures"
      Height          =   375
      Left            =   7080
      TabIndex        =   15
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton BtnPrintToPDF 
      Caption         =   "Create PDF"
      Height          =   375
      Left            =   5760
      TabIndex        =   11
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton BtnScreenshot 
      Caption         =   "Screenshot"
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton BtnGetWnd 
      Caption         =   "Set Wnd Rect"
      Height          =   375
      Left            =   3120
      TabIndex        =   13
      ToolTipText     =   "Move mouse over window & hit Enter"
      Top             =   120
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3120
      Top             =   240
   End
   Begin VB.TextBox TxtL 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   480
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
   Begin VB.ListBox LBPicList 
      Height          =   6300
      ItemData        =   "Form1.frx":1782
      Left            =   0
      List            =   "Form1.frx":1784
      TabIndex        =   10
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox TxtW 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox TxtH 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox TxtT 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   1320
      TabIndex        =   7
      Top             =   0
      Width           =   855
   End
   Begin VB.PictureBox PBScreenshot 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Height          =   6255
      Left            =   2160
      ScaleHeight     =   6195
      ScaleWidth      =   8355
      TabIndex        =   1
      Top             =   840
      Width           =   8415
   End
   Begin VB.Label Label4 
      Caption         =   "Width"
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   270
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Height"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   510
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Left"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   270
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Top"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   30
      Width           =   375
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "mnuPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuListMoveUp 
         Caption         =   "Move up ^"
      End
      Begin VB.Menu mnuListMoveDown 
         Caption         =   "Move down v"
      End
      Begin VB.Menu mnuListDeleteItem 
         Caption         =   "Delete"
         Shortcut        =   {DEL}
      End
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Screen As Screenshot
Private FNm As String
Private i As Long
'Private PicList As Collection
Private PicList As List
Private m_FocusRect As FocusRect
Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As WinAPIPoint) As Long

Private Sub Form_Load()
    Me.Caption = App.EXEName & " v" & App.Major & "." & App.Minor & "." & App.Revision
    Set PicList = MNew.List(vbObject)
    FNm = "C:\TestDir\"
    TxtL.Text = 1
    TxtT.Text = 84
    TxtW.Text = 672 'CLng(905 * CDbl(210) / CDbl(297))
    TxtH.Text = 913
    Set m_Screen = MNew.Screenshot(Me.PBScreenshot, GetWinAPIRect)
    Set m_FocusRect = MNew.FocusRect(MwinAPI.GetDC(0))
    BtnClear_Click
End Sub

Private Sub BtnInfo_Click()
    MsgBox App.CompanyName & " " & App.EXEName & " v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & App.FileDescription
End Sub

Private Sub BtnPrintToPDF_Click()
    Dim i As Long
    For i = 0 To Printers.Count - 1
        If Printers(i).DeviceName = "Microsoft Print to PDF" Then
            Set Printer = Printers(i)
            Exit For
        End If
    Next
    If Printer Is Nothing Then
        MsgBox "Printer not found: 'Microsoft Print to PDF'"
        Exit Sub
    End If
    Dim dpi    As Single:    dpi = 96   'dots per inch
    Dim ppi    As Single:    ppi = 72   'point per inch
    Dim mmpi   As Single:   mmpi = 25.4 'mm per inch
    Dim DA4_w  As Single:  DA4_w = 210  'mm
    Dim DA4_h  As Single:  DA4_h = 297  'mm
    Dim marg_L As Single: marg_L = 0 '5
    Dim marg_R As Single: marg_R = 0 '5
    Dim wA4    As Single:    wA4 = DA4_w - marg_L - marg_R
    Dim TPPX   As Single:   TPPX = Screen.TwipsPerPixelX
    Dim TPPY   As Single:   TPPX = Screen.TwipsPerPixelY
    
    Dim sc_w As Single
    Dim sc_h As Single
Try: On Error GoTo Catch
    'With Printer
        '.ScaleMode = ScaleModeConstants.vbMillimeters
        '.CurrentX = 5
        '.CurrentY = 5
        '.ScaleMode = ScaleModeConstants.vbPixels
        Dim w As Long
        Dim h As Long
        Dim pic As StdPicture
        Dim c As Long: c = PicList.Count
        For i = 0 To PicList.Count - 1
            Set pic = PicList.Item(i)
            w = PBScreenshot.ScaleX(pic.Width, ScaleModeConstants.vbHimetric, ScaleModeConstants.vbPixels)
            h = PBScreenshot.ScaleY(pic.Height, ScaleModeConstants.vbHimetric, ScaleModeConstants.vbPixels)
            
            sc_w = wA4 / ((w / dpi) * mmpi)
            'sc_h = wA4 / ((h / dpi) * mmpi)
            'Debug.Print sc
            'sc = 1.184628041
            'Debug.Print pic.Width
            Printer.PaintPicture pic, 0, 0, pic.Width * sc_w, pic.Height * sc_w, 0, 0, pic.Width, pic.Height
            If i < c Then
                Printer.NewPage
            End If
        Next
        Printer.EndDoc
        '.KillDoc
    'End With
    'Set Printer = Nothing
    Exit Sub
Catch:
    'If MsgBox("Retry?", vbInformation Or vbRetryCancel) = vbRetry Then GoTo Try
End Sub

Private Sub BtnSavePictures_Click()
    Dim i As Long
    For i = 0 To LBPicList.ListCount - 1
        LBPicList.ListIndex = i
        SavePicture PBScreenshot.Image, FNm & "\Bild_" & CStr(i) & ".bmp"
    Next
End Sub

Private Sub Form_Resize()
    Dim L As Single: L = 0
    Dim T As Single: T = LBPicList.Top
    Dim w As Single: w = LBPicList.Width
    Dim h As Single: h = Me.ScaleHeight - T
    If w > 0 And h > 0 Then LBPicList.Move L, T, w, h
End Sub

Private Sub LBPicList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KeyCodeConstants.vbKeyDelete Then
        mnuListDeleteItem_Click
    End If
End Sub

Private Sub LBPicList_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If LBPicList.ListCount > 0 Then
        If Button = MouseButtonConstants.vbRightButton Then
            PopupMenu mnuPopUp
        End If
    End If
End Sub

Private Sub mnuListMoveUp_Click()
    Dim c As Long: c = LBPicList.ListCount
    If c = 1 Then Exit Sub
    Dim i As Long: i = LBPicList.ListIndex
    If i <= 0 Or (c - 1) < i Then Exit Sub
    PicList.MoveUp i
    LBPicList_MoveUp i
    LBPicList.ListIndex = i - 1
End Sub
Private Sub mnuListMoveDown_Click()
    Dim c As Long: c = LBPicList.ListCount
    If c = 1 Then Exit Sub
    Dim i As Long: i = LBPicList.ListIndex
    If i < 0 Or (c - 1) <= i Then Exit Sub
    PicList.MoveDown i
    LBPicList_MoveDown i
    LBPicList.ListIndex = i + 1
End Sub
Private Sub mnuListDeleteItem_Click()
    Dim c As Long: c = LBPicList.ListCount
    Dim i As Long: i = LBPicList.ListIndex
    If i < 0 Or (c - 1) < i Then Exit Sub
    PicList.Remove i
    LBPicList.RemoveItem i
End Sub

Private Sub LBPicList_MoveUp(ByVal i As Long)
    LBPicList_Swap i - 1, i
End Sub
Private Sub LBPicList_MoveDown(ByVal i As Long)
    LBPicList_Swap i, i + 1
End Sub
Private Sub LBPicList_Swap(ByVal i1 As Long, ByVal i2 As Long)
    Dim tmp As String: tmp = LBPicList.List(i1)
    LBPicList.List(i1) = LBPicList.List(i2)
    LBPicList.List(i2) = tmp
End Sub

Private Sub BtnGetWnd_Click()
    Timer1.Enabled = Not Timer1.Enabled
    If Timer1.Enabled Then Exit Sub
    Set m_Screen = MNew.Screenshot(Me.PBScreenshot, GetWinAPIRect)
End Sub

Private Sub Timer1_Timer()
    Dim p As WinAPIPoint
    Dim hr As Long: hr = GetCursorPos(p)
    If hr = 0 Then Exit Sub
    Dim r As WinAPIRect:  r = m_Screen.RectFromPoint(p)
    TxtL.Text = r.Left
    TxtT.Text = r.Top
    TxtW.Text = r.Right - r.Left
    TxtH.Text = r.Bottom - r.Top
    m_FocusRect.Draw r
End Sub

'Private Sub BtnSet_Click()
'    Set m_Screen = MNew.Screenshot(Me.PBScreenshot, GetWinAPIRect)
'End Sub

Private Sub BtnScreenshot_Click()
    Dim pic As StdPicture: Set pic = m_Screen.Shot
    If pic Is Nothing Then
        MsgBox "pic is nothing"
        Exit Sub
    End If
    PicList.Add pic
    LBPicList.AddItem "Bild_" & LBPicList.ListCount
    LBPicList.ListIndex = LBPicList.ListCount - 1
End Sub

Private Sub BtnClear_Click()
    'Set PicList = mNew Collection
    PicList.Clear
    i = 0
    LBPicList.Clear
    'PBScreenshot.AutoRedraw = False
    Set PBScreenshot.Picture = Nothing
    PBScreenshot.Cls
    'PBScreenshot.AutoRedraw = True
End Sub

Private Function GetWinAPIRect() As WinAPIRect
    'sehr suboptimal
    If Not IsNumeric(TxtL.Text) Then Exit Function
    If Not IsNumeric(TxtT.Text) Then Exit Function
    If Not IsNumeric(TxtW.Text) Then Exit Function
    If Not IsNumeric(TxtH.Text) Then Exit Function
    
    Dim X As Long: X = CLng(TxtL.Text)
    Dim y As Long: y = CLng(TxtT.Text)
    Dim w As Long: w = CLng(TxtW.Text)
    Dim h As Long: h = CLng(TxtH.Text)
    GetWinAPIRect = MNew.WinAPIRect(X, y, w, h)
End Function

Private Sub LBPicList_Click()
    'PBScreenshot.Cls
    'PBScreenshot.AutoRedraw = False
    Dim pic As StdPicture
    Dim i As Long: i = LBPicList.ListIndex
    Set pic = PicList.Item(i)
    If pic Is Nothing Then
        MsgBox "pic is nothing"
    End If
    Set PBScreenshot.Picture = pic
    PBScreenshot.Refresh
    'PBScreenshot.AutoRedraw = True
End Sub

Private Sub TxtL_Change()
    If Timer1.Enabled Then Exit Sub
    Set m_Screen = MNew.Screenshot(Me.PBScreenshot, GetWinAPIRect)
End Sub

Private Sub TxtT_Change()
    If Timer1.Enabled Then Exit Sub
    Set m_Screen = MNew.Screenshot(Me.PBScreenshot, GetWinAPIRect)
End Sub

Private Sub TxtW_Change()
    If Timer1.Enabled Then Exit Sub
    Set m_Screen = MNew.Screenshot(Me.PBScreenshot, GetWinAPIRect)
End Sub

Private Sub TxtH_Change()
    If Timer1.Enabled Then Exit Sub
    Set m_Screen = MNew.Screenshot(Me.PBScreenshot, GetWinAPIRect)
End Sub

