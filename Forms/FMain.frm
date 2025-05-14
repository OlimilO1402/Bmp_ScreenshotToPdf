VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "FMain"
   ClientHeight    =   7140
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   13170
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "FMain"
   ScaleHeight     =   476
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   878
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnInfo 
      Caption         =   "Info"
      Height          =   375
      Left            =   11040
      TabIndex        =   14
      Top             =   360
      Width           =   615
   End
   Begin VB.CommandButton BtnClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   9720
      TabIndex        =   12
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton BtnSavePictures 
      Caption         =   "SavePictures"
      Height          =   375
      Left            =   8400
      TabIndex        =   15
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton BtnOpenPictures 
      Caption         =   "OpenPictures"
      Height          =   375
      Left            =   7080
      TabIndex        =   17
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton BtnPrintToPDF 
      Caption         =   "Create PDF"
      Height          =   375
      Left            =   5760
      TabIndex        =   11
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton BtnScreenshot 
      Caption         =   "Screenshot"
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton BtnDragWndRect 
      Caption         =   "Drag Wnd Rect"
      Height          =   375
      Left            =   3120
      TabIndex        =   16
      ToolTipText     =   "Move mouse over window & hit Enter"
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton BtnGetWnd 
      Caption         =   "Set Wnd Rect"
      Height          =   375
      Left            =   3120
      TabIndex        =   13
      ToolTipText     =   "Move mouse over window & hit Enter"
      Top             =   0
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2520
      Top             =   360
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
      ItemData        =   "FMain.frx":1782
      Left            =   0
      List            =   "FMain.frx":1784
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
      BorderStyle     =   0  'Kein
      Height          =   6255
      Left            =   2160
      ScaleHeight     =   6255
      ScaleWidth      =   8895
      TabIndex        =   1
      Top             =   840
      Width           =   8895
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
Private m_Screen As SScreen 'shot
Private FNm As String
Private i As Long
Private m_ScsList   As List ' As Collection 'Of Screenshot
Private m_FocusRect As FocusRect
Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As WinAPIPoint) As Long
Private OldRect As WndRect
Private bInit As Boolean


'the class SScreen has a function Shot that returns a Screenshot-object
'm_Screen.Shot As Screenshot
'Screenshot.Picture
'What data does it store?
'What data does Screen store?
'What data does Screenshot store?
'The screen stores the source-hwnd
'so screen is basically the desktop so it has (GetDesktopWindow) Dektop_hWnd and Desktop_hDC
'the screenshot stores the source-rect and returns an StdPicBmp-object
'every time the screenshot-button is pressed, a screenshot-object is created
'screen has the function
Private Sub Form_Load()
    bInit = True
    Me.Caption = App.EXEName & " v" & App.Major & "." & App.Minor & "." & App.Revision
    Set m_ScsList = MNew.List(vbObject) 'Of Screenshot)
    FNm = "C:\TestDir\"
    TxtL.Text = 1
    TxtT.Text = 84
    TxtW.Text = 672 'CLng(905 * CDbl(210) / CDbl(297))
    TxtH.Text = 913
    Set m_Screen = MNew.SScreen(Me.PBScreenshot, GetWndRect) ' GetWinAPIRect)
    Set m_FocusRect = MNew.FocusRect(m_Screen.DesktophDC)
    BtnClear_Click
    bInit = False
End Sub

Private Sub Form_LostFocus()
    Debug.Print "Form_LostFocus"
    If Timer1.Enabled Then BtnGetWnd_Click
End Sub
Private Sub Form_Deactivate()
    Debug.Print "Form_Deactivate"
    If Timer1.Enabled Then BtnGetWnd_Click
End Sub

Private Sub Form_Resize()
    Dim L As Single: L = 0
    Dim T As Single: T = LBPicList.Top
    Dim W As Single: W = LBPicList.Width
    Dim H As Single: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then LBPicList.Move L, T, W, H
End Sub

Private Sub BtnInfo_Click()
    MsgBox App.CompanyName & " " & App.EXEName & " v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & App.FileDescription
End Sub

'Private Function SelectPrinter(ByVal PrinterName As String) As Printer
'    Dim i As Long
'    For i = 0 To Printers.Count - 1
'        If Printers(i).DeviceName = PrinterName Then '"Microsoft Print to PDF" Then
'            Set SelectPrinter = Printers(i)
'            Exit For
'        End If
'    Next
'End Function

'Private Function Millimeter_ToTwips(ByVal mm As Double) As Single
'    Dim dpi    As Single:    dpi = 96   ' dots per inch
'    Dim ppi    As Single:    ppi = 72   'point per inch
'    'Dim mmpi   As Single:   mmpi = 25.4 '  mm  per inch
'    Dim TPPX   As Single:   TPPX = Screen.TwipsPerPixelX
'    'Dim sc     As Single. SC 0
'    Millimeter_ToTwips = mm * TPPX * dpi / ppi
'End Function

Private Sub BtnPrintToPDF_Click()
    Dim pn As String: pn = "Microsoft Print to PDF"
    Set Printer = SelectPrinter(pn)
    If Printer Is Nothing Then
        MsgBox "Printer not found: '" & pn & "'"
        Exit Sub
    End If
    
    Dim s As String: s = "Portrait- or landscape-format? (Portrait = Yes, Landscape = No)"
    Dim mbr As VbMsgBoxResult: mbr = MsgBox(s, vbYesNoCancel)
    If mbr = vbCancel Then Exit Sub
    Dim isFmtPortrait As Boolean: isFmtPortrait = mbr = vbYes
    Dim DA4_w  As Single:  DA4_w = IIf(isFmtPortrait, 210, 297)  'mm
    Dim DA4_h  As Single:  DA4_h = IIf(isFmtPortrait, 297, 210)  'mm
    Dim marg_L As Single: marg_L = 0 '5 '0 '5 'mm
    Dim marg_R As Single: marg_R = 0 '5 '0 '5 'mm
    Dim marg_T As Single: marg_T = 0
    Dim marg_B As Single: marg_B = 0
    Dim wA4    As Single:    wA4 = DA4_w - marg_L - marg_R
    Dim hA4    As Single:    hA4 = DA4_h - marg_T - marg_B
    Dim aar    As AARect:    aar = MPGeom.New_AARect(MPGeom.New_Point(marg_L, marg_T), MPGeom.New_Size(wA4, hA4))
    Printer.ScaleMode = ScaleModeConstants.vbPixels
    Dim sc As Double: sc = MPrinter.Millimeter_Scale(Printer.ScaleMode, 1)
    MPGeom.AARect_Mul aar, sc
Try: On Error GoTo Catch
    Printer.Orientation = IIf(isFmtPortrait, PrinterObjectConstants.vbPRORPortrait, PrinterObjectConstants.vbPRORLandscape)
    Dim scs As Screenshot
    Dim pic As StdPicture
    Dim u As Long: u = m_ScsList.Count - 1
    For i = 0 To u
        Set scs = m_ScsList.Item(i)
        Set pic = scs.Picture.StdPicture
        MPrinter.PaintPictureFit pic, aar.Pt.X, aar.Pt.Y, aar.Sz.Width, aar.Sz.Height
        If i < u Then
            Printer.NewPage
        End If
    Next
    Printer.EndDoc
    Exit Sub
Catch:
    'If MsgBox("Retry?", vbInformation Or vbRetryCancel) = vbRetry Then GoTo Try
End Sub

'Private Sub BtnPrintToPDF_Click()
'    Dim pn As String: pn = "Microsoft Print to PDF"
'    Set Printer = SelectPrinter(pn)
'    If Printer Is Nothing Then
'        MsgBox "Printer not found: '" & pn & "'"
'        Exit Sub
'    End If
'
'    Dim S As String: S = "Portrait- or landscape-format? (Portrait = Yes, Landscape = No)"
'    Dim mbr As VbMsgBoxResult: mbr = MsgBox(S, vbYesNoCancel)
'    If mbr = vbCancel Then Exit Sub
'    Dim isFmtPortrait As Boolean: isFmtPortrait = mbr = vbYes
'    'Dim dpi    As Single:    dpi = 96   'dots per inch
'    'Dim ppi    As Single:    ppi = 72   'point per inch
'    'Dim mmpi   As Single:   mmpi = 25.4 'mm per inch
'    Dim DA4_w  As Single:  DA4_w = IIf(isFmtPortrait, 210, 297)  'mm
'    Dim DA4_h  As Single:  DA4_h = IIf(isFmtPortrait, 297, 210)  'mm
'    Dim marg_L As Single: marg_L = 0 '5 '0 '5 'mm
'    Dim marg_R As Single: marg_R = 0 '5 '0 '5 'mm
'    Dim marg_T As Single: marg_T = 0
'    Dim marg_B As Single: marg_B = 0
'    Dim wA4    As Single:    wA4 = DA4_w - marg_L - marg_R
'    Dim hA4    As Single:    hA4 = DA4_h - marg_T - marg_B
'    Dim aar    As AARect:    aar = MPGeom.New_AARect(MPGeom.New_Point(marg_L, marg_T), MPGeom.New_Size(wA4, hA4))
'    'Dim TPPX   As Single:   TPPX = Screen.TwipsPerPixelX
'    'Dim TPPY   As Single:   TPPY = Screen.TwipsPerPixelY
'
'    'Dim sc_w As Single
'    'Dim sc_h As Single
'    Dim sc As Double: sc = MPrinter.Millimeter_Scale(Printer.ScaleMode, 1)
'    MPGeom.AARect_Mul aar, sc
'Try: On Error GoTo Catch
'    'With Printer
'        '.ScaleMode = ScaleModeConstants.vbMillimeters
'        '.CurrentX = 5
'        '.CurrentY = 5
'        '.ScaleMode = ScaleModeConstants.vbPixels
'
'        Printer.Orientation = IIf(isFmtPortrait, PrinterObjectConstants.vbPRORPortrait, PrinterObjectConstants.vbPRORLandscape)
'        'Printer.ScaleMode = ScaleModeConstants.vbPixels
'        'Dim W As Long
'        'Dim H As Long
'        Dim scs As Screenshot
'        Dim pic As StdPicture
'        Dim u As Long: u = m_ScsList.Count - 1
'        For i = 0 To u
'            Set scs = m_ScsList.Item(i)
'            Set pic = scs.Picture.StdPicture
'            'W = PBScreenshot.ScaleX(pic.Width, ScaleModeConstants.vbHimetric, ScaleModeConstants.vbPixels)
'            'H = PBScreenshot.ScaleY(pic.Height, ScaleModeConstants.vbHimetric, ScaleModeConstants.vbPixels)
'
'            'sc_w = wA4 / ((W / dpi) * mmpi)
'            'sc_h = sc_w 'wA4 / ((h / dpi) * mmpi)
'            'Debug.Print sc
'            'sc = 1.184628041
'            'Debug.Print pic.Width
'            'Printer.PSet (100, 100)
'            'Printer.PaintPicture pic, marg_L * TPPX * sc_w * dpi / ppi, marg_L * TPPX * sc_w * dpi / ppi, pic.Width * sc_w, pic.Height * sc_h, 0, 0, pic.Width, pic.Height
'            'Dim X1 As Single: X1 = marg_L * TPPX * sc_w * dpi / ppi
'            'Dim Y1 As Single: Y1 = marg_L * TPPX * sc_w * dpi / ppi
'            'Printer.PaintPicture pic, marg_L * TPPX * sc_w * dpi / ppi, marg_L * TPPX * sc_w * dpi / ppi, pic.Width * sc_w, pic.Height * sc_h, 0, 0, pic.Width, pic.Height
'            MPrinter.PaintPictureFit pic, aar.Pt.X, aar.Pt.Y, aar.Sz.Width, aar.Sz.Height
'            If i < u Then
'                Printer.NewPage
'            End If
'        Next
'        Printer.EndDoc
'        '.KillDoc
'    'End With
'    'Set Printer = Nothing
'    'Debug.Print Printer.DeviceName
'    'Debug.Print Printer.DriverName
'    'Debug.Print Printer.
'    Exit Sub
'Catch:
'    'If MsgBox("Retry?", vbInformation Or vbRetryCancel) = vbRetry Then GoTo Try
'End Sub




'Private Sub mnuFileOpen_Click()
'    BtnOpenPictures_Click
'End Sub

Private Sub BtnOpenPictures_Click()
    'Dim OFD As OpenFileDialog: Set OFD = New OpenFileDialog
    'OFD.Filter = "Bitmaps (*.bmp)|*.bmp|All files (*.*)|*.*"
    'If OFD.ShowDialog(Me) = vbCancel Then Exit Sub
    'Dim PFN As String: PFN = OFD.FileName
    'Dim FD As New OpenFileDialog
    'If Not m_Bmp Is Nothing Then FD.FileName = m_Bmp.FileName
    Dim aPFN As String
    Dim aPFNList As List ': If Not m_bmp Is Nothing Then aPFN = m_bmp.FileName
    Set aPFNList = MMain.GetOpenFileNames(Me, aPFN)
    'If Len(aPFN) = 0 Then Exit Sub
    If aPFNList.IsEmpty Then Exit Sub
    Dim i As Long
    For i = 0 To aPFNList.Count - 1
        aPFN = aPFNList.Item(i)
        Dim pos As Long: pos = InStrRev(aPFN, ".")
        Dim ext As String: ext = LCase(Right(aPFN, Len(aPFN) - pos))
        Dim pic As StdPicture
    'If ext = "bmp" Then
    '    Set m_bmp = MNew.Bitmap(aPFN)
    'Else
        Select Case ext
        Case "png": Set pic = MLoadPng.LoadPictureGDIp(aPFN)
        'Case "gif"
                    'Set PBBitmap.Picture = LoadPicture(aPFN)
                    'Dim ipd As IPictureDisp: Set ipd = LoadPicture(aPFN)
                    'Set PBBitmap.Picture = ipd
                    'Dim sdp As StdPicture: Set sdp = LoadPicture(aPFN)
                    'Set pic = LoadPicture(aPFN)
                    'Set PBBitmap.Picture = sdp
                    'UpdateView
                    'Exit Sub
        ', "jpg": Set pic = LoadPicture(aPFN)
        Case Else:  Set pic = LoadPicture(aPFN)
        End Select
        PBScreenshot.ScaleMode = ScaleModeConstants.vbPixels
        Set PBScreenshot.Picture = pic
        PBScreenshot.ScaleMode = ScaleModeConstants.vbPixels
        Set pic = PBScreenshot.Picture
        Debug.Print pic.Width & " - " & pic.Height
        'Dim r As WndRect: Set r = MNew.WndRect(MwinAPI.New_WinAPIRect(0, 0, pic.Width / Screen.TwipsPerPixelX, pic.Height / Screen.TwipsPerPixelY))
        Dim R As WndRect: Set R = MNew.WndRect(MwinAPI.New_WinAPIRect(0, 0, pic.Width, pic.Height))
        Dim scs As Screenshot: Set scs = MNew.Screenshot(MNew.StdPicBmp(pic), R)
        scs.Name = "Bild_" & m_ScsList.Count + 1
        m_ScsList.Add scs
        'Set m_bmp = MNew.BitmapSP(pic)
        
    'm_FocusRect.Delete
    'If scs Is Nothing Then
    '    MsgBox "screenshot is nothing"
    '    Exit Sub
    'End If
    'm_ScsList.Add scs
        LBPicList.AddItem scs.Name
        LBPicList.ListIndex = LBPicList.ListCount - 1
        
    Next
    'UpdateView
End Sub


Private Sub BtnSavePictures_Click()
    Dim i As Long
    For i = 0 To LBPicList.ListCount - 1
        LBPicList.ListIndex = i
        SavePicture PBScreenshot.image, FNm & "\Bild_" & CStr(i) & ".bmp"
    Next
End Sub

Private Sub LBPicList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = KeyCodeConstants.vbKeyDelete Then
        mnuListDeleteItem_Click
    End If
End Sub

Private Sub LBPicList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
    m_ScsList.MoveUp i
    LBPicList_MoveUp i
    LBPicList.ListIndex = i - 1
End Sub
Private Sub mnuListMoveDown_Click()
    Dim c As Long: c = LBPicList.ListCount
    If c = 1 Then Exit Sub
    Dim i As Long: i = LBPicList.ListIndex
    If i < 0 Or (c - 1) <= i Then Exit Sub
    m_ScsList.MoveDown i
    LBPicList_MoveDown i
    LBPicList.ListIndex = i + 1
End Sub
Private Sub mnuListDeleteItem_Click()
    Dim c As Long: c = LBPicList.ListCount
    Dim i As Long: i = LBPicList.ListIndex
    If i < 0 Or (c - 1) < i Then Exit Sub
    m_ScsList.Remove i
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
    Set m_Screen = MNew.SScreen(Me.PBScreenshot, GetWndRect) ' GetWinAPIRect)
End Sub

Private Sub BtnDragWndRect_Click()
    Dim FSB As New FScreenshotBorders
    'Dim R As WndRect: Set R = m_FocusRect.WndRect.Clone
    Dim R As WndRect: Set R = m_Screen.SrcRect.Clone
    If FSB.ShowDialog(Me, R) = vbCancel Then Exit Sub
    'm_FocusRect.WndRect.NewC R
    m_Screen.SrcRect.NewC R
    UpdateView
End Sub

Private Sub Timer1_Timer()
    Dim p As WinAPIPoint
    Dim hr As Long: hr = GetCursorPos(p)
    If hr = 0 Then Exit Sub
    'Dim R As WinAPIRect:  R = m_Screen.RectFromPoint(p)
    Dim R As WndRect:  Set R = MNew.WndRectFromMousePoint(p) ' m_Screen.RectFromPoint(p)
    TxtL.Text = R.Left
    TxtT.Text = R.Top
    TxtW.Text = R.Width  '.Right - R.Left
    TxtH.Text = R.Height '.Bottom - R.Top
    'm_FocusRect.WndRect.NewC R
    If Not OldRect Is Nothing Then
        If Not OldRect.Equals(R) Then m_FocusRect.Draw OldRect
    End If
    m_FocusRect.Draw R
    Set OldRect = R.Clone
End Sub

'Private Sub BtnSet_Click()
'    Set m_Screen = MNew.Screenshot(Me.PBScreenshot, GetWinAPIRect)
'End Sub

Private Sub BtnScreenshot_Click()
    m_FocusRect.Delete
    Dim scs As Screenshot: Set scs = m_Screen.Shot
    If scs Is Nothing Then
        MsgBox "screenshot is nothing"
        Exit Sub
    End If
    scs.Name = "Bild_" & LBPicList.ListCount + 1
    m_ScsList.Add scs
    LBPicList.AddItem scs.Name
    LBPicList.ListIndex = LBPicList.ListCount - 1
End Sub

Private Sub BtnClear_Click()
    'Set PicList = mNew Collection
    m_ScsList.Clear
    i = 0
    LBPicList.Clear
    'PBScreenshot.AutoRedraw = False
    Set PBScreenshot.Picture = Nothing
    PBScreenshot.Cls
    PBScreenshot.Move PBScreenshot.Left, PBScreenshot.Top, 593, 417
    'PBScreenshot.AutoRedraw = True
End Sub
'
'Private Function GetWinAPIRect() As WinAPIRect
'    'sehr suboptimal
'    If Not IsNumeric(TxtL.Text) Then Exit Function
'    If Not IsNumeric(TxtT.Text) Then Exit Function
'    If Not IsNumeric(TxtW.Text) Then Exit Function
'    If Not IsNumeric(TxtH.Text) Then Exit Function
'
'    Dim X As Long: X = CLng(TxtL.Text)
'    Dim Y As Long: Y = CLng(TxtT.Text)
'    Dim W As Long: W = CLng(TxtW.Text)
'    Dim H As Long: H = CLng(TxtH.Text)
'    GetWinAPIRect = MNew.WinAPIRect(X, Y, W, H)
'End Function

Private Function GetWndRect() As WndRect
    'sehr suboptimal
    If Not IsNumeric(TxtL.Text) Then Exit Function
    If Not IsNumeric(TxtT.Text) Then Exit Function
    If Not IsNumeric(TxtW.Text) Then Exit Function
    If Not IsNumeric(TxtH.Text) Then Exit Function
    
    Dim X As Long: X = CLng(TxtL.Text)
    Dim Y As Long: Y = CLng(TxtT.Text)
    Dim W As Long: W = CLng(TxtW.Text)
    Dim H As Long: H = CLng(TxtH.Text)
    Set GetWndRect = MNew.WndRect(MNew.WinAPIRect(X, Y, W, H))
End Function

Private Sub LBPicList_Click()
    'PBScreenshot.Cls
    'PBScreenshot.AutoRedraw = False
    Dim scs As Screenshot
    Dim i As Long: i = LBPicList.ListIndex
    Set scs = m_ScsList.Item(i)
    If scs Is Nothing Then
        MsgBox "pic is nothing"
    End If
    Set PBScreenshot.Picture = scs.Picture.StdPicture
    PBScreenshot.Refresh
    'PBScreenshot.AutoRedraw = True
End Sub
Private Sub LBPicList_DblClick()
    Dim scs As Screenshot
    Dim i As Long: i = LBPicList.ListIndex
    Set scs = m_ScsList.Item(i)
    If scs Is Nothing Then
        MsgBox "pic is nothing"
    End If
    Dim s As String: s = InputBox("Page name:", "Edit page name", scs.Name)
    If StrPtr(s) = 0 Then Exit Sub
    scs.Name = s
    LBPicList.List(i) = s
End Sub

Private Sub TxtL_Change(): TxtChange: End Sub
Private Sub TxtT_Change(): TxtChange: End Sub
Private Sub TxtW_Change(): TxtChange: End Sub
Private Sub TxtH_Change(): TxtChange: End Sub
Private Sub TxtChange()
    If Timer1.Enabled Or bInit Then Exit Sub
    'Timer1.Enabled = False
    Dim R As WndRect: Set R = GetWndRect
    If R Is Nothing Then Exit Sub
    'Set m_Screen = MNew.SScreen(Me.PBScreenshot, r)
    m_Screen.SrcRect.NewC R
    'm_FocusRect.Draw R
End Sub

Private Sub TxtL_KeyDown(KeyCode As Integer, Shift As Integer): TxtKeyDown 1, TxtL, KeyCode, Shift: End Sub
Private Sub TxtT_KeyDown(KeyCode As Integer, Shift As Integer): TxtKeyDown 2, TxtT, KeyCode, Shift: End Sub
Private Sub TxtW_KeyDown(KeyCode As Integer, Shift As Integer): TxtKeyDown 3, TxtW, KeyCode, Shift: End Sub
Private Sub TxtH_KeyDown(KeyCode As Integer, Shift As Integer): TxtKeyDown 4, TxtH, KeyCode, Shift: End Sub
Private Sub TxtKeyDown(ByVal prop As Long, tb As TextBox, KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    Case KeyCodeConstants.vbKeyUp, KeyCodeConstants.vbKeyDown, KeyCodeConstants.vbKeyLeft, KeyCodeConstants.vbKeyRight
        Dim d As Long: d = IIf(Shift, 5, 1) * IIf(KeyCode = KeyCodeConstants.vbKeyUp Or KeyCode = KeyCodeConstants.vbKeyLeft, -1, 1)
        'Select Case KeyCode
        'Case KeyCodeConstants.vbKeyUp, KeyCodeConstants.vbKeyLeft:    d = -d
        'Case KeyCodeConstants.vbKeyDown, KeyCodeConstants.vbKeyRight: v = v + d
        'Case KeyCodeConstants.vbKeyLeft:  v = v - 1
        'Case KeyCodeConstants.vbKeyRight: v = v + 1
        'End Select
    
        Dim R As WndRect: Set R = m_FocusRect.WndRect.Clone
        Dim v As Long
    
        Select Case prop
        Case 1: v = R.Left:   v = v + d: R.Left = v
        Case 2: v = R.Top:    v = v + d: R.Top = v
        'Case 3: v = R.Right:  v = v + d: R.Right = v
        Case 3: v = R.Width:  v = v + d: R.Width = v
        'Case 4: v = R.Bottom: v = v + d: R.Bottom = v
        Case 4: v = R.Height: v = v + d: R.Height = v
        End Select
        'Shift          = 1;
        'Strg           = 2;
        'Shift+Strg     = 1 + 2=3;
        'Alt            = 4;
        'Shift+Alt      = 1 + 4 = 5;
        'Shift+Strg+Alt = 1 + 2 + 4 = 7
        '
    
        'write the value in qestion
    '    Select Case prop
    '    Case 1: r.Left = v
    '    Case 2: r.Top = v
    '    Case 3: r.Right = v
    '    Case 4: r.Bottom = v
    '    End Select
    
        'bInit = True
        'tb.Text = CStr(v)
        'bInit = False
        m_Screen.SrcRect.NewC R
        
        UpdateView
    End Select
End Sub

Sub UpdateView()
    bInit = True
    TxtL.Text = m_Screen.SrcRect.Left
    TxtT.Text = m_Screen.SrcRect.Top
    TxtW.Text = m_Screen.SrcRect.Width
    TxtH.Text = m_Screen.SrcRect.Height
    m_FocusRect.Draw m_Screen.SrcRect
    bInit = False
End Sub
