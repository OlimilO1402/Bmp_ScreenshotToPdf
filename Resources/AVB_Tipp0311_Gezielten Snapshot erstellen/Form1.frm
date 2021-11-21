VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "www.activevb.de"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   ScaleHeight     =   1965
   ScaleWidth      =   3930
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   495
      Left            =   3120
      ScaleHeight     =   435
      ScaleWidth      =   675
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Snapshot erstellen"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":0000
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dieser Source stammt von http://www.activevb.de
'und kann frei verwendet werden. Für eventuelle Schäden
'wird nicht gehaftet.

'Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
'Ansonsten viel Spaß und Erfolg mit diesem Source !

'Autor: Florian Rittmeier
'E-Mail: Florian@ActiveVB.de
'Nach einer Idee von Lothar Kriegerow

Option Explicit

Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long

Private Declare Function GetDC Lib "user32.dll" ( _
     ByVal hwnd As Long) As Long

Private Declare Function ReleaseDC Lib "user32.dll" ( _
     ByVal hwnd As Long, _
     ByVal hdc As Long) As Long

Private Declare Function GetWindowRect Lib "user32.dll" ( _
     ByVal hwnd As Long, _
     lpRect As RECT) As Long

Private Declare Function BitBlt Lib "gdi32" ( _
    ByVal hDstDC As Long, ByVal xDst As Long, ByVal yDst As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const SRCCOPY As Long = &HCC0020

Private Sub Command2_Click()
    Dim deskWnd As Long, deskDC As Long, retval As Long
    Dim windowpos As RECT

    ' Device Context des Desktops(Bildschirms) ermitteln
    deskWnd = GetDesktopWindow
    deskDC = GetDC(deskWnd)

    ' Abmessungen des Formulars bestimmen
    retval = GetWindowRect(Me.hwnd, windowpos)
    If retval = 0 Then
        Call MsgBox("Die Abmessungen des Formulars konnten nicht bestimmt werden.", _
                    vbExclamation + vbOKOnly, App.Title)
        Exit Sub
    End If

    ' Größe der Picturebox
    Picture1.Width = Me.ScaleX(windowpos.Right - windowpos.Left, vbPixels, Me.ScaleMode)
    Picture1.Height = Me.ScaleY(windowpos.Bottom - windowpos.Top, vbPixels, Me.ScaleMode)

    Me.Visible = False ' Fenster unsichtbar machen
    DoEvents ' Dem Fenster Zeit geben, dass es verschwindet.

    ' Snapshot des entsprechenden Bereiches machen
    retval = BitBlt(Picture1.hdc, 0, 0, Picture1.Width, Picture1.Height, deskDC, windowpos.Left, windowpos.Top, SRCCOPY)

    ' Fenster wieder sichtbar machen
    Me.Visible = True

    ' Handle wieder freigeben
    Call ReleaseDC(deskWnd, deskDC)

    ' An dieser Stelle nun überprüfen,
    ' ob wir überhaupt einen erfolgreichen Screenshot gemacht haben
    If retval = 0 Then
        Call MsgBox("Das Snapshot konnte nicht erstellt werden.", vbExclamation + vbOKOnly, App.Title)
        Exit Sub
    End If

    Picture1.Refresh

    Clipboard.SetData Picture1.Image, vbCFBitmap
    Clipboard.SetData Picture1.Image, vbCFDIB
End Sub
