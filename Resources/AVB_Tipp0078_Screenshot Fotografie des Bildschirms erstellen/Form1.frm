VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "www.ActiveVB.de"
   ClientHeight    =   4170
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   4170
   ScaleWidth      =   4830
   Begin VB.CommandButton Command3 
      Caption         =   "Drucken"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Speichern"
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   3720
      Width           =   1215
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Width           =   4335
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3255
      Left            =   4440
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'Kein
      Height          =   3255
      Left            =   120
      ScaleHeight     =   3255
      ScaleWidth      =   4335
      TabIndex        =   1
      Top             =   120
      Width           =   4335
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'Kein
         Height          =   4935
         Left            =   120
         ScaleHeight     =   4935
         ScaleWidth      =   6015
         TabIndex        =   2
         Top             =   360
         Width           =   6015
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Screen Shot"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   1335
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

Option Explicit

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
        
Private Declare Function StretchBlt Lib "gdi32" ( _
    ByVal DstHDC As Long, ByVal xDst As Long, ByVal yDst As Long, ByVal DstW As Long, ByVal DstH As Long, _
    ByVal SrcHDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal SrcW As Long, ByVal SrcH As Long, ByVal dwRop As Long) As Long
        
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type

Const SRCCOPY = &HCC0020

Private Sub Form_Load()
    Picture2.Top = 0
    Picture2.Left = 0
    VScroll1.LargeChange = Picture1.Height / 4
    
    VScroll1.SmallChange = 120
    HScroll1.LargeChange = Picture1.Width / 4
    HScroll1.SmallChange = 120
End Sub

Private Sub Command1_Click()
    ScreenShot
End Sub

Private Sub Command2_Click()
    SavePicture Picture2.Image, App.Path & "\Test.bmp"
End Sub

Private Sub Command3_Click()
    Printer.Print
    Printer.PaintPicture Picture2.Image, 0, 0, _
                         Picture2.Width, Picture2.Height, _
                         0, 0, Picture2.Width * 2, _
                         Picture2.Height * 2
    Printer.EndDoc
End Sub

Private Sub HScroll1_Change()
    Picture2.Left = -HScroll1.Value
End Sub

Private Sub VScroll1_Change()
    Picture2.Top = -VScroll1.Value
End Sub


Private Sub ScreenShot()
    Dim hr As Long
    Dim SrcHWnd As Long
    Dim SrcHDC As Long
    Dim SrcRect As RECT
    
    Picture2.AutoRedraw = True
    
    '### Desktopgröße in Pixeln ermitteln
    SrcHWnd = GetDesktopWindow()
    SrcHDC = GetDC(SrcHWnd)
    hr = GetWindowRect(SrcHWnd, SrcRect)
    
    '### Zielbild und Scrollbalken der Desktopgröße anpassen
    Picture2.Width = SrcRect.Width * 15
    Picture2.Height = SrcRect.Height * 15
    VScroll1.Max = Picture2.Height - Picture1.Height + 15
    HScroll1.Max = Picture2.Width - Picture1.Width + 15
    
    '### Der eigentliche Screenshot
    hr = StretchBlt(Picture2.hdc, SrcRect.Left, SrcRect.Top, SrcRect.Width, SrcRect.Height, _
                    SrcHDC, 0, 0, SrcRect.Width, SrcRect.Height, SRCCOPY)
    
    
    '### Gerätekontext löschen
    hr = ReleaseDC(SrcHWnd, SrcHDC)
     
    Picture2.Refresh
    Picture2.AutoRedraw = False
End Sub
