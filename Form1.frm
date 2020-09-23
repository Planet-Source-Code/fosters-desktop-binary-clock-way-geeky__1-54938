VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1665
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":08CA
   ScaleHeight     =   61
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   111
   ShowInTaskbar   =   0   'False
   Begin VB.Timer timTime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3420
      Top             =   2400
   End
   Begin VB.PictureBox picS 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   5
      Left            =   240
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   19
      Top             =   540
      Width           =   135
   End
   Begin VB.PictureBox picM 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   5
      Left            =   240
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   18
      Top             =   360
      Width           =   135
   End
   Begin VB.PictureBox picH 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   5
      Left            =   240
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   17
      Top             =   180
      Width           =   135
   End
   Begin VB.PictureBox picS 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   4
      Left            =   420
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   16
      Top             =   540
      Width           =   135
   End
   Begin VB.PictureBox picM 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   4
      Left            =   420
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   15
      Top             =   360
      Width           =   135
   End
   Begin VB.PictureBox picH 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   4
      Left            =   420
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   14
      Top             =   180
      Width           =   135
   End
   Begin VB.PictureBox picS 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   3
      Left            =   600
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   13
      Top             =   540
      Width           =   135
   End
   Begin VB.PictureBox picM 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   3
      Left            =   600
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   12
      Top             =   360
      Width           =   135
   End
   Begin VB.PictureBox picH 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   3
      Left            =   600
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   11
      Top             =   180
      Width           =   135
   End
   Begin VB.PictureBox picS 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   2
      Left            =   780
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   10
      Top             =   540
      Width           =   135
   End
   Begin VB.PictureBox picM 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   2
      Left            =   780
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   9
      Top             =   360
      Width           =   135
   End
   Begin VB.PictureBox picH 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   2
      Left            =   780
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   8
      Top             =   180
      Width           =   135
   End
   Begin VB.PictureBox picS 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   1
      Left            =   960
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   7
      Top             =   540
      Width           =   135
   End
   Begin VB.PictureBox picM 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   1
      Left            =   960
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   6
      Top             =   360
      Width           =   135
   End
   Begin VB.PictureBox picH 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   1
      Left            =   960
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   5
      Top             =   180
      Width           =   135
   End
   Begin VB.PictureBox picS 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   0
      Left            =   1140
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   4
      Top             =   540
      Width           =   135
   End
   Begin VB.PictureBox picM 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   0
      Left            =   1140
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   3
      Top             =   360
      Width           =   135
   End
   Begin VB.PictureBox picH 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   0
      Left            =   1140
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   2
      Top             =   180
      Width           =   135
   End
   Begin VB.PictureBox picBase 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   1
      Left            =   3720
      Picture         =   "Form1.frx":591C
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   1
      Top             =   540
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox picBase 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   0
      Left            =   3720
      Picture         =   "Form1.frx":5A5A
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   0
      Top             =   300
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image Image9 
      Height          =   90
      Left            =   1320
      Picture         =   "Form1.frx":5B98
      Top             =   540
      Width           =   75
   End
   Begin VB.Image Image8 
      Height          =   105
      Left            =   1320
      Picture         =   "Form1.frx":5C3A
      Top             =   360
      Width           =   75
   End
   Begin VB.Image Image7 
      Height          =   105
      Left            =   1320
      Picture         =   "Form1.frx":5CEC
      Top             =   180
      Width           =   60
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const SWP_NOMOVE = &H2, SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1, HWND_NOTOPMOST = -2

Sub SetTopmostWindow(ByVal hwnd As Long, Optional topmost As Boolean = True)
    Const HWND_NOTOPMOST = -2
    Const HWND_TOPMOST = -1
    Const SWP_NOMOVE = &H2
    Const SWP_NOSIZE = &H1
    SetWindowPos hwnd, IIf(topmost, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, _
        SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub Form_DblClick()
    Unload Me
    End
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, _
    Y As Single)
    Const WM_NCLBUTTONDOWN = &HA1
    Const HTCAPTION = 2
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End Sub
Sub ResetLEDs(pic As Variant)
Dim I As Integer
    For I = 0 To pic.Count - 1
        BitBlt pic(I).hdc, 0, 0, picBase(0).Width, picBase(0).Height, picBase(0).hdc, 0, 0, vbSrcCopy
    Next
End Sub
Sub SetLED(pic As Variant, I As Integer)
    BitBlt pic(I).hdc, 0, 0, picBase(1).Width, picBase(1).Height, picBase(1).hdc, 0, 0, vbSrcCopy
End Sub
Sub SetNumber(pic As Variant, iNum As Integer)
Dim X As Integer
Dim c As Integer
    c = 1
    X = 1
    Do
        If iNum And X Then
            SetLED pic, c - 1
        End If
        X = X * 2
        c = c + 1
    Loop Until c > pic.Count
End Sub

Private Sub Form_Load()
Dim xPos As Long, yPos As Long
    
    timTime.Enabled = True

    SetColorTransparent Form1, RGB(255, 0, 255)
    
    SetTopmostWindow Me.hwnd
   ' Me.Show:   DoEvents
    If Len(GetSetting("BinClock", "Settings", "XPos")) > 0 Then
        xPos = CLng(GetSetting("BinClock", "Settings", "XPos"))
    Else
        xPos = Screen.Width - (Me.Width * 2.5)
    End If
    
    If Len(GetSetting("BinClock", "Settings", "YPos")) > 0 Then
        yPos = CLng(GetSetting("BinClock", "Settings", "YPos"))
    Else
        yPos = 0
    End If

    If (xPos + Me.Width) > Screen.Width Or xPos < 0 Then xPos = Screen.Width - (Me.Width * 2.5)
    If (yPos + Me.Height) > Screen.Height Or yPos < 0 Then yPos = 0
    
    Me.Left = xPos: Me.Top = yPos

    ResetLEDs picH
    ResetLEDs picM
    ResetLEDs picS
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "BinClock", "Settings", "XPos", Me.Left
    SaveSetting "BinClock", "Settings", "YPos", Me.Top
End Sub

Private Sub timTime_Timer()
Dim dTime As Date
    dTime = Now
    ResetLEDs picH
    ResetLEDs picM
    ResetLEDs picS
    
    SetNumber picH, CInt(Format(dTime, "hh"))
    SetNumber picM, CInt(Format(dTime, "nn"))
    SetNumber picS, CInt(Format(dTime, "ss"))
End Sub
