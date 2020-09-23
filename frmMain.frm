VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'Kein
   Caption         =   "SineScreenSaver by www.melaxis.com"
   ClientHeight    =   5370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8205
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00404040&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmMain.frx":044A
   ScaleHeight     =   358
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   547
   StartUpPosition =   2  'Bildschirmmitte
   WindowState     =   2  'Maximiert
   Begin VB.Timer tmrPreview 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6240
      Top             =   3360
   End
   Begin VB.Timer tmrModifier 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6240
      Top             =   2880
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST As Long = -1
Private Const SWP_SHOWWINDOW As Long = &H40
Private Const SWP_NOACTIVATE As Long = &H10
Private Const GWL_STYLE As Long = (-16)
Private Const WS_CHILD As Long = &H40000000
Private Const GWL_HWNDPARENT As Long = (-8)
Private Const HWND_TOP As Long = 0

  Private Declare Function GetClientRect Lib "user32" _
    (ByVal hwnd As Long, lpRect As RECT) As Long

  Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
  End Type
  Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
  Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long) As Long
  Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, _
    ByVal hWndNewParent As Long) As Long

Dim myW As Long
Dim myH As Long
Const pi = 3.1415
Const DIF = 0.1
Dim SinPos As Long
Dim BallPos As Long
Dim BallDir As Boolean
Dim federa As Double
Dim smaller As Boolean

Dim QuitOnMouseMove As Boolean
Dim IsMouseMove As Boolean
Dim LastX As Long
Dim LastY As Long

Dim PreviewHwnd As Long
Dim PreviewRect As RECT


'Shell "RunDll32.exe desk.cpl,InstallScreenSaver C:\Windows\MeinScreenSaver.scr"

Private Sub Form_Click()

    End

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    End

End Sub

Private Sub Form_Load()

    federa = 1
    
    If LCase(Command) = "/mel" Then
        ' Starten :-)
        If App.PrevInstance Then
            End
            Exit Sub
        End If
        Me.Visible = True
        tmrModifier.Enabled = True
        'tmrPreview.Enabled = True
    ElseIf Mid(LCase(Command), 1, 2) = "/s" Then
        ' Normaler Modus
        QuitOnMouseMove = True
        If App.PrevInstance Then
            End
            Exit Sub
        End If
        Me.Visible = True
        SetWindowPos Me.hwnd, HWND_TOPMOST, _
            0, 0, Screen.Width \ Screen.TwipsPerPixelX, _
            Screen.Height \ Screen.TwipsPerPixelY, _
            SWP_SHOWWINDOW
        tmrModifier.Enabled = True
        Me.MousePointer = 99
    ElseIf Mid(LCase(Command), 1, 2) = "/p" Then
        ' Preview
        PreviewHwnd = Val(Mid(Command, 4))
        GetClientRect PreviewHwnd, PreviewRect
        Dim Style As Long
        Style = GetWindowLong(Me.hwnd, GWL_STYLE)
        Style = Style Or WS_CHILD
        SetWindowLong Me.hwnd, GWL_STYLE, Style
        SetParent Me.hwnd, PreviewHwnd
        SetWindowLong Me.hwnd, GWL_HWNDPARENT, PreviewHwnd
        SetWindowPos Me.hwnd, HWND_TOP, _
            0&, 0&, PreviewRect.Right, PreviewRect.Bottom, _
            SWP_NOACTIVATE Or SWP_SHOWWINDOW
        Me.Visible = True
        tmrPreview.Enabled = True
    ElseIf Mid(LCase(Command), 1, 2) = "/c" Then
        ' Config
        frmConfig.Show vbModal
        End
        Exit Sub
    Else
        End
        Exit Sub
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Not IsMouseMove Then
        IsMouseMove = True
        LastX = X
        LastY = Y
        Exit Sub
    End If

    If QuitOnMouseMove Then
        'End
        Dim DiffX As Long
        Dim DiffY As Long
        
        DiffX = Abs(LastX - X)
        DiffY = Abs(LastY - Y)
        
        If (DiffX > 3) Or (DiffY > 3) Then
            End
        End If
        
        LastX = X
        LastY = Y
    End If

End Sub

Private Sub Form_Resize()

    myW = Me.ScaleWidth
    myH = Me.ScaleHeight

End Sub

Private Sub tmrModifier_Timer()

    Dim cx As Long, cy As Long
    Dim h As Integer, m As Integer, s As Integer
    Dim xs As Long, ys As Long, xm As Long, ym As Long, xh As Long, yh As Long
    Dim a As Integer, x1 As Long, y1 As Long, x2 As Long, y2 As Long
    Dim i As Long
    Dim v As Double
    Dim f As Double
    Dim c As Long
    Dim lx As Long
    Dim ly As Long
    
    Me.Cls
    
    ' Bewegende Sinuskurve
    For i = 0 To Me.ScaleWidth
        v = Sin((i + SinPos) * (pi / 180))
        If v >= 0 Then
            f = v
        Else
            f = v * (-1)
        End If
        c = RGB(f * 255, f * 128, f * 64)
        'Me.PSet (i, 80 - (v * (180 / pi))), c
        SetPixel Me.hdc, i, 80 - (v * (180 / pi)), c
    Next
    SinPos = SinPos + 1
    
    ' Ball
    v = Sin((SinPos + BallPos) * (pi / 180))
    Me.Circle (BallPos, 80 - (v * (180 / pi))), 10, vbRed
    If BallDir Then
        BallPos = BallPos + 3
        If BallPos >= (Me.ScaleWidth - 10) Then
            BallDir = Not BallDir
        End If
    Else
        BallPos = BallPos - 4
        If BallPos <= 10 Then
            BallDir = Not BallDir
        End If
    End If
    
    ' Feder
    lx = 0
    ly = 175
    For i = 0 To Me.ScaleWidth
        v = Sin((i * federa) * (pi / 180))
        If v >= 0 Then
            f = v
        Else
            f = v * (-1)
        End If
        'c = RGB(f * 255, f * 128, f * 64)
        c = RGB(f * 64, f * 128, f * 255)
        Me.Line (lx, ly)-(i, 350 - (v * (180 / pi))), c
        lx = i
        ly = 350 - (v * (180 / pi))
    Next
    If smaller Then
        federa = federa + DIF
        If federa > 10 Then
            smaller = False
        End If
    Else
        federa = federa - DIF
        If federa < -10 Then
            smaller = True
        End If
    End If
    
    ' Uhr
    h = Hour(Now)
    m = Minute(Now)
    s = Second(Now)
    cx = myW - 70 - 20
    cy = myH - 70 - 20
    Me.Circle (cx, cy), 70, vbRed
    xs = Cos(s * pi / 30 - pi / 2) * 60 + cx
    ys = Sin(s * pi / 30 - pi / 2) * 60 + cy
    xm = Cos(m * pi / 30 - pi / 2) * 50 + cx
    ym = Sin(m * pi / 30 - pi / 2) * 50 + cy
    xh = Cos((h * 30 + m / 2) * (pi / 180) - pi / 2) * 35 + cx
    yh = Sin((h * 30 + m / 2) * (pi / 180) - pi / 2) * 35 + cy
    Me.Line (cx, cy)-(xh, yh), vbWhite
    Me.Line (cx, cy)-(xm, ym), vbBlue
    Me.Line (cx, cy)-(xs, ys), vbGreen
    Me.CurrentY = myH - 20
    Me.CurrentX = myW - 115
    Me.Print Time
    
    ' Text
    Me.CurrentX = 5
    Me.CurrentY = myH - 20
    Me.Print "Copyright 2002 Pablo Hoch, www.melaxis.com"
    
    Me.Refresh

End Sub

Private Sub tmrPreview_Timer()

    Dim cx As Long, cy As Long
    Dim h As Integer, m As Integer, s As Integer
    Dim xs As Long, ys As Long, xm As Long, ym As Long, xh As Long, yh As Long
    Dim a As Integer, x1 As Long, y1 As Long, x2 As Long, y2 As Long
    Dim i As Long
    Dim v As Double
    Dim f As Double
    Dim c As Long
    Dim lx As Long
    Dim ly As Long
    
    Me.Cls
    
    ' Bewegende Sinuskurve
    For i = 0 To Me.ScaleWidth
        v = Sin((i + SinPos) * (pi / 180))
        If v >= 0 Then
            f = v
        Else
            f = v * (-1)
        End If
        c = RGB(f * 255, f * 128, f * 64)
        'Me.PSet (i, 80 - (v * (180 / pi))), c
        SetPixel Me.hdc, i, 40 - (v * (180 / pi) * 0.5), c
    Next
    SinPos = SinPos + 1

End Sub
