Attribute VB_Name = "Draw"
Option Explicit

Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal Hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal DwRop As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function InvertRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetSysColorA Lib "user32" Alias "GetSysColor" (ByVal nIndex As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyHeight As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Const SRCCOPY = &HCC0020

Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Type POINTAPI
  X As Long
  Y As Long
End Type

Public Enum EColores
  vbLightRed = &HC0C0FF
  vbdarkred = 128
  vblightYellow = &HDFFFFF
  vbdarkGreen = 32768
  vbDarkBlue = &H800000
  vblightgrey = &HE0E0E0
  vbgrey = &HC0C0C0
  vbdarkgrey = &H808080
  vbLightcyan = &HFFFFC0
End Enum

Public Enum EPosition
  UpLeft = 0
  UpCenter = 1
  UpRight = 2
  CenterLeft = 4
  Center = 5
  CenterRight = 6
  DownLeft = 8
  DownCenter = 9
  DownRigth = 10
End Enum

Public Enum EAlignment
  Izquierda = 0
  Derecha = 1
  Centro = 2
End Enum

Public Enum EBorderStyle
  BsNone = 0
  BsRaisedOuter = 1
  BsSunkenOuter = 2
  BsFlat = 3
  BsRaisedInner = 4
  BsRaised = 5
  BsEtched = 6
  BsSunkenInner = 8
  BsBump = 9
  BsSunken = 10
End Enum

Public Sub DrawBox(hdc As Long, Left As Integer, Top As Integer, Width As Integer, Height As Integer, Optional BackColor As Long = -1, Optional BorderColor As Long = -1, Optional Borderstyle As EBorderStyle = BsNone, Optional Texto As String, Optional Position As EPosition = UpLeft, Optional X1 As Integer, Optional Y1 As Integer, Optional X2 As Integer, Optional Y2 As Integer)
  Dim R As RECT
  Dim H As Long
  Dim T As Long
  R.Left = Left: R.Top = Top
  R.Right = Left + Width - 1: R.Bottom = Top + Height - 1
  If BackColor <> -1 Then
    H = CreateSolidBrush(GetSysColor(BackColor))
    FillRect hdc, R, H
    DeleteObject H
  End If
  If Borderstyle = BsFlat Then
    If BorderColor <> -1 Then
      H = CreateSolidBrush(GetSysColor(BorderColor))
      FrameRect hdc, R, H
      DeleteObject H
    Else
      DrawEdge hdc, R, Borderstyle, 16399
    End If
  Else
    DrawEdge hdc, R, Borderstyle, 15
  End If
  If Texto <> "" Then
    R.Left = R.Left + X1: R.Top = R.Top + Y1
    R.Right = R.Right + X2: R.Bottom = R.Bottom + Y2
    If InStr(Texto, vbCrLf) Then H = 0 Else H = &H20
    DrawText hdc, Texto, Len(Texto), R, Position + H
  End If
End Sub

Public Sub DrawBoxInvert(hdc As Long, Left As Integer, Top As Integer, Width As Integer, Height As Integer)
  Dim R As RECT
  R.Left = Left: R.Top = Top
  R.Right = Left + Width - 1: R.Bottom = Top + Height - 1
  InvertRect hdc, R
End Sub

Public Sub DrawControl(hdc As Long, Left As Integer, Top As Integer, Width As Integer, Height As Integer, Tipo As Long, Estado As Long)
  Dim R As RECT
  R.Left = Left: R.Top = Top
  R.Right = Left + Width - 1: R.Bottom = Top + Height - 1
  DrawFrameControl hdc, R, Tipo, Estado
End Sub

Public Sub DrawLinea(hdc As Long, X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer)
  Dim P As POINTAPI
  MoveToEx hdc, X1, Y1, P
  LineTo hdc, X2, Y2
End Sub

Public Function GetSysColor(Color As Long) As Long
  If Color <= -2147483624# Then
    GetSysColor = GetSysColorA(Color + 2147483648#)
  Else
    GetSysColor = Color
  End If
End Function








