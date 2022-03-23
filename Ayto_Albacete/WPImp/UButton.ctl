VERSION 5.00
Begin VB.UserControl UButton 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000016&
   ClientHeight    =   630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1170
   ClipControls    =   0   'False
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   42
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   78
   ToolboxBitmap   =   "UButton.ctx":0000
   Begin VB.Timer OverTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   600
      Top             =   120
   End
   Begin VB.Timer Rep 
      Enabled         =   0   'False
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "UButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const C1BackColor = 15002603
Private Const C1Light = 16777215
Private Const C1Shadow = 12238783
Private Const C1Dark = 8421504
Private Const C2BackColor = 15452586
Private Const C2Light = 15454137
Private Const C2Shadow = 14330501
Private Const C2Dark = 11569519
Private Const C3BackColor = 14268829
Private Const C3Light = 15922676
Private Const C3Shadow = 14476001
Private Const C3Dark = 12632256

Dim g_HasFocus As Boolean
Dim g_MouseDown As Boolean, g_MouseIn As Boolean
Dim g_Button As Integer, g_Shift As Integer, g_X As Single, g_Y As Single
Dim g_KeyPressed As Boolean

Dim m_BackColor As Long
Dim m_ForeColor As Long
Dim m_Text As String
Dim m_AlignPicture As EPosition
Dim m_AlignText As EPosition
Dim m_Margen As Integer
Dim m_Picture As StdPicture

Dim Pa As POINTAPI
Dim RPic As RECT

Dim m_Switch As Boolean
Dim Pulsado As Boolean

Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub UserControl_InitProperties()
  Set UserControl.Font = Ambient.Font
  m_Margen = 3
  m_AlignText = Center
  m_AlignPicture = CenterLeft
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_BackColor = PropBag.ReadProperty("BackColor", C1BackColor)
  m_ForeColor = PropBag.ReadProperty("ForeColor", vbBlack)
  m_Text = PropBag.ReadProperty("Text", "")
  m_AlignPicture = PropBag.ReadProperty("AlignPicture", 0)
  m_AlignText = PropBag.ReadProperty("AlignText", 0)
  Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
  Set m_Picture = PropBag.ReadProperty("Picture", Nothing)
  m_Margen = PropBag.ReadProperty("Margen", 3)
  UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
  m_Switch = PropBag.ReadProperty("Switch", False)
  Rep.Interval = PropBag.ReadProperty("Interval", 0)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("BackColor", m_BackColor, C1BackColor)
  Call PropBag.WriteProperty("ForeColor", m_ForeColor, vbBlack)
  Call PropBag.WriteProperty("Text", m_Text, "")
  Call PropBag.WriteProperty("AlignPicture", m_AlignPicture, 0)
  Call PropBag.WriteProperty("AlignText", m_AlignText, 0)
  Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
  Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
  Call PropBag.WriteProperty("Margen", m_Margen, 3)
  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
  Call PropBag.WriteProperty("Switch", m_Switch, False)
  Call PropBag.WriteProperty("Interval", Rep.Interval, 0)
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
  If Not Me.Enabled Then Exit Sub
  If KeyAscii = 13 Or KeyAscii = 27 Then RaiseEvent Click
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
  Refresh
End Sub

Private Sub UserControl_Resize()
  AlignPic
  Refresh
End Sub

Public Sub Refresh()
Attribute Refresh.VB_UserMemId = -550
  If m_BackColor <> C1BackColor Then
    UserControl.BackColor = m_BackColor
  Else
    If g_MouseDown Or Pulsado Then
      If g_MouseIn Then
        UserControl.BackColor = C3BackColor
      Else
        UserControl.BackColor = C1BackColor
      End If
    Else
      If g_MouseIn Then
        UserControl.BackColor = C2BackColor
      Else
        UserControl.BackColor = C1BackColor
      End If
    End If
  End If
  UserControl.Cls
  DrawPicture
  DrawCaption
  DrawBorder
End Sub

Private Sub UserControl_EnterFocus()
  g_HasFocus = True
  Refresh
End Sub

Private Sub UserControl_ExitFocus()
  g_HasFocus = False
  g_MouseDown = False
  Refresh
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  If Not g_HasFocus Then Exit Sub
  If Not g_KeyPressed Then
    Select Case KeyCode
      Case 40: KeyCode = 0: SendKey vbKeyTab
      Case 38: KeyCode = 0: SendKey vbKeyTab, True
    End Select
    If (KeyCode = vbKeyReturn) Or (KeyCode = vbKeySpace) Then
      g_MouseDown = True
      g_KeyPressed = True
      Refresh
    End If
  End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
  If Not g_KeyPressed Then Exit Sub
  g_KeyPressed = False
  g_MouseDown = False
  Refresh
  If (KeyCode = vbKeyReturn) Or (KeyCode = vbKeySpace) Then RaiseEvent Click
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  g_Button = Button: g_Shift = Shift: g_X = X: g_Y = Y
  If Button = vbLeftButton Then
    g_MouseDown = True
    If m_Switch Then Pulsado = Not Pulsado
    Refresh
    RaiseEvent MouseDown(Button, Shift, X, Y)
    Rep.Enabled = (Rep.Interval <> 0)
  End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If (X >= 0 And Y >= 0) And (X < ScaleWidth And Y < ScaleHeight) Then
    If g_MouseIn = False Then
      OverTimer.Enabled = True
      g_MouseIn = True
      Refresh
    End If
    RaiseEvent MouseMove(Button, Shift, X, Y)
  End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Not g_MouseDown Then Exit Sub
  Rep.Enabled = False
  g_MouseDown = False
  If Button = vbLeftButton Then
    Refresh
    If (X >= 0 And Y >= 0) And (X < ScaleWidth And Y < ScaleHeight) Then RaiseEvent Click
  End If
End Sub

Private Sub UserControl_DblClick()
  UserControl_MouseDown g_Button, g_Shift, g_X, g_Y
End Sub

Public Property Get Text() As String
Attribute Text.VB_UserMemId = -517
  Text = m_Text
End Property

Public Property Let Text(ByVal New_Text As String)
  m_Text = New_Text
  PropertyChanged "Text"
  Refresh
End Property

Public Property Get AlignPicture() As EPosition
  AlignPicture = m_AlignPicture
End Property

Public Property Let AlignPicture(ByVal New_AlignPicture As EPosition)
  m_AlignPicture = New_AlignPicture
  PropertyChanged "AlignPicture"
  AlignPic
  Refresh
End Property

Public Property Get AlignText() As EPosition
  AlignText = m_AlignText
End Property

Public Property Let AlignText(ByVal New_AlignText As EPosition)
  m_AlignText = New_AlignText
  PropertyChanged "AlignText"
  Refresh
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
Attribute Enabled.VB_UserMemId = -514
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  UserControl.Enabled() = New_Enabled
  PropertyChanged "Enabled"
  Refresh
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Devuelve un objeto Font."
Attribute Font.VB_ProcData.VB_Invoke_Property = "StandardFont;Apariencia"
Attribute Font.VB_UserMemId = -512
  Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
  Set UserControl.Font = New_Font
  Refresh
  PropertyChanged "Font"
End Property

Public Property Get Hwnd() As Long
Attribute Hwnd.VB_UserMemId = -515
  Hwnd = UserControl.Hwnd
End Property

Public Property Get Picture() As Picture
Attribute Picture.VB_ProcData.VB_Invoke_Property = "StandardPicture;Apariencia"
  Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
  On Error Resume Next
  If Not IsZero(New_Picture) Then
    If New_Picture.Type <> vbPicTypeIcon Then
      MsgBox "La imagen debe ser un icono", vbExclamation
      Exit Property
    End If
  End If
  Set m_Picture = New_Picture
  PropertyChanged "Picture"
  AlignPic
  Refresh
  err.Clear
End Property

Public Property Get Margen() As Integer
  Margen = m_Margen
End Property

Public Property Let Margen(ByVal New_Margen As Integer)
  m_Margen = New_Margen
  AlignPic
  PropertyChanged "Margen"
  Refresh
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_ProcData.VB_Invoke_Property = "StandardColor;Apariencia"
Attribute BackColor.VB_UserMemId = -501
  BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  m_BackColor = New_BackColor
  Refresh
  PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = "StandardColor;Apariencia"
Attribute ForeColor.VB_UserMemId = -513
  ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
  m_ForeColor = New_ForeColor
  Refresh
  PropertyChanged "ForeColor"
End Property

Public Property Get Switch() As Boolean
  Switch = m_Switch
End Property

Public Property Let Switch(ByVal New_Switch As Boolean)
  m_Switch = New_Switch
  PropertyChanged "Switch"
End Property

Public Property Get Interval() As Long
  Interval = Rep.Interval
End Property

Public Property Let Interval(ByVal New_Interval As Long)
  Rep.Interval = New_Interval
  PropertyChanged "Interval"
End Property

Public Property Get Value() As Boolean
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "200"
  Value = Pulsado
End Property

Public Property Let Value(ByVal New_Value As Boolean)
  Pulsado = New_Value
  Refresh
End Property

Private Sub DrawPicture()
  Dim X As Integer, Y As Integer
  Dim W As Integer, H As Integer
        
  If m_Picture Is Nothing Then Exit Sub
  If m_Picture = 0 Then Exit Sub
  
  X = RPic.Left: Y = RPic.Top
  W = RPic.Right - RPic.Left
  H = RPic.Bottom - RPic.Top
   
  If g_MouseDown Then
    X = X + 1: Y = Y + 1
  End If
    
  DrawIconEx hdc, X, Y, m_Picture.handle, W, H, 0, 0, 3
   
  If UserControl.Enabled Then Exit Sub
 
  Dim P As Long
  Dim V() As Boolean
  ReDim V(0 To W, 0 To H)
  For Y = 0 To H
    For X = 0 To W
      P = GetPixel(hdc, RPic.Left + X, RPic.Top + Y)
      If P <> UserControl.BackColor Then
        V(X, Y) = (((P And 255) + (P \ 256 And 255) + (P \ 65536)) < 500)
      End If
    Next X
  Next Y
  UserControl.Cls
  For Y = 0 To H
    For X = 0 To W
      If V(X, Y) Then
        SetPixel hdc, RPic.Left + X, RPic.Top + Y, C1Dark
        SetPixel hdc, RPic.Left + X + 1, RPic.Top + Y + 1, C1Light
      End If
    Next X
  Next Y

End Sub

Private Sub DrawCaption()
  If m_Text = "" Then Exit Sub
  Dim H As Integer
  Dim R As RECT
  R.Left = m_Margen: R.Top = m_Margen
  R.Right = ScaleWidth - m_Margen: R.Bottom = ScaleHeight - m_Margen
  
  If InStr(m_Text, vbCrLf) Then H = 0 Else H = &H20
  
  If Me.Enabled Then
    UserControl.ForeColor = m_ForeColor
    If g_MouseDown Or Pulsado Then
      R.Left = R.Left + 1: R.Top = R.Top + 1
      R.Bottom = R.Bottom + 1: R.Right = R.Right + 1
      DrawText hdc, m_Text, -1, R, m_AlignText + H
    Else
      DrawText hdc, m_Text, -1, R, m_AlignText + H
    End If
  Else
    R.Left = R.Left + 1: R.Top = R.Top + 1
    UserControl.ForeColor = C1Light: DrawText hdc, m_Text, -1, R, m_AlignText + H
    R.Left = R.Left - 1: R.Top = R.Top - 1
    UserControl.ForeColor = C1Dark: DrawText hdc, m_Text, -1, R, m_AlignText + H
  End If

End Sub

Private Sub DrawBorder()
  Dim C1 As Long
  Dim C2 As Long
  Dim C3 As Long
       
  If g_MouseDown Or Pulsado Then
    If g_MouseIn Then
      C1 = C2Dark: C2 = C2Light: C3 = C2Shadow
    Else
      C1 = C1Dark: C2 = C1Light: C3 = C1Shadow
    End If
  Else
    If g_HasFocus Then
      C1 = C2Dark: C2 = C2BackColor: C3 = -1
    Else
      If g_MouseIn Then
        C1 = C2Dark: C2 = C2Shadow: C3 = C2Light
      Else
        If UserControl.Enabled Then
          C1 = C1Dark: C2 = C1Shadow: C3 = C1Light
        Else
          C1 = C3Dark: C2 = C3Shadow: C3 = C3Light
        End If
      End If
    End If
  End If
  
  UserControl.ForeColor = C1: MoveToEx hdc, 0, 0, Pa: LineTo hdc, ScaleWidth - 1, 0
  LineTo hdc, ScaleWidth - 1, ScaleHeight - 1: LineTo hdc, 0, ScaleHeight - 1: LineTo hdc, 0, 0
  UserControl.ForeColor = C2
  
  If C3 = -1 Then
    MoveToEx hdc, 1, 1, Pa
    LineTo hdc, ScaleWidth - 2, 1: LineTo hdc, ScaleWidth - 2, ScaleHeight - 2
    LineTo hdc, 1, ScaleHeight - 2: LineTo hdc, 1, 2
    LineTo hdc, ScaleWidth - 3, 2: LineTo hdc, ScaleWidth - 3, ScaleHeight - 3
    LineTo hdc, 2, ScaleHeight - 3: LineTo hdc, 2, 2
  Else
    MoveToEx hdc, ScaleWidth - 2, 2, Pa
    LineTo hdc, ScaleWidth - 2, ScaleHeight - 2: LineTo hdc, 1, ScaleHeight - 2
    UserControl.ForeColor = C3: LineTo hdc, 1, 1: LineTo hdc, ScaleWidth - 2, 1
  End If
  
  SetPixel hdc, 0, 0, C1Shadow
  SetPixel hdc, 0, ScaleHeight - 1, C1Shadow
  SetPixel hdc, ScaleWidth - 1, 0, C1Shadow
  SetPixel hdc, ScaleWidth - 1, ScaleHeight - 1, C1Shadow
  
End Sub

Private Sub AlignPic()
  On Error Resume Next
  Dim UW As Integer
  Dim UH As Integer
  Dim PW As Integer
  Dim PH As Integer
  
  If m_Picture Is Nothing Then Exit Sub
  UW = ScaleWidth
  UH = ScaleHeight
  
  PW = Int(m_Picture.Width / 26.455) + 1
  PH = Int(m_Picture.Height / 26.455) + 1
  
  Select Case m_AlignPicture
    Case 0: RPic.Left = m_Margen: RPic.Top = m_Margen
    Case 1: RPic.Left = (UW - PW) \ 2: RPic.Top = m_Margen
    Case 2: RPic.Left = UW - PW - m_Margen: RPic.Top = m_Margen
    Case 4: RPic.Left = m_Margen: RPic.Top = (UH - PH) \ 2
    Case 5: RPic.Left = (UW - PW) \ 2: RPic.Top = (UH - PH) \ 2
    Case 6: RPic.Left = UW - PW - m_Margen: RPic.Top = (UH - PH) \ 2
    Case 8: RPic.Left = 3: RPic.Top = UH - PH - m_Margen
    Case 9: RPic.Left = (UW - PW) \ 2: RPic.Top = UH - PH - m_Margen
    Case 10: RPic.Left = UW - PW - m_Margen: RPic.Top = UH - PH - m_Margen
  End Select
  RPic.Right = RPic.Left + PW
  RPic.Bottom = RPic.Top + PH
  err.Clear
End Sub

Private Sub OverTimer_Timer()
  GetCursorPos Pa
  If Hwnd <> WindowFromPoint(Pa.X, Pa.Y) Then
    OverTimer.Enabled = False
    g_MouseIn = False
    Refresh
    If g_MouseDown = True Then
      g_MouseDown = False
      Refresh
      g_MouseDown = True
    End If
  End If
End Sub

Private Sub Rep_Timer()
  If g_MouseDown Then
    RaiseEvent MouseDown(g_Button, g_Shift, g_X, g_Y)
    DoEvents
  Else
    Rep.Enabled = False
    Refresh
  End If
End Sub















