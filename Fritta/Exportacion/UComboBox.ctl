VERSION 5.00
Begin VB.UserControl UComboBox 
   BackColor       =   &H0000FF00&
   ClientHeight    =   1725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3420
   DataBindingBehavior=   1  'vbSimpleBound
   FillStyle       =   0  'Solid
   ScaleHeight     =   115
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   228
   ToolboxBitmap   =   "UComboBox.ctx":0000
   Begin VB.TextBox Texto 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   135
      Width           =   975
   End
   Begin VB.Timer Tim 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2760
      Top             =   120
   End
   Begin VB.PictureBox Lis 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   120
      ScaleHeight     =   67
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   207
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   3135
      Begin VB.VScrollBar VScroll 
         Height          =   915
         Left            =   2835
         Max             =   0
         TabIndex        =   1
         Top             =   0
         Width           =   285
      End
   End
End
Attribute VB_Name = "UComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private bDataType As EDataType
Private bBorderStyle As EBorderStyle
Private bColor As OLE_COLOR
Private bColorFocus As OLE_COLOR
Private bColorDisabled As OLE_COLOR
Private bOldValue As String
Attribute bOldValue.VB_VarMemberFlags = "400"

Private m_DataChanged As Boolean

Private IsKeyPress As Boolean
Private IsMousedown As Boolean

Private Buscar As Boolean

Private m_SQL As String
Private m_Order As String
Private m_Where As String
Private m_RowSource As String

Private Push As Boolean

Private ListCols As Integer
Private ColW() As Integer   'Ancho de las columnnas
Private LSel As Integer
Private MSel As Integer
Private LCount As Integer

Private Btop As Integer
Public List As Variant
Private TH As Integer

Public Event AfterUpdate(Cancel As Boolean)
Public Event Change() 'MappingInfo=Texto,Texto,-1,Change
Public Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=Texto,Texto,-1,KeyDown
Public Event KeyPress(KeyAscii As Integer) 'MappingInfo=Texto,Texto,-1,KeyPress
Public Event AfterListUpdate()

Dim m_ColumnWidths As String
Dim m_ListRows As Byte
Dim nListRows As Byte
Dim m_BoundColumn As Byte
Dim m_ListWidth As Integer
Private m_TextColumns As Byte
Private m_LimitToList As Boolean
Private m_BorderColor As OLE_COLOR
Private Factor As Single

Private Sub Texto_Change()
  
  If IsKeyPress Then m_DataChanged = (Texto.Text <> bOldValue)
  If m_DataChanged Then RaiseEvent Change
    
  '************************ SELECCION ****************************
  If Buscar Then
    Dim H As Integer
    
    FindList Texto.Text
    MSel = LSel: If MSel < 0 Then MSel = 0
    If MSel > (LCount - nListRows) Then
      VScroll = (LCount - nListRows)
    Else
      VScroll = MSel
    End If
    If LSel < 0 Then
      UserControl_Paint
      Exit Sub
    End If
    
    H = Len(Texto.Text)
    Texto.Text = Texto.Text & Mid$(Nz(List(m_BoundColumn, LSel), ""), H + 1)
    Texto.SelStart = H
    Texto.SelLength = Len(Texto.Text)
  Else
    BuscaSel
  End If
  UserControl_Paint
  '************************ SELECCION ****************************
  If CanPropertyChange("Value") Then PropertyChanged "Value"
  
End Sub

Private Sub Texto_KeyDown(KeyCode As Integer, Shift As Integer)
  Buscar = False
  If Lis.Visible Then
    Select Case KeyCode
      Case 13: KeyCode = 0
               CloseLista
               If LSel > -1 Then Texto.Text = List(m_BoundColumn, LSel)
               BuscaSel
               MSel = LSel
               m_DataChanged = True
               Texto_Change
               Update
               RaiseEvent AfterListUpdate
      Case 27: KeyCode = 0
               m_DataChanged = False
               CloseLista
      Case 38:
        KeyCode = 0: LSel = MSel
        If LSel > 0 Then
          LSel = LSel - 1: MSel = LSel
          If LSel < VScroll Then VScroll = VScroll - 1 Else Lis_Paint
        End If
      Case 40:
        KeyCode = 0: LSel = MSel
        If LSel < LCount - 1 Then
          LSel = LSel + 1: MSel = LSel
          If LSel >= VScroll + nListRows Then VScroll = VScroll + 1 Else Lis_Paint
        End If
      Case 8, 46: IsKeyPress = True
      Case 17: KeyCode = 0: CloseLista
      Case Else: Buscar = True
    End Select
  Else
    RaiseEvent KeyDown(KeyCode, Shift)
    Select Case KeyCode
      Case 0: Exit Sub
      Case 27: KeyCode = 0: Texto.Text = bOldValue
      Case 13, 40: KeyCode = 0: SendKeys "{TAB}"
      Case 38: KeyCode = 0: SendKeys "+{TAB}"
      Case 8, 46: IsKeyPress = True
      Case 17: KeyCode = 0: OpenLista
      Case Else: Buscar = True
    End Select
  End If
End Sub

Private Sub Texto_KeyPress(KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii)
  If Not Texto.Locked Then Texto.Text = TextoKeyPress(Texto.Text, KeyAscii, bDataType)
  If KeyAscii = 9 Then KeyAscii = 0
  IsKeyPress = True
End Sub

Public Sub AddRow(Row As Integer, ParamArray F())
  On Error Resume Next
  Dim I As Integer
  Dim J As Integer
  
  ReDim Preserve List(AUBound(List, 1), AUBound(List, 2) + 1)
  If err Then Exit Sub
  For I = AUBound(List, 2) To Row Step -1
    For J = 0 To AUBound(List, 1)
      If I = Row Then
        If J <= UBound(F) Then List(J, I) = F(J)
      Else
        List(J, I) = List(J, I - 1)
      End If
    Next J
  Next I
  
  RedimLista
  BuscaSel
  UserControl_Paint
  err.Clear
End Sub

Public Sub UpdateRow(Row As Integer, ParamArray F())
  On Error Resume Next
  Dim J As Integer
  For J = 0 To UBound(List, 1)
    List(J, Row) = F(J)
  Next J
  err.Clear
End Sub

Private Sub Lis_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next

  If Button <> 1 Then Exit Sub
  
  Tim.Enabled = False
  
  LSel = VScroll + (Y \ TH) '13
  
  Texto.Text = List(m_BoundColumn, LSel)

  If Texto.Text <> bOldValue Then
    m_DataChanged = True
    Lis.Visible = False
    Texto_Change
    Update
  End If
  
  Texto.SelStart = 0
  Texto.SelLength = 0

  CloseLista
  
  RaiseEvent AfterListUpdate
  err.Clear
End Sub

Private Sub Lis_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim OLSel As Integer
  OLSel = MSel
  MSel = VScroll + (Y \ TH) '13
  If OLSel <> MSel Then Lis_Paint
End Sub

Private Sub BuscaSel()
  On Error Resume Next
  Dim I As Integer
  Dim V As Variant
  LSel = -1
  If IsEmpty(List) Or Trim$(Texto) = "" Then Exit Sub
    
  If bDataType = DtNumTexto Then
    V = Val(Texto)
    For I = 0 To LCount - 1
      If V = Val(List(m_BoundColumn, I)) Then
        LSel = I
        Exit Sub
      End If
      DoEvents
    Next I
    Exit Sub
  End If
      
  If bDataType = DtStrTexto Then
    V = GetCodigo(Texto)
    For I = 0 To LCount - 1
      If V = GetCodigo(CStr(List(m_BoundColumn, I))) Then
        LSel = I
        Exit Sub
      End If
      DoEvents
    Next I
    Exit Sub
  End If
      
  V = UCase$(Texto)
  For I = 0 To LCount - 1
    If V = UCase$(List(m_BoundColumn, I)) Then
      LSel = I
      Exit Sub
    End If
    DoEvents
  Next I
  
  err.Clear
End Sub

Private Sub Lis_Paint()
  On Error Resume Next

  Dim X As Integer
  Dim Y As Integer
  Dim I As Integer
  Dim P As Integer
    
  Lis.BackColor = bColorFocus
  Lis.Cls
  
  Lis.ForeColor = Texto.ForeColor
  For P = 0 To nListRows - 1
    X = 0
    For I = 0 To ListCols - 1
      If I = ListCols - 1 Then
        DrawBox Lis.hdc, X, Y, Lis.Width - X, TH + 2, bColorFocus, , , " " & CStr(Nz(List(I, VScroll + P), ""))
      Else
        DrawBox Lis.hdc, X, Y, (ColW(I) + 1) * Factor, TH + 2, bColorFocus, , , " " & CStr(Nz(List(I, VScroll + P), ""))
      End If
      X = X + ColW(I) * Factor
    Next I
    Y = Y + TH
  Next P
  
  Lis.ForeColor = m_BorderColor
  X = 0
  For I = 0 To ListCols - 2
    X = X + (ColW(I) * Factor)
    DrawLinea Lis.hdc, X, 0, X, Lis.Height \ Screen.TwipsPerPixelY
  Next I
 
  Y = (MSel - VScroll) * TH
  If Y < (Lis.Height \ Screen.TwipsPerPixelY) Then
    DrawBoxInvert Lis.hdc, 0, Y, Lis.Width, TH + 2   '15
  End If
  
  err.Clear
End Sub

Private Sub Lis_Resize()
  VScroll.Move Lis.Width - 18, 0, 16, Lis.Height - 2
End Sub

Private Sub Texto_GotFocus()
  OldValue = Texto.Text
End Sub

Public Function Update() As Boolean
  
  On Error Resume Next
  Dim Cancel As Boolean
  Dim S As String

  Static Into As Boolean
  
  Update = True
 ' If Texto.Locked Then Exit Function
  If Not m_DataChanged Then Exit Function
  
  If Into Then Exit Function Else Into = True
  IsKeyPress = False
   
  BuscaSel
  If LSel > -1 Then Texto.Text = List(m_BoundColumn, LSel)
  
 ' If (Texto <> "") And m_LimitToList And (LSel = -1) Then
  If m_LimitToList And (LSel = -1) Then

    Extender.Visible = True
    Extender.ZOrder: Texto.SetFocus: DoEvents
    Beep
    MsgBox "El valor introducido no es un elemento de la lista.", vbInformation, Texto.Text
    DoEvents
    Extender.ZOrder: Texto.SetFocus
    Texto.Text = bOldValue
    Update = False
    GoTo Fin
  End If

  If m_DataChanged Then
  
    S = IsDataType(CByte(bDataType), Texto.Text)
    If S <> "" Then
      Extender.Visible = True
      Extender.ZOrder
      Texto.SetFocus
      DoEvents
      MsgBox S, vbExclamation, Texto.Text
      DoEvents
      Extender.ZOrder
      Texto.SetFocus
      Texto.Text = ""
      Update = False
      GoTo Fin
    End If
       
    S = Extender.DataFormat.Format
    If S <> "" Then
      Texto.Text = Format$(Texto.Text, S)
    Else
      Texto.Text = CText(Texto.Text, bDataType)
    End If
     
    IsKeyPress = False
    RaiseEvent AfterUpdate(Cancel)
    If Cancel Then
      Update = False
      Texto.SetFocus
      Texto.Text = bOldValue
      GoTo Fin
    End If
    
    DoEvents
  End If

  bOldValue = ""
  
Fin:
  m_DataChanged = False
  Into = False
  
End Function

Private Sub OpenLista()
  On Error Resume Next
  Dim R As RECT
 ' Dim R1 As RECT
 ' Dim B As Boolean
  If (Not Extender.Enabled) Or (LCount < 1) Then Exit Sub
   
  GetWindowRect UserControl.Hwnd, R
   
'  B = Extender.Parent.MDIChild

'  If B Then
'    SetParent Lis.Hwnd, MDIHwnd
'    GetWindowRect MDIHwnd, R1
'    R.Left = R.Left - R1.Left - 4
'    R.Bottom = R.Bottom - R1.Top - (42 * Factor)
'   Else
    SetParent Lis.Hwnd, GetDesktopWindow
'  End If
  
  If (R.Left + Lis.Width \ Screen.TwipsPerPixelX) > (Screen.Width \ Screen.TwipsPerPixelX) Then
    R.Left = R.Right - (Lis.Width \ Screen.TwipsPerPixelX)
  End If
  
  If (R.Bottom + Lis.Height \ Screen.TwipsPerPixelY) > (Screen.Height \ Screen.TwipsPerPixelY) Then
    R.Bottom = R.Top - (Lis.Height \ Screen.TwipsPerPixelY)
  End If

  MoveWindow Lis.Hwnd, R.Left, R.Bottom, Lis.Width \ Screen.TwipsPerPixelX, Lis.Height \ Screen.TwipsPerPixelY, True
  BuscaSel
  
  MSel = LSel: If MSel < 0 Then MSel = 0
  If MSel > (LCount - nListRows) Then
    VScroll = (LCount - nListRows)
  Else
    VScroll = MSel
  End If
  Texto.BackColor = bColorFocus
  UserControl_Paint
  Texto.SelStart = 0
  Texto.SelLength = Len(Texto)
  Lis_Paint
  Lis.Visible = True
End Sub

Private Sub CloseLista()
  Lis.Visible = False
  SetParent Lis.Hwnd, UserControl.Hwnd
  Tim.Enabled = False
End Sub

Private Sub Texto_LostFocus()
  Update
End Sub

Private Sub Texto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  IsMousedown = True
  If Not Extender.Enabled Or Locked Then Exit Sub
End Sub

Private Sub Tim_Timer()
  Dim H As Long
  H = GetCapture
  If H <> 0 Then
    If H <> VScroll.Hwnd Then
    '  LSel = LastSel
      Tim.Enabled = False
      CloseLista
    End If
  End If
End Sub

Private Sub UserControl_EnterFocus()
  If Not Enabled Then
    If Not IsMousedown Then SendKeys "{TAB}"
    Exit Sub
  End If
  
  Texto.BackColor = bColorFocus
  bOldValue = Texto.Text
 
  If Not IsMousedown Then
    Texto.SelStart = 0
    Texto.SelLength = Len(Texto)
  End If
  
  UserControl_Paint
  
  IsKeyPress = False
  IsMousedown = False
  m_DataChanged = False
  Buscar = True
End Sub

Private Sub UserControl_ExitFocus()
  BuscaSel 'revisar
  CloseLista
  Texto.BackColor = bColor
  UserControl_Paint
End Sub

Private Sub UserControl_Initialize()
  LSel = -1
  Texto.Left = 4
  Texto.Top = 3
  TH = Lis.TextHeight("0")
  Factor = 15 / Screen.TwipsPerPixelX
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  IsMousedown = True
  If Not Extender.Enabled Then Exit Sub
  Tim.Enabled = False
  If X > (Width \ Screen.TwipsPerPixelX - 18) Then
    Push = True
    If Lis.Visible Then CloseLista Else OpenLista
  Else
    If Lis.Visible Then CloseLista
  End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Not Extender.Enabled Then Exit Sub
  If Push Then
    Push = False: UserControl_Paint
    Tim.Enabled = True
  End If
End Sub

Private Sub UserControl_Paint()
  On Error Resume Next

  Dim I As Long
  Dim X As Integer
  Dim W As Integer
  X = 1

  If Not Texto.Enabled Then Texto.BackColor = bColorDisabled

  If AutoDraw Then UserControl.Cls

  DrawBox hdc, 0, 0, Width \ Screen.TwipsPerPixelX + 1, Height \ Screen.TwipsPerPixelY + 1, Texto.BackColor, m_BorderColor, bBorderStyle

  If (Not Extender.Enabled) Or (LCount < 1) Then
    I = &H100
  Else
    If Push Then I = &H4000 Else I = 0
  End If

  X = 17: W = Height \ Screen.TwipsPerPixelY - 3: If X > W Then X = W
  DrawControl hdc, Width \ Screen.TwipsPerPixelX - X - 1, 2, X, W, 3, 1 + I

  X = 1
  For I = 0 To m_TextColumns - 2
    X = X + ColW(I) * Factor
    UserControl.ForeColor = m_BorderColor
    DrawLinea hdc, X, 0, X, Height \ Screen.TwipsPerPixelY
    If Extender.Enabled Then
      UserControl.ForeColor = Texto.ForeColor
    Else
      UserControl.ForeColor = &H80000011
    End If
    If LSel > -1 Then
      W = ColW(I + 1)
      If (X + 4 + W) > (UserControl.Width \ Screen.TwipsPerPixelX) Then W = (UserControl.Width \ Screen.TwipsPerPixelX) - (X + 23)
      DrawBox hdc, X + 4, 3, W, Height \ Screen.TwipsPerPixelY, , , , CStr(Nz(List(I + 1, LSel), ""))
    End If
  Next I

  DrawBox hdc, 0, 0, Width \ Screen.TwipsPerPixelX + 1, Height \ Screen.TwipsPerPixelY + 1, , m_BorderColor, bBorderStyle
  If Lis.Visible Then Lis_Paint

  err.Clear

End Sub

Private Sub UserControl_Resize()
  On Error Resume Next
  Dim W As Integer
  W = Width \ Screen.TwipsPerPixelX - 23
  If m_TextColumns < 2 Then
    Texto.Width = W
  Else
    Texto.Width = (ColW(0) - 3) * Factor
    If Texto.Width > W Then Texto.Width = W
  End If
  Texto.Height = Height \ Screen.TwipsPerPixelY - 5
  Btop = Height \ Screen.TwipsPerPixelY + 5
  UserControl_Paint
  err.Clear
End Sub

Private Sub UserControl_Terminate()
  Tim.Enabled = False
End Sub

Private Sub VScroll_Change()
  Lis_Paint
End Sub

Private Sub VScroll_Scroll()
  Lis_Paint
End Sub

'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
  bBorderStyle = 10
  bColorFocus = vblightYellow
  BackColor = vbWhite
  UserControl.BackColor = vbWhite
  m_ListRows = 8
  m_ListWidth = Width \ Screen.TwipsPerPixelX
  m_TextColumns = 1
  m_ColumnWidths = CStr(Width \ Screen.TwipsPerPixelX)
  m_LimitToList = False
End Sub

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  On Error Resume Next
  UserControl.AutoRedraw = AutoDraw
  Lis.AutoRedraw = AutoDraw
  bDataType = PropBag.ReadProperty("DataType", 0)
  If bDataType = DtPassword Then Texto.PasswordChar = "*"
  bColor = PropBag.ReadProperty("BackColor", vbWhite)
  Texto.BackColor = bColor: UserControl.BackColor = bColor
  bColorFocus = PropBag.ReadProperty("BackColorFocus", &HDFFFFF)
  bColorDisabled = PropBag.ReadProperty("BackColorDisabled", vbgrey)
  Texto.ForeColor = PropBag.ReadProperty("ForeColor", vbBlack)
  bBorderStyle = PropBag.ReadProperty("BorderStyle", 10)
  m_BorderColor = PropBag.ReadProperty("BorderColor", 0)
  Set Font = PropBag.ReadProperty("Font", Ambient.Font)
  Texto.Alignment = PropBag.ReadProperty("Alignment", 0)
  Texto.MaxLength = PropBag.ReadProperty("MaxLength", 0)
  Texto.Enabled = PropBag.ReadProperty("Enabled", True)
  Texto.Locked = PropBag.ReadProperty("Locked", False)
  Texto.Text = PropBag.ReadProperty("Text", "")
  m_ListRows = PropBag.ReadProperty("ListRows", 10)
  m_BoundColumn = PropBag.ReadProperty("BoundColumn", 0)
  ColumnWidths = PropBag.ReadProperty("ColumnWidths", CStr(Width \ Screen.TwipsPerPixelX))
  ListWidth = PropBag.ReadProperty("ListWidth", Width \ Screen.TwipsPerPixelX)
  TextColumns = PropBag.ReadProperty("TextColumns", 1)
  Lis.BackColor = bColorFocus
  RowSource = PropBag.ReadProperty("RowSource", "")
  m_LimitToList = PropBag.ReadProperty("LimitToList", False)
  m_Where = PropBag.ReadProperty("Where", "")
  m_Order = PropBag.ReadProperty("Order", "")

  If Ambient.UserMode Then Requery
  
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "DataType", bDataType, 0
  PropBag.WriteProperty "BackColor", bColor, vbWhite
  PropBag.WriteProperty "BackColorFocus", bColorFocus, bColor
  PropBag.WriteProperty "BackColorDisabled", bColorDisabled, vbgrey
  PropBag.WriteProperty "ForeColor", Texto.ForeColor, vbBlack
  PropBag.WriteProperty "BorderStyle", bBorderStyle, 10
  PropBag.WriteProperty "BorderColor", m_BorderColor, 0
  PropBag.WriteProperty "Font", Texto.Font, Ambient.Font
  PropBag.WriteProperty "Alignment", Texto.Alignment, 0
  PropBag.WriteProperty "MaxLength", Texto.MaxLength, 0
  PropBag.WriteProperty "Enabled", Texto.Enabled, True
  PropBag.WriteProperty "Locked", Texto.Locked, False
  PropBag.WriteProperty "Text", Texto.Text, ""
  PropBag.WriteProperty "RowSource", m_RowSource, ""
  PropBag.WriteProperty "ColumnWidths", m_ColumnWidths, CStr(Width \ Screen.TwipsPerPixelX)
  PropBag.WriteProperty "ListRows", m_ListRows, 10
  PropBag.WriteProperty "BoundColumn", m_BoundColumn, 0
  PropBag.WriteProperty "ListWidth", m_ListWidth, Width \ Screen.TwipsPerPixelX
  PropBag.WriteProperty "TextColumns", m_TextColumns, 1
  PropBag.WriteProperty "LimitToList", m_LimitToList, False
  PropBag.WriteProperty "Where", m_Where, ""
  PropBag.WriteProperty "Order", m_Order, ""
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Texto,Texto,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
  BackColor = bColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
Attribute BackColor.VB_Description = "Devuelve o establece el color de fondo usado para mostrar texto y gráficos en un objeto."
  bColor = New_BackColor
  Texto.BackColor = bColor
  UserControl_Paint
  PropertyChanged "BackColor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=10,0,0,0
Public Property Get BackColorFocus() As OLE_COLOR
  BackColorFocus = bColorFocus
End Property

Public Property Let BackColorFocus(ByVal New_BackColorFocus As OLE_COLOR)
  If New_BackColorFocus = 0 Then New_BackColorFocus = vblightYellow
  bColorFocus = New_BackColorFocus
  PropertyChanged "BackColorFocus"
End Property

Public Property Get BackColorDisabled() As OLE_COLOR
  BackColorDisabled = bColorDisabled
End Property

Public Property Let BackColorDisabled(ByVal New_BackColorDisabled As OLE_COLOR)
  bColorDisabled = New_BackColorDisabled
  UserControl_Paint
  PropertyChanged "BackColorDisabled"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=1,0,0,0
Public Property Get Datatype() As EDataType
  Datatype = bDataType
End Property

Public Property Let Datatype(ByVal New_DataType As EDataType)
  bDataType = New_DataType
  If bDataType = DtPassword Then Texto.PasswordChar = "*"
  PropertyChanged "DataType"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Texto,Texto,-1,Alignment
Public Property Get Alignment() As EAlignment
  Alignment = Texto.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As EAlignment)
  Texto.Alignment() = New_Alignment
  PropertyChanged "Alignment"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Texto,Texto,-1,MaxLength
Public Property Get MaxLength() As Long
  MaxLength = Texto.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
  Texto.MaxLength() = New_MaxLength
  PropertyChanged "MaxLength"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Texto,Texto,-1,Enabled
Public Property Get Enabled() As Boolean
  Enabled = Texto.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  Texto.Enabled() = New_Enabled
  If Texto.Enabled Then Texto.BackColor = bColor
  UserControl_Paint
  PropertyChanged "Enabled"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Texto,Texto,-1,Locked
Public Property Get Locked() As Boolean
  Locked = Texto.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
  Texto.Locked() = New_Locked
  PropertyChanged "Locked"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Texto,Texto,-1,Text
Public Property Get Text() As String
  Text = Texto.Text
End Property

Public Property Let Text(ByVal New_Text As String)
'  If Not IsKeyPress Then
    m_DataChanged = False
    Texto.Text = New_Text
    bOldValue = Texto.Text
    PropertyChanged "Text"
'  End If
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Texto,Texto,-1,Text
Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "62c"
 ' If Not IsKeyPress Then Value = CValue(Texto.Text, bDataType) Else Value = Texto.Text
  Value = CValue(Texto.Text, bDataType)
End Property

Public Property Let Value(ByVal New_Value As Variant)
  Dim S As String
  If Not IsKeyPress Then
    m_DataChanged = False
    S = Extender.DataFormat.Format
    If S <> "" Then
      Texto.Text = Format$(New_Value, S)
    Else
      Texto.Text = CText(New_Value, bDataType)
    End If
    bOldValue = Texto.Text
  End If
End Property

Public Property Get OldValue() As Variant
  OldValue = CValue(bOldValue, bDataType)
End Property

Public Property Let OldValue(ByVal New_OldValue As Variant)
  bOldValue = CText(New_OldValue, bDataType)
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Texto,Texto,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
  ForeColor = Texto.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
  Texto.ForeColor() = New_ForeColor
  PropertyChanged "ForeColor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Texto,Texto,-1,BorderStyle
Public Property Get Borderstyle() As EBorderStyle
  Borderstyle = bBorderStyle
End Property

Public Property Let Borderstyle(ByVal New_BorderStyle As EBorderStyle)
  bBorderStyle = New_BorderStyle
  UserControl_Paint
  PropertyChanged "BorderStyle"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Texto,Texto,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_MemberFlags = "400"
  SelLength = Texto.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
  Texto.SelLength() = New_SelLength
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Texto,Texto,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_MemberFlags = "400"
  SelStart = Texto.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
  Texto.SelStart() = New_SelStart
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Texto,Texto,-1,SelText
Public Property Get SelText() As String
Attribute SelText.VB_MemberFlags = "400"
  SelText = Texto.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
  Texto.SelText() = New_SelText
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,0,0,""
Public Property Get RowSource() As String
  RowSource = m_RowSource
End Property

Public Property Let RowSource(ByVal New_RowSource As String)
  m_RowSource = New_RowSource
  PropertyChanged "RowSource"
End Property

Private Sub RedimLista()
  LCount = 0: nListRows = m_ListRows
  If Not IsEmpty(List) Then LCount = UBound(List, 2) + 1
  If LCount < nListRows Then nListRows = LCount
  If m_ListRows < 1 Then m_ListRows = 1
  Lis.Height = nListRows * TH + 2
  If LCount > nListRows Then
    VScroll.Left = Lis.Width - VScroll.Width - 2
  Else
    VScroll.Left = Lis.Width
  End If
  VScroll.Max = LCount - ListRows
  Lis_Paint
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,0,0,""
Public Property Get ColumnWidths() As String
  ColumnWidths = m_ColumnWidths
End Property

Public Property Let ColumnWidths(ByVal New_ColumnWidths As String)
  On Error Resume Next
  m_ColumnWidths = New_ColumnWidths
  ColW() = IntSplit(m_ColumnWidths, ";")
  ListWidth = SumArray(ColW)
  ListCols = UBound(ColW) + 1
  UserControl_Paint
  PropertyChanged "ColumnWidths"
  err.Clear
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=1,0,0,8
Public Property Get ListRows() As Byte
  ListRows = m_ListRows
End Property

Public Property Let ListRows(ByVal New_ListRows As Byte)
  m_ListRows = New_ListRows
  PropertyChanged "ListRows"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=1,0,0,0
Public Property Get BoundColumn() As Byte
  BoundColumn = m_BoundColumn
End Property

Public Property Let BoundColumn(ByVal New_BoundColumn As Byte)
  m_BoundColumn = New_BoundColumn
  PropertyChanged "BoundColumn"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=14,0,0,0
Public Property Get Column(Col As Byte) As Variant
  If LSel > -1 Then Column = List(Col, LSel) Else Column = ""
End Property

Public Property Let Column(Col As Byte, Valor As Variant)
  List(Col, LSel) = Valor
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Texto,Texto,-1,Font
Public Property Get Font() As StdFont
  Set Font = Texto.Font
End Property

Public Property Set Font(ByVal New_Font As StdFont)
  Set Texto.Font = New_Font
  Set UserControl.Font = Texto.Font
  UserControl_Paint
  PropertyChanged "Font"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=7,0,0,0
Public Property Get ListWidth() As Integer
  ListWidth = m_ListWidth
End Property

Public Property Let ListWidth(ByVal New_ListWidth As Integer)
  If New_ListWidth = 0 Then New_ListWidth = SumArray(ColW)
  m_ListWidth = New_ListWidth
  If m_ListWidth < ScaleWidth Then m_ListWidth = ScaleWidth
  Lis.Width = m_ListWidth
  PropertyChanged "ListWidth"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=1,0,0,0
Public Property Get TextColumns() As Byte
  TextColumns = m_TextColumns
End Property

Public Property Let TextColumns(ByVal New_TextColumns As Byte)
  Dim I As Integer
  m_TextColumns = New_TextColumns
  If m_TextColumns < 1 Then m_TextColumns = 1
  If m_TextColumns > ListCols Then m_TextColumns = ListCols
  UserControl_Resize
  PropertyChanged "TextColumns"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,0
Public Property Get LimitToList() As Boolean
  LimitToList = m_LimitToList
End Property

Public Property Let LimitToList(ByVal New_LimitToList As Boolean)
  m_LimitToList = New_LimitToList
  PropertyChanged "LimitToList"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,0,0,
Public Property Get Where() As String
Attribute Where.VB_Description = "Filtro del ComboBox"
  Where = m_Where
End Property

Public Property Let Where(ByVal New_Where As String)
  m_Where = New_Where
  PropertyChanged "Where"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,0,0,
Public Property Get Order() As String
Attribute Order.VB_Description = "Ordenación del RecordSet"
  Order = m_Order
End Property

Public Property Let Order(ByVal New_Order As String)
  m_Order = New_Order
  PropertyChanged "Order"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=5
Public Sub Requery()
    
  On Error Resume Next
  Dim S() As String
  Dim I As Integer
  Dim X As Integer
  Dim Y As Integer

  If m_RowSource = "" Then Exit Sub

  m_SQL = UCase$(m_RowSource)
  If InStr(m_SQL, "SELECT") Then
    If m_Where <> "" Then m_SQL = m_SQL & " WHERE " & m_Where
    If m_Order <> "" Then m_SQL = m_SQL & " ORDER BY " & m_Order
    List = DB.GetRows(m_SQL)
  Else
    S() = Split(m_RowSource & ";", ";")
    ReDim List(ListCols - 1, (UBound(S) \ ListCols) - 1)
    For Y = 0 To UBound(List, 2)
      For X = 0 To UBound(List)
        If I <= UBound(S) Then List(X, Y) = S(I)
        I = I + 1
      Next X
    Next Y
  End If
  RedimLista
  BuscaSel
  UserControl_Paint
  err.Clear
End Sub

Private Sub FindList(Texto As String)
  On Error Resume Next
  Dim I As Integer
  Dim Lng As Integer
  LSel = -1
  Lng = Len(Texto)
  If IsEmpty(List) Or Trim$(Texto) = "" Then Exit Sub
  For I = 0 To LCount - 1
    If UCase$(Texto) = UCase$(Left$(List(m_BoundColumn, I), Lng)) Then
      LSel = I
      Exit Sub
    End If
    DoEvents
  Next I
  Exit Sub
End Sub

Public Function FindCodigo(ByVal Codigo As Variant) As String
  On Error Resume Next
  Dim S As String
  Dim I As Integer
  Dim Lng As Integer
  If IsNull(Codigo) Then Exit Function
  S = Trim$(Codigo) & " - "
  If IsEmpty(List) Or S = "" Then Exit Function
  Lng = Len(S)
  For I = 0 To LCount - 1
    If S = Left$(List(m_BoundColumn, I), Lng) Then
      FindCodigo = List(m_BoundColumn, I)
      Exit Function
    End If
  Next I
  FindCodigo = Codigo
  Exit Function
End Function

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=10,0,0,0
Public Property Get BorderColor() As OLE_COLOR
  BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
  m_BorderColor = New_BorderColor
  UserControl_Paint
  PropertyChanged "BorderColor"
End Property

Public Property Get TextBoxLeft() As Integer
  TextBoxLeft = UserControl.Texto.Left
End Property

Public Property Let TextBoxLeft(ByVal New_Left As Integer)
  UserControl.Texto.Left = New_Left
End Property

Public Property Get TextBoxTop() As Integer
  TextBoxTop = UserControl.Texto.Top
End Property

Public Property Let TextBoxTop(ByVal New_Top As Integer)
  UserControl.Texto.Top = New_Top
End Property

