VERSION 5.00
Begin VB.UserControl UTextBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0000FF00&
   ClientHeight    =   960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3075
   ClipControls    =   0   'False
   DataBindingBehavior=   1  'vbSimpleBound
   ScaleHeight     =   64
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   205
   ToolboxBitmap   =   "UTextBox.ctx":0000
   Begin VB.TextBox Texto 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "UTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private bDataType As EDataType
Private bBorderStyle As EBorderStyle
Private bColor As OLE_COLOR
Private bOldValue As String
Attribute bOldValue.VB_VarMemberFlags = "40"
Private m_DataChanged As Boolean
Private IsKeyPress As Boolean
Private IsMousedown As Boolean

Public Event AfterUpdate(Cancel As Boolean)
Event Change() 'MappingInfo=Texto,Texto,-1,Change
Attribute Change.VB_Description = "Ocurre cuando cambia el contenido de un control."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=Texto,Texto,-1,KeyDown
Attribute KeyDown.VB_Description = "Ocurre cuando el usuario presiona una tecla mientras un objeto tiene el enfoque."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=Texto,Texto,-1,KeyPress
Attribute KeyPress.VB_Description = "Ocurre cuando el usuario presiona y libera una tecla ANSI."
Event Click() 'MappingInfo=Texto,Texto,-1,Click
Attribute Click.VB_Description = "Ocurre cuando el usuario presiona y libera un botón del mouse encima de un objeto."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Texto,Texto,-1,MouseDown
Attribute MouseDown.VB_Description = "Ocurre cuando el usuario presiona el botón del mouse mientras un objeto tiene el enfoque."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=Texto,Texto,-1,MouseUp
Attribute MouseUp.VB_Description = "Ocurre cuando el usuario libera el botón del mouse mientras un objeto tiene el enfoque."
Private m_BorderColor As OLE_COLOR

Private Sub Texto_Change()
  If IsKeyPress Then m_DataChanged = (Texto.Text <> bOldValue)
  If m_DataChanged Then RaiseEvent Change
  If CanPropertyChange("Value") Then PropertyChanged "Value"
End Sub

Private Sub Texto_GotFocus()
  IsKeyPress = False
End Sub

Private Sub Texto_KeyDown(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyDown(KeyCode, Shift)
  Select Case KeyCode
    Case 27: KeyCode = 0: Texto.Text = bOldValue: m_DataChanged = False
    Case 13, 40: KeyCode = 0: SendKey vbKeyTab
    Case 38: KeyCode = 0: SendKey vbKeyTab, True
    Case 46: IsKeyPress = True
  End Select
End Sub

Private Sub Texto_KeyPress(KeyAscii As Integer)
  Dim S As String
  RaiseEvent KeyPress(KeyAscii)
  IsKeyPress = True
  If Not Texto.Locked Then
    S = Texto.Text
    Texto.Text = TextoKeyPress(Texto.Text, KeyAscii, bDataType)
    If Texto.Text <> S Then Texto_Change
  End If
  If KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub Texto_LostFocus()
  Update
End Sub

Public Function Update() As Boolean
  On Error Resume Next
  Dim Cancel As Boolean
  Dim S As String
  Static Into As Boolean
 
  Update = True
  IsKeyPress = False
  If Texto.Locked Then Exit Function
  If Into Then Exit Function Else Into = True
  
  If m_DataChanged Then
                      
    S = IsDataType(bDataType, Texto.Text)
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
        
    RaiseEvent AfterUpdate(Cancel)
    If Cancel Then
      Texto.SetFocus
      Texto.Text = bOldValue
      Update = False
      GoTo Fin
    End If
          
    'DoEvents  REVISAR QUITADO EL 15-2-2002
  End If
  
  bOldValue = ""

Fin:
  m_DataChanged = False
  Into = False
End Function

Private Sub Texto_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  IsMousedown = True
  IsKeyPress = False
  If Not Enabled Then Exit Sub
  RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_EnterFocus()
  If Not Enabled Then
    If Not IsMousedown Then SendKey vbKeyTab
    Exit Sub
  End If
  Texto.BackColor = BackColorFocus
  bOldValue = Texto.Text
  UserControl_Paint
  If Not IsMousedown Then
    Texto.SelStart = 0
    Texto.SelLength = Len(Texto)
  End If
  IsMousedown = False
  m_DataChanged = False
End Sub

Private Sub UserControl_ExitFocus()
  Texto.BackColor = bColor
  UserControl_Paint
End Sub

Private Sub UserControl_Initialize()
  Texto.Left = 4
  Texto.Top = 3
End Sub

Private Sub UserControl_Paint()
  UserControl.Cls
  If Not Texto.Enabled Then Texto.BackColor = BackColorDisabled
  DrawBox UserControl.hdc, 0, 0, UserControl.Width \ Screen.TwipsPerPixelX + 1, UserControl.Height \ Screen.TwipsPerPixelX + 1, Texto.BackColor, m_BorderColor, bBorderStyle
End Sub

Private Sub UserControl_Resize()
  On Error Resume Next
  Texto.Height = UserControl.Height \ Screen.TwipsPerPixelY - 5
  Texto.Width = UserControl.Width \ Screen.TwipsPerPixelX - 6
  UserControl_Paint
  Exit Sub
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Texto,Texto,-1,Font
Public Property Get Font() As StdFont
  Set Font = Texto.Font
End Property

Public Property Set Font(ByVal New_Font As StdFont)
  Set Texto.Font = New_Font
  PropertyChanged "Font"
End Property

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  bDataType = PropBag.ReadProperty("DataType", 0)
  If bDataType = DtPassword Then Texto.PasswordChar = "*"
  bColor = PropBag.ReadProperty("BackColor", vbWhite)
  Texto.BackColor = bColor
  Texto.ForeColor = PropBag.ReadProperty("ForeColor", vbBlack)
  Borderstyle = PropBag.ReadProperty("BorderStyle", 10)
  m_BorderColor = PropBag.ReadProperty("BorderColor", 0)
  Set Texto.Font = PropBag.ReadProperty("Font", Ambient.Font)
  Texto.Alignment = PropBag.ReadProperty("Alignment", 0)
  Texto.MaxLength = PropBag.ReadProperty("MaxLength", 0)
  Texto.Enabled = PropBag.ReadProperty("Enabled", True)
  Texto.Locked = PropBag.ReadProperty("Locked", False)
  Texto.Text = PropBag.ReadProperty("Text", "")
  UserControl_Paint
End Sub


'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "DataType", bDataType, 0
  PropBag.WriteProperty "BackColor", bColor, vbWhite
  PropBag.WriteProperty "ForeColor", Texto.ForeColor, vbBlack
  PropBag.WriteProperty "BorderStyle", bBorderStyle, 10
  PropBag.WriteProperty "BorderColor", m_BorderColor, 0
  PropBag.WriteProperty "Font", Texto.Font, Ambient.Font
  PropBag.WriteProperty "Alignment", Texto.Alignment, 0
  PropBag.WriteProperty "MaxLength", Texto.MaxLength, 0
  PropBag.WriteProperty "Enabled", Texto.Enabled, True
  PropBag.WriteProperty "Locked", Texto.Locked, False
  PropBag.WriteProperty "Text", Texto.Text, ""
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Texto,Texto,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Devuelve o establece el color de fondo usado para mostrar texto y gráficos en un objeto."
  BackColor = bColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  bColor = New_BackColor
  Texto.BackColor() = bColor
  UserControl.BackColor = bColor
  UserControl_Paint
  PropertyChanged "BackColor"
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

'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
  bBorderStyle = 10
  BackColor = vbWhite
  Set Texto.Font = Ambient.Font
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Texto,Texto,-1,Alignment
Public Property Get Alignment() As EAlignment
Attribute Alignment.VB_Description = "Devuelve o establece la alineación de un control CheckBox u OptionButton, o el texto de un control."
  Alignment = Texto.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As EAlignment)
  Texto.Alignment() = New_Alignment
  PropertyChanged "Alignment"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Texto,Texto,-1,MaxLength
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Devuelve o establece el número máximo de caracteres que se puede escribir en un control."
  MaxLength = Texto.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
  Texto.MaxLength() = New_MaxLength
  PropertyChanged "MaxLength"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Texto,Texto,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
  Enabled = Texto.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  Texto.Enabled() = New_Enabled
  If Texto.Enabled Then Texto.BackColor = bColor Else Texto.BackColor = BackColorDisabled
  UserControl_Paint
  PropertyChanged "Enabled"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Texto,Texto,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determina si se puede modificar un control."
  Locked = Texto.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
  Texto.Locked() = New_Locked
  PropertyChanged "Locked"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Texto,Texto,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Devuelve o establece el texto contenido en el control."
  Text = Texto.Text
End Property

Public Property Let Text(ByVal New_Text As String)
  If Not IsKeyPress Then
    m_DataChanged = False
    Texto.Text = New_Text
    bOldValue = Texto.Text
    PropertyChanged "Text"
  End If
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Texto,Texto,-1,Text
Public Property Get Value() As Variant
Attribute Value.VB_Description = "Devuelve o establece el texto contenido en el control."
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "62c"
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
Attribute ForeColor.VB_Description = "Devuelve o establece el color de primer plano usado para mostrar textos y gráficos en un objeto."
  ForeColor = Texto.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
  Texto.ForeColor() = New_ForeColor
  PropertyChanged "ForeColor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Texto,Texto,-1,BorderStyle
Public Property Get Borderstyle() As EBorderStyle
Attribute Borderstyle.VB_Description = "Devuelve o establece el estilo del borde de un objeto."
  Borderstyle = bBorderStyle
End Property

Public Property Let Borderstyle(ByVal New_BorderStyle As EBorderStyle)
  bBorderStyle = New_BorderStyle
  UserControl_Paint
  PropertyChanged "BorderStyle"
End Property

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

Private Sub Texto_Click()
  RaiseEvent Click
End Sub

Private Sub Texto_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Property Get TBox(ProcName As String) As Variant
  TBox = CallByName(Texto, ProcName, VbGet)
End Property

Public Property Let TBox(ProcName As String, Valor As Variant)
  CallByName Texto, ProcName, VbLet, Valor
End Property

