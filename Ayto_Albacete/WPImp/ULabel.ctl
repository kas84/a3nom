VERSION 5.00
Begin VB.UserControl ULabel 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0000FF00&
   CanGetFocus     =   0   'False
   ClientHeight    =   1515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3240
   ClipControls    =   0   'False
   DataBindingBehavior=   1  'vbSimpleBound
   ForwardFocus    =   -1  'True
   PropertyPages   =   "ULabel.ctx":0000
   ScaleHeight     =   101
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   216
   ToolboxBitmap   =   "ULabel.ctx":000F
End
Attribute VB_Name = "ULabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Private m_BorderStyle As EBorderStyle
Private m_BackColor As OLE_COLOR
Private m_ForeColor As OLE_COLOR
Private m_BorderColor As OLE_COLOR
Private m_AlignText As EPosition
Private m_Text As String
Private m_Datatype As EDataType

Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Ocurre cuando el usuario presiona el botón del mouse mientras un objeto tiene el enfoque."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Ocurre cuando el usuario libera el botón del mouse mientras un objeto tiene el enfoque."
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Ocurre cuando el usuario presiona y libera un botón del mouse encima de un objeto."

Private Sub UserControl_Paint()
  UserControl.Cls
  DrawBox UserControl.hdc, 0, 0, UserControl.Width \ Screen.TwipsPerPixelX + 1, UserControl.Height \ Screen.TwipsPerPixelY + 1, m_BackColor, m_BorderColor, m_BorderStyle
  If m_Text <> "" Then
    If UserControl.Enabled Then
      UserControl.ForeColor = m_ForeColor
      DrawBox UserControl.hdc, 4, 1, UserControl.Width \ Screen.TwipsPerPixelX - 6, UserControl.Height \ Screen.TwipsPerPixelY - 2, , , , m_Text, m_AlignText
    Else
      UserControl.ForeColor = vbWhite
      DrawBox UserControl.hdc, 4, 1, UserControl.Width \ Screen.TwipsPerPixelX - 6, UserControl.Height \ Screen.TwipsPerPixelY - 2, , , , m_Text, m_AlignText
      UserControl.ForeColor = vbdarkgrey
      DrawBox UserControl.hdc, 3, 0, UserControl.Width \ Screen.TwipsPerPixelX - 6, UserControl.Height \ Screen.TwipsPerPixelY - 2, , , , m_Text, m_AlignText
    End If
  End If
End Sub

'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
  If m_AlignText = 0 Then m_AlignText = 4
  If m_BackColor = 0 Then m_BackColor = &H8000000F
  If UserControl.Font Is Nothing Then Set UserControl.Font = Ambient.Font
End Sub

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_BackColor = PropBag.ReadProperty("BackColor", vbWhite)
  UserControl.BackColor = m_BackColor
  m_ForeColor = PropBag.ReadProperty("ForeColor", vbBlack)
  UserControl.ForeColor = m_ForeColor
  m_BorderStyle = PropBag.ReadProperty("BorderStyle", 10)
  m_BorderColor = PropBag.ReadProperty("BorderColor", vbBlack)
  UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
  Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
  m_AlignText = PropBag.ReadProperty("AlignText", 0)
  m_Datatype = PropBag.ReadProperty("DataType", 0)
  m_Text = PropBag.ReadProperty("Text", "")
  UserControl_Paint
End Sub

Private Sub UserControl_Resize()
  UserControl_Paint
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "BackColor", m_BackColor, vbWhite
  PropBag.WriteProperty "ForeColor", m_ForeColor, vbBlack
  PropBag.WriteProperty "BorderColor", m_BorderColor, vbBlack
  PropBag.WriteProperty "BorderStyle", m_BorderStyle, 10
  PropBag.WriteProperty "Font", UserControl.Font, Ambient.Font
  PropBag.WriteProperty "AlignText", m_AlignText, 0
  PropBag.WriteProperty "DataType", m_Datatype, 0
  PropBag.WriteProperty "Text", m_Text, ""
  PropBag.WriteProperty "Enabled", UserControl.Enabled, True
End Sub

Public Property Get Alignment() As EAlignment
  Select Case m_AlignText
    Case UpLeft, CenterLeft, DownLeft: Alignment = Izquierda
    Case UpCenter, Center, DownCenter: Alignment = Centro
    Case UpRight, CenterRight, DownRigth: Alignment = Derecha
  End Select
End Property

Public Property Let Alignment(ByVal New_Alignment As EAlignment)
  Select Case New_Alignment
    Case Izquierda: m_AlignText = 4
    Case Centro: m_AlignText = 5
    Case Derecha: m_AlignText = 6
  End Select
  UserControl_Paint
End Property

Public Property Get AlignText() As EPosition
  AlignText = m_AlignText
End Property

Public Property Let AlignText(ByVal New_AlignText As EPosition)
  m_AlignText = New_AlignText
  UserControl_Paint
  PropertyChanged "Alignment"
End Property

Public Property Get Text() As String
Attribute Text.VB_MemberFlags = "2c"
  Text = m_Text
End Property

Public Property Let Text(ByVal New_Text As String)
  m_Text = New_Text
  UserControl_Paint
  PropertyChanged "Text"
End Property

Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = "ULabelP1"
  Caption = m_Text
End Property

Public Property Let Caption(ByVal New_Text As String)
  m_Text = New_Text
  UserControl_Paint
  PropertyChanged "Text"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Devuelve o establece el color de primer plano usado para mostrar textos y gráficos en un objeto."
  ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
  m_ForeColor = New_ForeColor
  UserControl.ForeColor = New_ForeColor
  UserControl_Paint
  PropertyChanged "ForeColor"
End Property

Public Property Get Font() As StdFont
  Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As StdFont)
  Set UserControl.Font = New_Font
  UserControl_Paint
  PropertyChanged "Font"
End Property

Public Property Get Borderstyle() As EBorderStyle
Attribute Borderstyle.VB_Description = "Devuelve o establece el estilo del borde de un objeto."
  Borderstyle = m_BorderStyle
End Property

Public Property Let Borderstyle(ByVal New_BorderStyle As EBorderStyle)
  m_BorderStyle = New_BorderStyle
  UserControl_Paint
  PropertyChanged "BorderStyle"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=10,0,0,
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Devuelve o establece el color de fondo usado para mostrar texto y gráficos en un objeto."
  BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  m_BackColor = New_BackColor
  UserControl.BackColor = New_BackColor
  UserControl_Paint
  PropertyChanged "BackColor"
End Property

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Click()
  RaiseEvent Click
End Sub

Public Property Get Enabled() As Boolean
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  If New_Enabled = UserControl.Enabled Then Exit Property
  UserControl.Enabled() = New_Enabled
  UserControl_Paint
  PropertyChanged "Enabled"
End Property

Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
  Value = CValue(m_Text, m_Datatype)
End Property

Public Property Let Value(ByVal New_Value As Variant)
  Dim S As String
  S = Extender.DataFormat.Format
  If S <> "" Then
    Text = Format$(New_Value, S)
  Else
    Text = CText(New_Value, m_Datatype)
  End If
  UserControl_Paint
End Property

Public Property Get Datatype() As EDataType
  Datatype = m_Datatype
End Property

Public Property Let Datatype(ByVal New_DataType As EDataType)
  m_Datatype = New_DataType
  UserControl_Paint
  PropertyChanged "DataType"
End Property

Public Property Get BorderColor() As OLE_COLOR
  BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
  m_BorderColor = New_BorderColor
  UserControl_Paint
  PropertyChanged "BorderColor"
End Property

Public Function Update() As Boolean
  UserControl_Paint
  Update = True
End Function
