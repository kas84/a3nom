Attribute VB_Name = "Config"
Option Explicit

Public Cfg As New CConfig

Public Const bIbmterm = "IBMTERM"     ' Identificador del bloque del fichero IBMTERM
Public Const bIbmser = "IBMSER"       ' Identificador del bloque del fichero IBMSER
Public Const bCOM = "COMUNICACIONES"  ' Identificador del bloque de la configuración general de comunicaciones

Public Sub InitCfg()
  On Error GoTo ErrInitcfg
  
  Cfg.Datos.Add bPC, Nz(DB.Lookup("Datos", "Cfg", "ID=" & DB.IStr(bPC)), "")
  
  If (Cfg.Datos(bPC) = "") Then
    If MsgBox("Es la primera vez que se arrancan las comunicaciones en este PC [" & PCName & "]. ¿Desea continuar con la operación?", vbInformation + vbYesNo, App.Title) = vbNo Then End
    DB.Insert "Cfg", "ID,Datos", DB.IStr(bPC) & ",'KeyFile = '"
    Cfg.Datos.Item(bPC) = Nz(DB.Lookup("Datos", "Cfg", "ID=" & DB.IStr(bPC)), "")
  End If
  Exit Sub
  
ErrInitcfg:
  MsgErr "InitCfg"
  End
End Sub

Public Sub InitINI()
  On Error GoTo ErrInitIni
  Dim KeyFile As String
  
  bUSR = "USER." & UserName
  Cfg.Datos.RemoveAll
  KeyFile = App.Path & "\winplus.ini"
  'KeyFile = GetSetting("WinPLUS.Com", "General", "Keyfile", "")
  'If Not FExist(KeyFile) Then
  '  Dlg.DialogTitle = "¿ Donde está el fichero WINPLUS.INI ?"
  '  Dlg.Filter = "WINPLUS.INI"
  '  Dlg.InitDir = App.Path
  '  Do
  '    Dlg.ShowOpen
  '    If Dlg.Cancel Then End
  '    KeyFile = Dlg.FileName
  '    If FExist(KeyFile) Then Exit Do
  '  Loop
  '  SaveSetting App.Title, "General", "Keyfile", KeyFile
  'End If
  
  Cfg.Datos.Add "INI", Cfg.Decode(FLoad(KeyFile))
  Exit Sub
  
ErrInitIni:
  MsgErr "InitIni"
  End
End Sub

Public Sub LoadBlock(Block As Variant, Machacar As Boolean)
  Dim S As String
  
  S = CStr(Block)
  If Cfg.Datos.Exists(S) Then
    If Machacar Then Cfg.Datos.Item(S) = Nz(DB.Lookup("Datos", "Cfg", "ID=" & DB.IStr(S)), "")
  Else
    Cfg.Datos.Add S, Nz(DB.Lookup("Datos", "Cfg", "ID=" & DB.IStr(S)), "")
  End If
End Sub

Public Property Get KeyConfig(ByVal Clave As String) As Variant
  Dim V As Variant
  V = DB.GetRow("SELECT Valor,Tipo FROM Config WHERE Clave=" & DB.IStr(Clave))
  If IsEmpty(V) Then Exit Property
  If IsNull(V(0)) Then Exit Property
  KeyConfig = CValue(V(0), Nz(V(1), 0))
End Property

Public Property Let KeyConfig(ByVal Clave As String, ByVal Valor As Variant)
  Dim T As Variant
  T = DB.Lookup("Tipo", "Config", "Clave=" & DB.IStr(Clave))
  If IsEmpty(T) Or IsNull(T) Then
    DB.Insert "Config", "Clave,Valor,Tipo,Descripcion", DB.IStr(Clave), DB.IStr(CText(Valor, 0)), 0, DB.IStr("Nueva clave (Creada el " & Now & ")")
  Else
    DB.Update "Config", "Valor=" & DB.IStr(CText(Valor, T)), "Clave=" & DB.IStr(Clave)
  End If
End Property
