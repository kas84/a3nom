Attribute VB_Name = "General"
Option Explicit
Public DB As New CData
Public TipoCon As Byte '0=Open ; 1=Previsiones
Public SrvMensaje As String
Public SrvActual As Long
Public SrvTotal As Long

Type TypeCon
  num As Byte
  val As Long
  Text As String
  
End Type


Public Function FExist(Fichero As String) As Boolean
'Devuelve true si la cadena que se le pasa coincide con un fichero del sistema
  On Error Resume Next
  If Len(Trim$(Fichero)) < 1 Then Exit Function
  If Len(Dir$(Fichero)) > 0 Then FExist = True
  If err <> 0 Then FExist = False
  Exit Function
End Function

Public Function IFec(Fecha As Variant) As String
  If IsZero(Fecha) Then
    IFec = "NULL"
    Exit Function
  End If
  'IFec = "TO_DATE('" & Format$(Fecha, "ddmmyyyyhhnnss") & "','DDMMYYYYHH24MISS')"
  IFec = Format$(Fecha, "\'dd\/mm\/yyyy hh\:mm\'")
  'If DBType = 0 Then 'ACCESS
    'IFec = Format$(Fecha, "\#mm\/d\/yyyy hh\:mm\ s AM/PM\#")
    'IFec = Format$(Fecha, "\#mm\/d\/yyyy hh\:mm\  AM/PM\#")
  'ElseIf DBType = 1 Then 'SQL SERVER
    'IFec = Format$(Fecha, "\'dd\/mm\/yyyy hh\:mm\ s\'")
  'ElseIf DBType = 2 Then 'ORACLE
    'IFec = "TO_DATE('" & Format$(Fecha, "ddmmyyyyhhnnss") & "','DDMMYYYYHH24MISS')"
  'End If
End Function

Public Function IStr(V As Variant) As String
  On Error Resume Next
  If IsNull(V) Then
    IStr = "'NULL'"
  Else
    IStr = "'" & CStr(V) & "'"
  End If
End Function

Private Sub AddBackup(S As String)

    Open App.Path & "\Exportacion.log" For Append Access Write Shared As #2
    Print #2, S
    Close #2

End Sub

Public Sub PrintLog(Texto As String)
 
On Error Resume Next
Texto = String(130, "-") & vbCrLf & " " & Now & " -- " & Texto & vbCrLf & String(130, "-")
  
AddBackup Texto
err.Clear
End Sub

Public Sub MsgErr(Optional Texto As String = "Error")
  If err Then
    MsgBox err.Description, 16, err.Source & "  (" & Texto & ":" & CStr(err) & ")"
    'App.LogEvent Err.Source & ":" & Texto & ":" & CStr(Err) & ":" & Err.Description, 1
  End If
  Exit Sub
End Sub

Public Function SrvPorcentaje() As Single
  On Error Resume Next
  SrvPorcentaje = (SrvActual * 100) \ IIf(SrvTotal = 0, 1, SrvTotal)
  err.Clear
End Function
Public Sub ChangeSource(CCombo As UComboBox, SQL As String, PorCodigo As Boolean, ColumnWidths As String)
  CCombo.RowSource = SQL
  If PorCodigo Then
    CCombo.Order = "Codigo"
    CCombo.ColumnWidths = ColumnWidths
    CCombo.MaxLength = 100
    CCombo.BoundColumn = 0
    CCombo.Datatype = DtMayusculas
  Else
    CCombo.Order = "OPERARIO"
    CCombo.ColumnWidths = ColumnWidths
    CCombo.MaxLength = 0
    CCombo.BoundColumn = 1
    CCombo.Datatype = DtGeneral
  End If
  CCombo.Requery
End Sub



Public Function GetDatos(ByVal S As Variant, N As Byte) As String
'Obtenemos la información necesaria del campo Datos de Validación
'  N=1 : Incidencias/Causas
'  N=2 : Contadores
'  N=3 : Entradas y Salidas
'  N=4 : Incidencias
'  N=5 : Causas

  Dim L As Integer
  Dim L1 As Integer
  Dim L2 As Integer
  Dim L3 As Integer

  If IsNull(S) Then Exit Function
  L = Len(S): If L < 3 Then Exit Function
  
  L1 = Asc(Mid$(S, 1, 1)) * 2  'Longitud en Datos de incidencias y causas
  L2 = Asc(Mid$(S, 2, 1)) * 3  'Longitud en Datos de Contadores
  L3 = Asc(Mid$(S, 3, 1)) * 8  'Longitud en Datos de Entradas y Salidas
  If L < (3 + L1 + L2 + L3) Then Exit Function

  Select Case N
    Case 1, 4, 5: S = Mid$(S, 4, L1)
    Case 2: S = Mid$(S, 4 + L1, L2)
    Case 3: S = Mid$(S, 4 + L1 + L2, L3)
    Case Else: S = ""
  End Select

  If N = 4 Then S = Left$(S, L1 \ 2)
  If N = 5 Then S = Right$(S, L1 \ 2)

  GetDatos = S

End Function

