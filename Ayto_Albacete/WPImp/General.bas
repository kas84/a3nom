Attribute VB_Name = "General"
Option Explicit

Public DB As New CData
Public DBCoput As New CData
Public SrvActual As Long
Public SrvTotal As Long
Public SrvMensaje As String
Public ContImport As Byte
Public TipoCon As Byte '0=Open ; 1=Previsiones

Public Function SrvPorcentaje() As Single
  On Error Resume Next
  SrvPorcentaje = (SrvActual * 100) \ IIf(SrvTotal = 0, 1, SrvTotal)
  err.Clear
End Function

Public Function FormatFecha(F As Variant) As String
  If IsNull(F) Then
    FormatFecha = "NULL"
  Else
    FormatFecha = "TO_DATE('" & Format$(F, "ddmmyyyyHhNnSs") & "','DDMMYYYYHH24MISS')"
  End If
End Function

Public Sub AddBackup(S As String)
'  If Backup Then
    Open App.Path & "\WPImp.log" For Append Access Write Shared As #2
    Print #2, S
    Close #2
'  End If
End Sub

Public Sub MsgErr(Optional Texto As String = "Error")
  If err Then
    MsgBox err.Description, 16, err.Source & "  (" & Texto & ":" & CStr(err) & ")"
    'App.LogEvent Err.Source & ":" & Texto & ":" & CStr(Err) & ":" & Err.Description, 1
  End If
  Exit Sub
End Sub

Public Function Filtrar(rstTemp As ADODB.Recordset, strField As String, strFilter As String) As ADODB.Recordset
  ' Le quitamos todos los filtros al recordset rstemp, por si acaso.
  rstTemp.Filter = adFilterNone
  ' Establece un filtro sobre el objeto Recordset especificado y,
  ' después, abre un nuevo objeto Recordset.
  rstTemp.Filter = strField & " = '" & strFilter & "'"
  Set Filtrar = rstTemp
End Function

