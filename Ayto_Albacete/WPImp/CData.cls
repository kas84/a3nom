VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public DbType As Byte

Private WithEvents CnEv As ADODB.Connection
Attribute CnEv.VB_VarHelpID = -1
Public Cn As New ADODB.Connection
Public UserID As String
Public Password As String
Public ServerName As String
Public DatabaseName As String
Public Errors As ADODB.Errors
Public Status As Long
Private SQL As String
Private V As Variant

Public Function Lookup(Campo As String, Tabla As String, Optional Where As String = "0=0") As Variant
  On Error Resume Next
  V = GetRows("SELECT " & Campo & " FROM " & Tabla & " WHERE " & Where, 1)
  If Not IsEmpty(V) Then Lookup = V(0, 0)
End Function

Public Function Count(Campo As String, Tabla As String, Optional Where As String = "0=0") As Long
  On Error Resume Next
  V = GetRows("SELECT COUNT(" & Campo & ") AS CUENTA FROM " & Tabla & " WHERE " & Where, 1)
  If Not IsEmpty(V) Then Count = CLng(V(0, 0))
End Function

Public Function Min(Campo As String, Tabla As String, Optional Where As String = "0=0") As Variant
  On Error Resume Next
  V = GetRows("SELECT MIN(" & Campo & ") AS MINIMO FROM " & Tabla & " WHERE " & Where, 1)
  If Not IsEmpty(V) Then Min = V(0, 0)
End Function

Public Function Max(Campo As String, Tabla As String, Optional Where As String = "0=0") As Variant
  On Error Resume Next
  V = GetRows("SELECT MAX(" & Campo & ") AS MAXIMO FROM " & Tabla & " WHERE " & Where, 1)
  If Not IsEmpty(V) Then Max = V(0, 0)
End Function

Public Function Find(Tabla As String, Optional Where As String = "0=0") As Boolean
  On Error Resume Next
  Find = Not IsEmpty(GetRows("SELECT * FROM " & Tabla & " WHERE " & Where, 1))
End Function

Public Sub Insert(Tabla As String, Campos As String, ParamArray Valores() As Variant)
  On Error Resume Next
  Execute "INSERT INTO " & Tabla & " (" & Campos & ") VALUES (" & Join(Valores, ",") & ")"
End Sub

Public Sub Delete(Tabla As String, Where As String)
  On Error Resume Next
  Execute "DELETE FROM " & Tabla & " WHERE " & Where
End Sub

Public Sub Update(Tabla As String, Datos As String, Optional Where As String = "0=0")
  On Error Resume Next
  Execute "UPDATE " & Tabla & " SET " & Datos & " WHERE " & Where
End Sub

Public Function FieldType(Campo As String, Tabla As String) As Long
  If Not DBug Then On Error Resume Next
  Dim R As New ADODB.Recordset
  Cn.Errors.Clear
  If DbType = 1 Then 'SQL SERVER
    R.Open "SELECT TOP 0 " & Campo & " FROM " & Tabla, Cn
  ElseIf DbType = 2 Then 'ORACLE
    R.Open "SELECT " & Campo & " FROM " & Tabla & " WHERE RowNum=0", Cn
  End If
  FieldType = R(0).Type
  R.Close
End Function

Public Function FieldsNames(Source As String) As Variant
  If Not DBug Then On Error Resume Next
  Dim S As String
  Dim I As Integer
  Dim V() As Variant
  Dim R As New ADODB.Recordset
  Cn.Errors.Clear
  
  'BUSCAR ENTRE SELECT Y FROM
  
  R.Open Source, Cn
  ReDim V(R.Fields.Count - 1)
  For I = 0 To R.Fields.Count - 1
    V(I) = R.Fields(I).Name
  Next I
  FieldsNames = V()
  R.Close
End Function

Public Function GetRow(SQL As String, ParamArray F() As Variant) As Variant
  On Error Resume Next
  Cn.Errors.Clear
  Dim I As Integer
  Dim W() As Variant
  Dim V As Variant
  V = GetRows(SQL, 1)
  If IsEmpty(V) Then
    If UBound(F) > 0 Then GetRow = False
  Else
    If UBound(F) > 0 Then
      For I = 0 To UBound(V)
        If I > UBound(F) Then Exit For
        F(I) = V(I, 0)
      Next I
      GetRow = True
    Else
      ReDim W(UBound(V))
      For I = 0 To UBound(V)
        W(I) = V(I, 0)
      Next I
      GetRow = W
    End If
  End If
  DB.Error
End Function

Public Function GetRows(SQL As String, Optional Rows As Long = -1) As Variant
  On Error Resume Next
  GetRows = Execute(SQL).GetRows(Rows)
  DB.Error
End Function

Public Function Execute(SQL As String, Optional recordsAffected, Optional Options As Long = -1) As Recordset
  If Not DBug Then On Error Resume Next
  Dim S As String
  Cn.Errors.Clear
  'DBug = True
  If DBug Then
    S = vbCrLf & "----------------------------------------------------------" & vbCrLf & Now & vbCrLf & SQL
    Debug.Print S
    LogEvent S
    V = Timer
  End If
  
  Set Execute = Cn.Execute(SQL, recordsAffected, Options)
  
  If DBug Then
    Debug.Print CLng((Timer - V) * 1000) & " ms."
  End If
  
  DB.Error
End Function

Public Function Error(Optional Mensaje As String) As Long
  If Cn.Errors.Count > 0 Then
    With Cn.Errors(0)
      If (Mensaje <> "") Or DBug Then
        MsgBox .Description, vbCritical, .Source & "  (" & Mensaje & ":" & CStr(.Number) & ")"
        LogEvent "(" & Now & ") [DB.Error] Error " & CStr(.Number) & " : " & CStr(.Description), 1
      End If
      Error = .Number
    End With
  End If
End Function

Public Function ErrorStr(Optional ByVal ErrorNum As Long) As String
If Not DBug Then On Error Resume Next
  If ErrorNum = 0 Then ErrorNum = Error
  If DbType = 0 Then 'ACCESS
    Select Case ErrorNum:
      Case -2147467259: ErrorStr = "RegDuplicado"
      Case Else: ErrorStr = CStr(ErrorNum)
    End Select
  ElseIf DbType = 1 Then 'SQL SERVER
    Select Case ErrorNum:
      Case -2147217900: ErrorStr = "RegDuplicado"
      Case Else: ErrorStr = CStr(ErrorNum)
    End Select
  ElseIf DbType = 2 Then 'ORACLE
    Select Case ErrorNum:
      Case -2147217900: ErrorStr = "RegDuplicado" 'REVISAR CODIGO EN ORACLE
      Case Else: ErrorStr = CStr(ErrorNum)
    End Select
  End If
End Function

Public Function IFec(fecha As Variant) As String
  If IsZero(fecha) Then
    IFec = "NULL"
    Exit Function
  End If
  If DbType = 0 Then 'ACCESS
    IFec = Format$(fecha, "\#mm\/d\/yyyy hh\:mm\:ss AM/PM\#")
  ElseIf DbType = 1 Then 'SQL SERVER
    IFec = Format$(fecha, "\'dd\/mm\/yyyy hh\:mm\:ss\'")
  ElseIf DbType = 2 Then 'ORACLE
    IFec = "TO_DATE('" & Format$(fecha, "ddmmyyyyhhnnss") & "','DDMMYYYYHH24MISS')"
  End If
End Function

Public Function INum(ByVal Numero As Variant) As String
  If IsNull(Numero) Then
    INum = "NULL"
  Else
    Numero = Replace(CStr(Numero), ",", ".")
    If IsNumeric(Numero) Then INum = Numero Else INum = "0"
  End If
End Function

Public Function IStr(ByVal Cadena As Variant) As String
  If IsNull(Cadena) Then
    IStr = "NULL"
  Else
    Cadena = CStr(Cadena)
    If InStr(Cadena, "'") Then Cadena = Replace(Cadena, "'", "")
    IStr = "'" & Cadena & "'"
  End If
End Function

Public Function IField(Value As Variant, Tipo As Long) As String
   If IsNull(Value) Then
    IField = "NULL"
    Exit Function
  End If
  Select Case Tipo
    Case adBigInt, adCurrency, adDecimal, adDouble, adInteger, _
    adNumeric, adSingle, adSmallInt, adTinyInt, adUnsignedBigInt, _
    adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt:
      IField = INum(Value)
    Case adBSTR, adChar, adLongVarChar, adLongVarWChar, adVarChar, adVarWChar, adWChar:
      IField = IStr(Value)
    Case adDate, adDBDate, adDBTime, adDBTimeStamp:
      IField = IFec(Value)
    Case Else
      IField = CStr(Value)
  End Select
End Function

Public Sub OpenConnection(UDLFile As String)
  If Not DBug Then On Error GoTo OpenConnectionErr
  Dim S As String
  Dim S1 As String
  Dim I As Integer
  
  DbType = 255
  
  If Dir$(UDLFile) = "" Then
    MsgBox "Archivo no encontrado:" & vbCrLf & UDLFile, vbCritical, App.Title
    Exit Sub
  End If
  
  Open UDLFile For Input As #1
  Input #1, S: Input #1, S1: S = S & ";" & S1: Input #1, S1: S = S & ";" & S1
  Close #1
        
  UserID = GetValue("User ID", S)
  ServerName = GetValue("Data Source", S)
  DatabaseName = GetValue("Initial Catalog", S)
  S = GetValue("Provider", S)
  Password = ""
  If TipoCon = 0 Then
    Password = Cfg.Key("INI", "Password")
  ElseIf TipoCon = 1 Then
    Password = Cfg.Key("INI", "PasswordPrev")
  End If
                
  Cn.ConnectionTimeout = 10
  Status = 0
  If InStr(S, "Jet.OLEDB") Then ' ACCESS
    DbType = 0
    If Password = "" Then Password = "Pbjjajlp5h4m1~"
    Cn.CursorLocation = adUseClient
    Cn.Open "File name=" & UDLFile, UserID, Password ', adAsyncConnect
  ElseIf InStr(S, "SQLOLEDB") Then ' SQL SERVER
    DbType = 1
    If Password = "" Then Password = "Pbjjajlp5h4m1"
    Cn.CursorLocation = adUseServer
    Cn.Open "File name=" & UDLFile, UserID, Password ', adAsyncConnect
  ElseIf InStr(S, "MSDAORA") Then ' ORACLE
    DbType = 2
    If Password = "" Then Password = "winplus"
    Cn.CursorLocation = adUseClient
    Cn.Open "File name=" & UDLFile, UserID, Password ', adAsyncConnect
  Else
    err.Raise 1, , "Origen de datos no válido"
  End If
  
  Exit Sub
  
OpenConnectionErr:
    MsgBox "Error " & err.Number & ": " & err.Description, vbCritical, "OpenConnection"
    err.Clear
End Sub

Private Function GetValue(Key As String, ByVal S As String) As String
  Dim I As Integer
  I = InStr(1, S, Key, vbTextCompare)
  If I Then
    S = Mid$(S, I + Len(Key) - 1)
    I = InStr(S, "=")
    If I Then
      S = Mid$(S, I + 1)
      I = InStr(S, ";")
      If I Then S = Left$(S, I - 1)
      GetValue = Trim$(S)
    End If
  End If
End Function

Private Sub Class_Initialize()
  Set CnEv = Cn
End Sub

Private Sub Class_Terminate()
  If Cn.State = adStateOpen Then Cn.Close
End Sub

Private Sub CnEv_ConnectComplete(ByVal err As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
  Status = adStatus
  If Status <> adStatusErrorsOccurred Then
    If DbType = 2 Then Execute "ALTER SESSION SET NLS_DATE_FORMAT='ddddd HH24:MI:SS'"
    Cn.Properties.Refresh
    Cn.CommandTimeout = 0
    Set Errors = Cn.Errors
  End If
  Set CnEv = Nothing
End Sub

