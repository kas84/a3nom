Attribute VB_Name = "Global"
Option Explicit

Public bPC As String         ' Identificador del bloque de la configuración del PC
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public PCName As String
Public UserName As String
Public bUSR As String        ' Identificador del bloque de la configuración del Usuario
Public DBug As Boolean
Public ExportArchivo As String
Public bCancelar As Boolean
Public sTipoExportacion As String


Public Function Nz(Value As Variant, Default As Variant) As Variant
  On Error Resume Next
  Nz = Default
  If IsEmpty(Value) Then Exit Function
  If IsNull(Value) Then Exit Function
  Nz = Value
  Err.Clear
End Function

Public Function IsZero(Valor As Variant) As Boolean
  On Error Resume Next
  IsZero = True
  If IsEmpty(Valor) Then Exit Function
  If IsNull(Valor) Then Exit Function
  If IsDate(Valor) Then
    If Valor = 0 Then Exit Function
  End If
  If IsNumeric(Valor) Then
    If Valor = 0 Then Exit Function
  Else
    If Trim$(Valor) = "" Then Exit Function
  End If
  IsZero = False
  Err.Clear
End Function

Public Function MinToHor(ByVal M As Variant, Optional Max24 As Boolean) As String
  On Error Resume Next
  Dim S As String
  If Not IsNumeric(M) Then M = HorToMin(M, Max24)
  If M < 0 Then
    S = "-": M = Abs(M)
    If Max24 Then M = 1440 - M
  ElseIf M >= 1440 And Max24 Then
    S = "+": M = M - 1440
  End If
  MinToHor = S & CStr(M \ 60) & ":" & Format$(M Mod 60, "00")
  Err.Clear
End Function

Public Function HorToMin(ByVal S As String, Optional Max24 As Boolean) As Long
  On Error Resume Next
  Dim i As Long
  S = Trim$(S)
  i = InStrChr(S, ":., ")
  If i Then
    i = Abs(Val(Left$(S, i - 1))) * 60 + Val(Mid$(S, i + 1, 2))
    Select Case Left$(S, 1)
      Case "-":
        If Max24 Then i = -1440 + i Else i = -i
      Case "+":
        If Max24 Then i = 1440 + i
    End Select
    HorToMin = i
  End If
End Function

Public Function InStrChr(cadena As String, Cars As String) As Integer
  Dim i As Integer
  Dim P As Integer
  For i = 1 To Len(Cars)
    P = InStr(cadena, Mid$(Cars, i, 1))
    If P Then Exit For
  Next i
  InStrChr = P
End Function

Public Function IntSplit(cadena As String, Delimiter As String) As Variant
  On Error GoTo ErrIntSplit
  Dim S() As String
  Dim i() As Integer
  Dim N As Integer
  S() = Split(cadena, Delimiter)
  If UBound(S) > -1 Then
    ReDim i(UBound(S))
    For N = 0 To UBound(S)
      i(N) = Val(S(N))
    Next N
    IntSplit = i()
  End If
  Exit Function
ErrIntSplit:
  ReDim i(0)
  IntSplit = i()
  Exit Function
End Function

Public Function AUBound(SourceArray As Variant, Optional Dimension As Long = 1) As Long
  On Error Resume Next
  Dim N As Long
  N = UBound(SourceArray, Dimension)
  If Err = 0 Then AUBound = N Else AUBound = -1
  Err.Clear
End Function

Public Function SumArray(SourceArray As Variant) As Long
  On Error GoTo ErrSumArray
  Dim N As Integer
  Dim L As Long

  For N = LBound(SourceArray) To UBound(SourceArray)
    L = L + SourceArray(N)
  Next N

  SumArray = L
  Exit Function
ErrSumArray:
  SumArray = 0
  Exit Function
End Function

Public Function FDelete(File As String, Optional Forzar As Boolean) As Boolean
  On Error Resume Next
  Kill File
  FDelete = (Err = 0)
  Exit Function
End Function

Public Function FCopy(Origen As String, Destino As String, Optional Machacar As Boolean) As Boolean
  On Error Resume Next
  If Machacar Then Kill Destino
  FileCopy Origen, Destino
  FCopy = (Err = 0)
  Exit Function
End Function

Public Function GetCodigo(S As String) As String
  Dim P As Integer
  P = InStr(S, " - ")
  If P < 1 Then P = InStr(S, " -")
  If P > 0 Then GetCodigo = Trim$(Left$(S, P - 1)) Else GetCodigo = S
End Function
Public Sub SendKey(KeyCode As Byte, Optional Shift As Boolean)
  On Error Resume Next
  If Shift Then keybd_event &H10, 0, 0, 0
  keybd_event KeyCode, 0, 0, 0
  keybd_event KeyCode, 0, &H2, 0
  If Shift Then keybd_event &H10, 0, &H2, 0
  Err.Clear
End Sub
Public Function FLoad(File As String) As String
  On Error Resume Next
  Dim S As String
  If Not FExist(File) Then Exit Function
  S = String$(FileLen(File), " ")
  Open File For Binary Access Read Shared As #1
  Get #1, , S
  Close #1
  FLoad = S
  Exit Function
End Function
Public Function GetToken(ByVal Pos As Byte, Cad As String, Optional Sep As String = ";") As String
  On Error Resume Next
  Dim S() As String
  If Pos < 1 Then Exit Function
  S() = Split(Cad, Sep)
  If (Pos - 1) <= UBound(S) Then GetToken = S(Pos - 1)
  Err.Clear
End Function

Public Function XDate(Fecha As String) As Date
  On Error Resume Next
  Dim i As Byte
  Dim S As String
  Dim F(0 To 2) As String
  Dim P As Byte
  For i = 1 To Len(Fecha)
    S = Mid$(Fecha, i, 1)
    If IsNumeric(S) Then
      F(P) = F(P) & S
    Else
      P = P + 1
      If P > 2 Then Exit For
    End If
  Next i
  XDate = DateSerial(F(2), F(1), F(0))
  Err.Clear
End Function


