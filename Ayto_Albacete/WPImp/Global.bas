Attribute VB_Name = "Global"
Option Explicit

Public DBug As Boolean
Public PCName As String
Public UserName As String
Public UserLevel As Byte
Public WindowsDir As String
Public SystemDir As String
Public bINI As String        ' Identificador del bloque del fichero INI
Public bSYS As String        ' Identificador del bloque de la configuración del Sistema
Public bPC As String         ' Identificador del bloque de la configuración del PC
Public bUSR As String        ' Identificador del bloque de la configuración del Usuario

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwAccess As Long, ByVal fInherit As Integer, ByVal hObject As Long) As Long
Private Declare Function GetWindowsDirectoryA Lib "kernel32" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectoryA Lib "kernel32" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, LParam As Long) As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Public Function ToStr(ByVal S As String) As String
  On Error Resume Next
  ToStr = Chr$(34) & CStr(S) & Chr$(34)
  err.Clear
End Function

Public Function Nz(Value As Variant, Default As Variant) As Variant
  On Error Resume Next
  Nz = Default
  If IsEmpty(Value) Then Exit Function
  If IsNull(Value) Then Exit Function
  Nz = Value
  err.Clear
End Function

Public Function CStrB(Cadena As String) As String
  On Error Resume Next
  Dim P As Integer
  P = InStr(Cadena, vbNullChar)
  If P > 0 Then CStrB = RTrim$(Left$(Cadena, P - 1)) Else CStrB = RTrim$(Cadena)
  err.Clear
End Function

Public Function Zz(Value As Variant, Default As Variant) As Variant
  Zz = Default
  If IsEmpty(Value) Then Exit Function
  If IsNull(Value) Then Exit Function
  If IsNumeric(Value) Then
    If Value = 0 Then Exit Function
  Else
    If Trim$(Value) = "" Then Exit Function
  End If
  Zz = Value
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
  err.Clear
End Function

Public Function ExecCmd(cmdline As String, Optional WindowMode As Integer = vbNormalFocus) As Boolean
  On Error GoTo ErrExecCmd
  
  Dim ret As Long
  Dim ProcessHandle As Long
  Dim ProcessID As Long

  Screen.MousePointer = 11
  ProcessID = Shell(cmdline, WindowMode)
  ProcessHandle = OpenProcess(1048576, True, ProcessID)

  DoEvents
  ret = WaitForSingleObject(ProcessHandle, -1)

  Screen.MousePointer = 0
  ExecCmd = True
  Exit Function
ErrExecCmd:
  Screen.MousePointer = 0
  MsgBox "Se ha producido un error al ejecutar " & cmdline, vbCritical, "Error " & CStr(err.Number) & " - " & err.Description
  err = 0
  Exit Function
End Function

Public Function StrMes(Mes As Byte, Optional Abreviar As Boolean = False) As String
  Dim S As String
  S = MonthName(((Mes - 1) Mod 12) + 1, Abreviar)
  StrMes = UCase$(Left$(S, 1)) & Mid$(S, 2)
End Function

Public Function StrDia(Dia As Byte, Optional Abreviar As Boolean = False) As String
  Dim S As String
  S = WeekdayName(((Dia - 1) Mod 7) + 1, Abreviar)
  StrDia = UCase$(Left$(S, 1)) & Mid$(S, 2)
End Function

Public Function DiasMes(Mes As Byte, Optional Anyo As Integer = 0) As Byte
  Dim A As Integer
  Select Case Mes
    Case 2: If Anyo = 0 Then A = Year(Now) Else A = Anyo
            If A Mod 4 = 0 Then DiasMes = 29 Else DiasMes = 28
    Case 4, 6, 9, 11: DiasMes = 30
    Case Else: DiasMes = 31
  End Select
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
  err.Clear
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

Public Function IsHora(S As String) As Boolean
  Dim i As Integer
  For i = 1 To Len(S)
    If InStr("-+0123456789:., ", Mid$(S, i, 1)) < 1 Then Exit Function
  Next i
  IsHora = True
End Function

Public Function InStrChr(Cadena As String, Cars As String) As Integer
  Dim i As Integer
  Dim P As Integer
  For i = 1 To Len(Cars)
    P = InStr(Cadena, Mid$(Cars, i, 1))
    If P Then Exit For
  Next i
  InStrChr = P
End Function

Public Function InStrChrRev(Cadena As String, Cars As String) As Integer
  On Error Resume Next
  Dim i As Integer
  Dim P As Integer
  For i = 1 To Len(Cars)
    P = InStrRev(Cadena, Mid$(Cars, i, 1))
    If P Then Exit For
  Next i
  InStrChrRev = P
  err.Clear
End Function

Public Function IntSplit(Cadena As String, Delimiter As String) As Variant
  On Error GoTo ErrIntSplit
  Dim S() As String
  Dim i() As Integer
  Dim N As Integer
  S() = Split(Cadena, Delimiter)
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

Public Function IntJoin(SourceArray As Variant, Delimiter As String) As String
On Error GoTo ErrIntJoin
  Dim S As String
  Dim N As Integer
  For N = LBound(SourceArray) To UBound(SourceArray)
    S = S & CStr(SourceArray(N)) & Delimiter
  Next N
  IntJoin = Left$(S, Len(S) - 1)
  Exit Function
ErrIntJoin:
  IntJoin = ""
  Exit Function
End Function

Public Function AFind(Matriz As Variant, Valor As Variant, BuscaCol As Byte, Optional DevuelveCol As Integer = -1) As Variant
  On Error Resume Next
  Dim i As Integer
  If DevuelveCol > -1 Then AFind = "" Else AFind = -1
  For i = 0 To AUBound(Matriz, 2)
    If Matriz(BuscaCol, i) = Valor Then
      If DevuelveCol > -1 Then
        AFind = Matriz(DevuelveCol, i)
      Else
        AFind = i
      End If
      Exit Function
    End If
  Next i
End Function

Public Function AUBound(SourceArray As Variant, Optional Dimension As Long = 1) As Long
  On Error Resume Next
  Dim N As Long
  N = UBound(SourceArray, Dimension)
  If err = 0 Then AUBound = N Else AUBound = -1
  err.Clear
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

Public Sub FillArray(SourceArray As Variant, Value As Variant)
  Dim N As Integer
  For N = LBound(SourceArray) To UBound(SourceArray)
    SourceArray(N) = Value
  Next N
End Sub

Public Function Random(Optional Min As Integer, Optional Max As Integer) As Integer
  Random = Int((Max - Min + 1) * Rnd + Min)
End Function

Public Sub Limita(Valor As Variant, Min As Variant, Max As Variant)
  If Valor < Min Then Valor = Min
  If Valor > Max Then Valor = Max
End Sub

Public Function ValStr(S As String) As Long
  On Error Resume Next
  Dim i As Integer
  Dim N As String
  Dim C As String
  For i = 1 To Len(S)
    C = Mid$(S, i, 1)
    If IsNumeric(C) Then N = N & C
  Next i
  ValStr = CLng(Val(N))
  err.Clear
End Function

Public Function LeftNot(Cadena As String, Largo As Byte) As String
  On Error Resume Next
  LeftNot = Left$(Cadena, Len(Cadena) - Largo)
  err.Clear
End Function

Public Function LetraNIF(ByVal DNI As Long) As String
  LetraNIF = Mid$("TRWAGMYFPDXBNJZSQVHLCKET", DNI Mod 23 + 1, 1)
End Function

Public Function FormatNIF(DNI As String) As String
  Dim C As String
  Dim S As String
  Dim L As Long
  C = UCase$(Left$(DNI, 1))
  L = ValStr(DNI)
  S = DNI
  If L > 0 Then
    If IsNumeric(C) Then
      S = Format$(L, "00,000,000") & "-" & LetraNIF(L)
    Else
      If (C >= "A") And (C <= "Z") Then S = UCase$(C) & "-" & Format$(L, "00,000,000")
    End If
  End If
  If Len(S) > 12 Then FormatNIF = Left$(DNI, 12) Else FormatNIF = S
End Function

'----------------------------------------------------------------------------
' FUNCIONES DE FICHEROS

Public Function FName(File As String) As String
  Dim P As Integer
  P = InStrChrRev(File, "\/")
  If P Then FName = Mid$(File, P + 1)
End Function

Public Function FPath(File As String) As String
  Dim P As Integer
  P = InStrChrRev(File, "\/")
  If P Then FPath = Left$(File, P)
End Function

Public Function FExist(File As String) As Boolean
  On Error Resume Next
  Dim S As String
  If File = "" Then Exit Function
  S = Dir(File)
  FExist = (err = 0) And (S <> "")
  Exit Function
End Function

Public Function FDelete(File As String, Optional Forzar As Boolean) As Boolean
  On Error Resume Next
  Kill File
  FDelete = (err = 0)
  Exit Function
End Function

Public Function FCopy(Origen As String, Destino As String, Optional Machacar As Boolean) As Boolean
  On Error Resume Next
  If Machacar Then Kill Destino
  FileCopy Origen, Destino
  FCopy = (err = 0)
  Exit Function
End Function

Public Function FMove(Origen As String, Destino As String) As Boolean
  On Error Resume Next
  Kill Destino
  FileCopy Origen, Destino
  If err = 0 Then
    Kill Origen
    FMove = (err = 0)
  End If
  Exit Function
End Function

Public Function FParent(ByVal Cadena As String, Optional Nivel As Byte = 0, Optional Separator As String = "\") As String
  Dim P As Integer
  Dim i As Integer
  For i = 0 To Nivel
    P = InStrChrRev(Cadena, Separator)
    If P = 0 Then Exit For Else Cadena = Left$(Cadena, P - 1)
  Next i
  FParent = Cadena & "\"
End Function

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

Public Sub FSave(File As String, Data As String)
  On Error Resume Next
  If FExist(File) Then Kill File
  Open File For Binary As #1
  Put #1, , Data
  Close #1
  Exit Sub
End Sub

'----------------------------------------------------------------------------

Public Function Bit(NBit As Byte, ByVal Valor As Long) As Boolean
  Dim Aux As Long
  Aux = 2 ^ NBit
  Bit = ((Valor And Aux) = Aux)
End Function

Public Function Round(Num As Variant, Optional Decimales As Byte = 2) As Variant
  Round = Format(Nz(Num, 0), "0." & String$(Decimales, "0"))
End Function

Public Sub Wait(ByVal Segundos As Single, Optional DoEvent As Boolean = True)
  Segundos = Segundos + Timer
  Do Until Segundos < Timer
    If DoEvent Then DoEvents
  Loop
End Sub

Public Function Succ(N As String) As String
  Dim S As String
  Dim i As Integer
  For i = Len(N) To 1 Step -1
    If Not IsNumeric(Mid$(N, i, 1)) Then Exit For
  Next i
  S = Format$(CStr(Val(Mid$(N, i + 1)) + 1), String$(Len(N) - i, "0"))
  Succ = Mid$(N, 1, i) & S
End Function

Public Function RVal(N As String) As Double
  Dim i As Integer
  Dim Aux As Double
  For i = Len(N) To 1 Step -1
    If Not IsNumeric(Mid$(N, i, 1)) Then Exit For
  Next i
  RVal = Val(Mid$(N, i + 1))
End Function

Public Sub SendMsg(ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal LParam As Long)
  Dim Tmp
  Tmp = SendMessage(Hwnd, wMsg, wParam, LParam)
End Sub

Public Sub AppKill()
  ExitProcess 0
End Sub

Public Function GetCodigo(S As String) As String
  Dim P As Integer
  P = InStr(S, " - ")
  If P < 1 Then P = InStr(S, " -")
  If P > 0 Then GetCodigo = Trim$(Left$(S, P - 1)) Else GetCodigo = S
End Function

Public Function GetDescripcion(S As String) As String
  Dim P As Integer
  GetDescripcion = S
  P = InStr(S, " - ")
  If P < 1 Then
    P = InStr(S, " -")
    If P > 0 Then GetDescripcion = Trim$(Mid$(S, P + 2))
  Else
    GetDescripcion = Trim$(Mid$(S, P + 3))
  End If
End Function

Public Function GetComputerName() As String
  On Error Resume Next
  Dim L As Long
  Dim Size As Long
  Dim Buf As String * 255
  Size = 255
  L = GetComputerNameA(Buf, Size)
  GetComputerName = UCase$(Left$(Buf, Size))
  err.Clear
End Function

Public Function GetUserName() As String
  On Error Resume Next
  Dim L As Long
  Dim Buf As String * 255
  L = GetUserNameA(Buf, 25)
  GetUserName = Left(Buf, InStr(Buf, Chr(0)) - 1)
  err.Clear
End Function

Public Function GetWindowsDirectory() As String
  On Error Resume Next
  Dim L As Long
  Dim Buf As String * 255
  L = GetWindowsDirectoryA(Buf, 255)
  GetWindowsDirectory = Left$(Buf, L)
  err.Clear
End Function

Public Function GetSystemDirectory() As String
  On Error Resume Next
  Dim L As Long
  Dim Buf As String * 255
  L = GetSystemDirectoryA(Buf, 255)
  GetSystemDirectory = Left$(Buf, L)
  err.Clear
End Function

Public Sub InitGlobal()
  PCName = GetComputerName
  WindowsDir = GetWindowsDirectory
  SystemDir = GetSystemDirectory
  UserName = GetUserName
  bINI = "INI"
  bSYS = "SYSTEM"
  bPC = "PC." & PCName
End Sub

Public Sub SendKey(KeyCode As Byte, Optional Shift As Boolean)
  On Error Resume Next
  If Shift Then keybd_event &H10, 0, 0, 0
  keybd_event KeyCode, 0, 0, 0
  keybd_event KeyCode, 0, &H2, 0
  If Shift Then keybd_event &H10, 0, &H2, 0
  err.Clear
End Sub

Public Function GetToken(ByVal Pos As Byte, Cad As String, Optional Sep As String = ";") As String
  On Error Resume Next
  Dim S() As String
  If Pos < 1 Then Exit Function
  S() = Split(Cad, Sep)
  If (Pos - 1) <= UBound(S) Then GetToken = S(Pos - 1)
  err.Clear
End Function

Public Function XDate(fecha As String) As Date
  On Error Resume Next
  Dim i As Byte
  Dim S As String
  Dim F(0 To 2) As String
  Dim P As Byte
  For i = 1 To Len(fecha)
    S = Mid$(fecha, i, 1)
    If IsNumeric(S) Then
      F(P) = F(P) & S
    Else
      P = P + 1
      If P > 2 Then Exit For
    End If
  Next i
  XDate = DateSerial(F(2), F(1), F(0))
  err.Clear
End Function


