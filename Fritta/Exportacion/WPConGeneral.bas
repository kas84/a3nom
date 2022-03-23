Attribute VB_Name = "WPConGeneral"
Option Explicit

Public Const BackColorFocus = &HEFC6AD      '&HDFFFFF
Public Const BackColorDisabled = &HE0E0E0

Public MDIHwnd As Long

Public Const AutoDraw = True

Public Declare Function SetCapture Lib "user32" (ByVal Hwnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetCapture Lib "user32" () As Long

Public Enum EDataType
  DtGeneral = 0     ' General
  DtNumEntero = 1   ' 10000
  DtNumDecimal = 2  ' 10,34             / Devuelve 10.34
  DtMoneda = 3      ' 10.460,32         / Devuelve 10460.32
  DtFecha = 4       ' 01-12-2000
  DtHora = 5        ' 08:30
  DtFechaHora = 6   ' 01-12-2000 08:30
  DtPassword = 7    ' ********          / Devuelve la Password introducida
  DtMayusculas = 8  ' MAYUSCULAS
  DtColor = 9       '                   / Devuelve el nº de color
  DtMinutos1 = 10   ' 01:30             / Devuelve 90
  DtMinutos2 = 11   ' 1,30              / Devuelve 90
  DtMinutos3 = 14   ' +01:30            / Devuelve 1530
  DtBooleano = 12   ' Si / No           / Devuelve True / False
  DtNumTexto = 13   ' 56 - Ejemplo      / Devuelve 56
  DtStrTexto = 15   ' EJ - Ejemplo      / Devuelve EJ
  DtDescripcion = 16 ' Usado para comboboxes
End Enum

Public Function IsDataType(ByVal Datatype As EDataType, ByVal Value As Variant) As String
  Dim L As Long

  If Trim$(Nz(Value, "")) = "" Then Exit Function

  Select Case Datatype
    Case 4:
            If Not IsDate(Value) Then
              IsDataType = "La Fecha introducida no es correcta"
            Else
              If (Year(Value) < 1900) Or (Year(Value) > 2100) Then
                IsDataType = "La Fecha introducida no es correcta"
              End If
            End If
    Case 5: If (Left$(Value, 1) = "+") Or (Left$(Value, 1) = "-") Then Value = Mid$(Value, 2)
       If Not IsDate(Value) Then
         IsDataType = "La Hora introducida no es correcta"
       End If

    Case 10: If (Left$(Value, 1) = "+") Or (Left$(Value, 1) = "-") Then Value = Mid$(Value, 2)

       L = HorToMin(Value)
       If Value <> MinToHor(L) Then
         IsDataType = "La Hora introducida no es correcta"
       End If

    Case 6: If Not IsDate(Value) Then IsDataType = "La Fecha/Hora introducida no es correcta"

  End Select

End Function

Public Function TextoKeyPress(Texto As String, KeyAscii As Integer, Datatype As EDataType) As String
  Dim S As String
  Dim C As String

  If KeyAscii = 13 Or KeyAscii = 27 Then KeyAscii = 0
  
  TextoKeyPress = Texto

  If KeyAscii > 31 Then
    C = Chr$(KeyAscii)

    Select Case Datatype
      Case 1: If InStr("0123456789-+", C) = 0 Then C = vbNullChar
      Case 2, 3:
        If C = "." Then C = ","
        If InStr("0123456789-+,", C) = 0 Then C = vbNullChar
      Case 4:
        If InStr("Hh", C) Then
          TextoKeyPress = Format$(Now, "ddddd")
          KeyAscii = 0: Exit Function
        End If
        If InStr(".,-/\ ", C) Then C = "/"
        If InStr("0123456789/", C) = 0 Then C = vbNullChar
      Case 5, 10:
        If InStr("Hh", C) Then
          TextoKeyPress = Format$(Now - (Fix(Now)), "hh:nn")
          KeyAscii = 0: Exit Function
        End If
        If InStr("., ", C) Then C = ":"
        If InStr("0123456789:-+", C) = 0 Then C = vbNullChar
      Case 6:
        If InStr("Hh", C) Then
          TextoKeyPress = Format$(Now, "ddddd hh:nn")
          KeyAscii = 0: Exit Function
        End If
        If InStr("0123456789.,-:/\ ", C) = 0 Then C = vbNullChar
      Case 8: C = UCase$(C)
      Case 14:
        If InStr("Hh", C) Then
          TextoKeyPress = Format$(Now - (Fix(Now)), "hh:nn")
          KeyAscii = 0: Exit Function
        End If
        If InStr("., ", C) Then C = ":"
        If InStr("0123456789:-+", C) = 0 Then C = vbNullChar
    End Select

    KeyAscii = Asc(C)
        
    If KeyAscii = 0 Then Beep
    
  End If

End Function

Public Function CValue(ByVal Valor As String, ByVal tipo As EDataType) As Variant
  On Error Resume Next
    
  Select Case tipo
    Case DtFecha, DtHora, DtFechaHora: CValue = Null
    Case DtNumEntero, DtNumDecimal, DtMoneda, DtMinutos1, DtMinutos2, DtMinutos3, DtColor: CValue = 0
    Case DtBooleano: CValue = False
    Case Else: CValue = ""
  End Select
  
  If Trim$(Valor) = "" Then Exit Function
  
  Select Case tipo
    Case DtNumEntero, DtColor: CValue = CLng(Val(Valor))
    Case DtNumDecimal: CValue = CSng(Val(Replace(CStr(Valor), ",", ".")))
    Case DtMoneda: CValue = CCur(Valor)
    Case DtFecha: CValue = Fix(CDate(Valor))
    Case DtHora: CValue = CDate(Valor) - Fix(CDate(Valor))
    Case DtFechaHora: CValue = CDate(Valor)
    Case DtPassword: CValue = Valor
    Case DtMayusculas: CValue = Trim$(UCase$(Valor))
    Case DtColor: CValue = CLng(Val(Valor))
    Case DtMinutos1: CValue = HorToMin(Valor)
    Case DtMinutos2: CValue = Val(Valor) * 60
    Case DtMinutos3: CValue = HorToMin(Valor, True)
    Case DtBooleano: CValue = CBool(Valor)
    Case DtNumTexto: CValue = CLng(Val(Valor))
    Case DtStrTexto: CValue = GetCodigo(Valor)
    Case Else: CValue = Trim$(CStr(Valor))
  End Select
  Err.Clear
End Function

Public Function CText(ByVal Valor As Variant, ByVal tipo As EDataType) As String
  On Error Resume Next
  If IsNull(Valor) Or IsEmpty(Valor) Then Exit Function
  Select Case tipo
    Case DtNumEntero, DtNumDecimal, DtColor, DtNumTexto, DtStrTexto: CText = Trim$(CStr(Valor))
    Case DtMoneda: CText = Format$(Valor, "#,###.00")
    Case DtFecha: CText = Format$(Valor, "ddddd")
    Case DtHora: CText = Format$(Valor, "hh:nn")
    Case DtFechaHora: CText = Format$(Valor, "ddddd hh:nn:ss") 'REM
    Case DtPassword: CText = Valor
    Case DtMayusculas: CText = UCase$(Trim$(CStr(Valor)))
    Case DtMinutos1: CText = MinToHor(Valor)
    Case DtMinutos2: CText = CStr(Valor / 60)
    Case DtMinutos3: CText = MinToHor(Valor, True)
    Case DtBooleano: If Valor Then CText = "Si" Else CText = "No"
    Case Else: CText = Trim$(CStr(Valor))
  End Select
  Err.Clear
End Function
Public Sub LogEvent(Texto As String, Optional EventType As Integer = 4)
  If DBug Then
    App.LogEvent Texto, EventType
    DPrint Texto
  End If
End Sub
Public Sub DPrint(ParamArray Texto() As Variant)
  If Not DBug Then Exit Sub
  Dim S As String
  Dim i As Integer
  For i = 0 To UBound(Texto)
    S = S & CStr(Texto(i)) & vbTab
  Next i
  'FDebug.FDebugPrint S
End Sub


