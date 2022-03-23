VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Compare Text
Option Explicit

Public Refresh As Boolean
Public Datos As New CDictionary

Public Event AfterUpdateKey(Id As Variant, Clave As String, Valor As String)

Public Property Let Key(Id As Variant, ByVal Clave As String, Optional Token As Byte, Optional Separator As String = ",", ByVal Valor As Variant)
  Dim P As Integer
  Dim S As String
  Dim S1 As String
  Dim S2 As String
  Dim T() As String
      
  P = GetKeyPos(Id, Clave, S)
     
  If P > 0 Then
    
    S1 = Left$(Datos(Id), P - 1)
    If Len(S1) > 1 Then If Right$(S1, 2) <> vbCrLf Then S1 = S1 & vbCrLf
    P = InStr(P, Datos(Id), vbCrLf, 1)
    If P > 0 Then S2 = Mid$(Datos(Id), P + 2)
    
    If Token > 0 Then
      T() = Split(S, Separator)
      T(Token - 1) = Trim$(Nz(Valor, ""))
      S = Join(T(), Separator)
    Else
      S = Trim$(Nz(Valor, ""))
    End If
      
    Datos(Id) = S1 & Clave & " = " & S & vbCrLf & S2
  
    If Refresh Then RaiseEvent AfterUpdateKey(Id, Clave, S)
        
  End If
        
End Property

Public Property Get Key(Id As Variant, ByVal Clave As String, Optional Token As Byte, Optional Separator As String = ",") As Variant
Attribute Key.VB_UserMemId = 0
  On Error Resume Next
  
  Dim Tipo As String * 1
  Dim S As String
    
  Key = ""
  Tipo = Right$(Clave, 1)
  If InStr("?#%*$", Tipo) Then Clave = Left$(Clave, Len(Clave) - 1)

  GetKeyPos Id, Clave, S
  If Token > 0 Then S = GetToken(Token, S, Separator)
  
  Select Case Tipo
    Case "?": Key = False: If S <> "" Then Key = CBool(S)
    Case "#": Key = 0: If S <> "" Then Key = CLng(S)
    Case "%": Key = 0: If S <> "" Then Key = CDbl(S)
    Case "*": Key = Null: If S <> "" Then Key = XDate(S)
    Case Else: Key = S
  End Select
  
  Exit Property
End Property

Private Function GetKeyPos(Id As Variant, Clave As String, Value As String) As Integer
  Dim P As Integer
  Dim P2 As Integer
  Dim S As String

  Value = ""
  P = InStr(Clave, "\")
  If P Then
    S = "[" & Left$(Clave, P - 1) & "]"
    Clave = Mid$(Clave, P + 1)
    P = InStr(Datos(Id), S)
    If P Then P = P + Len(S) - 1 Else Exit Function
    P2 = InStr(P, Datos(Id), "[")
  End If

  P = InStr(P + 1, Datos(Id), Clave & " =")
  If P = 0 Then Exit Function
  
  If P2 > 0 Then If P2 < P Then Exit Function

  GetKeyPos = P
  S = Mid$(Datos(Id), P + Len(Clave) + 2)
  P = InStr(S, vbCrLf): If P Then S = Left$(S, P - 1)
  Value = Trim$(S)
End Function

Private Function CodeINI(Cadena As String, Patron As String) As String
  On Error Resume Next
  Dim Linea As String
  Dim LineaCod As String
  Dim I As Integer
  Dim Lng As Integer
  Dim Pos As Integer
  Dim NCar As Integer
  Dim C1 As Byte
  Dim C2 As Byte
  Lng = Len(Cadena)
  If (Lng <= 15) Or (Len(Patron) < 15) Then
    CodeINI = Cadena
    Exit Function
  End If
  Pos = 1
  Do Until Pos > Lng
    If (Pos + 14) <= Lng Then
      NCar = 15
    Else
      NCar = Lng - Pos + 1
    End If
    Linea = Mid$(Cadena, Pos, NCar)
    For I = 1 To NCar
      C1 = Asc(Mid$(Patron, I, 1))
      C2 = Asc(Mid$(Linea, I, 1))
      LineaCod = LineaCod & Chr$(C1 Eqv C2)
    Next I
    Pos = Pos + 15
  Loop
  CodeINI = LineaCod
  Exit Function
End Function

Private Function CalcularCRC(Cadena As String) As Byte
  On Error Resume Next
  Dim I As Integer
  Dim CRC As Byte
  For I = 1 To Len(Cadena)
    CRC = CRC Xor Asc(Mid$(Cadena, I, 1))
  Next I
  CalcularCRC = CRC
  Exit Function
End Function

Public Function Code(ByVal S As String) As String
  Dim Patron As String
  S = Replace(Trim$(S), vbCrLf, "ÿ")
  Patron = Left$(S, InStr(S, "ÿ") - 1)
  S = Mid$(S, Len(Patron) + 2)
  If Right$(S, 1) <> "ÿ" Then S = S & "ÿ"
  S = CodeINI(S, Patron)
  S = Chr$(Len(Patron)) & Patron & S
  Code = Chr$(CalcularCRC(S)) & S
End Function

Public Function Decode(ByVal S As String) As String
  Dim CRC As Byte
  Dim Lng As Integer
  Dim Patron As String
  CRC = Asc(Left$(S, 1))
  Lng = Asc(Mid$(S, 2, 1))
  Patron = Mid$(S, 3, Lng)
  S = Mid$(S, 2)
  If CalcularCRC(S) <> CRC Then Exit Function
  S = CodeINI(Mid$(S, Lng + 2), Patron)
  Decode = Replace(Patron & "ÿ" & S, "ÿ", vbCrLf)
End Function

Private Sub Class_Initialize()
  Refresh = True
End Sub

