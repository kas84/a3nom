VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private MKey() As Variant
Private MDat() As Variant
Private Cont As Integer

Public Function Count() As Integer
  Count = Cont
End Function

Public Sub Remove(Id As Variant)
  Dim i As Integer
  If Cont <= 1 Then
    RemoveAll
    Exit Sub
  End If
  i = IIndex(Id): If i < 1 Then Exit Sub
  Cont = Cont - 1
  For i = i To Cont
    MKey(i) = MKey(i + 1)
    If IsObject(MDat(i + 1)) Then
      Set MDat(i) = MDat(i + 1)
    Else
      MDat(i) = MDat(i + 1)
    End If
  Next i
  ReDim Preserve MKey(Cont)
  ReDim Preserve MDat(Cont)
End Sub

Public Function Keys() As Variant
  Keys = MKey()
End Function

Public Function Items() As Variant
  Items = MDat()
End Function

Public Function Exists(Id As Variant) As Boolean
  Exists = (IIndex(Id) > 0)
End Function

Public Sub Add(Id As Variant, Data As Variant)
  Cont = Cont + 1
  ReDim Preserve MKey(Cont)
  ReDim Preserve MDat(Cont)
  MKey(Cont) = Id
  If IsObject(Data) Then
    Set MDat(Cont) = Data
  Else
    MDat(Cont) = Data
  End If
End Sub

Public Sub RemoveAll()
  Cont = 0
  Erase MKey
  Erase MDat
End Sub

Public Property Let Item(Id As Variant, Data As Variant)
  Dim i As Integer
  i = IIndex(Id)
  If i < 1 Then Add Id, Data Else MDat(IIndex(Id)) = Data
End Property

Public Property Set Item(Id As Variant, Data As Variant)
  Dim i As Integer
  i = IIndex(Id)
  If i < 1 Then Add Id, Data Else Set MDat(IIndex(Id)) = Data
End Property

Public Property Get Item(Id As Variant) As Variant
  Dim i As Integer
  Item = ""
  i = IIndex(Id)
  If i < 1 Then Exit Property
  If IsObject(MDat(i)) Then
    Set Item = MDat(i)
  Else
    Item = MDat(i)
  End If
End Property

Private Function IIndex(Id As Variant) As Integer
  On Error Resume Next
  Dim i As Integer
  For i = 1 To UBound(MKey)
    If MKey(i) = Id Then
      IIndex = i
      Exit Function
    End If
  Next i
  Exit Function
End Function


