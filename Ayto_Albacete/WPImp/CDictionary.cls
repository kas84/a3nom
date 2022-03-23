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
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Base 1

Private MKey() As Variant
Private MDat() As Variant
Private Cont As Integer

Public Function Count() As Integer
  Count = Cont
End Function

Public Sub Remove(Id As Variant)
  Dim I As Integer
  If Cont <= 1 Then
    RemoveAll
    Exit Sub
  End If
  I = IIndex(Id): If I < 1 Then Exit Sub
  Cont = Cont - 1
  For I = I To Cont
    MKey(I) = MKey(I + 1)
    If IsObject(MDat(I + 1)) Then
      Set MDat(I) = MDat(I + 1)
    Else
      MDat(I) = MDat(I + 1)
    End If
  Next I
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
  Dim I As Integer
  I = IIndex(Id)
  If I < 1 Then Add Id, Data Else MDat(IIndex(Id)) = Data
End Property

Public Property Set Item(Id As Variant, Data As Variant)
  Dim I As Integer
  I = IIndex(Id)
  If I < 1 Then Add Id, Data Else Set MDat(IIndex(Id)) = Data
End Property

Public Property Get Item(Id As Variant) As Variant
Attribute Item.VB_UserMemId = 0
  Dim I As Integer
  Item = ""
  I = IIndex(Id)
  If I < 1 Then Exit Property
  If IsObject(MDat(I)) Then
    Set Item = MDat(I)
  Else
    Item = MDat(I)
  End If
End Property

Private Function IIndex(Id As Variant) As Integer
  On Error Resume Next
  Dim I As Integer
  For I = 1 To UBound(MKey)
    If MKey(I) = Id Then
      IIndex = I
      Exit Function
    End If
  Next I
  Exit Function
End Function

