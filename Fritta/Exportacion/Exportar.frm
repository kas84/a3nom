VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportación Datos Nomina"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6285
   Icon            =   "Exportar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   6285
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtdestino 
      Height          =   375
      Left            =   1800
      TabIndex        =   21
      Top             =   4200
      Width           =   3975
   End
   Begin VB.TextBox txtcomplemento 
      Height          =   375
      Left            =   1800
      TabIndex        =   20
      Top             =   3600
      Width           =   4455
   End
   Begin Proyecto1.UButton cmdbuscar 
      Height          =   375
      Left            =   5880
      TabIndex        =   7
      Top             =   4200
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BackColor       =   14737632
      Text            =   "..."
      AlignPicture    =   4
      AlignText       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog cmdialog 
      Left            =   4320
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Fichero texto|*.txt"
   End
   Begin MSComctlLib.ProgressBar barra 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   5280
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker dtpfechainicio 
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   1200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   67371009
      CurrentDate     =   42852
   End
   Begin Proyecto1.UComboBox cbopersonal 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   1800
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      BackColorFocus  =   14680063
      BackColorDisabled=   0
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnWidths    =   "228"
      ListRows        =   8
      ListWidth       =   313
   End
   Begin Proyecto1.UComboBox cbodepartamento 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   2400
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      BackColorFocus  =   14680063
      BackColorDisabled=   0
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnWidths    =   "228"
      ListRows        =   8
      ListWidth       =   313
   End
   Begin Proyecto1.UComboBox cboseccion 
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   3000
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      BackColorFocus  =   14680063
      BackColorDisabled=   0
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnWidths    =   "228"
      ListRows        =   8
      ListWidth       =   313
   End
   Begin MSComCtl2.DTPicker dtpfechafin 
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   1200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   67371009
      CurrentDate     =   42852
   End
   Begin Proyecto1.UButton cmdaceptar 
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   4800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      BackColor       =   14737632
      ForeColor       =   16711680
      Text            =   "Aceptar "
      AlignPicture    =   4
      AlignText       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Proyecto1.UButton cmdcancelar 
      Height          =   255
      Left            =   4920
      TabIndex        =   9
      Top             =   4800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      BackColor       =   14737632
      ForeColor       =   16711680
      Text            =   "Cancelar"
      AlignPicture    =   4
      AlignText       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC6AD&
      BackStyle       =   0  'Transparent
      Caption         =   "Limite Percepcion Complementaria :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   120
      TabIndex        =   19
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC6AD&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Fin :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3480
      TabIndex        =   18
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC6AD&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Inicio :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC6AD&
      BackStyle       =   0  'Transparent
      Caption         =   "Personal :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC6AD&
      BackStyle       =   0  'Transparent
      Caption         =   "Departamento :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC6AD&
      BackStyle       =   0  'Transparent
      Caption         =   "Sección :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC6AD&
      BackStyle       =   0  'Transparent
      Caption         =   "Fichero :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Control de Presencia, Producción y Accesos © 2017 por Informática del Este, s.l."
      Height          =   660
      Left            =   960
      TabIndex        =   12
      Top             =   0
      Width           =   2655
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   -360
      Picture         =   "Exportar.frx":030A
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00EFC7AE&
      Caption         =   "                                                                    EXPORTACIÓN DATOS A3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   6360
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6280
      Y1              =   740
      Y2              =   740
   End
   Begin VB.Label lestado 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC6AD&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   5520
      Width           =   3495
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC6AD&
      BackStyle       =   0  'Transparent
      Caption         =   "Exportar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   5040
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_sPath As String
Private bCancelar  As Boolean
Private Archivo As String
Private co() As TypeCon


Private Sub Exportar()
Dim RS_Vali As New ADODB.Recordset
Dim F As String
Dim NRegistros As Long
Dim nregistro As Long
Dim acumExtras As Long
Dim acumExtras2 As Long
Dim acumNoctur As Long
Dim acumComple As Long
Dim TipoRegistro As String
Dim empresa As String
Dim centro As String
Dim PerA3Nom As String
Dim FechaActual As String
Dim codIncidencia As String
Dim importe As String
Dim limite As Long 'num max de horas extras a partir del cual consideramos complementeario
Dim Linea As String
Dim ArchivoFinal As String

    If Not Validacion Then Exit Sub
    
  
   
    'If Dir(Trim$(txtdestino.Text)) <> "" Then
   
    '  If MsgBox("El fichero ya existe ¿Desea sobreescribir el fichero?", vbYesNo, "EXPORTACION") = vbYes Then
       ' m_sPath = txtdestino.Text
    ' End If
 
   ' Else
      ' m_sPath = txtdestino.Text
      ' bCancelar = False
    'End If
        
        
    Archivo = txtdestino.Text
   
    Set RS_Vali = DB.Execute(CadenaSQL())
    
   
    
    NRegistros = 0
  
    If Not RS_Vali.EOF Then
      RS_Vali.MoveFirst
      NRegistros = ContadorRegistros
    End If
    
  barra.Value = 0
  nregistro = 0
  If NRegistros > 0 Then
    barra.Max = NRegistros
  Else
    barra.Max = 1
    MsgBox "No existen datos para exportar", vbInformation, "EXPORTACIÓN DATOS VARIABLES"
    'Exit Sub
    GoTo Salir
  End If
  

  If FExist(Archivo) Then
    If MsgBox("El Archivo " & Archivo & " ya existe" & vbCrLf & "¿ Desea eliminarlo ?", vbQuestion + vbYesNo, "EXPORTACION DATOS VARIABLES") = vbYes Then
      Kill Archivo
    Else
      'Exit Sub
      GoTo Salir
    End If
  End If
  
  PrintLog "Los Datos serán exportados sobre el Fichero " & Archivo
  PrintLog "Exportando Datos -----------------------------------"

  Open Archivo For Output As #1
  Screen.MousePointer = 13
  
  acumExtras = 0
  acumNoctur = 0
  acumComple = 0
  acumExtras2 = 0
  limite = Me.txtcomplemento.Text
  
  Do Until RS_Vali.EOF
 
     nregistro = nregistro + 1
     barra.Value = barra.Value + 1
     Me.lestado.Caption = nregistro & " de " & NRegistros
   
        acumExtras = RS_Vali!hextras
        acumNoctur = RS_Vali!hnocturnas
        acumExtras2 = RS_Vali!hextras2
        TipoRegistro = "03"
        empresa = "00999"
        centro = "00000"
        PerA3Nom = Right$("000000" & CStr(RS_Vali!Personal), 6)
        FechaActual = Format$(Now(), "YYYYMMDD")
        codIncidencia = "001"
        importe = "0000000+"
        
      
        
        If (acumExtras > limite * 60) Then  'procesamos horas extras de limite
            acumComple = acumExtras - limite * 60
            acumExtras = limite * 60
        End If
                                                 
        If acumExtras <> 0 Then
                       ' buscamos el codigo que utiliza A3nom y que se encuentra en contadoresequivalentes
                                  
         Linea = TipoRegistro & empresa & centro & PerA3Nom & FechaActual & codIncidencia _
                 & Right$("000" & CStr(CodigoA3(6)), 3) & importe & formatContA3Nom(acumExtras)
               
                 
         Print #1, Linea
        End If
        
       If acumExtras2 <> 0 Then
                       ' buscamos el codigo que utiliza A3nom y que se encuentra en contadoresequivalentes
                                  
         Linea = TipoRegistro & empresa & centro & PerA3Nom & FechaActual & codIncidencia _
                 & Right$("000" & CStr(CodigoA3(7)), 3) & importe & formatContA3Nom(acumExtras2)
               
                 
         Print #1, Linea
        End If
        
        If acumComple <> 0 Then
            'TablaEquiv.FindFirst "contador=8"
           
           Linea = TipoRegistro & empresa & centro & PerA3Nom & FechaActual & codIncidencia _
                    & Right$("000" & CStr(CodigoA3(8)), 3) & importe & formatContA3Nom(acumComple)
                   
           Print #1, Linea
        End If
        
        If acumNoctur >= 480 Then
            'TablaEquiv.FindFirst "contador=9"
            
            Linea = TipoRegistro & empresa & centro & PerA3Nom & FechaActual & codIncidencia _
                    & Right$("000" & CStr(CodigoA3(9)), 3) & importe & formatContA3Nom_dias(acumNoctur)
                   
           Print #1, Linea
        End If

         acumExtras = 0
         acumComple = 0
         acumNoctur = 0
         acumExtras2 = 0
         
     RS_Vali.MoveNext
   
  Loop
  
  PrintLog "Se han exportado " & nregistro
  
  If Not err Then PrintLog "La Exportación de " & Archivo & " se ha realizado con éxito"
    MsgBox "La exportación se ha realizado con exito", vbInformation, "EXPORTACION DATOS VARIABLES"
    
Salir:
  Close #1
  Screen.MousePointer = vbDefault
  Set RS_Vali = Nothing
  
  Exit Sub

ExportarPresencia_Err:
  PrintLog "Error en la Exportación de " & Mid$(F, 1, Len(F) - 1) & vbCrLf & "Numero: " & err.Number & " Descripcion: " & err.Description
  err.Clear
  Screen.MousePointer = vbDefault
  GoTo Salir
End Sub
Private Function CodigoA3(iconta As Integer) As Long
Dim sSQL As String
Dim rs As New ADODB.Recordset


sSQL = " SELECT ContadorA3nom AS valor from ContadoresEquivalente where contador = " & iconta


Set rs = DB.Execute(sSQL)

rs.MoveFirst

CodigoA3 = rs!Valor

rs.Close
Set rs = Nothing

End Function

Function formatContA3Nom_dias(Valor As Long)
Dim horas As String, formatDecimal As String, Min As String, signo As String

    formatDecimal = Format$(Valor / (60 * 8), "###,###,##0.00")
    signo = IIf(Valor >= 0, "+", "-")
    horas = Right$("0000000" & Mid(formatDecimal, 1, InStr(formatDecimal, ",") - 1), 7)
   ' min = Right$("00" & Right$(formatDecimal, InStr(formatDecimal, ",")), 2)
    formatContA3Nom_dias = horas & "00" & signo

End Function
'formatea según la especificación del campo unidades de concepto de A3Nom: 7cifras enteras+2decimales+signo
Function formatContA3Nom(Valor As Long)
Dim horas As String, formatDecimal As String, Min As String, signo As String

    formatDecimal = Format$(Valor / 60, "###,###,##0.00")
    signo = IIf(Valor >= 0, "+", "-")
    horas = Right$("0000000" & Mid(formatDecimal, 1, InStr(formatDecimal, ",") - 1), 7)
    Min = Right$("00" & Right$(formatDecimal, InStr(formatDecimal, ",")), 2)
    formatContA3Nom = horas & Min & signo

End Function

Private Function ContadorRegistros() As Long
 Dim consulta As String  '4-8-10-14-16-17-35-36-37  22/11/2006
 Dim rs As New ADODB.Recordset
  
  consulta = "SELECT Count(*) as valor from ("
  consulta = consulta & "SELECT F.personal,isnull(SUM(F.hextras),0) as hextras,isnull(SUM(F.hextras2),0) as hextras2,"
  consulta = consulta & "isnull(SUM(hnocturnas),0) as hnocturnas FROM ("
  consulta = consulta & "Select v.personal, case when vc.contador=6 then sum(vc.valor) end as hextras,"
  consulta = consulta & " case when vc.contador=7 then sum(vc.valor) end as hextras2,"
  consulta = consulta & " case when vc.Contador=9 then SUM(vc.valor) end as hnocturnas from Validacion V "
  consulta = consulta & "INNER JOIN PERSONAL P ON V.Personal=P.Codigo "
  consulta = consulta & " INNER JOIN  vcontadores VC ON VC.Personal=P.Codigo and v.Fecha=vc.fecha"
  consulta = consulta & " WHERE  (P.FechaBaja is null OR P.FechaBaja>= '" & Me.dtpfechainicio.Value & "')"
  consulta = consulta & " AND P.Calendario Not in('8','10','17')"
  consulta = consulta & " AND V.Horario not in ('4','8','10','14','16','17','35','36','37')"
  consulta = consulta & " AND V.Fecha>= '" & Me.dtpfechainicio.Value & "'"
  consulta = consulta & " AND V.Fecha<='" & Me.dtpfechafin & "'"
  consulta = consulta & " AND VC.Contador in('6','7','9')"
  
  If Me.cbodepartamento.Text <> "" Then
  consulta = consulta & " AND P.Departamento='" & Me.cbodepartamento.Column(1) & "'"
  End If
  If Me.cboseccion.Text <> "" Then
   consulta = consulta & " AND P.Seccion='" & Me.cboseccion.Column(1) & "'"
  End If
  If Me.cbopersonal.Text <> "" Then
    consulta = consulta & " AND P.Codigo='" & Me.cbopersonal.Column(1) & "'"
  End If
  consulta = consulta & " GROUP BY V.personal,vc.contador) F "
  consulta = consulta & " GROUP BY F.Personal) X"
    
 Set rs = DB.Execute(consulta)

  rs.MoveFirst
  
  ContadorRegistros = rs!Valor

  rs.Close
  Set rs = Nothing

End Function
Private Function CadenaSQL() As String
  
  Dim consulta As String  '4-8-10-14-16-17-35-36-37  22/11/2006
  
  
  consulta = "SELECT F.personal,isnull(SUM(F.hextras),0) as hextras,isnull(SUM(F.hextras2),0) as hextras2,"
  consulta = consulta & "isnull(SUM(hnocturnas),0) as hnocturnas FROM ("
  consulta = consulta & "Select v.personal, case when vc.contador=6 then sum(vc.valor) end as hextras,"
  consulta = consulta & " case when vc.contador=7 then sum(vc.valor) end as hextras2,"
  consulta = consulta & " case when vc.Contador=9 then SUM(vc.valor) end as hnocturnas from Validacion V "
  consulta = consulta & "INNER JOIN PERSONAL P ON V.Personal=P.Codigo "
  consulta = consulta & " INNER JOIN  vcontadores VC ON VC.Personal=P.Codigo and v.Fecha=vc.fecha"
  consulta = consulta & " WHERE  (P.FechaBaja is null OR P.FechaBaja>= '" & Me.dtpfechainicio.Value & "')"
  consulta = consulta & " AND P.Calendario Not in('8','10','17')"
  consulta = consulta & " AND V.Horario not in ('4','8','10','14','16','17','35','36','37')"
  consulta = consulta & " AND V.Fecha>= '" & Me.dtpfechainicio.Value & "'"
  consulta = consulta & " AND V.Fecha<='" & Me.dtpfechafin & "'"
  consulta = consulta & " AND VC.Contador in('6','7','9')"
  
  If Me.cbodepartamento.Text <> "" Then
  consulta = consulta & " AND P.Departamento='" & Me.cbodepartamento.Column(1) & "'"
  End If
  If Me.cboseccion.Text <> "" Then
   consulta = consulta & " AND P.Seccion='" & Me.cboseccion.Column(1) & "'"
  End If
  If Me.cbopersonal.Text <> "" Then
    consulta = consulta & " AND P.Codigo='" & Me.cbopersonal.Column(1) & "'"
  End If
  consulta = consulta & " GROUP BY V.personal,vc.contador) F "
  consulta = consulta & " GROUP BY F.Personal"
    
  CadenaSQL = consulta
         
End Function

Private Function Validacion() As Boolean
Dim bok As Boolean

bok = True
If Me.dtpfechainicio > Me.dtpfechafin Then
   bok = False
   MsgBox "la fecha de inicio no puede ser mayor que la fecha de fin", vbOKOnly, "ERROR INTRODUCCION FECHAS"
   Exit Function
End If

If Me.txtdestino.Text = "" Then
   bok = False
   MsgBox "Tiene que introducir la ruta del fichero", vbOKOnly, "INTRODUCIR RUTA FICHERO"
   Exit Function
End If

If Me.txtcomplemento.Text = "" Then
   bok = False
   MsgBox "Introduzca límite para percepción Complementaria", vbOKOnly, "INTRODUCIR PERCEPCION COMPLEMENTARIA"
   Exit Function
End If

If Me.txtcomplemento.Text = 0 Then
   bok = False
   MsgBox "Introduzca límite valido para percepción Complementaria", vbOKOnly, "INTRODUCCION ERRONEA"
   Exit Function
End If
Validacion = bok

End Function

Function Inicializa()

 
  Me.dtpfechainicio.Value = "01/" & Month(Now()) & "/" & Year(Now())
  Me.dtpfechafin.Value = "01/" & Month(Now()) & "/" & Year(Now())
  
End Function

Private Sub Conectar()
On Error GoTo ErrConectar
  
  DB.OpenConnection App.Path & "\WPOpen.UDL"
  
  If DB.error <> 0 Then
    MsgBox "No se ha podido conectar con Winplus.", vbCritical, App.Title
    End
  End If
  
  Exit Sub
  
ErrConectar:
  MsgBox "Error " & err.Number & " : " & err.Description
  PrintLog "Error " & err.Number & " : " & err.Description
End Sub

Private Sub cmdaceptar_Click()
Call Exportar
End Sub

Private Sub cmdbuscar_Click()
    cmdialog.ShowOpen
    txtdestino.Text = Me.cmdialog.FileName
End Sub

Private Sub cmdcancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()

On Error GoTo ErrCargar

    InitINI
    Conectar
    Inicializa

    ChangeSource Me.cbopersonal, "SELECT  APELLIDO1 +' '+APELLIDO2 + ', '+ NOMBRE AS OPERARIO,codigo FROM Personal", True, "250;100"
    ChangeSource Me.cbodepartamento, "SELECT DESCRIPCION,CODIGO FROM departamentos", True, "250;100"
    ChangeSource Me.cboseccion, "SELECT DESCRIPCION,CODIGO FROM secciones", True, "250;100"
    Me.txtdestino.Text = "C:\Winplus\expA3Nom" & Format$(Now(), "YYYYMMDD") & ".txt"
    Me.txtcomplemento.Text = 6
 
    Exit Sub
    
ErrCargar:
  MsgBox "Error " & err.Number & " : " & err.Description
  PrintLog "Error " & err.Number & " : " & err.Description

End Sub



Private Sub txtcomplemento_KeyPress(KeyAscii As Integer)
  Dim bok As Boolean
  
  bok = False
  If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
    bok = True
  Else
    KeyAscii = 0
  End If
  
End Sub
