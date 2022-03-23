VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MDI 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WinPLUS.Open - Integración Nómina"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MDI.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   2  'Custom
   ScaleHeight     =   314
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   436
   StartUpPosition =   1  'CenterOwner
   Begin WPImp.UButton BForzar 
      Height          =   285
      Left            =   5145
      TabIndex        =   8
      Top             =   1125
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   503
      BackColor       =   12632256
      ForeColor       =   8388608
      Text            =   "Activar"
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
   Begin WPImp.UTextBox HoraAct 
      Height          =   285
      Left            =   2160
      TabIndex        =   7
      Top             =   1125
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   503
      DataType        =   5
      ForeColor       =   4194304
      BorderStyle     =   4
      BorderColor     =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
   End
   Begin VB.TextBox TLog 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1695
      Left            =   210
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2295
      Width           =   6105
   End
   Begin VB.Timer TMotor 
      Interval        =   60000
      Left            =   720
      Top             =   5310
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   225
      Left            =   210
      TabIndex        =   2
      Top             =   4305
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin WPImp.UButton BActualizar 
      Height          =   285
      Left            =   5145
      TabIndex        =   10
      Top             =   1500
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   503
      BackColor       =   12632256
      ForeColor       =   8388608
      Text            =   "Actualizar"
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
   Begin WPImp.UTextBox Inter 
      Height          =   285
      Left            =   2160
      TabIndex        =   11
      Top             =   1530
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   503
      DataType        =   1
      ForeColor       =   4194304
      BorderStyle     =   4
      BorderColor     =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Intervalo :"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   390
      TabIndex        =   12
      Top             =   1575
      Width           =   1740
   End
   Begin VB.Label txtEstado 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado : Activo"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   3090
      TabIndex        =   9
      Top             =   1155
      Width           =   1575
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC6AD&
      BackStyle       =   0  'Transparent
      Caption         =   "Motor"
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
      Height          =   225
      Left            =   255
      TabIndex        =   6
      Top             =   690
      Width           =   855
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFC6AD&
      BackStyle       =   0  'Transparent
      Caption         =   "Log"
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
      Left            =   255
      TabIndex        =   5
      Top             =   2010
      Width           =   855
   End
   Begin VB.Label LEstado 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
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
      Height          =   225
      Left            =   255
      TabIndex        =   4
      Top             =   4095
      Width           =   6105
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00D67339&
      BorderWidth     =   2
      X1              =   13
      X2              =   419
      Y1              =   151
      Y2              =   151
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00D67339&
      BorderWidth     =   2
      X1              =   14
      X2              =   420
      Y1              =   63
      Y2              =   63
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   5835
      Picture         =   "MDI.frx":030A
      Top             =   45
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Control de Presencia y Accesos © 2008 por Informática del Este, s.l."
      Height          =   540
      Left            =   945
      TabIndex        =   1
      Top             =   75
      Width           =   3270
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   -270
      Picture         =   "MDI.frx":0614
      Top             =   105
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Hora de Importación :"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   390
      TabIndex        =   0
      Top             =   1155
      Width           =   1740
   End
   Begin VB.Image ILed 
      Height          =   240
      Index           =   0
      Left            =   4710
      Picture         =   "MDI.frx":0A76
      Top             =   1140
      Width           =   240
   End
   Begin VB.Image ILed 
      Height          =   240
      Index           =   1
      Left            =   2925
      Picture         =   "MDI.frx":1000
      Top             =   5355
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ILed 
      Height          =   240
      Index           =   2
      Left            =   2610
      Picture         =   "MDI.frx":158A
      Top             =   5340
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ILed 
      Height          =   240
      Index           =   3
      Left            =   2295
      Picture         =   "MDI.frx":1B14
      Top             =   5355
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00EFC6AD&
      FillStyle       =   0  'Solid
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   6570
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1035
      Left            =   210
      Top             =   945
      Width           =   6105
   End
End
Attribute VB_Name = "MDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim I As Long
Dim Estado As Byte ' 1= En espera; 2= En proceso; 3=Parado
Dim strsql As String
Dim ImportArchivo As String
Dim UltimaImportacion As Variant
Dim TipoEjecucion As String

Private Sub BActualizar_Click()


    'If Estado = 1 Then
        ContImport = ContImport + 1
        If ContImport > 5 Then
            Me.TLog.Text = ""
            ContImport = 1
        End If
        'Lanzar el proceso de importación
        PrintLog "Comienzo del proceso de Importación Manual.", 1
  
        If ImportarNomina() Then
            PrintLog "Proceso de Importación Manual Finalizado, sin anomalías.", 1
        Else
            PrintLog "Proceso de Importación Manual Finalizado, con anomalías.", 1
        End If
        Estado = 1
        Led Estado
        ActBarra 0
    'End If

End Sub

Private Sub Form_Load()
On Error GoTo err_conect

  InitINI
  TipoCon = 0

  DB.OpenConnection App.Path & "\WPOpen.UDL"
  If DB.Error <> 0 Then
    MsgBox "No se ha podido conectar con Winplus.", vbCritical, App.Title
    End
  End If
 
  Screen.MousePointer = 0
  ImportArchivo = DB.Lookup("VALOR", "CONFIG", "CLAVE='DirImportNomina'")
 
  If DB.Count("*", "config", "CLAVE='HImportNomina'") = 0 Then
    DB.Execute "INSERT INTO CONFIG(clave,Valor,Tipo,Descripcion) Values ('HImportNomina','00:00',10,'Hora Inicio para el proceso de Actualización con Nómina')"
    If DB.Error <> 0 Then
      MsgBox "No se ha podido insertar la clave HImportNomina en la tabla CONFIG."
      End
    End If
  End If
  
  Me.HoraAct = Nz(DB.Lookup("valor", "CONFIG", "CLAVE='HImportNomina'"), "00:00")
  
  Me.Inter = Nz(DB.Lookup("valor", "CONFIG", "CLAVE='InterImportNomina'"), 0)

  UltimaImportacion = Now
  
 
    Estado = 1
   ' Me.TMotor.Enabled = False
    'Me.BForzar.Enabled = False
    'Me.Inter.Enabled = False
    'Me.HoraAct.Enabled = False
  
  Led Estado
  ContImport = 0
  'Iniciamos el contenido de la lista de mensajes
  PrintLog "Inicialización del Servidor de Importación de Nómina.", 1

  Exit Sub
err_conect:
    If err.Number <> 0 Then
        MsgBox err.Description, , App.Title
        err.Clear
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.TMotor.Enabled = False
  PrintLog "Cierre del servidor..."
  DoEvents
End Sub

Private Function Led(Estado As Byte)
  On Error Resume Next
  Set Me.ILed(0).Picture = Me.ILed(Estado).Picture
  Select Case Estado
  Case 1 'Verde
    Me.txtEstado = "Estado: Activo"
    Me.BForzar.Text = "Stop"
  Case 2 'Naranja
    Me.txtEstado = "Estado: Importando"
    Me.BForzar.Text = "Stop"
  Case 3 'Rojo
    Me.txtEstado = "Estado: Parado"
    Me.BForzar.Text = "Activar"
  End Select
  DoEvents
  err.Clear
End Function

Private Sub HoraAct_AfterUpdate(Cancel As Boolean)
  Dim mihora As String
  mihora = DB.Lookup("VALOR", "CONFIG", "CLAVE='HImportNomina'")
  If DB.Errors.Count <> 0 Then
    PrintLog "No se puede acceder a la base de datos. Se recomienda Cerrar y volver a abrir la aplicación. Si el problema persiste consulte con el administrador de la base de datos."
  Else
    If Me.HoraAct <> Me.HoraAct.OldValue Then
      I = MsgBox("Intenta cambiar la hora de activación del proceso de Importacion. ¿Está seguro?", vbYesNo + vbQuestion, App.Title)
      If I = 6 Then
        DB.Update "CONFIG", "VALOR='" & Format(Me.HoraAct, "HH:MM") & "'", "CLAVE='HImportNomina'"
        If DB.Errors.Count <> 0 Then
          PrintLog "No se puede actualizar el cambio de hora de activación. Se recomienda Cerrar y volver a abrir la aplicación. Si el problema persiste consulte con el administrador de la base de datos."
        Else
          PrintLog "Actualización del cambio de hora de activación completada. El proceso de importación se realizará a partir de ahora a las :  " & Format(Me.HoraAct, "HH:MM")
        End If
      Else
        Cancel = True
      End If
    End If
  End If
End Sub

Private Sub Inter_AfterUpdate(Cancel As Boolean)
  Dim mihora As String
  mihora = DB.Lookup("VALOR", "CONFIG", "CLAVE='InterImportNomina'")
  If DB.Errors.Count <> 0 Then
    PrintLog "No se puede acceder a la base de datos. Se recomienda Cerrar y volver a abrir la aplicación. Si el problema persiste consulte con el administrador de la base de datos."
  Else
    If Me.Inter <> Me.Inter.OldValue Then
      I = MsgBox("Intenta cambiar el intervalo del proceso de Importacion. ¿Está seguro?", vbYesNo + vbQuestion, App.Title)
      If I = 6 Then
        DB.Update "CONFIG", "VALOR='" & Me.Inter & "'", "CLAVE='InterImportNomina'"
        If DB.Errors.Count <> 0 Then
          PrintLog "No se puede actualizar el intervalo de activación. Se recomienda Cerrar y volver a abrir la aplicación. Si el problema persiste consulte con el administrador de la base de datos."
        Else
          PrintLog "Actualización del intervalo de activación completada. El proceso de importación se realizará intervalos de :  " & Me.Inter & " Minutos"
        End If
      Else
        Cancel = True
      End If
    End If
  End If
End Sub

Private Sub TMotor_Timer()
Dim TocaIntervalo As Boolean


  TocaIntervalo = False
  TocaIntervalo = (DateDiff("n", CDate(UltimaImportacion), CDate(Now)) >= Me.Inter) And (Not (Format(Now, "HH:MM") = Format(Me.HoraAct, "HH:MM"))) And (Me.Inter > 0)

  If Estado = 1 Then
    If (Format(Now, "HH:MM") = Format(Me.HoraAct, "HH:MM")) Or TocaIntervalo Then
      'Vaciamos la lista de mensajes cada 5 importaciones.
      ContImport = ContImport + 1
      If ContImport > 5 Then
        Me.TLog.Text = ""
        ContImport = 1
      End If
      'Lanzar el proceso de importación
      PrintLog "Comienzo del proceso de Importación Desatendido.", 1
      UltimaImportacion = Now

      If ImportarNomina() Then
        PrintLog "Proceso de Importación Desatendido Finalizado, sin anomalías.", 1
      Else
        PrintLog "Proceso de Importación Desatendido Finalizado, con anomalías.", 1
      End If
      Estado = 1
      Led Estado
      ActBarra 0
    End If
  End If
End Sub

Private Sub BForzar_Click()
  If Estado <> 3 Then 'En funcionamiento
    Estado = 3
    Me.TMotor.Enabled = False
    SrvMensaje = "Se ha parado manualmente el Servidor de Importación."
    BForzar.Text = "Inicializar"
  Else
    Estado = 1
    BForzar.Text = "Stop"
    Me.TMotor.Enabled = True
    SrvMensaje = "Se ha Inicializado manualmente el Servidor de Importación."
  End If
  PrintLog SrvMensaje, 1
  Led Estado
End Sub



Private Sub PrintLog(Texto As String, Optional control As Byte = 0)
  Select Case control
  Case 0
    On Error Resume Next
     Me.TLog.Text = Me.TLog.Text & vbCrLf & " " & Now & " -- " & Texto
    If err Then Me.TLog.Text = " " & Now & " -- " & Texto
    On Error GoTo 0
    Texto = " " & Now & " -- " & Texto
  Case 1
    On Error Resume Next
    Me.TLog.Text = Me.TLog.Text & vbCrLf & String(130, "-") & vbCrLf & " " & Now & " -- " & Texto & vbCrLf & String(130, "-")
    If err Then Me.TLog.Text = String(130, "-") & vbCrLf & " " & Now & " -- " & Texto & vbCrLf & String(130, "-")
    On Error GoTo 0
    Texto = String(130, "-") & vbCrLf & " " & Now & " -- " & Texto & vbCrLf & String(130, "-")
  End Select
  AddBackup Texto
  Me.TLog.SelStart = Len(Me.TLog)
  err.Clear
End Sub

Private Sub ActBarra(Optional Tipo As Byte = 1)
  If Tipo = 1 Then
    Me.LEstado.Caption = SrvActual & " de " & SrvTotal
    Me.PBar.Value = SrvPorcentaje
  Else
    Me.LEstado.Caption = "Estado: "
    Me.PBar.Value = 0
  End If
  DoEvents
End Sub
Private Function ImportarNomina() As Boolean
  
  Dim Importacion As Boolean

  
  Estado = 2
  Led Estado
  Importacion = True

    If FExist(ImportArchivo & "\EMPRESAS.TXT") Then
        Importacion = Actualizar_Tabla("EMPRESAS")
        If Not Importacion Then GoTo Salir
        PrintLog "EMPRESAS Actualizadas."
    Else
        PrintLog "No se importarán EMPRESAS. No existe el fichero."
    End If

  
    If FExist(ImportArchivo & "\CENTROS.TXT") Then
        Importacion = Actualizar_Tabla("CENTROS")
        If Not Importacion Then GoTo Salir
        PrintLog "CENTROS Actualizadas."
    Else
        PrintLog "No se importarán CENTROS. No existe el fichero."
    End If
  
    If FExist(ImportArchivo & "\DEPARTAMENTOS.TXT") Then
      Importacion = Actualizar_Tabla("DEPARTAMENTOS")
      If Not Importacion Then GoTo Salir
      PrintLog "DEPARTAMENTOS Actualizadas."
    Else
      PrintLog "No se importarán DEPARTAMENTOS. No existe el fichero."
    End If
  
    'If FExist(ImportArchivo & "\SECCIONES.TXT") Then
      'Importacion = Actualizar_Tabla("SECCIONES")
      'If Not Importacion Then GoTo Salir
      'PrintLog "SECCIONES Actualizadas."
    'Else
      'PrintLog "No se importarán SECCIONES. No existe el fichero."
    'End If
  
   If FExist(ImportArchivo & "\CATEGORIAS.TXT") Then
        Importacion = Actualizar_Tabla("CATEGORIAS")
        If Not Importacion Then GoTo Salir
        PrintLog "CATEGORIAS Actualizadas."
    Else
        PrintLog "No se importarán CATEGORIAS. No existe el fichero."
    End If

  
    If FExist(ImportArchivo & "\PERSONAL.TXT") Then
        Importacion = Actualizar_Personal
        If Not Importacion Then GoTo Salir
        PrintLog "PERSONAL Actualizadas."
    Else
        PrintLog "No se importará PERSONAL. No existe el fichero."
    End If

    If FExist(ImportArchivo & "\ABSENTISMO.TXT") Then
        Importacion = Actualizar_Absentismo
        If Not Importacion Then GoTo Salir
        PrintLog "ABSENTISMOS Actualizados."
    Else
        PrintLog "No se importarán ABSENTISMOS. No existe el fichero."
    End If
  
  
Salir:
  ImportarNomina = Importacion
End Function
Private Function Actualizar_Personal() As Boolean
On Error GoTo Actualizar_PersonalErr
  
  Dim Lin As String
  Dim CODIGO As String
  Dim NOMBRE As String
  Dim Apellidos, APELLIDO1, APELLIDO2 As String
  Dim TARJETA As String
  Dim TELEFONO As String
  Dim EMPRESA, CENTRO, DEPARTAMENTO, SECCION, CATEGORIA, PUESTO, CALIFICACION, COSTE, CALENDARIO, CENTROCOSTE As String
  Dim DNI As String
  Dim FECHAINGRESO, FECHABAJA As String
  Dim OBSERVACIONES As String
  Dim SEXO As String
  Dim TIPOCONTROL As Byte
  Dim Sep As Integer
  Dim mitrans As Long
  Dim Activo As Boolean
  Dim Fichero As String



  If Dir(ImportArchivo & "\Backup", vbDirectory) <> "" Then
    Fichero = ImportArchivo & "\Backup\Personal" & Format$(Now, "yyyymmddhhmm") & ".TXT"
  Else
    Fichero = ImportArchivo & "\Personal" & Format$(Now, "yyyymmddhhmm") & ".TXT"
  End If
  

  Name ImportArchivo & "\Personal.TXT" As Fichero




  Actualizar_Personal = True

  SrvTotal = NumLineas(Fichero)
  SrvActual = 0
  PBar.Value = 0
  Screen.MousePointer = vbHourglass
  mitrans = DB.Cn.BeginTrans

  Open Fichero For Input Access Read Shared As 1

  
  Do While Not EOF(1)
      Sep = 1
      SrvActual = SrvActual + 1
      ActBarra
      Line Input #1, Lin
      
      If IsNull(Lin) Or Lin = "" Then
        Close (1)
        PrintLog "Importacion interrumpida por datos erroneos en Fichero Personal."
        Exit Function
      End If
      
      CODIGO = Right$("00000000" & Left$(Mid$(Lin, 1, InStr(Sep, Lin, ";") - 1), 10), 8)
      Sep = InStr(Sep, Lin, ";")
      
      NOMBRE = Left$(Mid$(Lin, Sep + 1, InStr(Sep + 1, Lin, ";") - Sep - 1), 20)
      Sep = InStr(Sep + 1, Lin, ";")
      
      APELLIDO1 = Left$(Mid$(Lin, Sep + 1, InStr(Sep + 1, Lin, ";") - Sep - 1), 20)
      Sep = InStr(Sep + 1, Lin, ";")
      
      APELLIDO2 = Left$(Mid$(Lin, Sep + 1, InStr(Sep + 1, Lin, ";") - Sep - 1), 20)
      Sep = InStr(Sep + 1, Lin, ";")
      
      TARJETA = Nz(Left$(Mid$(Lin, Sep + 1, InStr(Sep + 1, Lin, ";") - Sep - 1), 13), Null)
      If TARJETA <> "" Then TARJETA = Right$("0000000000000" & TARJETA, 13)
      Sep = InStr(Sep + 1, Lin, ";")
      
      DNI = Left$(Mid$(Lin, Sep + 1, InStr(Sep + 1, Lin, ";") - Sep - 1), 12)
      Sep = InStr(Sep + 1, Lin, ";")
      
      TELEFONO = Left$(Mid$(Lin, Sep + 1, InStr(Sep + 1, Lin, ";") - Sep - 1), 30)
      Sep = InStr(Sep + 1, Lin, ";")
      
      EMPRESA = Left$(Mid$(Lin, Sep + 1, InStr(Sep + 1, Lin, ";") - Sep - 1), 10)
      If EMPRESA = "" Then EMPRESA = Null
      Sep = InStr(Sep + 1, Lin, ";")
      
      CENTRO = Left$(Mid$(Lin, Sep + 1, InStr(Sep + 1, Lin, ";") - Sep - 1), 10)
      If CENTRO = "" Then CENTRO = Null
      Sep = InStr(Sep + 1, Lin, ";")
      
      DEPARTAMENTO = Left$(Mid$(Lin, Sep + 1, InStr(Sep + 1, Lin, ";") - Sep - 1), 10)
      If DEPARTAMENTO = "" Then DEPARTAMENTO = Null
      Sep = InStr(Sep + 1, Lin, ";")
      
      SECCION = Left$(Mid$(Lin, Sep + 1, InStr(Sep + 1, Lin, ";") - Sep - 1), 10)
      If SECCION = "" Then SECCION = Null
      Sep = InStr(Sep + 1, Lin, ";")
         
      CATEGORIA = Left$(Mid$(Lin, Sep + 1, InStr(Sep + 1, Lin, ";") - Sep - 1), 10)
      If CATEGORIA = "" Then CATEGORIA = Null
      Sep = InStr(Sep + 1, Lin, ";")
      
      FECHAINGRESO = Mid$(Mid$(Lin, Sep + 1, InStr(Sep + 1, Lin, ";") - Sep - 1), 1, 2) & "/" & Mid$(Mid$(Lin, Sep + 1, InStr(Sep + 1, Lin, ";") - Sep - 1), 4, 2) & "/" & Mid$(Mid$(Lin, Sep + 1, InStr(Sep + 1, Lin, ";") - Sep - 1), 7, 4)
      If Not IsDate(FECHAINGRESO) Then FECHAINGRESO = ""
      Sep = InStr(Sep + 1, Lin, ";")
      
      FECHABAJA = Nz(Mid$(Mid$(Lin, Sep + 1, InStr(Sep + 1, Lin, ";") - Sep - 1), 1, 2) & "/" & Mid$(Mid$(Lin, Sep + 1, InStr(Sep + 1, Lin, ";") - Sep - 1), 4, 2) & "/" & Mid$(Mid$(Lin, Sep + 1, InStr(Sep + 1, Lin, ";") - Sep - 1), 7, 4), "")
      If Not IsDate(FECHABAJA) Then FECHABAJA = ""
      Sep = InStr(Sep + 1, Lin, ";")
      
      OBSERVACIONES = Mid$(Lin, Sep + 1, InStr(Sep + 1, Lin, ";") - Sep - 1)
      Sep = InStr(Sep + 1, Lin, ";")
      
      SEXO = Left$(Mid$(Lin, Sep + 1), 1)
      
      TIPOCONTROL = 4
      

      Activo = True
      If Not IsZero(FECHABAJA) Then
        If CDate(FECHABAJA) < Fix(Now) Then Activo = False
      End If
      

      
        If DB.Find("PERSONAL", "CODIGO = '" & CODIGO & "'") Then
            strsql = "UPDATE PERSONAL SET " & _
            "NOMBRE='" & Left$(NOMBRE, 60) & "'," & _
            "APELLIDO1='" & Left$(APELLIDO1, 20) & "'," & "APELLIDO2='" & Left$(APELLIDO2, 20) & "'," & _
            "TELEFONO='" & TELEFONO & "'," & _
            "EMPRESA=" & DB.IStr(EMPRESA) & "," & "CENTRO=" & DB.IStr(CENTRO) & "," & _
            "DEPARTAMENTO=" & DB.IStr(DEPARTAMENTO) & "," & "SECCION=" & DB.IStr(SECCION) & "," & _
            "DNI='" & DNI & "'," & _
            "FECHAINGRESO=" & DB.IFec(FECHAINGRESO) & "," & "FECHABAJA=" & DB.IFec(FECHABAJA) & _
            " WHERE CODIGO = '" & CODIGO & "'"

            DB.Execute strsql
            If DB.Error Then
                PrintLog "Error al Actualizar Personal: " & vbCrLf & _
                "Numero: " & DB.Errors(0).Number & " Descripcion: " & DB.Errors(0).Description & vbCrLf & _
                strsql
                Actualizar_Personal = False
            End If
        Else
          If Activo Then
            strsql = "INSERT INTO PERSONAL " & _
            "(CODIGO, NOMBRE, APELLIDO1, APELLIDO2, TARJETA, TELEFONO, EMPRESA," & _
            " CENTRO, DEPARTAMENTO, SECCION," & _
            " DNI,FECHAINGRESO, FECHABAJA, OBSERVACIONES, SEXO, TIPOCONTROL) " & _
            " Values('" & CODIGO & "','" & Left$(NOMBRE, 60) & "'," & _
            "'" & Left$(APELLIDO1, 20) & "','" & Left$(APELLIDO2, 20) & "'," & _
            "'" & TARJETA & "','" & TELEFONO & "'," & DB.IStr(EMPRESA) & "," & _
            DB.IStr(CENTRO) & "," & DB.IStr(DEPARTAMENTO) & "," & DB.IStr(SECCION) & "," & _
            "'" & DNI & "'," & _
            DB.IFec(FECHAINGRESO) & "," & DB.IFec(FECHABAJA) & ",'" & OBSERVACIONES & "'," & _
            "'" & SEXO & "'," & TIPOCONTROL & ")"

            DB.Execute strsql
            If DB.Error Then
                PrintLog "Error al Insertar en Personal: " & vbCrLf & _
                "Numero: " & DB.Errors(0).Number & " Descripcion: " & DB.Errors(0).Description & vbCrLf & _
                strsql
                Actualizar_Personal = False
            End If
          End If
        End If
        
    'End If
  Loop
  

Salir:
  If mitrans <> 0 Then DB.Cn.CommitTrans
  Close (1)
  Screen.MousePointer = vbDefault
  Exit Function

Actualizar_PersonalErr:
  If mitrans <> 0 Then DB.Cn.RollbackTrans
  mitrans = 0
  PrintLog "Numero: " & err.Number & " Descripcion: " & err.Description
  PrintLog "Revise los registros del Fichero de Personal generados por el software de Nómina o la conexion con su base de datos Winplus", 0
  err.Clear
  Screen.MousePointer = vbDefault
  Actualizar_Personal = False
  GoTo Salir
  
End Function

Private Function Actualizar_Tabla(Tabla As String) As Boolean
On Error GoTo Actualizar_TablaErr
  Dim Lin As String
  Dim CODIGO As String
  Dim DESCRIPCION As String
  Dim Sep As Integer
  Dim mitrans As Long
  Dim Fichero As String
  
  Fichero = Tabla
  Fichero = Tabla & Format$(Now, "yyyymmddhhmm")
  

  If Dir(ImportArchivo & "\Backup", vbDirectory) <> "" Then
    Name ImportArchivo & "\" & Tabla & ".TXT" As ImportArchivo & "\Backup\" & Fichero & ".TXT"
  Else
    Name ImportArchivo & "\" & Tabla & ".TXT" As ImportArchivo & "\" & Fichero & ".TXT"
  End If
   
  Actualizar_Tabla = True
  


  If Dir(ImportArchivo & "\Backup", vbDirectory) <> "" Then
    SrvTotal = NumLineas(ImportArchivo & "\Backup\" & Fichero & ".TXT")
  Else
    SrvTotal = NumLineas(ImportArchivo & "\" & Fichero & ".TXT")
  End If


  SrvActual = 0
  PBar.Value = 0
  Screen.MousePointer = vbHourglass
  mitrans = DB.Cn.BeginTrans
  

  If Dir(ImportArchivo & "\Backup", vbDirectory) <> "" Then
    Open ImportArchivo & "\Backup\" & Fichero & ".TXT" For Input Access Read Shared As 1
  Else
    Open ImportArchivo & "\" & Fichero & ".TXT" For Input Access Read Shared As 1
  End If

  
  Do While Not EOF(1)
      SrvActual = SrvActual + 1
      ActBarra
      Sep = 1
      Line Input #1, Lin
      If IsNull(Lin) Or Lin = "" Then
        Close (1)
        PrintLog "Actualización interrumpida por datos erroneos en Fichero " & Tabla & ".TXT"
        Exit Function
      End If
      CODIGO = Mid$(Lin, 1, InStr(Sep, Lin, ";") - 1)
      Sep = InStr(Sep, Lin, ";")
      DESCRIPCION = Left$(Mid$(Lin, Sep + 1, Len(Lin)), 50)
      
      
      Select Case Tabla
        Case "CALENDARIOS": CODIGO = Right$("000" & CODIGO, 3)
        Case Else:
      End Select
      

      If DB.Find(Tabla, "CODIGO = '" & CODIGO & "'") Then
            strsql = "UPDATE " & Tabla & " SET " & _
            "DESCRIPCION='" & Left$(DESCRIPCION, 30) & "'" & _
            " WHERE CODIGO = '" & CODIGO & "'"
            DB.Execute strsql
            If DB.Error Then
                PrintLog "Error al Actualizar la Tabla: " & Tabla & vbCrLf & _
                "Numero: " & DB.Errors(0).Number & " Descripcion: " & DB.Errors(0).Description & vbCrLf & _
                strsql
                Actualizar_Tabla = False
            End If
      Else
            If CODIGO <> "" Then
                strsql = "INSERT INTO " & Tabla & " (CODIGO, DESCRIPCION) " & _
                " Values('" & CODIGO & "','" & Left$(DESCRIPCION, 30) & "')"
                DB.Execute strsql
                If DB.Error Then
                    PrintLog "Error al Insertar en  la Tabla: " & Tabla & vbCrLf & _
                    "Numero: " & DB.Errors(0).Number & " Descripcion: " & DB.Errors(0).Description & vbCrLf & _
                    strsql
                    Actualizar_Tabla = False
                End If
            End If
      End If
  Loop
  
  'Close (1)
  
Salir:
  If mitrans <> 0 Then DB.Cn.CommitTrans
  Close (1)
  Screen.MousePointer = vbDefault
  Exit Function

Actualizar_TablaErr:
  If mitrans <> 0 Then DB.Cn.RollbackTrans
  mitrans = 0
  PrintLog "Numero: " & err.Number & " Descripcion: " & err.Description
  PrintLog "Revise los registros de los Ficheros generados por el software de Nómina o la conexion con su base de datos Winplus", 0
  err.Clear
  Screen.MousePointer = vbDefault
  Actualizar_Tabla = False
  GoTo Salir
End Function
Private Function NumLineas(Fichero As String)
Dim NLin As Long
Dim Lin As String

Open Fichero For Input Access Read Shared As 1
NLin = 0
Do While Not EOF(1)
  Line Input #1, Lin
  NLin = NLin + 1
Loop
NumLineas = NLin
Close (1)

End Function


Private Function Actualizar_Absentismo() As Boolean
On Error GoTo Actualizar_AbsentismoErr

  'Dim RPerDestino As New ADODB.Recordset
  
  Dim Lin As String
  Dim PERSONAL As String
  Dim EMPRESA, CENTRO  As String
  Dim CAUSA As String
  Dim FECHAINICIO, FECHAFIN As String
  Dim OPERACION As Byte
  Dim Sep As Integer
  Dim mitrans As Long
  Dim Activo As Boolean
  Dim Fichero As String



  If Dir(ImportArchivo & "\Backup", vbDirectory) <> "" Then
    Fichero = ImportArchivo & "\Backup\Absentismo" & Format$(Now, "yyyymmddhhmm") & ".TXT"
  Else
    Fichero = ImportArchivo & "\Absentismo" & Format$(Now, "yyyymmddhhmm") & ".TXT"
  End If
  
  Name ImportArchivo & "\Absentismo.TXT" As Fichero

  Actualizar_Absentismo = True
  SrvTotal = NumLineas(Fichero)
  SrvActual = 0
  PBar.Value = 0
  Screen.MousePointer = vbHourglass
  mitrans = DB.Cn.BeginTrans
  Open Fichero For Input Access Read Shared As 1

  
  Do While Not EOF(1)
      Sep = 1
      SrvActual = SrvActual + 1
      ActBarra
      Line Input #1, Lin
      
      If IsNull(Lin) Or Lin = "" Then
        Close (1)
        PrintLog "Importacion interrumpida por datos erroneos en Fichero Absentismo."
        Exit Function
      End If
      
      OPERACION = Nz(Left$(Mid$(Lin, 1, InStr(Sep, Lin, ";") - 1), 10), 1)
      Sep = InStr(Sep, Lin, ";")
      
      PERSONAL = Right$("00000000" & Left$(Mid$(Lin, Sep + 1, InStr(Sep + 1, Lin, ";") - Sep - 1), 10), 8)
      Sep = InStr(Sep + 1, Lin, ";")
      
      EMPRESA = Left$(Mid$(Lin, Sep + 1, InStr(Sep + 1, Lin, ";") - Sep - 1), 10)
      Sep = InStr(Sep + 1, Lin, ";")
      
      CENTRO = Left$(Mid$(Lin, Sep + 1, InStr(Sep + 1, Lin, ";") - Sep - 1), 10)
      Sep = InStr(Sep + 1, Lin, ";")
      
      CAUSA = Nz(Mid$(Lin, Sep + 1, InStr(Sep + 1, Lin, ";") - Sep - 1), 6)
      Sep = InStr(Sep + 1, Lin, ";")

      FECHAINICIO = Mid$(Lin, Sep + 1, InStr(Sep + 1, Lin, ";") - Sep - 1)
      If Not IsDate(FECHAINICIO) Then FECHAINICIO = ""
      Sep = InStr(Sep + 1, Lin, ";")
      
      FECHAFIN = Mid$(Lin, Sep + 1)
      If Not IsDate(FECHAFIN) Then FECHAFIN = ""
      If IsZero(FECHAFIN) Then FECHAFIN = "31/12/2050"

      Activo = True
      If Not IsZero(FECHAINICIO) Then
        If (CDate(FECHAINICIO) > CDate(FECHAFIN)) Or CAUSA = 0 Or Nz(DB.Lookup("CODIGO", "PERSONAL", "CODIGO = '" & PERSONAL & "'"), 0) = 0 Then Activo = False
      Else
        Activo = False
      End If


      If OPERACION = 1 Then
'        If DB.Find("PREVISIONES", "PERSONAL = '" & PERSONAL & "'" & " AND FECHAINICIO = " & DB.IFec(FECHAINICIO) & " AND FECHAFIN = " & DB.IFec(FECHAFIN) & " AND CAUSA = " & CAUSA & " AND INCIDENCIA = 12 AND HORAINICIO = 0 AND HORAFIN = 0") Then
'            strsql = "UPDATE PREVISIONES SET " & _
'            "PERSONAL='" & PERSONAL & "'," & _
'            "FECHAINICIO=" & DB.IFec(FECHAINICIO) & "," & "FECHAFIN=" & DB.IFec(FECHAFIN) & "," & _
'            "CAUSA=" & CAUSA & _
'            " WHERE PERSONAL = '" & PERSONAL & "'" & " AND FECHAINICIO = " & DB.IFec(FECHAINICIO) & _
'            " AND FECHAFIN = " & DB.IFec(FECHAFIN) & " AND CAUSA = " & CAUSA & " AND INCIDENCIA = 12 AND HORAINICIO = 0 AND HORAFIN = 0"
'
'            DB.Execute strsql
'            If DB.Error Then
'                PrintLog "Error al Actualizar Previsiones: " & vbCrLf & _
'                "Numero: " & DB.Errors(0).Number & " Descripcion: " & DB.Errors(0).Description & vbCrLf & _
'                strsql
'                Actualizar_Absentismo = False
'            End If
'        Else
          If Activo Then
            strsql = "INSERT INTO PREVISIONES " & _
            "(PERSONAL, FECHAINICIO, FECHAFIN, HORAINICIO, HORAFIN, INCIDENCIA, CAUSA)" & _
            " VALUES ('" & PERSONAL & "', " & DB.IFec(FECHAINICIO) & "," & DB.IFec(FECHAFIN) & ", " & _
            "0, 0, 12, " & CAUSA & ")"

            DB.Execute strsql
            If DB.Error Then
                PrintLog "Error al Insertar en Previsiones: " & vbCrLf & _
                "Numero: " & DB.Errors(0).Number & " Descripcion: " & DB.Errors(0).Description & vbCrLf & _
                strsql
                Actualizar_Absentismo = False
            Else
                'realizo un update en validacion para que el servidor de calculos recalcule las fichas afectadas
                strsql = "UPDATE Validacion set estado=0 where Personal='" & PERSONAL & "'"
                strsql = strsql & " AND Fecha >=" & DB.IFec(FECHAINICIO) & ""
                strsql = strsql & " AND  Fecha<=" & DB.IFec(FECHAFIN) & ""
                DB.Execute strsql
                If DB.Error Then
                   PrintLog "Error al actualizar las fichas de validacion: " & vbCrLf & _
                           "Numero: " & DB.Errors(0).Number & " Descripcion: " & DB.Errors(0).Description & vbCrLf & _
                    strsql
                    Actualizar_Absentismo = False
                End If
                 
            End If
          End If
      Else
        strsql = "DELETE FROM PREVISIONES " & _
        " WHERE PERSONAL = '" & PERSONAL & "'" & " AND FECHAINICIO = " & DB.IFec(FECHAINICIO) & _
        " AND FECHAFIN = " & DB.IFec(FECHAFIN) & " AND CAUSA = " & CAUSA & " AND INCIDENCIA = 12 AND HORAINICIO = 0 AND HORAFIN = 0"
        
        DB.Execute strsql
        If DB.Error Then
            PrintLog "Error al  Eliminar en Previsiones: " & vbCrLf & _
            "Numero: " & DB.Errors(0).Number & " Descripcion: " & DB.Errors(0).Description & vbCrLf & _
            strsql
            Actualizar_Absentismo = False
        End If
      End If
        
  Loop
  

Salir:
  If mitrans <> 0 Then DB.Cn.CommitTrans
  Close (1)
  Screen.MousePointer = vbDefault
  Exit Function

Actualizar_AbsentismoErr:
  If mitrans <> 0 Then DB.Cn.RollbackTrans
  mitrans = 0
  PrintLog "Numero: " & err.Number & " Descripcion: " & err.Description
  PrintLog "Revise los registros del Fichero de Absentismo generado por el software de Nómina o la conexion con su base de datos Winplus", 0
  err.Clear
  Screen.MousePointer = vbDefault
  Actualizar_Absentismo = False
  GoTo Salir
  
End Function


