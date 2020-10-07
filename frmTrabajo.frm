VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTrabajo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos del trabajo"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   Icon            =   "frmTrabajo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   60
      Top             =   7680
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   180
      TabIndex        =   30
      Top             =   8760
      Visible         =   0   'False
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
      Max             =   50
   End
   Begin VB.CommandButton btnVolverAInstalar 
      Caption         =   "Pasar a ""Para programar"""
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   26
      ToolTipText     =   "La orden vuelve a la primera solapa del sistema"
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Frame frmDatosUsuario 
      Height          =   8715
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   6255
      Begin VB.CheckBox chkEnviarCorreoOrden 
         Caption         =   "Enviar orden por correo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   29
         Top             =   6000
         Visible         =   0   'False
         Width           =   3255
      End
      Begin MSComDlg.CommonDialog cdImpresora 
         Left            =   5520
         Top             =   6000
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CheckBox chkImprimirOrden 
         Caption         =   "Imprimir orden de trabajo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   28
         Top             =   5520
         Width           =   3135
      End
      Begin VB.CommandButton btnImprimirOrden 
         Caption         =   "Imprimir orden de trabajo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3720
         TabIndex        =   27
         Top             =   5520
         Width           =   2295
      End
      Begin VB.TextBox txtObsConex 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         MaxLength       =   50
         TabIndex        =   24
         Top             =   6960
         Width           =   5775
      End
      Begin VB.TextBox txtFechaPedido 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   2280
         Width           =   3735
      End
      Begin VB.TextBox txtObs 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   240
         MaxLength       =   50
         TabIndex        =   22
         Top             =   3360
         Width           =   5775
      End
      Begin VB.ComboBox cmbTipoConexion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "frmTrabajo.frx":030A
         Left            =   240
         List            =   "frmTrabajo.frx":031A
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   4200
         Width           =   2775
      End
      Begin VB.TextBox txtTlfn 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1800
         Width           =   3735
      End
      Begin VB.TextBox txtUsInternet 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox txtDomi 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   840
         Width           =   3735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   360
         Width           =   3735
      End
      Begin VB.ComboBox cmbCuadrilla 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   4200
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker dtHoraInst 
         Height          =   375
         Left            =   3600
         TabIndex        =   15
         Top             =   5040
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "hh:mm tt"
         Format          =   95879171
         UpDown          =   -1  'True
         CurrentDate     =   44076
      End
      Begin MSComCtl2.DTPicker dtFechaInst 
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   5040
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   95879169
         CurrentDate     =   44076
      End
      Begin VB.CommandButton btnEliminar 
         Caption         =   "Eliminar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   3240
         Picture         =   "frmTrabajo.frx":0354
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   7560
         Width           =   1245
      End
      Begin VB.CommandButton btnVolver 
         Caption         =   "Volver"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   4920
         Picture         =   "frmTrabajo.frx":0C1E
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   7560
         Width           =   1125
      End
      Begin VB.CommandButton btnActualizar 
         Caption         =   "Terminar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   120
         Picture         =   "frmTrabajo.frx":0D30
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   7560
         Width           =   2475
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   6120
         Y1              =   7440
         Y2              =   7440
      End
      Begin VB.Label lblObsConex 
         Caption         =   "Observaciones finales"
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   6720
         Width           =   2175
      End
      Begin VB.Label lblObs 
         Caption         =   "Observaciones"
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Line linea 
         X1              =   120
         X2              =   6120
         Y1              =   6600
         Y2              =   6600
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   6120
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label lblCuadrilla 
         Caption         =   "Cuadrilla"
         Height          =   375
         Left            =   3120
         TabIndex        =   10
         Top             =   3960
         Width           =   2175
      End
      Begin VB.Label lblHora 
         Caption         =   "Hora de instalación"
         Height          =   375
         Left            =   3600
         TabIndex        =   9
         Top             =   4800
         Width           =   2415
      End
      Begin VB.Label lblFechaDeInstalacion 
         Caption         =   "Fecha de instalación"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   4800
         Width           =   2415
      End
      Begin VB.Label lblTelefono 
         Caption         =   "Teléfono"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label lblFechaDePedido 
         Caption         =   "Fecha de pedido"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label lblTipoDeConexion 
         Caption         =   "Tipo de conexión"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   3960
         Width           =   2175
      End
      Begin VB.Label lblNombreDeUsuario 
         Caption         =   "Cuenta de Internet"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label lblDomicilio 
         Caption         =   "Domicilio de instalación"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblNombreCompleto 
         Caption         =   "Apellido y nombre"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmTrabajo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private idTrabajo As Long

Private Sub Form_Load()

    Call cargarCuadrillas
    
    If main.tabTrabajos.Tab = 0 Then
        Call cargarFormProgramar
    ElseIf main.tabTrabajos.Tab = 1 Then
        Call cargarFormInstalar
    ElseIf main.tabTrabajos.Tab = 2 Then
        Call cargarFormTerminado
    End If
    
    Call cargarObs
    
    ' Ver si se puede cargar directamente en las props.
    cmbTipoConexion.ItemData(0) = 1
    cmbTipoConexion.ItemData(1) = 2
    cmbTipoConexion.ItemData(2) = 3
    cmbTipoConexion.ItemData(3) = 4
End Sub

Private Sub cargarFormProgramar()
    btnActualizar.Caption = "Ingresar"
    
    With main.tablaTrabajosAProgramar
        txtNombre = .TextMatrix(.Row, 0)
        txtDomi = .TextMatrix(.Row, 1)
        txtUsInternet = .TextMatrix(.Row, 2)
        txtTlfn = .TextMatrix(.Row, 5)
        txtFechaPedido = .TextMatrix(.Row, 4)
        cmbTipoConexion.Text = .TextMatrix(.Row, 3)
        idTrabajo = Val(.TextMatrix(.Row, 6))
        dtFechaInst.Value = DateTime.Now
        dtHoraInst = DateTime.Now
    End With
        
    btnImprimirOrden.Visible = False
    lblObsConex.Visible = False
    txtObsConex.Visible = False
    btnVolverAInstalar.Visible = False
    linea.Visible = False
    
    ' chkEnviarCorreoOrden.Value = 1
    
End Sub

Private Sub cargarFormInstalar()
    With main.tablaTrabajosAInstalar
        txtNombre = .TextMatrix(.Row, 1)
        txtDomi = .TextMatrix(.Row, 2)
        txtUsInternet = .TextMatrix(.Row, 3)
        txtTlfn = .TextMatrix(.Row, 6)
        txtFechaPedido = .TextMatrix(.Row, 5)
        cmbTipoConexion.Text = .TextMatrix(.Row, 4)
        idTrabajo = Val(.TextMatrix(.Row, 10))
        dtFechaInst = .TextMatrix(.Row, 7)
        dtHoraInst = .TextMatrix(.Row, 8)
        cmbCuadrilla.Text = main.tablaTrabajosAInstalar.TextMatrix(main.tablaTrabajosAInstalar.Row, 9)
    End With
    
    chkImprimirOrden.Visible = False
    chkEnviarCorreoOrden.Visible = False
End Sub

Private Sub cargarFormTerminado()
    btnActualizar.Enabled = False
    
    With main.tablaTrabajosTerminados
        txtNombre = .TextMatrix(.Row, 0)
        txtDomi = .TextMatrix(.Row, 1)
        txtUsInternet = .TextMatrix(.Row, 2)
        txtTlfn = .TextMatrix(.Row, 5)
        txtFechaPedido = .TextMatrix(.Row, 4)
        cmbTipoConexion.Text = .TextMatrix(.Row, 3)
        idTrabajo = Val(.TextMatrix(.Row, 9))
        dtFechaInst = .TextMatrix(.Row, 6)
        dtHoraInst = .TextMatrix(.Row, 7)
        cmbCuadrilla.Text = .TextMatrix(.Row, 8)
    End With
    
    cmbTipoConexion.Enabled = False
    cmbCuadrilla.Enabled = False
    txtObs.Enabled = False
    txtObsConex.Enabled = False
    dtHoraInst.Enabled = False
    dtFechaInst.Enabled = False
    
    chkImprimirOrden.Visible = False
    chkEnviarCorreoOrden.Visible = False
End Sub


Private Sub cargarObs()
    With main.vTrabInternet
        .IndexNumber = 0
        .FieldValue("id_trabajo") = idTrabajo
        .GetEqual
        
        If .status = 0 Then
            If Not IsNull(.FieldValue("obs")) Then
                txtObs.Text = .FieldValue("obs")
            End If
            If Not IsNull(.FieldValue("reserva")) Then
                txtObsConex.Text = .FieldValue("reserva")
            End If
        End If
    End With
End Sub

Private Sub dialogoImpresion(idTrabajo As Long)
    Dim copia As Integer
    Dim defPrinter As String
    defPrinter = Printer.DeviceName
  
    With frmTrabajo.cdImpresora
      .Flags = cdlPDNoSelection Or cdlPDHidePrintToFile Or cdlPDUseDevModeCopies
      On Error GoTo FinImpresion
      .ShowPrinter
      On Error GoTo 0
      
      For copia = 1 To .Copies
        Call imprimirOrden(idTrabajo)
      Next
      
    End With
    
    Call SetDefaultPrinter(defPrinter)
        
FinImpresion:
    On Error GoTo 0
End Sub

Private Sub btnActualizar_Click()

    Dim st As Integer
    
    If cmbCuadrilla.Text = vbNullString Then
        MsgBox "¡Recordá seleccionar una cuadrilla!", vbOKOnly + vbExclamation, "Datos incompletos"
    Else
    
        With main.vTrabInternet
            .IndexNumber = 0
            .FieldValue("id_trabajo") = idTrabajo
            
            st = .GetEqual
            
            If st = 0 Then
                .FieldValue("tipo_conexion") = cmbTipoConexion.ItemData(cmbTipoConexion.ListIndex)
                .FieldValue("fecha_inst") = dtFechaInst.Value
                .FieldValue("hora_inst") = dtHoraInst.Value
                .FieldValue("idcuadrilla") = cmbCuadrilla.ItemData(cmbCuadrilla.ListIndex)
                .FieldValue("estado") = IIf(main.tabTrabajos.Tab = 0, Estados.PROGRAMADO, Estados.TERMINADO)
                .FieldValue("obs") = txtObs.Text
                .FieldValue("reserva") = txtObsConex.Text
                .Update
            End If
            
            If main.tabTrabajos.Tab = 1 Then
                cambiarNoFacturar .FieldValue("nroOrden"), "SIFACTURAR"
            End If
            
        End With
        
        If chkImprimirOrden.Value = 1 Then
            Call dialogoImpresion(idTrabajo)
        End If
        
'        If chkEnviarCorreoOrden.Value = 1 Then
'            ProgressBar1.Value = 1
'            ProgressBar1.Visible = True
'            ' Timer1.Enabled = True
'            Call prepararCorreo(idTrabajo)
'            ProgressBar1.Visible = False
'            Timer1.Enabled = False
'        End If
     
        Unload Me
        
    End If

End Sub

Private Sub btnEliminar_Click()
    If MsgBox("Se eliminará este trabajo de la base de datos, ¿está seguro?", vbYesNo + vbQuestion, "Eliminar trabajo") = vbYes Then
    
        Dim st As Integer
        
        With main.vTrabInternet
            .FieldValue("id_trabajo") = idTrabajo
            st = .GetEqual
            If st = 0 Then
                .Delete
            End If
        End With
        Unload Me
    End If
End Sub

Private Sub btnVolverAInstalar_Click()
    With main.vTrabInternet
        .IndexNumber = 0
        .FieldValue("id_trabajo") = idTrabajo
        .GetEqual
        
        If .status = 0 Then
            .FieldValue("estado") = Estados.NUEVO
            .Update
            
            cambiarNoFacturar .FieldValue("nroOrden"), "NOFACTURAR"
        End If
    End With
    
    Unload Me
End Sub

Private Sub btnVolver_Click()
    Unload Me
End Sub

Private Sub cargarCuadrillas()
    Dim status As Integer
    
    With main.VCuadrillas
        .IndexNumber = 0
        status = .GetFirst
        
        While status = 0
            If .FieldValue("habilitado") = 1 Or main.tabTrabajos.Tab = 1 Then
                cmbCuadrilla.AddItem (.FieldValue("miembros"))
                cmbCuadrilla.ItemData(cmbCuadrilla.NewIndex) = .FieldValue("idcuadrilla")
            End If
            
            status = .GetNext
        Wend
    End With
End Sub

Private Sub btnImprimirOrden_Click()
    Call dialogoImpresion(idTrabajo)
End Sub

Private Sub Timer1_Timer()
    If ProgressBar1.Value = ProgressBar1.Max Then
        ProgressBar1.Value = ProgressBar1.Min
    End If
    
    ProgressBar1.Value = ProgressBar1.Value + 1
End Sub
