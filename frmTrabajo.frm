VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTrabajo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos del trabajo"
   ClientHeight    =   10650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9210
   Icon            =   "frmTrabajo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10650
   ScaleWidth      =   9210
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8760
      Top             =   10320
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   315
      Left            =   360
      TabIndex        =   18
      Top             =   10320
      Visible         =   0   'False
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
      Max             =   50
   End
   Begin VB.CommandButton btnVolverAInstalar 
      BackColor       =   &H00CCF2FF&
      Caption         =   "Pasar a ""Para programar"""
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "La orden vuelve a la primera solapa del sistema"
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Frame frmDatos 
      BorderStyle     =   0  'None
      Height          =   10275
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   8895
      Begin MSComDlg.CommonDialog cdImpresora 
         Left            =   8040
         Top             =   5280
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame frmBotones 
         BorderStyle     =   0  'None
         Caption         =   "Acciones"
         Height          =   1095
         Left            =   240
         TabIndex        =   42
         Top             =   9120
         Width           =   8415
         Begin VB.CommandButton btnActualizar 
            BackColor       =   &H00DAEFE2&
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
            Left            =   0
            Picture         =   "frmTrabajo.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   0
            Width           =   2475
         End
         Begin VB.CommandButton btnVolver 
            BackColor       =   &H00D9D9D9&
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
            Left            =   7200
            Picture         =   "frmTrabajo.frx":0BD4
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   0
            Width           =   1125
         End
         Begin VB.CommandButton btnEliminar 
            BackColor       =   &H00ADCBF8&
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
            Left            =   5520
            Picture         =   "frmTrabajo.frx":0CE6
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   0
            Width           =   1245
         End
         Begin VB.CheckBox chkImprimirOrden 
            Caption         =   "Imprimir orden de trabajo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2640
            TabIndex        =   10
            Top             =   0
            Width           =   2775
         End
         Begin VB.CheckBox chkEnviarCorreoOrden 
            Caption         =   "Enviar orden por correo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2640
            TabIndex        =   11
            Top             =   480
            Visible         =   0   'False
            Width           =   2775
         End
      End
      Begin VB.CommandButton btnImprimirOrden 
         BackColor       =   &H00E6C29B&
         Caption         =   "Imprimir orden de trabajo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   5040
         Width           =   2175
      End
      Begin VB.Frame frmDatosTrabajo 
         Caption         =   "Datos del trabajo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   240
         TabIndex        =   36
         Top             =   3120
         Width           =   8415
         Begin VB.ComboBox cmbCuadrilla 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3120
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   1320
            Width           =   2775
         End
         Begin VB.ComboBox cmbTipoConexion 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "frmTrabajo.frx":15B0
            Left            =   240
            List            =   "frmTrabajo.frx":15C0
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1320
            Width           =   2775
         End
         Begin VB.TextBox txtObs 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   240
            MaxLength       =   50
            TabIndex        =   1
            Top             =   600
            Width           =   7935
         End
         Begin MSComCtl2.DTPicker dtHoraInst 
            Height          =   375
            Left            =   3120
            TabIndex        =   4
            Top             =   2040
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "hh:mm tt"
            Format          =   81592323
            UpDown          =   -1  'True
            CurrentDate     =   44076
         End
         Begin MSComCtl2.DTPicker dtFechaInst 
            Height          =   375
            Left            =   240
            TabIndex        =   3
            Top             =   2040
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   81592321
            CurrentDate     =   44076
         End
         Begin VB.Label lblTipoDeConexion 
            Caption         =   "Tipo de conexión"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   41
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label lblFechaDeInstalacion 
            Caption         =   "Fecha de instalación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   1800
            Width           =   2775
         End
         Begin VB.Label lblHora 
            Caption         =   "Hora de instalación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3120
            TabIndex        =   39
            Top             =   1800
            Width           =   2655
         End
         Begin VB.Label lblCuadrilla 
            Caption         =   "Cuadrilla"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3120
            TabIndex        =   38
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label lblObs 
            Caption         =   "Observaciones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   37
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame frmDatosConexion 
         Caption         =   "Datos de la conexión"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   240
         TabIndex        =   30
         Top             =   6120
         Width           =   8415
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
            TabIndex        =   9
            Top             =   2040
            Width           =   7935
         End
         Begin VB.ComboBox cmbTarifas 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   240
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   600
            Width           =   3735
         End
         Begin VB.TextBox txtDirMAC 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4200
            TabIndex        =   6
            Top             =   600
            Width           =   3975
         End
         Begin VB.TextBox txtUbFis 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   7
            Top             =   1320
            Width           =   3735
         End
         Begin VB.TextBox txtUbLog 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4200
            TabIndex        =   8
            Top             =   1320
            Width           =   3975
         End
         Begin VB.Label lblObsConex 
            Caption         =   "Observaciones finales"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Top             =   1800
            Width           =   2175
         End
         Begin VB.Label lblTarifas 
            Caption         =   "Nueva tarifa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblDirMAC 
            Caption         =   "Dirección MAC"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   33
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label lblUbFis 
            Caption         =   "Ubicación física"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label lblUbLog 
            Caption         =   "Ubicación lógica"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   31
            Top             =   1080
            Width           =   1575
         End
      End
      Begin VB.Frame frmDatosAbonado 
         Caption         =   "Datos del abonado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   8415
         Begin VB.TextBox txtNombre 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   360
            Width           =   5295
         End
         Begin VB.TextBox txtDomi 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   840
            Width           =   5295
         End
         Begin VB.TextBox txtUsInternet 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   1320
            Width           =   5295
         End
         Begin VB.TextBox txtTlfn 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   1800
            Width           =   5295
         End
         Begin VB.TextBox txtFechaPedido 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   2280
            Width           =   5295
         End
         Begin VB.Label lblNombreCompleto 
            Caption         =   "Apellido y nombre"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   29
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label lblDomicilio 
            Caption         =   "Domicilio de instalación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   28
            Top             =   840
            Width           =   2535
         End
         Begin VB.Label lblNombreDeUsuario 
            Caption         =   "Cuenta de Internet"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   27
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label lblFechaDePedido 
            Caption         =   "Fecha de pedido"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   26
            Top             =   2280
            Width           =   1815
         End
         Begin VB.Label lblTelefono 
            Caption         =   "Teléfono"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   25
            Top             =   1800
            Width           =   1815
         End
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
    Call cargarTarifasTodas(cmbTarifas)
    
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
    
    ' Modificar dimensiones porque sino queda mucho espacio vacío
    Me.Height = 8000
    frmDatos.Height = 7400
    frmBotones.top = 6100
    
    btnImprimirOrden.Visible = False
    btnVolverAInstalar.Visible = False
    frmDatosConexion.Visible = False
    
    ' chkEnviarCorreoOrden.Value = 1
    
End Sub

Private Sub cargarFormInstalar()
    With main.tablaTrabajosAInstalar
        idTrabajo = Val(.TextMatrix(.Row, 10))
        
        txtNombre = .TextMatrix(.Row, 1)
        txtDomi = .TextMatrix(.Row, 2)
        txtUsInternet = .TextMatrix(.Row, 3)
        txtTlfn = .TextMatrix(.Row, 6)
        txtFechaPedido = .TextMatrix(.Row, 5)
        cmbTipoConexion.Text = .TextMatrix(.Row, 4)
        dtFechaInst = .TextMatrix(.Row, 7)
        dtHoraInst = .TextMatrix(.Row, 8)
        cmbCuadrilla.Text = main.tablaTrabajosAInstalar.TextMatrix(main.tablaTrabajosAInstalar.Row, 9)
    End With
    
    ' Seleccionar en cmbTarifas la tarifa que ya tiene asignado el trabajo.
    Call seleccionarPorItemData(getIdTarifa(idTrabajo), cmbTarifas)
    
    chkImprimirOrden.Visible = False
    chkEnviarCorreoOrden.Visible = False
End Sub

Private Sub cargarFormTerminado()
    btnActualizar.Enabled = False
    
    With main.tablaTrabajosTerminados
        idTrabajo = Val(.TextMatrix(.Row, 9))
        
        txtNombre = .TextMatrix(.Row, 0)
        txtDomi = .TextMatrix(.Row, 1)
        txtUsInternet = .TextMatrix(.Row, 2)
        txtTlfn = .TextMatrix(.Row, 5)
        txtFechaPedido = .TextMatrix(.Row, 4)
        cmbTipoConexion.Text = .TextMatrix(.Row, 3)
        dtFechaInst = .TextMatrix(.Row, 6)
        dtHoraInst = .TextMatrix(.Row, 7)
        cmbCuadrilla.Text = .TextMatrix(.Row, 8)
    End With
    
    ' Seleccionar en cmbTarifas la tarifa que ya tiene asignado el trabajo.
    Call seleccionarPorItemData(getIdTarifa(idTrabajo), cmbTarifas)
    
    With main
        .vTrabInternet.IndexNumber = 0
        .vTrabInternet.FieldValue("id_trabajo") = idTrabajo
        
        If .vTrabInternet.GetEqual = 0 Then
            Dim codAlumbrado As Long
            codAlumbrado = getCodAlumbrado(.vTrabInternet.FieldValue("NroOrden"))
            
            .VDatosConexInet.IndexNumber = 0
            .VDatosConexInet.FieldValue("codAlumbrado") = codAlumbrado
            
            If .VDatosConexInet.GetEqual = 0 Then
                txtDirMAC.Text = .VDatosConexInet.FieldValue("direc_MAC") & vbNullString
                txtUbFis.Text = .VDatosConexInet.FieldValue("ubic_fisica") & vbNullString
                txtUbLog.Text = .VDatosConexInet.FieldValue("ubic_logica") & vbNullString
            End If
            
        End If
    End With
    
    frmDatosAbonado.Enabled = False
    frmDatosTrabajo.Enabled = False
    frmDatosConexion.Enabled = False
    
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
            
            If .GetEqual = 0 Then
                .FieldValue("tipo_conexion") = cmbTipoConexion.ItemData(cmbTipoConexion.ListIndex)
                .FieldValue("fecha_inst") = dtFechaInst.Value
                .FieldValue("hora_inst") = dtHoraInst.Value
                .FieldValue("idcuadrilla") = cmbCuadrilla.ItemData(cmbCuadrilla.ListIndex)
                .FieldValue("estado") = IIf(main.tabTrabajos.Tab = 0, Estados.PROGRAMADO, Estados.TERMINADO)
                .FieldValue("obs") = txtObs.Text
                .FieldValue("reserva") = txtObsConex.Text
                
                If main.tabTrabajos.Tab = 1 Then ' Si el trabajo pasa a terminado entonces...
                    .FieldValue("ancho_banda") = cmbTarifas.ItemData(cmbTarifas.ListIndex)
                    Call cambiarNoFacturar(.FieldValue("nroOrden"), "SIFACTURAR")
                    Call actualizarTarifa(.FieldValue("nroOrden"), .FieldValue("ancho_banda"))
                    Call actualizarDatosConexInet(getCodAlumbrado(.FieldValue("nroOrden")))
                End If
                
                .Update
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

Private Sub actualizarDatosConexInet(codAlumbrado As Long)
    With main.VDatosConexInet
        .IndexNumber = 0
        .FieldValue("CodAlumbrado") = codAlumbrado
        
        If .GetEqual = 0 Then
            ' Hay que actualizar
            .FieldValue("direc_MAC") = txtDirMAC.Text
            .FieldValue("ubic_fisic") = txtUbFis.Text
            .FieldValue("ubic_logic") = txtUbLog.Text
            .Update
        Else
            ' Hay que agregarlo nuevo
            .FieldValue("CodAlumbrado") = codAlumbrado
            .FieldValue("direc_MAC") = txtDirMAC.Text
            .FieldValue("ubic_fisica") = txtUbFis.Text
            .FieldValue("ubic_logica") = txtUbLog.Text
            .Insert
        End If
    End With
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
