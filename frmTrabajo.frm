VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmTrabajo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos del trabajo"
   ClientHeight    =   10245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9210
   Icon            =   "frmTrabajo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10245
   ScaleWidth      =   9210
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame frmDatos 
      BorderStyle     =   0  'None
      Height          =   10155
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   8895
      Begin VB.Frame frmBotones 
         BorderStyle     =   0  'None
         Caption         =   "Acciones"
         Height          =   1215
         Left            =   180
         TabIndex        =   39
         Top             =   9060
         Width           =   8535
         Begin VB.CommandButton btnSinTerminar 
            BackColor       =   &H00CCF2FF&
            Caption         =   "Sin terminar"
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
            Left            =   3540
            Picture         =   "frmTrabajo.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   0
            Visible         =   0   'False
            Width           =   1600
         End
         Begin VB.CommandButton btnModificar 
            BackColor       =   &H00F2E1D9&
            Caption         =   "Actualizar"
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
            Left            =   1800
            Picture         =   "frmTrabajo.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   0
            Width           =   1600
         End
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
            Left            =   60
            Picture         =   "frmTrabajo.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   0
            Width           =   1600
         End
         Begin VB.CommandButton btnVolver 
            BackColor       =   &H00D9D9D9&
            Cancel          =   -1  'True
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
            Left            =   7320
            Picture         =   "frmTrabajo.frx":11E8
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
            Left            =   5880
            Picture         =   "frmTrabajo.frx":12FA
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   0
            Width           =   1245
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
            Height          =   555
            Left            =   3540
            TabIndex        =   10
            Top             =   0
            Width           =   2175
         End
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
            Height          =   555
            Left            =   3540
            TabIndex        =   11
            Top             =   540
            Visible         =   0   'False
            Width           =   2295
         End
      End
      Begin VB.Frame frmDatosTrabajo 
         Caption         =   "Datos del trabajo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2955
         Left            =   240
         TabIndex        =   33
         Top             =   2880
         Width           =   8415
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
            Left            =   6000
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "La orden vuelve a la primera solapa del sistema"
            Top             =   1260
            Width           =   2175
         End
         Begin VB.CommandButton btnImprimirOrden 
            BackColor       =   &H00E6C29B&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   6000
            Picture         =   "frmTrabajo.frx":1BC4
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   1980
            Width           =   975
         End
         Begin VB.CommandButton btnPDFTrabajo 
            BackColor       =   &H00E6C29B&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   7200
            Picture         =   "frmTrabajo.frx":77D6
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   1980
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.ComboBox cmbPrioridad 
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
            Left            =   3600
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   2400
            Width           =   2295
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
            Top             =   1560
            Width           =   2775
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
            ItemData        =   "frmTrabajo.frx":7AE0
            Left            =   240
            List            =   "frmTrabajo.frx":7AF0
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1560
            Width           =   2775
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
            Height          =   405
            Left            =   240
            MaxLength       =   50
            TabIndex        =   1
            Top             =   720
            Width           =   7935
         End
         Begin MSComCtl2.DTPicker dtHoraInst 
            Height          =   375
            Left            =   2040
            TabIndex        =   4
            Top             =   2400
            Width           =   1455
            _ExtentX        =   2566
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
            Format          =   95617027
            UpDown          =   -1  'True
            CurrentDate     =   44076
         End
         Begin MSComCtl2.DTPicker dtFechaInst 
            Height          =   375
            Left            =   240
            TabIndex        =   3
            Top             =   2400
            Width           =   1695
            _ExtentX        =   2990
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
            Format          =   95617025
            CurrentDate     =   44076
         End
         Begin MSComDlg.CommonDialog cdImpresora 
            Left            =   6600
            Top             =   2400
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label1 
            Caption         =   "Prioridad"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3600
            TabIndex        =   40
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label lblTipoDeConexion 
            Caption         =   "Tipo de conexi�n"
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
            TabIndex        =   38
            Top             =   1200
            Width           =   2175
         End
         Begin VB.Label lblFechaDeInstalacion 
            Caption         =   "Fecha de inst."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Top             =   2040
            Width           =   1575
         End
         Begin VB.Label lblHora 
            Caption         =   "Hora de inst."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   36
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label lblCuadrilla 
            Caption         =   "Cuadrilla"
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
            Left            =   3120
            TabIndex        =   35
            Top             =   1200
            Width           =   2175
         End
         Begin VB.Label lblObs 
            Caption         =   "Observaciones"
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
            TabIndex        =   34
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame frmDatosConexion 
         Caption         =   "Datos de la conexi�n"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2955
         Left            =   240
         TabIndex        =   27
         Top             =   5880
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
            Top             =   2400
            Width           =   7935
         End
         Begin VB.ComboBox cmbTarifas 
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
            Left            =   240
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   720
            Width           =   3735
         End
         Begin VB.TextBox txtDirMAC 
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
            Left            =   4200
            TabIndex        =   6
            Top             =   720
            Width           =   3975
         End
         Begin VB.TextBox txtUbFis 
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
            TabIndex        =   7
            Top             =   1560
            Width           =   3735
         End
         Begin VB.TextBox txtUbLog 
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
            Left            =   4200
            TabIndex        =   8
            Top             =   1560
            Width           =   3975
         End
         Begin VB.Label lblObsConex 
            Caption         =   "Observaciones finales"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   2040
            Width           =   2175
         End
         Begin VB.Label lblTarifas 
            Caption         =   "Nueva tarifa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label lblDirMAC 
            Caption         =   "Direcci�n MAC"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   30
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label lblUbFis 
            Caption         =   "Ubicaci�n f�sica"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   1200
            Width           =   2295
         End
         Begin VB.Label lblUbLog 
            Caption         =   "Ubicaci�n l�gica"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   28
            Top             =   1200
            Width           =   2295
         End
      End
      Begin VB.Frame frmDatosAbonado 
         Caption         =   "Datos del abonado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   240
         TabIndex        =   16
         Top             =   60
         Width           =   8415
         Begin VB.TextBox txtNombre 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   360
            Width           =   5295
         End
         Begin VB.TextBox txtDomi 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   840
            Width           =   5295
         End
         Begin VB.TextBox txtUsInternet 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   1320
            Width           =   5295
         End
         Begin VB.TextBox txtTlfn 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   1800
            Width           =   5295
         End
         Begin VB.TextBox txtFechaPedido 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   2280
            Width           =   5295
         End
         Begin VB.Label lblNombreCompleto 
            Caption         =   "Apellido y nombre"
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
            TabIndex        =   26
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label lblDomicilio 
            Caption         =   "Domicilio de instalaci�n"
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
            TabIndex        =   25
            Top             =   840
            Width           =   2535
         End
         Begin VB.Label lblNombreDeUsuario 
            Caption         =   "Cuenta de Internet"
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
            TabIndex        =   24
            Top             =   1320
            Width           =   2295
         End
         Begin VB.Label lblFechaDePedido 
            Caption         =   "Fecha de pedido"
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
            TabIndex        =   23
            Top             =   2280
            Width           =   1815
         End
         Begin VB.Label lblTelefono 
            Caption         =   "Tel�fono"
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
            TabIndex        =   22
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
  Call cargarComboPrioridad(cmbPrioridad)

  If main.tabTrabajos.Tab = 0 Then
    Call cargarFormProgramar
  ElseIf main.tabTrabajos.Tab = 1 Then
    Call cargarFormInstalar
  ElseIf main.tabTrabajos.Tab = 2 Then
    Call cargarFormTerminado
  ElseIf main.tabTrabajos.Tab = 3 Then
    Call cargarFormSinTerminar
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
    idTrabajo = Val(.TextMatrix(.Row, 6))

    txtNombre = .TextMatrix(.Row, 0)
    txtDomi = .TextMatrix(.Row, 1)
    txtUsInternet = .TextMatrix(.Row, 2)
    txtTlfn = .TextMatrix(.Row, 5)
    txtFechaPedido = .TextMatrix(.Row, 4)
    cmbTipoConexion.Text = .TextMatrix(.Row, 3)
  End With

  ' Otros datos que no pueden sacarse de la grilla
  Call cargarDatosExtrasFormProgramar(idTrabajo)

  ' Modificar dimensiones porque sino queda mucho espacio vac�o
  Me.Height = 7800
  frmDatos.Height = 8000
  frmBotones.top = 6100

  btnImprimirOrden.Visible = False
  btnPDFTrabajo.Visible = False
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
    cmbCuadrilla.Text = .TextMatrix(.Row, 9)
  End With

  Call seleccionarPrioridad
  
  ' Seleccionar en cmbTarifas la tarifa que ya tiene asignado el trabajo.
  Call seleccionarPorItemData(getIdTarifa(idTrabajo), cmbTarifas)
  Call cargarDatosConexion

  chkImprimirOrden.Visible = False
  chkEnviarCorreoOrden.Visible = False
  btnSinTerminar.Visible = True
End Sub

Private Sub cargarFormTerminado()
  btnActualizar.Enabled = False
  btnModificar.Enabled = False

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

  Call seleccionarPrioridad

  ' Seleccionar en cmbTarifas la tarifa que ya tiene asignado el trabajo.
  Call seleccionarPorItemData(getIdTarifa(idTrabajo), cmbTarifas)
  Call cargarDatosConexion

  frmDatosAbonado.Enabled = False
  frmDatosTrabajo.Enabled = False
  frmDatosConexion.Enabled = False
  
  btnVolverAInstalar.Enabled = False
  btnImprimirOrden.Enabled = False

  chkImprimirOrden.Visible = False
  chkEnviarCorreoOrden.Visible = False
End Sub

Private Sub cargarFormSinTerminar()
  With main.tablaTrabajosSinTerminar
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

  Call seleccionarPrioridad
  
  ' Seleccionar en cmbTarifas la tarifa que ya tiene asignado el trabajo.
  Call seleccionarPorItemData(getIdTarifa(idTrabajo), cmbTarifas)
  Call cargarDatosConexion

  chkImprimirOrden.Visible = False
  chkEnviarCorreoOrden.Visible = False
End Sub

Private Sub cargarDatosConexion()
  With main
    .vTrabInternet.IndexNumber = 0
    .vTrabInternet.FieldValue("id_trabajo") = idTrabajo

    If .vTrabInternet.GetEqual = 0 Then
      Dim CodAlumbrado As Long
      CodAlumbrado = getCodAlumbrado(.vTrabInternet.FieldValue("NroOrden"))

      .VDatosConexInet.IndexNumber = 0
      .VDatosConexInet.FieldValue("codAlumbrado") = CodAlumbrado

      If .VDatosConexInet.GetEqual = 0 Then
        txtDirMAC.Text = .VDatosConexInet.FieldValue("direc_MAC") & vbNullString
        txtUbFis.Text = .VDatosConexInet.FieldValue("ubic_fisica") & vbNullString
        txtUbLog.Text = .VDatosConexInet.FieldValue("ubic_logica") & vbNullString
      End If

    End If
  End With
End Sub

Private Sub cargarDatosExtrasFormProgramar(idTrabajo As Long)
  With main.vTrabInternet
    .IndexNumber = 0
    .FieldValue("id_trabajo") = idTrabajo
    If .GetEqual = 0 Then
      ' Cargar fecha y hora
      dtFechaInst.Value = IIf(.FieldValue("fecha_inst") = 0, DateTime.Now, .FieldValue("fecha_inst"))
      dtHoraInst = IIf(IsNull(.FieldValue("hora_inst")), DateTime.Now, .FieldValue("hora_inst"))
      
      ' Seleccionar la cuadrilla en caso de que se haya modificado antes
      If Not (IsNull(.FieldValue("idCuadrilla"))) Then
        Call seleccionarPorItemData(.FieldValue("idCuadrilla"), cmbCuadrilla)
      End If
      
      ' Cargar la prioridad del trabajo
      Call seleccionarPrioridad
      
    End If
  End With
End Sub

Private Sub seleccionarPrioridad()
  With main.vTrabInternet
    .IndexNumber = 0
    .FieldValue("id_trabajo") = idTrabajo
    .GetEqual

    Dim prioridad As Long
    prioridad = IIf(IsNull(.FieldValue("prioridad")), 0, .FieldValue("prioridad"))

    Call seleccionarPorItemData(prioridad, cmbPrioridad)
  End With

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

Private Sub btnActualizar_Click()

  Dim st As Integer

  If cmbCuadrilla.Text = vbNullString Then
    MsgBox "�Record� seleccionar una cuadrilla!", vbOKOnly + vbExclamation, "Datos incompletos"
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
        .FieldValue("prioridad") = cmbPrioridad.ItemData(cmbPrioridad.ListIndex)

        If main.tabTrabajos.Tab = 1 Then  ' Si el trabajo pasa a terminado entonces...
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

Private Sub btnModificar_Click()
  With main.vTrabInternet
    .IndexNumber = 0
    .FieldValue("id_trabajo") = idTrabajo

    If .GetEqual = 0 Then
    
        If Not (cmbCuadrilla.Text = vbNullString) Then
          .FieldValue("idcuadrilla") = cmbCuadrilla.ItemData(cmbCuadrilla.ListIndex)
        End If
    
        .FieldValue("tipo_conexion") = cmbTipoConexion.ItemData(cmbTipoConexion.ListIndex)
        .FieldValue("fecha_inst") = dtFechaInst.Value
        .FieldValue("hora_inst") = dtHoraInst.Value

        .FieldValue("obs") = txtObs.Text
        .FieldValue("reserva") = txtObsConex.Text
        .FieldValue("prioridad") = cmbPrioridad.ItemData(cmbPrioridad.ListIndex)
  
        If Not (main.tabTrabajos.Tab = 0) Then
          .FieldValue("ancho_banda") = cmbTarifas.ItemData(cmbTarifas.ListIndex)
          Call actualizarDatosConexInet(getCodAlumbrado(.FieldValue("nroOrden")))
        End If
  
        .Update
    End If
  End With
  
  Unload Me
End Sub

Private Sub btnSinTerminar_Click()
  ' Este bot�n solo deber�a accederse por los trabajos en estado PROGRAMADO
  With main.vTrabInternet
    .IndexNumber = 0
    .FieldValue("id_trabajo") = idTrabajo

    If .GetEqual = 0 Then
    
        .FieldValue("estado") = Estados.SIN_TERMINAR
    
        .FieldValue("idcuadrilla") = cmbCuadrilla.ItemData(cmbCuadrilla.ListIndex)
        .FieldValue("tipo_conexion") = cmbTipoConexion.ItemData(cmbTipoConexion.ListIndex)
        .FieldValue("fecha_inst") = dtFechaInst.Value
        .FieldValue("hora_inst") = dtHoraInst.Value

        .FieldValue("obs") = txtObs.Text
        .FieldValue("reserva") = txtObsConex.Text
        .FieldValue("prioridad") = cmbPrioridad.ItemData(cmbPrioridad.ListIndex)
        .FieldValue("ancho_banda") = cmbTarifas.ItemData(cmbTarifas.ListIndex)
  
        .Update
        
        Call actualizarDatosConexInet(getCodAlumbrado(.FieldValue("nroOrden")))
    End If
  End With
  
  Unload Me
End Sub

Private Sub actualizarDatosConexInet(CodAlumbrado As Long)
  With main.VDatosConexInet
    .IndexNumber = 0
    .FieldValue("CodAlumbrado") = CodAlumbrado

    If .GetEqual = 0 Then
      ' Hay que actualizar
      .FieldValue("direc_MAC") = txtDirMAC.Text
      .FieldValue("ubic_fisica") = txtUbFis.Text
      .FieldValue("ubic_logica") = txtUbLog.Text
      .Update
    Else
      ' Hay que agregarlo nuevo
      .Clear
      .FieldValue("CodAlumbrado") = CodAlumbrado
      .FieldValue("direc_MAC") = txtDirMAC.Text
      .FieldValue("ubic_fisica") = txtUbFis.Text
      .FieldValue("ubic_logica") = txtUbLog.Text
      .Insert
    End If
  End With
End Sub

Private Sub btnEliminar_Click()
  If MsgBox("Se eliminar� este trabajo de la base de datos, �est� seguro?", vbYesNo + vbQuestion, "Eliminar trabajo") = vbYes Then

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

Private Sub btnPDFTrabajo_Click()
  On Error GoTo algoSalioMal
  
  Dim rutaCompleta As String
  rutaCompleta = prepararPDFOrden(idTrabajo)
  
  Shell "explorer.exe /select, " & rutaCompleta, vbNormalFocus
  Exit Sub
  
algoSalioMal:
  MsgBox "Hubo un problema al generar el .pdf. Cont�ctenos para que lo ayudemos.", vbCritical, "No se puede generar el .pdf"
  Exit Sub
End Sub

Private Sub dialogoImpresion(idTrabajo As Long)
  Dim copia As Integer
  Dim defPrinter As String
  defPrinter = Printer.DeviceName

  With frmTrabajo.cdImpresora
    .Flags = cdlPDNoSelection Or cdlPDHidePrintToFile Or cdlPDUseDevModeCopies
    On Error GoTo FinImpresion
    .ShowPrinter
    On Error GoTo ErrorImpresion

    For copia = 1 To .Copies
      Sleep 1000
      Call imprimirOrden(idTrabajo)
    Next

  End With

  Call SetDefaultPrinter(defPrinter)
  Exit Sub

FinImpresion:
  On Error GoTo 0
  Exit Sub
ErrorImpresion:
  MsgBox "Algo sali� mal al imprimir el documento", vbCritical, "Error"
  Exit Sub
End Sub

