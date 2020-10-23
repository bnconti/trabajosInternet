VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCambioFTTH 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio a FTTH"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   Icon            =   "frmCambioFTTH.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6355.141
   ScaleMode       =   0  'User
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdVolver 
      BackColor       =   &H00D9D9D9&
      Caption         =   "&Volver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3645
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5640
      Width           =   2040
   End
   Begin VB.CommandButton cmdGuardarTrabajo 
      BackColor       =   &H00DAEFE2&
      Caption         =   "&Guardar trabajo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1125
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5640
      Width           =   2175
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
      Height          =   3135
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   6495
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
         ItemData        =   "frmCambioFTTH.frx":08CA
         Left            =   2760
         List            =   "frmCambioFTTH.frx":08CC
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Si no selecciona nada, se mantendrá la tarifa anterior"
         Top             =   720
         Width           =   3495
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
         TabIndex        =   2
         ToolTipText     =   "Si no selecciona nada, se mantendrá la tarifa anterior"
         Top             =   1680
         Width           =   6015
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
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   2640
         Width           =   6015
      End
      Begin MSComCtl2.DTPicker dtpInstInt 
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   720
         Width           =   2295
         _ExtentX        =   4048
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
         Format          =   109510657
         CurrentDate     =   44083
      End
      Begin VB.Label lblPrioridad 
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
         Left            =   2760
         TabIndex        =   15
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblAnchoDeBanda 
         Caption         =   "Ancho de banda a instalar"
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
         TabIndex        =   14
         Top             =   1320
         Width           =   2895
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
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   2280
         Width           =   1695
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
         TabIndex        =   10
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame frmSelCliente 
      Caption         =   "Seleccionar usuario y orden"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6495
      Begin VB.CommandButton cmdSelCli 
         Height          =   732
         Left            =   240
         Picture         =   "frmCambioFTTH.frx":08CE
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   600
         Width           =   945
      End
      Begin VB.Label lblTelefono 
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
         Left            =   1320
         TabIndex        =   17
         Top             =   1440
         Width           =   4695
      End
      Begin VB.Label lblDomicilio 
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
         Left            =   1320
         TabIndex        =   13
         Top             =   1080
         Width           =   4695
      End
      Begin VB.Label lblNombre 
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
         Left            =   1320
         TabIndex        =   12
         Top             =   720
         Width           =   4695
      End
      Begin VB.Label lblCodCli 
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
         Left            =   3960
         TabIndex        =   8
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblCodInternet 
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
         Left            =   1320
         TabIndex        =   7
         Top             =   360
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmCambioFTTH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const FTTH_20MB As Integer = 1002

Private mNroOrden As Long

Private Sub Form_Load()
    dtpInstInt = DateTime.Now
    Call cargarTarifasFTTH(cmbTarifas)
    Call seleccionarPorItemData(FTTH_20MB, cmbTarifas)
    Call cargarComboPrioridad(cmbPrioridad)
End Sub

Property Get nroOrden() As Long
    nroOrden = mNroOrden
End Property

Property Let nroOrden(NroOrdenNuevo As Long)
    mNroOrden = NroOrdenNuevo
End Property

Private Sub cmdGuardarTrabajo_Click()
    If lblCodInternet.Caption = vbNullString Then
        MsgBox "Seleccione un usuario antes de generar la orden.", vbOKOnly + vbInformation, "Faltan datos"
    Else
        With main.vTrabInternet
            .Clear
            .FieldValue("nroOrden") = mNroOrden
            .FieldValue("estado") = Estados.NUEVO
            .FieldValue("fecha_pedido") = dtpInstInt.Value
            .FieldValue("tipo_conexion") = 4 ' Cambio a FTTH
            .FieldValue("ancho_banda") = cmbTarifas.ItemData(cmbTarifas.ListIndex) ' idTarifa
            .FieldValue("obs") = txtObs.Text
            .FieldValue("prioridad") = cmbPrioridad.ItemData(cmbPrioridad.ListIndex)
            If .Insert = 0 Then
                Call cerrar
            Else
                Call MsgBox("Hubo un problema al guardar el trabajo.", vbCritical, "Resultado incorrecto")
            End If
        End With
    End If
End Sub

Private Sub cmdSelCli_Click()
    frmSelCli.Show 1, Me
End Sub

Private Sub cerrar()
    Unload Me
End Sub

Private Sub cmdVolver_Click()
    Call cerrar
End Sub

