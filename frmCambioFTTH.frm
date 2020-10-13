VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCambioFTTH 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio a FTTH"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6405
   Icon            =   "frmCambioFTTH.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleMode       =   0  'User
   ScaleWidth      =   6405
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdVolver 
      BackColor       =   &H00D9D9D9&
      Caption         =   "&Volver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3495
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton cmdGuardarTrabajo 
      BackColor       =   &H00DAEFE2&
      Caption         =   "&Guardar trabajo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1095
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4080
      Width           =   2055
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
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   6135
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
         Left            =   2040
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         ToolTipText     =   "Si no selecciona nada, se mantendrá la tarifa anterior"
         Top             =   600
         Width           =   3975
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
         TabIndex        =   9
         Top             =   1320
         Width           =   5775
      End
      Begin MSComCtl2.DTPicker dtpInstInt 
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   95748097
         CurrentDate     =   44083
      End
      Begin VB.Label lblAnchoDeBanda 
         Caption         =   "Ancho de banda a instalar"
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
         Left            =   2040
         TabIndex        =   13
         Top             =   360
         Width           =   3855
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
         TabIndex        =   5
         Top             =   1080
         Width           =   1935
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
         TabIndex        =   4
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame frmSelCliente 
      Caption         =   "Seleccionar usuario y orden"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.CommandButton cmdSelCli 
         Height          =   732
         Left            =   240
         Picture         =   "frmCambioFTTH.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   480
         Width           =   945
      End
      Begin VB.Label lblDomicilio 
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
         Left            =   1320
         TabIndex        =   11
         Top             =   1080
         Width           =   4695
      End
      Begin VB.Label lblNombre 
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
         Left            =   1320
         TabIndex        =   10
         Top             =   720
         Width           =   4695
      End
      Begin VB.Label lblCodCli 
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
         Left            =   3720
         TabIndex        =   2
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblCodInternet 
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
         Left            =   1320
         TabIndex        =   1
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
            .FieldValue("tipo_conexion") = 4
            .FieldValue("ancho_banda") = cmbTarifas.ItemData(cmbTarifas.ListIndex) ' idTarifa
            .FieldValue("obs") = txtObs.Text
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

