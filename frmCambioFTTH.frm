VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCambioFTTH 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio a FTTH"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   Icon            =   "frmCambioFTTH.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleMode       =   0  'User
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdVolver 
      Caption         =   "&Volver"
      Height          =   495
      Left            =   3000
      TabIndex        =   7
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton cmdGuardarTrabajo 
      Caption         =   "&Guardar trabajo"
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Frame frmDatosTrabajo 
      Caption         =   "Datos del trabajo"
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   5415
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
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   5055
      End
      Begin MSComCtl2.DTPicker dtpInstInt 
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
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
         CurrentDate     =   44083
      End
      Begin VB.Label lblObs 
         Caption         =   "Observaciones"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblFechaDePedido 
         Caption         =   "Fecha de pedido"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame frmSelCliente 
      Caption         =   "Seleccionar usuario y orden"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
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
            Size            =   12
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
         Width           =   3975
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
         TabIndex        =   10
         Top             =   720
         Width           =   3975
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
         Left            =   3720
         TabIndex        =   2
         Top             =   360
         Width           =   1575
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

Private mNroOrden As Long

Private Sub Form_Load()
    dtpInstInt = DateTime.Now
End Sub

Property Get NroOrden() As Long
    NroOrden = mNroOrden
End Property

Property Let NroOrden(NroOrdenNuevo As Long)
    mNroOrden = NroOrdenNuevo
End Property

Private Sub cmdGuardarTrabajo_Click()

    If lblCodInternet.Caption = vbNullString Then
        MsgBox "Seleccione un usuario antes de generar la orden.", vbOKOnly + vbInformation, "Faltan datos"
    Else
    
        Dim st As Integer
        With main.vTrabInternet
            .Clear
            .FieldValue("NroOrden") = mNroOrden
            .FieldValue("Estado") = Estados.NUEVO
            .FieldValue("Fecha_Pedido") = dtpInstInt.Value
            .FieldValue("tipo_conexion") = 4
            .FieldValue("obs") = txtObs.Text
            st = .Insert
        End With
    
        Call cerrar
        
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

