VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCambioFTTH 
   Caption         =   "Cambio a FTTH"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5085
   Icon            =   "frmCambioFTTH.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   5085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdVolver 
      Caption         =   "&Volver"
      Height          =   495
      Left            =   495
      TabIndex        =   7
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton cmdGuardarTrabajo 
      Caption         =   "&Guardar trabajo"
      Height          =   495
      Left            =   2775
      TabIndex        =   6
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Frame frmDatosTrabajo 
      Caption         =   "Datos del trabajo"
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   4815
      Begin VB.TextBox txtObs 
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   4455
      End
      Begin MSComCtl2.DTPicker dtpInstInt 
         Height          =   375
         Left            =   2640
         TabIndex        =   8
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   41025537
         CurrentDate     =   44083
      End
      Begin VB.Label lblObs 
         Caption         =   "Observaciones"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   840
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
      Width           =   4815
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
         Height          =   375
         Left            =   1320
         TabIndex        =   11
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label lblNombre 
         Height          =   375
         Left            =   1320
         TabIndex        =   10
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label lblCodCli 
         Height          =   375
         Left            =   3120
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblCodInternet 
         Height          =   375
         Left            =   1320
         TabIndex        =   1
         Top             =   360
         Width           =   1575
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
            .FieldValue("tipo_conexion") = Conexiones.CAMBIO_FTTH
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


