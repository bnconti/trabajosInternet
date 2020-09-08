VERSION 5.00
Begin VB.Form frmCambioFTTH 
   Caption         =   "Cambio a FTTH"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmDatosTrabajo 
      Caption         =   "Datos del trabajo"
      Height          =   1695
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   4455
      Begin VB.Label lblObs 
         Caption         =   "Observaciones"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblFechaDePedido 
         Caption         =   "Fecha de pedido"
         Height          =   375
         Left            =   240
         TabIndex        =   7
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
      Width           =   4455
      Begin VB.TextBox txtCodCliente 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         TabIndex        =   5
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtCodUs 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         TabIndex        =   4
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdSelCli 
         Caption         =   "Lupita"
         Height          =   615
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblNroOrden 
         Caption         =   "Cod. de cliente"
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblCodUs 
         Caption         =   "Cód. de usuario"
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmCambioFTTH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_Click()
    frmSelCli.Show 1, Me
End Sub
