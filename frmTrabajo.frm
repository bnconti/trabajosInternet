VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTrabajo 
   Caption         =   "Datos del trabajo"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmDatosUsuario 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.CheckBox chkImprimirOrden 
         Caption         =   "Imprimir orden de conexión"
         Height          =   375
         Left            =   3840
         TabIndex        =   22
         Top             =   5280
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2760
         TabIndex        =   21
         Text            =   "Combo1"
         Top             =   1800
         Width           =   3255
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2760
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   2280
         Width           =   3255
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2760
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   2760
         Width           =   3255
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2760
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   1320
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2760
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   840
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2760
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   360
         Width           =   3255
      End
      Begin VB.ComboBox cmbCuadrilla 
         Height          =   315
         Left            =   2760
         TabIndex        =   15
         Text            =   "Combo1"
         Top             =   4560
         Width           =   3255
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2760
         TabIndex        =   14
         Top             =   4080
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
         _Version        =   393216
         Format          =   118882305
         CurrentDate     =   44076
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2760
         TabIndex        =   13
         Top             =   3600
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
         _Version        =   393216
         Format          =   118882305
         CurrentDate     =   44076
      End
      Begin VB.CommandButton btnEliminar 
         Caption         =   "Eliminar"
         Height          =   450
         Left            =   2400
         TabIndex        =   12
         Top             =   6000
         Width           =   1600
      End
      Begin VB.CommandButton btnVolver 
         Caption         =   "Volver"
         Height          =   450
         Left            =   480
         TabIndex        =   11
         Top             =   6000
         Width           =   1600
      End
      Begin VB.CommandButton btnActualizar 
         Caption         =   "Actualizar"
         Height          =   450
         Left            =   4320
         TabIndex        =   10
         Top             =   6000
         Width           =   1600
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   6120
         Y1              =   5760
         Y2              =   5760
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   6120
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Label lblCuadrilla 
         Caption         =   "Cuadrilla"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   4560
         Width           =   2175
      End
      Begin VB.Label lblHora 
         Caption         =   "Hora"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   4080
         Width           =   2415
      End
      Begin VB.Label lblFechaDeInstalacion 
         Caption         =   "Fecha de instalación"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Label lblTelefono 
         Caption         =   "Teléfono"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label lblFechaDePedido 
         Caption         =   "Fecha de pedido"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label lblTipoDeConexion 
         Caption         =   "Tipo de conexión"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label lblNombreDeUsuario 
         Caption         =   "Nombre de usuario"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label lblDomicilio 
         Caption         =   "Domicilio de instalación"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label lblNombreCompleto 
         Caption         =   "Nombre completo"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmTrabajo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnActualizar_Click()

    ' Actualizar el trabajo en la BD.

    If chkImprimirOrden = True Then
        ' Llamar módulo de impresión
    End If
End Sub

Private Sub btnEliminar_Click()
    If MsgBox("Se eliminará este trabajo de la base de datos, ¿está seguro?", vbYesNo) = vbYes Then
        ' Borrar
    Else
        Unload Me
    End If
End Sub

Private Sub btnVolver_Click()
    Unload Me
End Sub
