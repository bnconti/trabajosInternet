VERSION 5.00
Object = "{50777BB0-FB9D-11D1-A76F-006097D2F089}#7.0#0"; "ACCTR732.OCX"
Begin VB.Form frmSelCli 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Búsqueda de clientes"
   ClientHeight    =   4710
   ClientLeft      =   1290
   ClientTop       =   345
   ClientWidth     =   6330
   Icon            =   "frmSelCli.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Seleccionar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4620
      TabIndex        =   3
      Top             =   4260
      Width           =   1575
   End
   Begin VB.TextBox BuscaCliente 
      Height          =   285
      Index           =   1
      Left            =   3000
      TabIndex        =   1
      Top             =   330
      Width           =   3075
   End
   Begin VB.TextBox BuscaCliente 
      Height          =   285
      Index           =   0
      Left            =   780
      TabIndex        =   0
      Top             =   330
      Width           =   2145
   End
   Begin ACCtrl7Lib.VAVScroll VAVScroll1 
      Height          =   3420
      Left            =   5940
      TabIndex        =   4
      Top             =   660
      Width           =   225
      _Version        =   458752
      _ExtentX        =   397
      _ExtentY        =   6032
      _StockProps     =   64
      VAccessName     =   "vAClientes"
   End
   Begin ACCtrl7Lib.VAList lstClientes 
      Height          =   3420
      Left            =   60
      TabIndex        =   2
      Top             =   660
      Width           =   5895
      _Version        =   458752
      _ExtentX        =   10398
      _ExtentY        =   6032
      _StockProps     =   77
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VAccessName     =   "VAClientes"
      VAFieldName     =   "CODCLI,APELLIDO,NOMBRE"
      VARecordList    =   -1  'True
      ColumnWidth     =   "50;146"
      VAAutoScroll    =   0   'False
   End
   Begin VB.Label CODCLI 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3975
      TabIndex        =   7
      Top             =   30
      Width           =   1530
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre/s"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   3030
      TabIndex        =   6
      Top             =   45
      Width           =   1155
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido / R. Social"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   840
      TabIndex        =   5
      Top             =   45
      Width           =   1785
   End
End
Attribute VB_Name = "frmSelCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PosicionAnterior As Long

Dim mCodLuz As Long

Property Get codLuz() As Long
  codLuz = mCodLuz
End Property

Private Sub Form_Load()
  With main.VAClientes
    PosicionAnterior = .Position
    .IndexNumber = 1  ' Apellido + Nombre
    .GetFirst
  End With

  mCodLuz = 0
End Sub

Private Sub BuscaCliente_Change(Index As Integer)
  With main.VAClientes
    .FieldValue("Apellido") = BuscaCliente(0)
    .FieldValue("Nombre") = BuscaCliente(1)
    .GetGreaterOrEqual
  End With
End Sub

Private Sub CmdAceptar_Click()
  Call lstClientes_DblClick
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  With main.VAClientes
    Select Case KeyCode
    Case vbKeyEscape
      .IndexNumber = 0
      .Position = PosicionAnterior
      .GetDirect
      Unload Me
    Case vbKeyReturn
      lstClientes_DblClick
    End Select
  End With
End Sub


Private Sub lstClientes_DblClick()
  frmSelOrd.Show 1
  mCodLuz = frmSelOrd.CodAlumbrado

  If mCodLuz > 0 Then
    frmCambioFTTH.lblNombre = main.VAClientes.FieldValue("nombre") & " " & main.VAClientes.FieldValue("apellido")
    frmCambioFTTH.lblCodCli = "Cód. cli. " & main.VAClientes.FieldValue("CodCli")
    Unload Me
  Else
    frmCambioFTTH.lblNombre = vbNullString
    frmCambioFTTH.lblCodCli = vbNullString
  End If

End Sub
