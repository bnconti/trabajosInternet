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
      Caption         =   "&Volver"
      Height          =   495
      Left            =   3375
      TabIndex        =   7
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton cmdGuardarTrabajo 
      Caption         =   "&Guardar trabajo"
      Height          =   495
      Left            =   1215
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
         Format          =   40435713
         CurrentDate     =   44083
      End
      Begin VB.Label Label1 
         Caption         =   "Actualizar tarifa"
         Height          =   375
         Left            =   2040
         TabIndex        =   13
         Top             =   360
         Width           =   1935
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
         TabIndex        =   10
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
         Left            =   3720
         TabIndex        =   2
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
    Call cargarTarifasFTTH
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
        
        cambiarNoFacturar mNroOrden, "NOFACTURAR"
        
        If cmbTarifas.Text <> vbNullString Then
            Dim idTarifa As Long
            idTarifa = cmbTarifas.ItemData(cmbTarifas.ListIndex)
            Call actualizarTarifa(idTarifa)
        End If
        
        Call cerrar
        
    End If
End Sub

Private Sub actualizarTarifa(idTarifa As Long)
    ' Cambia la tarifa a la de fibra seleccionada por el operador
    
    With main
        .VOrdenes.IndexNumber = 0
        .VOrdenes.FieldValue("nroOrden") = nroOrden
        
        If .VOrdenes.GetEqual = 0 Then
        
            If Not (IsNull(.VOrdenes.FieldValue("CodAlumbrado"))) Then
                .VAsumAlum.IndexNumber = 0
                .VAsumAlum.FieldValue("CodAlumbrado") = .VOrdenes.FieldValue("CodAlumbrado")
                
                If .VAsumAlum.GetEqual = 0 Then
                    .VAsumAlum.FieldValue("ID_Tarifa") = idTarifa
                    .VAsumAlum.Update
                End If
                
            End If
        End If
    End With
End Sub

Private Sub cargarTarifasFTTH()
    With main.VTarifas
        Dim st As Integer
    
        .IndexNumber = 2
        .FieldValue("Id_Servicio") = 3
        .FieldValue("Id_Tipo") = 0
        
        st = .GetGreaterOrEqual
        
        Do While st = 0 And .FieldValue("Id_Servicio") = 3
            If InStr(UCase(.FieldValue("descrip")), "FTTH") > 0 Then
                cmbTarifas.AddItem Format(.FieldValue("Id_Tarifa"), "0000") & " - " & (UCase(.FieldValue("descrip")))
                cmbTarifas.ItemData(cmbTarifas.NewIndex) = .FieldValue("Id_Tarifa")
            End If
            .GetNext
        Loop

    End With
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

