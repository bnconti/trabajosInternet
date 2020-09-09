VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTrabajo 
   Caption         =   "Datos del trabajo"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6555
   Icon            =   "frmTrabajo.frx":0000
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
         TabIndex        =   21
         Top             =   5280
         Width           =   2295
      End
      Begin VB.ComboBox cmbTipoConexion 
         Height          =   315
         ItemData        =   "frmTrabajo.frx":030A
         Left            =   2760
         List            =   "frmTrabajo.frx":031A
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   3120
         Width           =   3255
      End
      Begin VB.TextBox txtTlfn 
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   375
         Left            =   2760
         TabIndex        =   19
         Top             =   1800
         Width           =   3255
      End
      Begin VB.TextBox txtUsInternet 
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   375
         Left            =   2760
         TabIndex        =   18
         Top             =   1320
         Width           =   3255
      End
      Begin VB.TextBox txtDomi 
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   375
         Left            =   2760
         TabIndex        =   17
         Top             =   840
         Width           =   3255
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   375
         Left            =   2760
         TabIndex        =   16
         Top             =   360
         Width           =   3255
      End
      Begin VB.ComboBox cmbCuadrilla 
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   4560
         Width           =   3255
      End
      Begin MSComCtl2.DTPicker dtHoraInst 
         Height          =   375
         Left            =   2760
         TabIndex        =   14
         Top             =   4080
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
         _Version        =   393216
         Format          =   41025538
         CurrentDate     =   44076
      End
      Begin MSComCtl2.DTPicker dtFechaInst 
         Height          =   375
         Left            =   2760
         TabIndex        =   13
         Top             =   3600
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
         _Version        =   393216
         Format          =   41025537
         CurrentDate     =   44076
      End
      Begin VB.CommandButton btnEliminar 
         Caption         =   "Eliminar"
         Height          =   450
         Left            =   4320
         TabIndex        =   12
         Top             =   6000
         Width           =   1600
      End
      Begin VB.CommandButton btnVolver 
         Caption         =   "Volver"
         Height          =   450
         Left            =   240
         TabIndex        =   11
         Top             =   6000
         Width           =   1600
      End
      Begin VB.CommandButton btnActualizar 
         Caption         =   "Actualizar"
         Height          =   450
         Left            =   2280
         TabIndex        =   10
         Top             =   6000
         Width           =   1600
      End
      Begin MSComCtl2.DTPicker dtFechaPedido 
         Height          =   375
         Left            =   2760
         TabIndex        =   22
         Top             =   2280
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   41025537
         CurrentDate     =   44076
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
         Y1              =   2880
         Y2              =   2880
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
         Caption         =   "Hora de instalación"
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
         Top             =   1800
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
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label lblNombreDeUsuario 
         Caption         =   "Cuenta de Internet"
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
         Caption         =   "Apellido y nombre"
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
Private idTrabajo As Long


Private Sub Form_Load()

    With main.tablaTrabajosAProgramar
        txtNombre = .TextMatrix(.Row, 0)
        txtDomi = .TextMatrix(.Row, 1)
        txtUsInternet = .TextMatrix(.Row, 2)
        txtTlfn = .TextMatrix(.Row, 5)
        dtFechaPedido = .TextMatrix(.Row, 4)
        cmbTipoConexion.ListIndex = Val(Mid(.TextMatrix(.Row, 3), 1, 1)) - 1
        idTrabajo = Val(.TextMatrix(.Row, 6))
    End With
    
    dtFechaInst.Value = DateTime.Now
    
    Call cargarCuadrillas
    
    
    ' Ver si se puede cargar directamente en las props.
    cmbTipoConexion.ItemData(0) = 1
    cmbTipoConexion.ItemData(1) = 2
    cmbTipoConexion.ItemData(2) = 3
    cmbTipoConexion.ItemData(3) = 4

End Sub

Private Sub btnActualizar_Click()

    Dim st As Integer
    
    If cmbCuadrilla.Text = vbNullString Then
        MsgBox "¡Recordá seleccionar una cuadrilla!", vbOKOnly + vbExclamation, "Datos incompletos"
    Else
    
        With main.vTrabInternet
            .IndexNumber = 0
            .FieldValue("id_trabajo") = idTrabajo
            
            st = .GetEqual
            
            If st = 0 Then
                .FieldValue("tipo_conexion") = cmbTipoConexion.ItemData(cmbTipoConexion.ListIndex)
                .FieldValue("fecha_inst") = dtFechaInst.Value
                .FieldValue("hora_inst") = dtHoraInst.Value
                .FieldValue("idcuadrilla") = cmbCuadrilla.ItemData(cmbCuadrilla.ListIndex)
                .FieldValue("estado") = Estados.PROGRAMADO
                
                .Update
            End If
            
        End With

        If chkImprimirOrden = True Then
            ' Llamar módulo de impresión
        End If
        
        Unload Me
        
    End If

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
    End If
        
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
            If .FieldValue("habilitado") = 1 Then
                cmbCuadrilla.AddItem (.FieldValue("miembros"))
                cmbCuadrilla.ItemData(cmbCuadrilla.NewIndex) = .FieldValue("idcuadrilla")
            End If
            
            status = .GetNext
        Wend
    End With
End Sub


