VERSION 5.00
Object = "{47E7B6C9-8256-11CF-AB56-0000C04D1EB9}#7.0#0"; "ACBTR732.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7l.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trabajos de Internet"
   ClientHeight    =   8265
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   17415
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   17415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCambioFTTH 
      Caption         =   "Cambiar a FTTH"
      Height          =   375
      Left            =   14400
      TabIndex        =   27
      Top             =   480
      Width           =   2655
   End
   Begin VB.Frame frmBD 
      Height          =   1455
      Left            =   11520
      TabIndex        =   8
      Top             =   7200
      Visible         =   0   'False
      Width           =   2775
      Begin VAccessLib.VAccess VAsumAlumInte 
         Left            =   1080
         Top             =   240
         _Version        =   458752
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
         VAccessName     =   "VAsumAlumInte"
         TableName       =   "ASUMALUMINTE"
         Location        =   "\\servidor\D\Compu\SFS2000\Datos\ASUMALUMINTE.mkd"
         OpenMode        =   2
         DdfPath         =   "\\servidor\D\Compu\SFS2000\Datos"
         VAUDDDFInfo     =   "main.frx":27A2
      End
      Begin VAccessLib.VAccess VOrdenes 
         Left            =   120
         Top             =   240
         _Version        =   458752
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
         VAccessName     =   "VOrdenes"
         TableName       =   "ORDENES"
         Location        =   "\\servidor\compu\SFS2000\datos\ORDENES.mkd"
         OpenMode        =   2
         DdfPath         =   "\\servidor\compu\SFS2000\datos"
         HostConnect     =   0   'False
         VAUDDDFInfo     =   "main.frx":313D
      End
      Begin VAccessLib.VAccess VAClientes 
         Left            =   600
         Top             =   240
         _Version        =   458752
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
         VAccessName     =   "VAClientes"
         TableName       =   "ACLIENTES"
         Location        =   "\\servidor\D\Compu\SFS2000\Datos\ACLIENTES.MKD"
         OpenMode        =   2
         DdfPath         =   "\\servidor\D\Compu\SFS2000\Datos"
         VAUDDDFInfo     =   "main.frx":3DD0
      End
      Begin VAccessLib.VAccess vTrabInternet 
         Left            =   120
         Top             =   720
         _Version        =   458752
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
         VAccessName     =   "vTrabInternet"
         TableName       =   "TRABAJOINTERNET"
         Location        =   "\\servidor\D\Compu\SFS2000\Datos\TRABAJOINTERNET.mkd"
         DdfPath         =   "\\servidor\D\Compu\SFS2000\Datos"
         VAUDDDFInfo     =   "main.frx":4959
      End
      Begin VAccessLib.VAccess VCuadrillas 
         Left            =   1560
         Top             =   240
         _Version        =   458752
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
         VAccessName     =   "VCuadrillas"
         TableName       =   "CUADRILLASINTERNET"
         Location        =   "\\servidor\D\Compu\SFS2000\Datos\CUADRILLASINTERNET.mkd"
         DdfPath         =   "\\servidor\D\Compu\SFS2000\Datos"
         VAUDDDFInfo     =   "main.frx":53D8
      End
   End
   Begin TabDlg.SSTab tabTrabajos 
      Height          =   8250
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17400
      _ExtentX        =   30692
      _ExtentY        =   14552
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Para programar"
      TabPicture(0)   =   "main.frx":5D73
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "tablaTrabajosAProgramar"
      Tab(0).Control(1)=   "btnProgRecuperar"
      Tab(0).Control(2)=   "btnProgExcel"
      Tab(0).Control(3)=   "btnProgImprimir"
      Tab(0).Control(4)=   "frmFiltrar"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Para instalar"
      TabPicture(1)   =   "main.frx":5D8F
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "tablaTrabajosAInstalar"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Command5"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Command6"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "btnAInstalarRecuperar"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "btnGuardarFinalizados"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Terminados"
      TabPicture(2)   =   "main.frx":5DAB
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tablaTrabajosInstalados"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "frmFiltrado"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Command9"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Command10"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "btnInstaladosRecuperar"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin VB.CommandButton btnGuardarFinalizados 
         BackColor       =   &H8000000E&
         Caption         =   "Guardar finalizados"
         Height          =   615
         Left            =   360
         TabIndex        =   28
         Top             =   7200
         Width           =   1695
      End
      Begin VB.CommandButton btnInstaladosRecuperar 
         BackColor       =   &H8000000E&
         Caption         =   "Recuperar"
         Height          =   615
         Left            =   -74640
         TabIndex        =   25
         Top             =   6480
         Width           =   1695
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Exportar a Excel"
         Height          =   615
         Left            =   -59640
         TabIndex        =   24
         Top             =   6480
         Width           =   1695
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Imprimir"
         Height          =   615
         Left            =   -61560
         TabIndex        =   23
         Top             =   6480
         Width           =   1695
      End
      Begin VB.Frame frmFiltrado 
         Caption         =   "Filtrar por"
         Height          =   1875
         Left            =   -72600
         TabIndex        =   18
         Top             =   6240
         Width           =   4455
         Begin MSComCtl2.DTPicker dtDesdeTerminados 
            Height          =   375
            Left            =   2400
            TabIndex        =   31
            Top             =   600
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd/MM"
            Format          =   95551491
            CurrentDate     =   44089
         End
         Begin VB.ComboBox cmbConexionTerminados 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   1320
            Width           =   2055
         End
         Begin VB.ComboBox cmbCuadrillaTerminados 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   600
            Width           =   2055
         End
         Begin MSComCtl2.DTPicker dtHastaTerminados 
            Height          =   375
            Left            =   2400
            TabIndex        =   32
            Top             =   1320
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd/MM"
            Format          =   95551491
            CurrentDate     =   44089
         End
         Begin VB.Label lblFechaHastaInstalados 
            Caption         =   "Fecha hasta"
            Height          =   345
            Left            =   2400
            TabIndex        =   30
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lblFechaDesdeInstalados 
            Caption         =   "Fecha desde"
            Height          =   345
            Left            =   2400
            TabIndex        =   29
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Tipo de conexión"
            Height          =   345
            Left            =   120
            TabIndex        =   22
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Cuadrilla"
            Height          =   345
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.CommandButton btnAInstalarRecuperar 
         BackColor       =   &H8000000E&
         Caption         =   "Recuperar"
         Height          =   615
         Left            =   360
         TabIndex        =   16
         Top             =   6480
         Width           =   1695
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Exportar a Excel"
         Height          =   615
         Left            =   15360
         TabIndex        =   15
         Top             =   6480
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Imprimir"
         Height          =   615
         Left            =   13440
         TabIndex        =   14
         Top             =   6480
         Width           =   1695
      End
      Begin VB.Frame Frame1 
         Caption         =   "Filtrar por"
         Height          =   1875
         Left            =   2400
         TabIndex        =   9
         Top             =   6240
         Width           =   2535
         Begin VB.ComboBox cmbConexionAInstalar 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1320
            Width           =   2055
         End
         Begin VB.ComboBox cmbCuadrillaAInstalar 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo de conexión"
            Height          =   345
            Left            =   120
            TabIndex        =   13
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Cuadrilla"
            Height          =   345
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame frmFiltrar 
         Caption         =   "Filtrar por"
         Height          =   1155
         Left            =   -72600
         TabIndex        =   4
         Top             =   6240
         Width           =   2535
         Begin VB.ComboBox cmbConexionAProgramar 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label lblTipoDeConexion 
            Caption         =   "Tipo de conexión"
            Height          =   345
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.CommandButton btnProgImprimir 
         Caption         =   "Imprimir"
         Height          =   615
         Left            =   -61560
         TabIndex        =   3
         Top             =   6480
         Width           =   1695
      End
      Begin VB.CommandButton btnProgExcel 
         Caption         =   "Exportar a Excel"
         Height          =   615
         Left            =   -59640
         TabIndex        =   2
         Top             =   6480
         Width           =   1695
      End
      Begin VB.CommandButton btnProgRecuperar 
         BackColor       =   &H8000000E&
         Caption         =   "Recuperar"
         Height          =   615
         Left            =   -74640
         TabIndex        =   1
         Top             =   6480
         Width           =   1695
      End
      Begin VSFlex7LCtl.VSFlexGrid tablaTrabajosAProgramar 
         Height          =   5295
         Left            =   -74640
         TabIndex        =   7
         Top             =   960
         Width           =   16695
         _cx             =   29448
         _cy             =   9340
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"main.frx":5DC7
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VSFlex7LCtl.VSFlexGrid tablaTrabajosAInstalar 
         Height          =   5295
         Left            =   360
         TabIndex        =   17
         Top             =   960
         Width           =   16695
         _cx             =   29448
         _cy             =   9340
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"main.frx":5ED9
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VSFlex7LCtl.VSFlexGrid tablaTrabajosInstalados 
         Height          =   5295
         Left            =   -74640
         TabIndex        =   26
         Top             =   960
         Width           =   16695
         _cx             =   29448
         _cy             =   9340
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"main.frx":606A
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuCuadrilla 
         Caption         =   "Cuadrilla"
      End
   End
   Begin VB.Menu mnuProcesos 
      Caption         =   "Procesos"
      Begin VB.Menu mnuCambioFTTH 
         Caption         =   "Cambio a FTTH"
      End
   End
   Begin VB.Menu mnuListados 
      Caption         =   "Listados"
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum Estados
    NUEVO = 1          ' Recien se cargó la orden de trabajo
    PROGRAMADO = 2     ' Se le asignó una fecha, hora y cuadrilla
    TERMINADO = 0      ' La instalación fue realizada
End Enum

Private Const CHEQUEADO As Integer = 1

Private Const COL_ID_TRABAJO As Integer = 10

Private arrConexiones As Variant


Private Sub Form_Load()

    arrConexiones = Array("ALTA FTTH", "ALTA ANTENA", "ALTA EDIFICIO", "CAMBIO A FTTH")
    
    Call cargarCuadrillas
    Call cargarTiposConexion

End Sub

Private Sub btnCambioFTTH_Click()
    frmCambioFTTH.Show 1, Me
End Sub

Private Sub btnProgRecuperar_Click()
    Call cargarTablaTrabajosAProgramar
End Sub

Private Sub btnAInstalarRecuperar_Click()
    Call cargarTablaTrabajosAInstalar
End Sub

Private Sub btnInstaladosRecuperar_Click()
    Call cargarTablaTrabajosInstalados
End Sub

Private Sub cargarTablaTrabajosAProgramar()
    Dim st As Integer
    
    tablaTrabajosAProgramar.Rows = 1
    
    With vTrabInternet
        vTrabInternet.IndexNumber = 0
        st = .GetFirst
        
        While st = 0
        
            VOrdenes.IndexNumber = 0
            VAClientes.IndexNumber = 0
            VAsumAlumInte.IndexNumber = 0
            
            VOrdenes.FieldValue("NroOrden") = .FieldValue("Nroorden")
            VOrdenes.GetEqual
            
            VAClientes.FieldValue("CodCli") = VOrdenes.FieldValue("CodCli")
            VAClientes.GetEqual
            
            VAsumAlumInte.FieldValue("CodAlumbrado") = VOrdenes.FieldValue("CodAlumbrado")
            VAsumAlumInte.GetEqual
            
            If VOrdenes.status = 0 And _
                VAClientes.status = 0 And _
                VAsumAlumInte.status = 0 And _
                vTrabInternet.FieldValue("estado") = Estados.NUEVO Then
            
                tablaTrabajosAProgramar.AddItem (VAClientes.FieldValue("apellido") & ", " & VAClientes.FieldValue("nombre") & vbTab & _
                                         VOrdenes.FieldValue("domicilio") & vbTab & _
                                         VAsumAlumInte.FieldValue("UsInt") & vbTab & _
                                         vTrabInternet.FieldValue("Tipo_Conexion") & " - " & arrConexiones(vTrabInternet.FieldValue("Tipo_Conexion") - 1) & vbTab & _
                                         vTrabInternet.FieldValue("fecha_pedido") & vbTab & _
                                         VAClientes.FieldValue("reserva") & vbTab & _
                                         vTrabInternet.FieldValue("id_trabajo"))
                
            End If
                                     
            st = .GetNext
        
        Wend
        
        Call filtrarTablaAProgramar
        tablaTrabajosAProgramar.AutoSize 0, tablaTrabajosAProgramar.Cols - 1
        
    End With
End Sub

Private Sub cargarTablaTrabajosAInstalar()
    Dim st As Integer
    
    tablaTrabajosAInstalar.Rows = 1
    
    vTrabInternet.IndexNumber = 0
    st = vTrabInternet.GetFirst
    
    While st = 0
    
        VOrdenes.IndexNumber = 0
        VAClientes.IndexNumber = 0
        VAsumAlumInte.IndexNumber = 0
        VCuadrillas.IndexNumber = 0
        
        VOrdenes.FieldValue("NroOrden") = vTrabInternet.FieldValue("Nroorden")
        VOrdenes.GetEqual
        
        VAClientes.FieldValue("CodCli") = VOrdenes.FieldValue("CodCli")
        VAClientes.GetEqual
        
        VAsumAlumInte.FieldValue("CodAlumbrado") = VOrdenes.FieldValue("CodAlumbrado")
        VAsumAlumInte.GetEqual
        
        VCuadrillas.FieldValue("idcuadrilla") = vTrabInternet.FieldValue("idcuadrilla")
        VCuadrillas.GetEqual
        
        If VOrdenes.status = 0 And _
            VAClientes.status = 0 And _
            VAsumAlumInte.status = 0 And _
            VCuadrillas.status = 0 And _
            vTrabInternet.FieldValue("estado") = Estados.PROGRAMADO Then
        
            tablaTrabajosAInstalar.AddItem (vbNullString & vbTab & _
                                    VAClientes.FieldValue("apellido") & ", " & VAClientes.FieldValue("nombre") & vbTab & _
                                    VOrdenes.FieldValue("domicilio") & vbTab & _
                                    VAsumAlumInte.FieldValue("UsInt") & vbTab & _
                                    vTrabInternet.FieldValue("Tipo_Conexion") & " - " & arrConexiones(vTrabInternet.FieldValue("Tipo_Conexion") - 1) & vbTab & _
                                    vTrabInternet.FieldValue("fecha_pedido") & vbTab & _
                                    VAClientes.FieldValue("reserva") & vbTab & _
                                    vTrabInternet.FieldValue("fecha_inst") & vbTab & _
                                    vTrabInternet.FieldValue("hora_inst") & vbTab & _
                                    VCuadrillas.FieldValue("miembros") & vbTab & _
                                    vTrabInternet.FieldValue("id_trabajo"))
            tablaTrabajosAInstalar.Cell(flexcpChecked, tablaTrabajosAInstalar.Rows - 1, 0, tablaTrabajosAInstalar.Rows - 1, 0) = flexUnchecked
            
        End If
                                 
        st = vTrabInternet.GetNext
    
    Wend
    
    ' Call filtrarTablaAInstalar
    tablaTrabajosAInstalar.AutoSize 0, tablaTrabajosAInstalar.Cols - 1

End Sub

Private Sub cargarTablaTrabajosInstalados()
    Dim st As Integer
    
    tablaTrabajosInstalados.Rows = 1
    
    With vTrabInternet
        vTrabInternet.IndexNumber = 0
        st = .GetFirst
        
        While st = 0
        
            VOrdenes.IndexNumber = 0
            VAClientes.IndexNumber = 0
            VAsumAlumInte.IndexNumber = 0
            VCuadrillas.IndexNumber = 0
            
            VOrdenes.FieldValue("NroOrden") = .FieldValue("Nroorden")
            VOrdenes.GetEqual
            
            VAClientes.FieldValue("CodCli") = VOrdenes.FieldValue("CodCli")
            VAClientes.GetEqual
            
            VAsumAlumInte.FieldValue("CodAlumbrado") = VOrdenes.FieldValue("CodAlumbrado")
            VAsumAlumInte.GetEqual
            
            VCuadrillas.FieldValue("idcuadrilla") = vTrabInternet.FieldValue("idcuadrilla")
            VCuadrillas.GetFirst
            
            If VOrdenes.status = 0 And _
                VAClientes.status = 0 And _
                VAsumAlumInte.status = 0 And _
                VCuadrillas.status = 0 And _
                vTrabInternet.FieldValue("estado") = Estados.TERMINADO Then
            
                tablaTrabajosInstalados.AddItem (VAClientes.FieldValue("apellido") & ", " & VAClientes.FieldValue("nombre") & vbTab & _
                                         VOrdenes.FieldValue("domicilio") & vbTab & _
                                         VAsumAlumInte.FieldValue("UsInt") & vbTab & _
                                         vTrabInternet.FieldValue("Tipo_Conexion") & " - " & arrConexiones(vTrabInternet.FieldValue("Tipo_Conexion") - 1) & vbTab & _
                                         vTrabInternet.FieldValue("fecha_pedido") & vbTab & _
                                         VAClientes.FieldValue("reserva") & vbTab & _
                                         vTrabInternet.FieldValue("fecha_inst") & vbTab & _
                                         vTrabInternet.FieldValue("hora_inst") & vbTab & _
                                         VCuadrillas.FieldValue("miembros") & vbTab & _
                                         vTrabInternet.FieldValue("id_trabajo"))
                
            End If
                                     
            st = .GetNext
        
        Wend
        
        tablaTrabajosInstalados.AutoSize 0, tablaTrabajosInstalados.Cols - 1
        
    End With
End Sub

Private Sub btnGuardarFinalizados_Click()
    If tablaTrabajosAInstalar.Rows > 1 Then
    
        Dim idTrabajo As Integer

        Dim fila As Integer
        Dim ultimaFila As Integer
        ultimaFila = tablaTrabajosAInstalar.Rows - 1

        With tablaTrabajosAInstalar
            For fila = 1 To ultimaFila
                If .Cell(flexcpChecked, fila, 0, fila, 0) = CHEQUEADO Then
                    idTrabajo = .TextMatrix(fila, 10)
                    Call finalizarTrabajo(idTrabajo)
                End If
            Next fila
        End With
        
        Call cargarTablaTrabajosAInstalar

    Else
        MsgBox "Tiene que recuperar los trabajos antes de poder marcarlos como finalizados.", vbInformation + vbOKOnly, "No hay trabajos"
    End If
End Sub

Private Sub finalizarTrabajo(idTrabajo As Integer)
    With main.vTrabInternet
        .IndexNumber = 0
        .FieldValue("id_trabajo") = idTrabajo
        .GetEqual
        
        If .status = 0 Then
            .FieldValue("estado") = Estados.TERMINADO
            .Update
        End If
            
    End With
End Sub

Private Sub cmbConexionAProgramar_Click()
    Call filtrarTablaAProgramar
End Sub

Private Sub filtrarTablaAProgramar()

    Dim fila As Integer
    Dim ultimaFila As Integer
    ultimaFila = tablaTrabajosAProgramar.Rows - 1
    
    Dim conexion As String

    With tablaTrabajosAProgramar
        For fila = 1 To ultimaFila
            conexion = Mid(.TextMatrix(fila, 3), 5)
            If conexion = cmbConexionAProgramar.Text Or cmbConexionAProgramar.Text = "TODAS" Then
                .RowHidden(fila) = False
            Else
                .RowHidden(fila) = True
            End If
        Next fila
    End With
End Sub


Private Sub cmbConexionAInstalar_Click()
    If tablaTrabajosAInstalar.Rows > 1 Then Call filtrarTablaAInstalar
End Sub

Private Sub cmbCuadrillaAInstalar_Click()
    If tablaTrabajosAInstalar.Rows > 1 Then Call filtrarTablaAInstalar
End Sub

Private Sub filtrarTablaAInstalar()
    Dim fila As Integer
    Dim ultimaFila As Integer
    ultimaFila = tablaTrabajosAInstalar.Rows - 1
    
    Dim conexion As String
    Dim cuadrilla As String

    With tablaTrabajosAInstalar
        For fila = 1 To ultimaFila
            conexion = Mid(.TextMatrix(fila, 4), 5)
            cuadrilla = .TextMatrix(fila, 9)
            If (conexion = cmbConexionAInstalar.Text Or cmbConexionAInstalar.Text = "TODAS") _
                And (cuadrilla = cmbCuadrillaAInstalar.Text Or cmbCuadrillaAInstalar.Text = "TODAS") Then
                .RowHidden(fila) = False
            Else
                .RowHidden(fila) = True
            End If
        Next fila
    End With
End Sub

Private Sub cmbCuadrillaTerminados_Click()
    Call filtrarTablaTerminados
End Sub

Private Sub cmbConexionTerminados_Change()
    Call filtrarTablaTerminados
End Sub

Private Sub dtDesdeTerminados_Click()
    Call filtrarTablaTerminados
End Sub

Private Sub dtHastaTerminados_Click()
    Call filtrarTablaTerminados
End Sub

Private Sub filtrarTablaTerminados()
    Dim fila As Integer
    Dim ultimaFila As Integer
    ultimaFila = tablaTrabajosInstalados.Rows - 1
    
    Dim conexion As String
    Dim cuadrilla As String

    With tablaTrabajosInstalados
        For fila = 1 To ultimaFila
            conexion = Mid(.TextMatrix(fila, 3), 5)
            cuadrilla = .TextMatrix(fila, 8)
            If (conexion = cmbConexionAInstalar.Text Or cmbConexionAInstalar.Text = "TODAS") _
                And (cuadrilla = cmbCuadrillaAInstalar.Text Or cmbCuadrillaAInstalar.Text = "TODAS") Then
                .RowHidden(fila) = False
            Else
                .RowHidden(fila) = True
            End If
        Next fila
    End With
End Sub

Private Sub cargarTiposConexion()
    Dim i As Integer
    
    For i = 0 To UBound(arrConexiones)
        cmbConexionAProgramar.AddItem (arrConexiones(i))
        cmbConexionAInstalar.AddItem (arrConexiones(i))
        cmbConexionTerminados.AddItem (arrConexiones(i))
    Next i
    
    cmbConexionAProgramar.AddItem "TODAS"
    cmbConexionAInstalar.AddItem "TODAS"
    cmbConexionTerminados.AddItem "TODAS"
    
    cmbConexionAProgramar.ListIndex = cmbConexionAProgramar.ListCount - 1
    cmbConexionAInstalar.ListIndex = cmbConexionAInstalar.ListCount - 1
    cmbConexionTerminados.ListIndex = cmbConexionTerminados.ListCount - 1
    
End Sub

Private Sub cargarCuadrillas()
    Dim status As Integer
    
    With main.VCuadrillas
        .IndexNumber = 0
        status = .GetFirst
        
        While status = 0
            cmbCuadrillaAInstalar.AddItem (.FieldValue("miembros"))
            cmbCuadrillaAInstalar.ItemData(cmbCuadrillaAInstalar.NewIndex) = .FieldValue("idcuadrilla")
            
            cmbCuadrillaTerminados.AddItem (.FieldValue("miembros"))
            cmbCuadrillaTerminados.ItemData(cmbCuadrillaTerminados.NewIndex) = .FieldValue("idcuadrilla")
            status = .GetNext
        Wend
        
        cmbCuadrillaAInstalar.AddItem "TODAS"
        cmbCuadrillaAInstalar.ListIndex = cmbCuadrillaAInstalar.ListCount - 1
    
        cmbCuadrillaTerminados.AddItem "TODAS"
        cmbCuadrillaTerminados.ListIndex = cmbCuadrillaTerminados.ListCount - 1
    End With
End Sub


Private Sub mnuCuadrilla_Click()
    frmCuadrilla.Show 1, Me
End Sub

Private Sub tablaTrabajosAProgramar_DblClick()
    If tablaTrabajosAProgramar.MouseRow > 0 And tablaTrabajosAProgramar.MouseCol >= 0 Then
        Call abrirFrmTrabajo
    End If
End Sub

Private Sub tablaTrabajosAInstalar_DblClick()
    If tablaTrabajosAInstalar.MouseRow > 0 And tablaTrabajosAInstalar.MouseCol >= 1 Then
        Call abrirFrmTrabajo
    End If
End Sub


Private Sub abrirFrmTrabajo()
    frmTrabajo.Show 1, Me
    Call cargarTablaTrabajosAProgramar
End Sub




