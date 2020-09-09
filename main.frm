VERSION 5.00
Object = "{47E7B6C9-8256-11CF-AB56-0000C04D1EB9}#7.0#0"; "ACBTR732.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7l.ocx"
Begin VB.Form main 
   Caption         =   "Trabajos de Internet"
   ClientHeight    =   11400
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   18000
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11400
   ScaleWidth      =   18000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCambioFTTH 
      Caption         =   "Cambiar a FTTH"
      Height          =   375
      Left            =   7680
      TabIndex        =   29
      Top             =   480
      Width           =   2655
   End
   Begin VB.Frame frmBD 
      Height          =   1455
      Left            =   7800
      TabIndex        =   10
      Top             =   7440
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
      Height          =   8535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   15055
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Para programar"
      TabPicture(0)   =   "main.frx":5D73
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tablaTrabajosAProgramar"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "btnProgRecuperar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "btnProgExcel"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "btnProgImprimir"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "frmFiltrar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Para instalar"
      TabPicture(1)   =   "main.frx":5D8F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "VSFlexGrid1"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(2)=   "Command5"
      Tab(1).Control(3)=   "Command6"
      Tab(1).Control(4)=   "Command7"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Instalados"
      TabPicture(2)   =   "main.frx":5DAB
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command11"
      Tab(2).Control(1)=   "Command10"
      Tab(2).Control(2)=   "Command9"
      Tab(2).Control(3)=   "Frame2"
      Tab(2).Control(4)=   "VSFlexGrid2"
      Tab(2).ControlCount=   5
      Begin VB.CommandButton Command11 
         BackColor       =   &H8000000E&
         Caption         =   "Recuperar"
         Height          =   615
         Left            =   -74640
         TabIndex        =   27
         Top             =   6720
         Width           =   1695
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Exportar a Excel"
         Height          =   615
         Left            =   -66480
         TabIndex        =   26
         Top             =   6480
         Width           =   1695
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Imprimir"
         Height          =   615
         Left            =   -68400
         TabIndex        =   25
         Top             =   6480
         Width           =   1695
      End
      Begin VB.Frame Frame2 
         Caption         =   "Filtrar por"
         Height          =   1395
         Left            =   -72600
         TabIndex        =   20
         Top             =   6360
         Width           =   3855
         Begin VB.ComboBox cmbInstaladosConexion 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   840
            Width           =   2055
         End
         Begin VB.ComboBox cmbInstaladosCuadrilla 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label4 
            Caption         =   "Tipo de conexión"
            Height          =   345
            Left            =   120
            TabIndex        =   24
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Cuadrilla"
            Height          =   345
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H8000000E&
         Caption         =   "Recuperar"
         Height          =   615
         Left            =   -74640
         TabIndex        =   18
         Top             =   6720
         Width           =   1695
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Exportar a Excel"
         Height          =   615
         Left            =   -66480
         TabIndex        =   17
         Top             =   6480
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Imprimir"
         Height          =   615
         Left            =   -68400
         TabIndex        =   16
         Top             =   6480
         Width           =   1695
      End
      Begin VB.Frame Frame1 
         Caption         =   "Filtrar por"
         Height          =   1395
         Left            =   -72600
         TabIndex        =   11
         Top             =   6360
         Width           =   3855
         Begin VB.ComboBox cmbInstConexion 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   840
            Width           =   2055
         End
         Begin VB.ComboBox cmbInstCuadrilla 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo de conexión"
            Height          =   345
            Left            =   120
            TabIndex        =   15
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Cuadrilla"
            Height          =   345
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame frmFiltrar 
         Caption         =   "Filtrar por"
         Height          =   1395
         Left            =   2400
         TabIndex        =   4
         Top             =   6360
         Width           =   3855
         Begin VB.ComboBox cmbProgCuadrilla 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   360
            Width           =   2055
         End
         Begin VB.ComboBox cmbProgConexion 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   840
            Width           =   2055
         End
         Begin VB.Label lblCuadrilla 
            Caption         =   "Cuadrilla"
            Height          =   345
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label lblTipoDeConexion 
            Caption         =   "Tipo de conexión"
            Height          =   345
            Left            =   120
            TabIndex        =   7
            Top             =   840
            Width           =   1335
         End
      End
      Begin VB.CommandButton btnProgImprimir 
         Caption         =   "Imprimir"
         Height          =   615
         Left            =   6600
         TabIndex        =   3
         Top             =   6480
         Width           =   1695
      End
      Begin VB.CommandButton btnProgExcel 
         Caption         =   "Exportar a Excel"
         Height          =   615
         Left            =   8520
         TabIndex        =   2
         Top             =   6480
         Width           =   1695
      End
      Begin VB.CommandButton btnProgRecuperar 
         BackColor       =   &H8000000E&
         Caption         =   "Recuperar"
         Height          =   615
         Left            =   360
         TabIndex        =   1
         Top             =   6720
         Width           =   1695
      End
      Begin VSFlex7LCtl.VSFlexGrid tablaTrabajosAProgramar 
         Height          =   5295
         Left            =   360
         TabIndex        =   9
         Top             =   960
         Width           =   10815
         _cx             =   19076
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
      Begin VSFlex7LCtl.VSFlexGrid VSFlexGrid1 
         Height          =   5295
         Left            =   -74640
         TabIndex        =   19
         Top             =   960
         Width           =   9975
         _cx             =   17595
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
         Cols            =   9
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
      Begin VSFlex7LCtl.VSFlexGrid VSFlexGrid2 
         Height          =   5295
         Left            =   -74640
         TabIndex        =   28
         Top             =   960
         Width           =   9975
         _cx             =   17595
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
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"main.frx":6041
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

Public Enum Conexiones
    ALTA_FTTH = 1
    ALTA_ANTENA = 2
    ALTA_EDIFICIO = 3
    CAMBIO_FTTH = 4
End Enum

Private arrConexiones As Variant


Private Sub btnCambioFTTH_Click()
    frmCambioFTTH.Show 1, Me
End Sub

Private Sub btnProgRecuperar_Click()
    Call cargarTablaTrabajosAProgramar
End Sub

Private Sub Form_Load()

    arrConexiones = Array("ALTA FTTH", "ALTA ANTENA", "ALTA EDIFICIO", "CAMBIO A FTTH")
    
    Call cargarCuadrillas
    Call cargarTiposConexion

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
            
            If VOrdenes.status = 0 And VAClientes.status = 0 And VAsumAlumInte.status = 0 And vTrabInternet.FieldValue("estado") = 1 Then
            
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
        
        tablaTrabajosAProgramar.AutoSize 0, tablaTrabajosAProgramar.Cols - 1
        
    End With
End Sub


Private Sub cmbFiltroCuadrilla_Change()
    ' Filtrado por cuadrilla
End Sub

Private Sub cmbFiltroTipoDeConexion_Change()
    ' Filtrado por tipo de conexión
End Sub

Private Sub btnRecuperar_Click()
    Select Case tabTrabajos.Index
        Case 0: ' Traer trabajos para programar
        Case 1: ' Traer trabajos para instalar
        Case 2: ' Traer trabajos instalados
    End Select
    
End Sub

Private Sub cargarTiposConexion()
    Dim conexion As Variant
    For Each conexion In arrConexiones
        cmbProgConexion.AddItem (conexion)
        cmbInstConexion.AddItem (conexion)
        cmbInstaladosConexion.AddItem (conexion)
    Next
End Sub

Private Sub cargarCuadrillas()
    Dim status As Integer
    
    With main.VCuadrillas
        .IndexNumber = 0
        status = .GetFirst
        
        While status = 0
            cmbProgCuadrilla.AddItem (.FieldValue("miembros"))
            cmbProgCuadrilla.ItemData(cmbProgCuadrilla.NewIndex) = .FieldValue("idcuadrilla")
            
            cmbInstCuadrilla.AddItem (.FieldValue("miembros"))
            cmbInstCuadrilla.ItemData(cmbInstCuadrilla.NewIndex) = .FieldValue("idcuadrilla")
            
            cmbInstaladosCuadrilla.AddItem (.FieldValue("miembros"))
            cmbInstaladosCuadrilla.ItemData(cmbInstaladosCuadrilla.NewIndex) = .FieldValue("idcuadrilla")
            status = .GetNext
        Wend
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


Private Sub abrirFrmTrabajo()
    frmTrabajo.Show 1, Me
    Call cargarTablaTrabajosAProgramar
End Sub




