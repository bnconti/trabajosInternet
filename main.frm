VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7l.ocx"
Object = "{47E7B6C9-8256-11CF-AB56-0000C04D1EB9}#7.0#0"; "ACBTR732.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trabajos de Internet"
   ClientHeight    =   8475
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   21135
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   21135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E6E6E7&
      Cancel          =   -1  'True
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   19680
      Picture         =   "main.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6480
      Width           =   1125
   End
   Begin VB.CommandButton btnCambioFTTH 
      BackColor       =   &H00E6E6E7&
      Caption         =   "Cambiar a FTTH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   480
      Width           =   2295
   End
   Begin VB.Frame frmBD 
      Height          =   1215
      Left            =   14280
      TabIndex        =   5
      Top             =   7440
      Visible         =   0   'False
      Width           =   2775
      Begin VAccessLib.VAccess VDatosConexInet 
         Left            =   1080
         Top             =   720
         _Version        =   458752
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
         VAccessName     =   "VDatosConexInet"
         TableName       =   "DATOSCONEXINET"
         Location        =   "\\servidor\D\Compu\SFS2000\Datos\DATOSCONEXINET.mkd"
         DdfPath         =   "\\servidor\D\Compu\SFS2000\Datos"
         VAUDDDFInfo     =   "main.frx":2D6D
      End
      Begin VAccessLib.VAccess VTarifas 
         Left            =   600
         Top             =   720
         _Version        =   458752
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
         VAccessName     =   "VTarifas"
         TableName       =   "TARIFAS"
         Location        =   "\\servidor\D\Compu\SFS2000\Datos\TARIFAS.MKD"
         OpenMode        =   2
         DdfPath         =   "\\servidor\D\Compu\SFS2000\Datos"
         VAUDDDFInfo     =   "main.frx":372E
      End
      Begin VAccessLib.VAccess VAsumAlum 
         Left            =   2040
         Top             =   240
         _Version        =   458752
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
         VAccessName     =   "VAsumAlum"
         TableName       =   "ASUMALUM"
         Location        =   "\\servidor\D\Compu\SFS2000\Datos\ASUMALUM.mkd"
         DdfPath         =   "\\servidor\D\Compu\SFS2000\Datos"
         VAUDDDFInfo     =   "main.frx":40EF
      End
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
         VAUDDDFInfo     =   "main.frx":4C78
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
         DdfPath         =   "\\servidor\compu\SFS2000\datos"
         HostConnect     =   0   'False
         VAUDDDFInfo     =   "main.frx":5613
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
         VAUDDDFInfo     =   "main.frx":62A6
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
         VAUDDDFInfo     =   "main.frx":6E2F
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
         VAUDDDFInfo     =   "main.frx":78FA
      End
   End
   Begin TabDlg.SSTab tabTrabajos 
      Height          =   8490
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   21120
      _ExtentX        =   37253
      _ExtentY        =   14975
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Para programar"
      TabPicture(0)   =   "main.frx":8295
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tablaTrabajosAProgramar"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frmFiltrar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "btnExpExcelProgramar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "btnImprimirProgramar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "btnAProgramarRecuperar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Para instalar"
      TabPicture(1)   =   "main.frx":82B1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frmTotalesAInstalar"
      Tab(1).Control(1)=   "btnEnviarOrdenPorCorreo"
      Tab(1).Control(2)=   "btnAInstalarRecuperar"
      Tab(1).Control(3)=   "btnImprimirInstalar"
      Tab(1).Control(4)=   "btnExpExcelInstalar"
      Tab(1).Control(5)=   "btnGuardarFinalizados"
      Tab(1).Control(6)=   "Frame1"
      Tab(1).Control(7)=   "tablaTrabajosAInstalar"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Terminados"
      TabPicture(2)   =   "main.frx":82CD
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "tablaTrabajosTerminados"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "frmFiltrado"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "btnExpExcelTerminados"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "btnImprimirTerminados"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "btnInstaladosRecuperar"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Frame3"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      Begin VB.Frame Frame3 
         Caption         =   "Totales de trabajos terminados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -67440
         TabIndex        =   57
         Top             =   6360
         Width           =   6255
         Begin VB.Label Label19 
            Caption         =   "Total:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   67
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label18 
            Caption         =   "Alta FTTH:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   66
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label17 
            Caption         =   "Alta Altena:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   65
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label16 
            Caption         =   "Alta Edificio:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   64
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label15 
            Caption         =   "Cambio a FTTH:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   63
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label lbCambioFTTHTotalTerminados 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   4800
            TabIndex        =   62
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label lblAltaEdificioTotalTerminados 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   4800
            TabIndex        =   61
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lblAltaAntenaTotalTerminados 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   1440
            TabIndex        =   60
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label lblAltaFTTHTotalTerminados 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   1440
            TabIndex        =   59
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lblTrabajosTerminadosTotal 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   1440
            TabIndex        =   58
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Totales de trabajos para programar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   5880
         TabIndex        =   46
         Top             =   6360
         Width           =   6255
         Begin VB.Label Label14 
            Caption         =   "Total:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   56
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label13 
            Caption         =   "Alta FTTH:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   55
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label12 
            Caption         =   "Alta Altena:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   54
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label11 
            Caption         =   "Alta Edificio:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   53
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label10 
            Caption         =   "Cambio a FTTH:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   52
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label lbCambioFTTHTotalParaProgramar 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   4800
            TabIndex        =   51
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label lblAltaEdificioTotalParaProgramar 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   4800
            TabIndex        =   50
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lblAltaAntenaTotalParaProgramar 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   1440
            TabIndex        =   49
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label lblAltaFTTHTotalParaProgramar 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   1440
            TabIndex        =   48
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lblTrabajosParaProgramarTotal 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   1440
            TabIndex        =   47
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame frmTotalesAInstalar 
         Caption         =   "Totales de trabajos para instalar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -69120
         TabIndex        =   35
         Top             =   6360
         Width           =   6255
         Begin VB.Label lblTrabajosParaInstalarTotal 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   1440
            TabIndex        =   45
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label lblAltaFTTHTotalParaInstalar 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   1440
            TabIndex        =   44
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lblAltaAntenaTotalParaInstalar 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   1440
            TabIndex        =   43
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label lblAltaEdificioTotalParaInstalar 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   4800
            TabIndex        =   42
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lbCambioFTTHTotalParaInstalar 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   4800
            TabIndex        =   41
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label lblCambioFTTHParaInstalar 
            Caption         =   "Cambio a FTTH:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   40
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label lblAltaEdificioParaInstalar 
            Caption         =   "Alta Edificio:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   39
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblAltaAntenaParaInstalar 
            Caption         =   "Alta Altena:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   38
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label lblAltaFTTHParaInstalar 
            Caption         =   "Alta FTTH:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   37
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblTotalParaInstalar 
            Caption         =   "Total:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.CommandButton btnEnviarOrdenPorCorreo 
         BackColor       =   &H00F7EBDD&
         Caption         =   "Enviar correo a la cuadrilla"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -59760
         Picture         =   "main.frx":82E9
         TabIndex        =   34
         Top             =   6480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton btnInstaladosRecuperar 
         BackColor       =   &H00CCF2FF&
         Caption         =   "Recuperar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74640
         Picture         =   "main.frx":91B3
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   6480
         Width           =   2175
      End
      Begin VB.CommandButton btnAInstalarRecuperar 
         BackColor       =   &H00F7EBDD&
         Caption         =   "Recuperar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74640
         Picture         =   "main.frx":EDC5
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   6480
         Width           =   2175
      End
      Begin VB.CommandButton btnAProgramarRecuperar 
         BackColor       =   &H00DAEFE2&
         Caption         =   "Recuperar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         Picture         =   "main.frx":149D7
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   6480
         Width           =   2175
      End
      Begin VB.CommandButton btnImprimirProgramar 
         BackColor       =   &H00DAEFE2&
         Caption         =   "Imprimir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   16800
         Picture         =   "main.frx":1A5E9
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   6480
         Width           =   1215
      End
      Begin VB.CommandButton btnImprimirInstalar 
         BackColor       =   &H00F7EBDD&
         Caption         =   "Imprimir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -58200
         Picture         =   "main.frx":20873
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   6480
         Width           =   1215
      End
      Begin VB.CommandButton btnImprimirTerminados 
         BackColor       =   &H00CCF2FF&
         Caption         =   "Imprimir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -58200
         Picture         =   "main.frx":26AFD
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   6480
         Width           =   1215
      End
      Begin VB.CommandButton btnExpExcelTerminados 
         BackColor       =   &H00CCF2FF&
         Caption         =   "Exportar a Excel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -56760
         Picture         =   "main.frx":2CD87
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   6480
         Width           =   1215
      End
      Begin VB.CommandButton btnExpExcelInstalar 
         BackColor       =   &H00F7EBDD&
         Caption         =   "Exportar a Excel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -56760
         Picture         =   "main.frx":2D841
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   6480
         Width           =   1215
      End
      Begin VB.CommandButton btnExpExcelProgramar 
         BackColor       =   &H00DAEFE2&
         Caption         =   "Exportar a Excel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   18240
         Picture         =   "main.frx":2E2FB
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   6480
         Width           =   1215
      End
      Begin VB.CommandButton btnGuardarFinalizados 
         BackColor       =   &H00F7EBDD&
         Caption         =   "Guardar finalizados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74640
         Picture         =   "main.frx":2EDB5
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   7440
         Width           =   2175
      End
      Begin VB.Frame frmFiltrado 
         Caption         =   "Filtrar por"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1875
         Left            =   -72120
         TabIndex        =   22
         Top             =   6360
         Width           =   4455
         Begin MSComCtl2.DTPicker dtDesdeTerminados 
            Height          =   375
            Left            =   2400
            TabIndex        =   28
            Top             =   600
            Width           =   1695
            _ExtentX        =   2990
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
            CustomFormat    =   "dd/MM/yy"
            Format          =   94371843
            CurrentDate     =   44089
         End
         Begin VB.ComboBox cmbConexionTerminados 
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
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   1320
            Width           =   2055
         End
         Begin VB.ComboBox cmbCuadrillaTerminados 
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
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   600
            Width           =   2055
         End
         Begin MSComCtl2.DTPicker dtHastaTerminados 
            Height          =   375
            Left            =   2400
            TabIndex        =   30
            Top             =   1320
            Width           =   1695
            _ExtentX        =   2990
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
            CustomFormat    =   "dd/MM/yy"
            Format          =   94371843
            CurrentDate     =   44089
         End
         Begin VB.Label lblFechaHastaInstalados 
            Caption         =   "Fecha hasta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2400
            TabIndex        =   29
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label lblFechaDesdeInstalados 
            Caption         =   "Fecha desde"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2400
            TabIndex        =   27
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Tipo de conexión"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   25
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label Label3 
            Caption         =   "Cuadrilla"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Filtrar por"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1875
         Left            =   -72120
         TabIndex        =   13
         Top             =   6360
         Width           =   2775
         Begin VB.ComboBox cmbConexionAInstalar 
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
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   1320
            Width           =   2535
         End
         Begin VB.ComboBox cmbCuadrillaAInstalar 
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
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   600
            Width           =   2535
         End
         Begin VB.Label Label2 
            Caption         =   "Tipo de conexión"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   16
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Cuadrilla"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame frmFiltrar 
         Caption         =   "Filtrar por"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   2880
         TabIndex        =   2
         Top             =   6360
         Width           =   2775
         Begin VB.ComboBox cmbConexionAProgramar 
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
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   600
            Width           =   2295
         End
         Begin VB.Label lblTipoDeConexion 
            Caption         =   "Tipo de conexión"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   1815
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid tablaTrabajosAProgramar 
         Height          =   5295
         Left            =   360
         TabIndex        =   31
         Top             =   960
         Width           =   20415
         _cx             =   36010
         _cy             =   9340
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
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
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"main.frx":2F0BF
         ScrollTrack     =   -1  'True
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
         ExplorerBar     =   5
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
         Left            =   -74640
         TabIndex        =   10
         Top             =   960
         Width           =   20415
         _cx             =   36010
         _cy             =   9340
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
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
         Cols            =   12
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"main.frx":2F1F3
         ScrollTrack     =   -1  'True
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
         ExplorerBar     =   5
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
      Begin VSFlex7LCtl.VSFlexGrid tablaTrabajosTerminados 
         Height          =   5295
         Left            =   -74640
         TabIndex        =   20
         Top             =   960
         Width           =   20415
         _cx             =   36010
         _cy             =   9340
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
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
         FormatString    =   $"main.frx":2F3AA
         ScrollTrack     =   -1  'True
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
         ExplorerBar     =   5
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
         Caption         =   "Cuadrillas"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuProcesos 
      Caption         =   "Procesos"
      Begin VB.Menu mnuCambioFTTH 
         Caption         =   "Cambio a FTTH"
      End
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

Public Enum prioridad
  ALTA = 1
  MEDIA = 2
  BAJA = 3
End Enum

Public ini As New ArchIni

Private Const CHEQUEADO As Integer = 1

Private Const COL_ID_TRABAJO As Integer = 10

Public arrConexiones As Variant

Dim cVSFlex As New ClsVSFlex

Dim FUENTEDATOS As String

Private Sub btnEnviarOrdenPorCorreo_Click()
  With tablaTrabajosAInstalar
    If tablaTrabajosAInstalar.Row > 1 Then
      ' Dim idTrabajo =
      ' Call prepararCorreo(idTrabajo)
    Else
      MsgBox "Tiene que seleccionar el trabajo del cual se enviará la orden.", vbInformation + vbOKOnly, "No se seleccionó un trabajo"
    End If
  End With
End Sub

Private Sub cmdSalir_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  arrConexiones = Array("ALTA FTTH", "ALTA ANTENA", "ALTA EDIFICIO", "CAMBIO A FTTH")
  FUENTEDATOS = GetSetting("SFS2000", "Sistema", "FuenteDeDatos", "\\Servidor\D\Compu\SFS2000\Datos\")

  Call cargarCuadrillas
  Call cargarTiposConexion

  ini.Path_del_Ini = FUENTEDATOS & "SFS.ini"

  dtDesdeTerminados.Value = "01/01/2020"
  dtHastaTerminados.Value = DateTime.Day(DateTime.Now) & "/" & DateTime.Month(DateTime.Now) & "/" & DateTime.Year(DateTime.Now) + 1

  Call formatearEncabezados
  tabTrabajos.Tab = 0
End Sub

Private Sub btnCambioFTTH_Click()
  frmCambioFTTH.Show 1, Me
  Call cargarTablaTrabajosAProgramar
End Sub

Private Sub btnAProgramarRecuperar_Click()
  Call cargarTablaTrabajosAProgramar
End Sub

Private Sub btnAInstalarRecuperar_Click()
  Call cargarTablaTrabajosAInstalar
End Sub

Private Sub btnInstaladosRecuperar_Click()
  Call cargartablaTrabajosTerminados
End Sub

Private Sub cargarTablaTrabajosAProgramar()
  Dim st As Integer

  tablaTrabajosAProgramar.Rows = 1

  With vTrabInternet
    vTrabInternet.IndexNumber = 0
    st = .GetFirst

    Screen.MousePointer = 11
    While st = 0

      VOrdenes.IndexNumber = 0
      VAClientes.IndexNumber = 0
      VAsumAlumInte.IndexNumber = 0
      VAsumAlum.IndexNumber = 0

      VOrdenes.FieldValue("NroOrden") = .FieldValue("Nroorden")
      VOrdenes.GetEqual

      VAClientes.FieldValue("CodCli") = VOrdenes.FieldValue("CodCli")
      VAClientes.GetEqual

      VAsumAlumInte.FieldValue("CodAlumbrado") = VOrdenes.FieldValue("CodAlumbrado")
      VAsumAlumInte.GetEqual

      VAsumAlum.FieldValue("CodAlumbrado") = VOrdenes.FieldValue("CodAlumbrado")
      VAsumAlum.GetEqual

      If VOrdenes.status = 0 And _
         VAClientes.status = 0 And _
         VAsumAlumInte.status = 0 And _
         VAsumAlum.status = 0 And _
         vTrabInternet.FieldValue("estado") = Estados.NUEVO Then

        tablaTrabajosAProgramar.AddItem (VAClientes.FieldValue("apellido") & ", " & VAClientes.FieldValue("nombre") & vbTab & _
                                         VAsumAlum.FieldValue("cuenta") & vbTab & _
                                         VAsumAlumInte.FieldValue("UsInt") & vbTab & _
                                         arrConexiones(vTrabInternet.FieldValue("Tipo_Conexion") - 1) & vbTab & _
                                         vTrabInternet.FieldValue("fecha_pedido") & vbTab & _
                                         VAClientes.FieldValue("reserva") & vbTab & _
                                         vTrabInternet.FieldValue("id_trabajo")) & vbTab & _
                                         vTrabInternet.FieldValue("obs")


        On Error Resume Next
        If vTrabInternet.FieldValue("prioridad") = prioridad.ALTA Then
          ' Pintar fila rojo
          tablaTrabajosAProgramar.Cell(flexcpBackColor, tablaTrabajosAProgramar.Rows - 1, 0, tablaTrabajosAProgramar.Rows - 1, tablaTrabajosAProgramar.Cols - 1) = RGB(255, 122, 122)
        ElseIf vTrabInternet.FieldValue("prioridad") = prioridad.MEDIA Then
          ' Pintar fila amarillo
          tablaTrabajosAProgramar.Cell(flexcpBackColor, tablaTrabajosAProgramar.Rows - 1, 0, tablaTrabajosAProgramar.Rows - 1, tablaTrabajosAProgramar.Cols - 1) = RGB(255, 255, 122)
        End If
        On Error GoTo 0


      End If

      st = .GetNext

    Wend

    Call filtrarTablaAProgramar
    tablaTrabajosAProgramar.AutoSize 0, tablaTrabajosAProgramar.Cols - 1
    Screen.MousePointer = 0

  End With
End Sub

Private Sub cargarTablaTrabajosAInstalar()
  Dim st As Integer

  tablaTrabajosAInstalar.Rows = 1

  vTrabInternet.IndexNumber = 0
  st = vTrabInternet.GetFirst

  Screen.MousePointer = 11
  While st = 0

    VOrdenes.IndexNumber = 0
    VAClientes.IndexNumber = 0
    VAsumAlumInte.IndexNumber = 0
    VCuadrillas.IndexNumber = 0
    VAsumAlum.IndexNumber = 0

    VOrdenes.FieldValue("NroOrden") = vTrabInternet.FieldValue("Nroorden")
    VOrdenes.GetEqual

    VAClientes.FieldValue("CodCli") = VOrdenes.FieldValue("CodCli")
    VAClientes.GetEqual

    VAsumAlumInte.FieldValue("CodAlumbrado") = VOrdenes.FieldValue("CodAlumbrado")
    VAsumAlumInte.GetEqual

    VCuadrillas.FieldValue("idcuadrilla") = vTrabInternet.FieldValue("idcuadrilla")
    VCuadrillas.GetEqual

    VAsumAlum.FieldValue("CodAlumbrado") = VOrdenes.FieldValue("CodAlumbrado")
    VAsumAlum.GetEqual

    If VOrdenes.status = 0 And _
       VAClientes.status = 0 And _
       VAsumAlumInte.status = 0 And _
       VCuadrillas.status = 0 And _
       VAsumAlum.status = 0 And _
       vTrabInternet.FieldValue("estado") = Estados.PROGRAMADO Then

      tablaTrabajosAInstalar.AddItem (vbNullString & vbTab & _
                                      VAClientes.FieldValue("apellido") & ", " & VAClientes.FieldValue("nombre") & vbTab & _
                                      VAsumAlum.FieldValue("cuenta") & vbTab & _
                                      VAsumAlumInte.FieldValue("UsInt") & vbTab & _
                                      arrConexiones(vTrabInternet.FieldValue("Tipo_Conexion") - 1) & vbTab & _
                                      vTrabInternet.FieldValue("fecha_pedido") & vbTab & _
                                      VAClientes.FieldValue("reserva") & vbTab & _
                                      vTrabInternet.FieldValue("fecha_inst") & vbTab & _
                                      Format(vTrabInternet.FieldValue("hora_inst"), "hh:mm AMPM") & vbTab & _
                                      VCuadrillas.FieldValue("miembros") & vbTab & _
                                      vTrabInternet.FieldValue("id_trabajo")) & vbTab & _
                                      vTrabInternet.FieldValue("obs")
      tablaTrabajosAInstalar.Cell(flexcpChecked, tablaTrabajosAInstalar.Rows - 1, 0, tablaTrabajosAInstalar.Rows - 1, 0) = flexUnchecked

      If vTrabInternet.FieldValue("prioridad") = prioridad.ALTA Then
        ' Pintar fila rojo
        tablaTrabajosAInstalar.Cell(flexcpBackColor, tablaTrabajosAInstalar.Rows - 1, 0, tablaTrabajosAInstalar.Rows - 1, tablaTrabajosAInstalar.Cols - 1) = RGB(255, 122, 122)
      ElseIf vTrabInternet.FieldValue("prioridad") = prioridad.MEDIA Then
        ' Pintar fila amarillo
        tablaTrabajosAInstalar.Cell(flexcpBackColor, tablaTrabajosAInstalar.Rows - 1, 0, tablaTrabajosAInstalar.Rows - 1, tablaTrabajosAInstalar.Cols - 1) = RGB(255, 255, 122)
      End If

    End If

    st = vTrabInternet.GetNext

  Wend

  Call filtrarTablaAInstalar

  tablaTrabajosAInstalar.AutoSize 0, tablaTrabajosAInstalar.Cols - 1
  Screen.MousePointer = 0

End Sub

Private Sub actualizarLabelsTablaParaInstalar()

  lblTrabajosParaInstalarTotal.Caption = 0
  lblAltaFTTHTotalParaInstalar.Caption = 0
  lblAltaAntenaTotalParaInstalar.Caption = 0
  lblAltaEdificioTotalParaInstalar.Caption = 0
  lbCambioFTTHTotalParaInstalar.Caption = 0

  Dim i As Long

  With tablaTrabajosAInstalar
    If .Rows > 1 Then

      For i = 1 To .Rows - 1
        If Not (.RowHidden(i)) Then
          lblTrabajosParaInstalarTotal.Caption = Val(lblTrabajosParaInstalarTotal.Caption) + 1

          Dim tipoConexion As String
          tipoConexion = .TextMatrix(i, 4)

          Select Case tipoConexion
          Case "ALTA FTTH": lblAltaFTTHTotalParaInstalar.Caption = Val(lblAltaFTTHTotalParaInstalar.Caption) + 1
          Case "ALTA ANTENA": lblAltaAntenaTotalParaInstalar.Caption = Val(lblAltaAntenaTotalParaInstalar.Caption) + 1
          Case "ALTA EDIFICIO": lblAltaEdificioTotalParaInstalar.Caption = Val(lblAltaEdificioTotalParaInstalar.Caption) + 1
          Case "CAMBIO A FTTH": lbCambioFTTHTotalParaInstalar.Caption = Val(lbCambioFTTHTotalParaInstalar.Caption) + 1
          End Select
        End If
      Next
    End If
  End With

End Sub

Private Sub actualizarLabelsTablaParaProgramar()
  lblTrabajosParaProgramarTotal.Caption = 0
  lblAltaFTTHTotalParaProgramar.Caption = 0
  lblAltaAntenaTotalParaProgramar.Caption = 0
  lblAltaEdificioTotalParaProgramar.Caption = 0
  lbCambioFTTHTotalParaProgramar.Caption = 0

  Dim i As Long

  With tablaTrabajosAProgramar
    If .Rows > 1 Then

      For i = 1 To .Rows - 1
        If Not (.RowHidden(i)) Then
          lblTrabajosParaProgramarTotal.Caption = Val(lblTrabajosParaProgramarTotal.Caption) + 1

          Dim tipoConexion As String
          tipoConexion = .TextMatrix(i, 3)

          Select Case tipoConexion
          Case "ALTA FTTH": lblAltaFTTHTotalParaProgramar.Caption = Val(lblAltaFTTHTotalParaProgramar.Caption) + 1
          Case "ALTA ANTENA": lblAltaAntenaTotalParaProgramar.Caption = Val(lblAltaAntenaTotalParaProgramar.Caption) + 1
          Case "ALTA EDIFICIO": lblAltaEdificioTotalParaProgramar.Caption = Val(lblAltaEdificioTotalParaProgramar.Caption) + 1
          Case "CAMBIO A FTTH": lbCambioFTTHTotalParaProgramar.Caption = Val(lbCambioFTTHTotalParaProgramar.Caption) + 1
          End Select
        End If
      Next
    End If
  End With

End Sub

Private Sub actualizarLabelsTablaTerminados()
  lblTrabajosTerminadosTotal.Caption = 0
  lblAltaFTTHTotalTerminados.Caption = 0
  lblAltaAntenaTotalTerminados.Caption = 0
  lblAltaEdificioTotalTerminados.Caption = 0
  lbCambioFTTHTotalTerminados.Caption = 0

  Dim i As Long

  With tablaTrabajosTerminados
    If .Rows > 1 Then

      For i = 1 To .Rows - 1
        If Not (.RowHidden(i)) Then
          lblTrabajosTerminadosTotal.Caption = Val(lblTrabajosTerminadosTotal.Caption) + 1

          Dim tipoConexion As String
          tipoConexion = .TextMatrix(i, 3)

          Select Case tipoConexion
          Case "ALTA FTTH": lblAltaFTTHTotalTerminados.Caption = Val(lblAltaFTTHTotalTerminados.Caption) + 1
          Case "ALTA ANTENA": lblAltaAntenaTotalTerminados.Caption = Val(lblAltaAntenaTotalTerminados.Caption) + 1
          Case "ALTA EDIFICIO": lblAltaEdificioTotalTerminados.Caption = Val(lblAltaEdificioTotalTerminados.Caption) + 1
          Case "CAMBIO A FTTH": lbCambioFTTHTotalTerminados.Caption = Val(lbCambioFTTHTotalTerminados.Caption) + 1
          End Select
        End If
      Next
    End If
  End With

End Sub

Private Sub cargartablaTrabajosTerminados()
  Dim st As Integer

  tablaTrabajosTerminados.Rows = 1

  With vTrabInternet
    vTrabInternet.IndexNumber = 0
    st = .GetFirst

    Screen.MousePointer = 0
    While st = 0

      VOrdenes.IndexNumber = 0
      VAClientes.IndexNumber = 0
      VAsumAlumInte.IndexNumber = 0
      VCuadrillas.IndexNumber = 0
      VAsumAlum.IndexNumber = 0

      VOrdenes.FieldValue("NroOrden") = .FieldValue("Nroorden")
      VOrdenes.GetEqual

      VAClientes.FieldValue("CodCli") = VOrdenes.FieldValue("CodCli")
      VAClientes.GetEqual

      VAsumAlumInte.FieldValue("CodAlumbrado") = VOrdenes.FieldValue("CodAlumbrado")
      VAsumAlumInte.GetEqual

      VCuadrillas.FieldValue("idcuadrilla") = vTrabInternet.FieldValue("idcuadrilla")
      VCuadrillas.GetEqual

      VAsumAlum.FieldValue("CodAlumbrado") = VOrdenes.FieldValue("CodAlumbrado")
      VAsumAlum.GetEqual

      If VOrdenes.status = 0 And _
         VAClientes.status = 0 And _
         VAsumAlumInte.status = 0 And _
         VCuadrillas.status = 0 And _
         VAsumAlum.status = 0 And _
         vTrabInternet.FieldValue("estado") = Estados.TERMINADO Then

        tablaTrabajosTerminados.AddItem (VAClientes.FieldValue("apellido") & ", " & VAClientes.FieldValue("nombre") & vbTab & _
                                         VAsumAlum.FieldValue("cuenta") & vbTab & _
                                         VAsumAlumInte.FieldValue("UsInt") & vbTab & _
                                         arrConexiones(vTrabInternet.FieldValue("Tipo_Conexion") - 1) & vbTab & _
                                         vTrabInternet.FieldValue("fecha_pedido") & vbTab & _
                                         VAClientes.FieldValue("reserva") & vbTab & _
                                         vTrabInternet.FieldValue("fecha_inst") & vbTab & _
                                         Format(vTrabInternet.FieldValue("hora_inst"), "hh:mm AMPM") & vbTab & _
                                         VCuadrillas.FieldValue("miembros") & vbTab & _
                                         vTrabInternet.FieldValue("id_trabajo")) & vbTab & _
                                         vTrabInternet.FieldValue("obs")

        If vTrabInternet.FieldValue("prioridad") = prioridad.ALTA Then
          ' Pintar fila rojo
          tablaTrabajosTerminados.Cell(flexcpBackColor, tablaTrabajosTerminados.Rows - 1, 0, tablaTrabajosTerminados.Rows - 1, tablaTrabajosTerminados.Cols - 1) = RGB(255, 122, 122)
        ElseIf vTrabInternet.FieldValue("prioridad") = prioridad.MEDIA Then
          ' Pintar fila amarillo
          tablaTrabajosTerminados.Cell(flexcpBackColor, tablaTrabajosTerminados.Rows - 1, 0, tablaTrabajosTerminados.Rows - 1, tablaTrabajosTerminados.Cols - 1) = RGB(255, 255, 122)
        End If

      End If

      st = .GetNext

    Wend

    Call filtrarTablaTerminados
    tablaTrabajosTerminados.AutoSize 0, tablaTrabajosTerminados.Cols - 1
    Screen.MousePointer = 0

  End With
End Sub

Private Sub btnGuardarFinalizados_Click()
  If tablaTrabajosAInstalar.Rows > 1 Then

    Dim idTrabajo As Integer

    Dim fila As Integer
    Dim ultimaFila As Integer
    ultimaFila = tablaTrabajosAInstalar.Rows - 1

    With tablaTrabajosAInstalar
      Screen.MousePointer = 0
      For fila = 1 To ultimaFila
        If .Cell(flexcpChecked, fila, 0, fila, 0) = CHEQUEADO Then
          idTrabajo = .TextMatrix(fila, 10)
          Call finalizarTrabajo(idTrabajo)
        End If
      Next fila
      Screen.MousePointer = 11
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

    Call cambiarNoFacturar(.FieldValue("nroOrden"), "SIFACTURAR")

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
    Screen.MousePointer = 11
    For fila = 1 To ultimaFila
      conexion = .TextMatrix(fila, 3)
      If conexion = cmbConexionAProgramar.Text Or cmbConexionAProgramar.Text = "TODAS" Then
        .RowHidden(fila) = False
      Else
        .RowHidden(fila) = True
      End If
    Next fila

    Call actualizarLabelsTablaParaProgramar

    Screen.MousePointer = 0
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

    Screen.MousePointer = 11
    For fila = 1 To ultimaFila
      conexion = .TextMatrix(fila, 4)
      cuadrilla = .TextMatrix(fila, 9)
      If (conexion = cmbConexionAInstalar.Text Or cmbConexionAInstalar.Text = "TODAS") _
         And (cuadrilla = cmbCuadrillaAInstalar.Text Or cmbCuadrillaAInstalar.Text = "TODAS") Then
        .RowHidden(fila) = False
      Else
        .RowHidden(fila) = True
      End If
    Next fila

    Call actualizarLabelsTablaParaInstalar

    Screen.MousePointer = 0
  End With
End Sub

Private Sub cmbCuadrillaTerminados_Click()
  Call filtrarTablaTerminados
End Sub

Private Sub cmbConexionTerminados_Click()
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
  ultimaFila = tablaTrabajosTerminados.Rows - 1

  Dim conexion As String
  Dim cuadrilla As String
  Dim fechaInstalacion As Date

  With tablaTrabajosTerminados
    For fila = 1 To ultimaFila
      conexion = .TextMatrix(fila, 3)
      cuadrilla = .TextMatrix(fila, 8)
      fechaInstalacion = .TextMatrix(fila, 6)
      If (conexion = cmbConexionTerminados.Text Or cmbConexionTerminados.Text = "TODAS") _
         And (cuadrilla = cmbCuadrillaTerminados.Text Or cmbCuadrillaTerminados.Text = "TODAS") _
         And (dtDesdeTerminados <= fechaInstalacion) _
         And (dtHastaTerminados >= fechaInstalacion) _
         Then
        .RowHidden(fila) = False
      Else
        .RowHidden(fila) = True
      End If
    Next fila

    Call actualizarLabelsTablaTerminados

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

Public Sub cargarCuadrillas()
  Dim status As Integer

  ' Los borro por si llamo a la funcion de vuelta
  cmbCuadrillaAInstalar.Clear
  cmbCuadrillaTerminados.Clear

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


Private Sub Form_Unload(Cancel As Integer)
' Por si queda algo abierto o procesando algo, que se carga de vuelta el formulario y queda el proceso como zombie...
  End
End Sub


Private Sub mnuCuadrilla_Click()
  frmCuadrilla.Show 1, Me
End Sub

Private Sub mnuSalir_Click()
  Unload Me
End Sub

Private Sub tablaTrabajosAProgramar_DblClick()
  If tablaTrabajosAProgramar.MouseRow > 0 And tablaTrabajosAProgramar.MouseCol >= 0 Then
    Call abrirFrmTrabajo
    Call cargarTablaTrabajosAProgramar
  End If
End Sub

Private Sub tablaTrabajosAProgramar_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  If Col <> 1 Then
    Cancel = True
  End If
End Sub

Private Sub tablaTrabajosAInstalar_DblClick()
  If tablaTrabajosAInstalar.MouseRow > 0 And tablaTrabajosAInstalar.MouseCol >= 1 Then
    Call abrirFrmTrabajo
    Call cargarTablaTrabajosAInstalar
  End If
End Sub

Private Sub tablaTrabajosTerminados_DblClick()
  If tablaTrabajosTerminados.MouseRow > 0 And tablaTrabajosTerminados.MouseCol >= 0 Then
    Call abrirFrmTrabajo
    Call cargartablaTrabajosTerminados
  End If
End Sub

Private Sub abrirFrmTrabajo()
  frmTrabajo.Show 1, Me
End Sub

Private Sub btnExpExcelProgramar_Click()
  Call exportarExcel(tablaTrabajosAProgramar)
End Sub

Private Sub btnExpExcelInstalar_Click()
  Call exportarExcel(tablaTrabajosAInstalar)
End Sub

Private Sub btnExpExcelTerminados_Click()
  Call exportarExcel(tablaTrabajosTerminados)
End Sub


Private Sub exportarExcel(tabla As VSFlexGrid)
  If tabla.Rows < 2 Then
    MsgBox "No hay datos que exportar, ¿recuperó la tabla?", vbOKOnly + vbInformation, "Fallo al exportar"
  Else

    Dim rutaArchivo As String

    Dim rutaBase As String
    rutaBase = "C:\ExcelTrabajosInternet\"

    If Dir(rutaBase, vbDirectory) = "" Then
      MkDir rutaBase
    End If

    rutaArchivo = InputBox("Indique el destino del archivo ", "Exportar a Excel ", rutaBase & "LibroIVA" & Format(DateTime.Now, "ddMMyyhhss") & ".csv")

    If rutaArchivo <> vbNullString Then
      Screen.MousePointer = 11

      Open rutaArchivo For Output As #1

      Write #1, "sep=,"

      Dim Titulo As String
      Titulo = getTituloExcel(tabla.Name)

      Write #1, Titulo
      Write #1, "Listado generado el " & Format$(DateTime.Now, "dd/MM/yyyy") & " a las " & Format$(DateTime.Now, "hh:mm AMPM")

      If cmbCuadrillaAInstalar.Text <> "TODAS" And tabla.Name = "tablaTrabajosAInstalar" Then
        Write #1, "Trabajos filtrados para la cuadrilla " & UCase(cmbCuadrillaAInstalar.Text)
      End If

      Write #1, vbNullString

      With tabla
        Dim fila As Integer
        Dim ultimaFila As Integer
        ultimaFila = .Rows - 1
        For fila = 0 To ultimaFila
          If Not (tabla.RowHidden(fila)) Then
            If tabla.Name = "tablaTrabajosAProgramar" Then
              Write #1, .TextMatrix(fila, 0), .TextMatrix(fila, 1), .TextMatrix(fila, 2), .TextMatrix(fila, 3), .TextMatrix(fila, 4), .TextMatrix(fila, 5), .TextMatrix(fila, 7)
            ElseIf tabla.Name = "tablaTrabajosAInstalar" Then
              Write #1, .TextMatrix(fila, 1), .TextMatrix(fila, 2), .TextMatrix(fila, 3), .TextMatrix(fila, 4), .TextMatrix(fila, 6), .TextMatrix(fila, 7), .TextMatrix(fila, 8), .TextMatrix(fila, 11)
            ElseIf tabla.Name = "tablaTrabajosTerminados" Then
              Write #1, .TextMatrix(fila, 1), .TextMatrix(fila, 2), .TextMatrix(fila, 3), .TextMatrix(fila, 4), .TextMatrix(fila, 5), .TextMatrix(fila, 6), .TextMatrix(fila, 7), .TextMatrix(fila, 8), .TextMatrix(fila, 10)
            End If
          End If
        Next fila
      End With

      Close #1
      Screen.MousePointer = 0

      Shell "explorer.exe /select, " & rutaArchivo, vbNormalFocus
    End If
  End If

End Sub

Private Function getTituloExcel(nombreTabla As String) As String
  If nombreTabla = "tablaTrabajosAProgramar" Then
    getTituloExcel = "Trabajos sin programar"
  ElseIf nombreTabla = "tablaTrabajosAInstalar" Then
    getTituloExcel = "Trabajos pendientes"
  ElseIf nombreTabla = "tablaTrabajosTerminados" Then
    getTituloExcel = "Trabajos terminados"
  End If
End Function

Private Sub btnImprimirProgramar_Click()
  Call imprimirTabla(tablaTrabajosAProgramar)
End Sub

Private Sub btnImprimirInstalar_Click()
  tablaTrabajosAInstalar.ColHidden(0) = True
  Call imprimirTabla(tablaTrabajosAInstalar)
  tablaTrabajosAInstalar.ColHidden(0) = False
End Sub

Private Sub btnImprimirTerminados_Click()
  Call imprimirTabla(tablaTrabajosTerminados)
End Sub

Private Sub imprimirTabla(tabla As VSFlexGrid)
  If tabla.Rows > 1 Then
    With cVSFlex
      .grilla = tabla
      .RazonSocial = ini.GetVar("Empresa", "RazonSocial")

      .Titulo = getTituloExcel(tabla.Name)
      .Subtitulo = "Listado generado el " & Format$(DateTime.Now, "dd/MM/yyyy") & " a las " & Format$(DateTime.Now, "hh:mm AMPM")

      Call .Imprimir(, vbPRORLandscape, 6)
    End With
  Else
    MsgBox "No hay datos que exportar, ¿recuperó la tabla?", vbOKOnly + vbInformation, "Fallo al exportar"
  End If
End Sub

Private Sub formatearEncabezados()
  Call ponerEncabezadoEnNegrita(tablaTrabajosAProgramar)
  Call ponerEncabezadoEnNegrita(tablaTrabajosAInstalar)
  Call ponerEncabezadoEnNegrita(tablaTrabajosTerminados)
  tablaTrabajosAProgramar.AutoSize 0, tablaTrabajosAProgramar.Cols - 1
  tablaTrabajosAInstalar.AutoSize 0, tablaTrabajosAInstalar.Cols - 1
  tablaTrabajosTerminados.AutoSize 0, tablaTrabajosTerminados.Cols - 1
End Sub

