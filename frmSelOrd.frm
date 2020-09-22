VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7l.ocx"
Begin VB.Form frmSelOrd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordenes de Facturación de Servicios"
   ClientHeight    =   3345
   ClientLeft      =   480
   ClientTop       =   345
   ClientWidth     =   7830
   Icon            =   "frmSelOrd.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
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
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   1575
   End
   Begin VSFlex7LCtl.VSFlexGrid grilla 
      Height          =   2220
      Left            =   120
      TabIndex        =   0
      Top             =   540
      Width           =   7575
      _cx             =   13361
      _cy             =   3916
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
      BackColorFixed  =   32768
      ForeColorFixed  =   16777215
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
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   2
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSelOrd.frx":030A
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
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Órdenes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   180
      TabIndex        =   2
      Top             =   60
      Width           =   1290
   End
End
Attribute VB_Name = "frmSelOrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mCodAlumbrado As Long

Property Get CodAlumbrado() As Long
    CodAlumbrado = mCodAlumbrado
End Property

Private Sub CmdAceptar_Click()
     Call regreso
End Sub

Private Sub grilla_DblClick()
    Call regreso
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyReturn
      Call grilla_DblClick
    Case vbKeyEscape
      Unload Me
  End Select
End Sub

Private Sub Form_Load()
  lblTitulo.Caption = "Ordenes de " & main.VAClientes.FieldValue("Apellido") & ", " & main.VAClientes.FieldValue("Nombre")
  Call MostrarOrdenes(main.VAClientes.FieldValue("CodCli"))
  mCodAlumbrado = 0
End Sub

Private Sub MostrarOrdenes(Cli As Long)
  Dim st As Long
  
  With main
    .VOrdenes.IndexNumber = 1
    .VOrdenes.FieldValue("CodCli") = Cli
    .VOrdenes.GetEqual
  
    grilla.Rows = 1
    Do While (.VOrdenes.status = 0) And (.VOrdenes.FieldValue("CodCli") = Cli)
      If Not IsNull(.VOrdenes.FieldValue("CodAlumbrado")) Then
        If .VOrdenes.FieldValue("CodAlumbrado") > 0 Then
            grilla.AddItem .VOrdenes.FieldValue("NroOrden") & vbTab & _
              Format$(.VOrdenes.FieldValue("Ruta"), String$(2, "0")) & "-" & Format$(.VOrdenes.FieldValue("SubRuta"), String$(6, "0")) & vbTab & _
              .VOrdenes.FieldValue("CodAlumbrado") & vbTab & _
              .VOrdenes.FieldValue("domicilio")
          grilla.RowData(grilla.Rows - 1) = .VOrdenes.Position
        End If
      End If
      
      .VOrdenes.GetNext
    Loop
  End With
  
  grilla.AutoSize 0, grilla.Cols - 1
End Sub

Private Sub regreso()
  If grilla.Row <= 0 Then
    mCodAlumbrado = 0
    frmCambioFTTH.NroOrden = 0
  Else
    With main.VOrdenes
      ' dejar el formulario posicionado
      .IndexNumber = 4
      .FieldValue("CodAlumbrado") = mCodAlumbrado
      .GetEqual
    End With
    
    main.VAsumAlumInte.FieldValue("codalumbrado") = mCodAlumbrado
    main.VAsumAlumInte.GetEqual
    
    main.VAsumAlum.FieldValue("codalumbrado") = mCodAlumbrado
    main.VAsumAlum.GetEqual
    
    If main.VAsumAlum.status = 0 And main.VAsumAlumInte.status = 0 Then
        frmCambioFTTH.NroOrden = main.VOrdenes.FieldValue("NroOrden")
        mCodAlumbrado = Val(grilla.TextMatrix(grilla.Row, 2))
        frmCambioFTTH.lblCodInternet = "Cód. Internet " & mCodAlumbrado
        frmCambioFTTH.lblDomicilio = grilla.TextMatrix(grilla.Row, 3)
    Else
        MsgBox "El cliente no tiene un suministro de Internet - generarlo en la solapa de suministro del sistema de facturación.", vbExclamation + vbOKOnly, "Error"
    End If
    
    Unload Me
    
  End If
End Sub



