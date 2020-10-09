VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsFlex7l.ocx"
Begin VB.Form frmCuadrilla 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuadrillas"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9690
   Icon            =   "frmCuadrilla.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   9690
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnNuevaCuadrilla 
      Caption         =   "Nueva cuadrilla"
      Height          =   735
      Left            =   120
      Picture         =   "frmCuadrilla.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   1695
   End
   Begin VSFlex7LCtl.VSFlexGrid tablaCuadrillas 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      _cx             =   16748
      _cy             =   4471
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmCuadrilla.frx":09D2
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Para editar una cuadrilla, haga doble clic sobre la celda correspondiente y modifique el dato."
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   2760
      Width           =   6735
   End
End
Attribute VB_Name = "frmCuadrilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const COL_HABILITADO As Integer = 3
Private datoViejo As String
Private correoInvalido As Boolean

Private Sub Form_Load()
    Call ponerEncabezadoEnNegrita(tablaCuadrillas)
    Call cargarCuadrillas
    correoInvalido = False
End Sub

Private Sub btnNuevaCuadrilla_Click()
    frmCuadrillaNueva.Show 1, Me
    Call cargarCuadrillas
    ' Actualizarla tambien en el formulario principal
    Call main.cargarCuadrillas
End Sub

Private Sub cargarCuadrillas()
    Dim status As Integer
    tablaCuadrillas.Rows = 1
    
    With main.VCuadrillas
        .IndexNumber = 0
        status = .GetFirst
        
        While status = 0
            tablaCuadrillas.AddItem (.FieldValue("idcuadrilla") & vbTab & _
                                    .FieldValue("miembros") & vbTab & _
                                    .FieldValue("email") & vbTab)
            If .FieldValue("habilitado") Then
                tablaCuadrillas.Cell(flexcpChecked, tablaCuadrillas.Rows - 1, COL_HABILITADO, tablaCuadrillas.Rows - 1, COL_HABILITADO) = flexChecked
            Else
                tablaCuadrillas.Cell(flexcpChecked, tablaCuadrillas.Rows - 1, COL_HABILITADO, tablaCuadrillas.Rows - 1, COL_HABILITADO) = flexUnchecked
            End If
            
            status = .GetNext
        Wend
    End With
    
    tablaCuadrillas.AutoSize 1, 2
End Sub

Private Sub tablaCuadrillas_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not correoInvalido Then
        datoViejo = tablaCuadrillas.TextMatrix(Row, Col)
    End If
End Sub

Private Sub tablaCuadrillas_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = COL_HABILITADO Then
        modificarCuadrillaHabilitado Row, Col
    Else
        modificarCuadrilla Row, Col
    End If
End Sub

Private Sub modificarCuadrillaHabilitado(ByVal Row As Long, ByVal Col As Long)
    Dim estaHabilitado As Boolean
    estaHabilitado = (tablaCuadrillas.Cell(flexcpChecked, Row, COL_HABILITADO, Row, COL_HABILITADO) = flexChecked)
    
    Dim msj As String
    msj = "¿Está seguro de querer " & IIf(estaHabilitado, "habilitar", "deshabilitar") & " esta cuadrilla?"
    
    Dim resp As String
    resp = MsgBox(msj, vbYesNo, "Confirmación")
    
    If resp = vbYes Then
        Dim id As Integer
        id = tablaCuadrillas.TextMatrix(Row, 0)
        
        With main.VCuadrillas
            .IndexNumber = 0
            .FieldValue("idcuadrilla") = id
            .GetEqual
             
            If .status = 0 Then
                .FieldValue("habilitado") = estaHabilitado
                .Update
            End If
            
        End With
        
    ElseIf resp = vbNo Then
        ' Deja la celda como estaba antes de modificarla
        tablaCuadrillas.Cell(flexcpChecked, Row, COL_HABILITADO, Row, COL_HABILITADO) = IIf(Not estaHabilitado, flexChecked, flexUnchecked)
    End If
    
End Sub

Private Sub modificarCuadrilla(ByVal Row As Long, ByVal Col As Long)
    Dim datoNuevo As String
    
    ' Le saco los espacios iniciales y finales
    datoNuevo = Trim$(tablaCuadrillas.TextMatrix(Row, Col))
    
    If datoViejo = datoNuevo Then
        ' No actualizar de gusto para dejar la misma cosa
        Exit Sub
    End If
    
    If Col = 1 And datoNuevo = vbNullString Then
        ' No permitir sacar los miembros!
        tablaCuadrillas.TextMatrix(Row, Col) = datoViejo
        Exit Sub
    End If
    
    If Col = 2 Then
        ' correos
        ' sacar espacios
        tablaCuadrillas.TextMatrix(Row, Col) = Replace(tablaCuadrillas.TextMatrix(Row, Col), " ", vbNullString)
        If Not ValidarCorreos(tablaCuadrillas.TextMatrix(Row, Col)) Then
            Call MsgBox("Ingrese una dirección válida", vbOKOnly + vbInformation, Me.Caption)
            correoInvalido = True
            tablaCuadrillas.Select Row, Col
            tablaCuadrillas.EditCell
            Exit Sub
        End If
    End If
    
    correoInvalido = False
    
    Dim msj As String
    ' Detalle: encierro los datos entre comillas
    msj = "¿Está seguro de querer modificar """ & datoViejo & """ por """ & datoNuevo & """?"
    
    If MsgBox(msj, vbYesNo, "Confirmación") = vbYes Then
        Dim idCuadrilla As Integer
        idCuadrilla = tablaCuadrillas.TextMatrix(Row, 0)
        
        With main.VCuadrillas
            .IndexNumber = 0
            .FieldValue("idcuadrilla") = idCuadrilla
            .GetEqual
            
            If .status = 0 Then
                Select Case Col
                    Case 1: .FieldValue("miembros") = datoNuevo
                    Case 2: .FieldValue("email") = datoNuevo
                End Select
                
                .Update
                
                If .status = 0 Then
                    If Col = 1 Then
                        ' Actualizar comboboxes si cambiaron los miembros
                        ' (lo hice publico en main.....)
                        Call main.cargarCuadrillas
                    End If
                Else
                    ' si no se pudo actualizar por algun motivo, mostrar el dato viejo
                    ' (puede darse por ej si le pone los miembros que a otra cuadrilla)
                    tablaCuadrillas.TextMatrix(Row, Col) = datoViejo
                End If
            End If
            
        End With
    Else
        tablaCuadrillas.TextMatrix(Row, Col) = datoViejo
    End If
End Sub



Private Sub tablaCuadrillas_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        If Col = 2 Then
            If correoInvalido Then
                tablaCuadrillas.TextMatrix(Row, Col) = datoViejo
                ' o algo asi
            End If
        End If
    End If
End Sub
