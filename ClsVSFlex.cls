VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsVSFlex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Franco 2019

Private anchoCols() As Integer

Private mGrilla As VSFlexGrid

' Para los listados
Private mTitulo As String
Private mSubtitulo As String
Private mRazonSocial As String


Public Property Get grilla() As VSFlexGrid
  Set grilla = mGrilla
End Property

Public Property Let grilla(ByRef vNewValue As VSFlexGrid)
  Set mGrilla = vNewValue
  Call Inicializar
End Property

Public Property Get Titulo() As String
  Titulo = mTitulo
End Property

Public Property Let Titulo(ByRef vNewValue As String)
  mTitulo = vNewValue
End Property

Public Property Get Subtitulo() As String
  Subtitulo = mSubtitulo
End Property

Public Property Let Subtitulo(ByRef vNewValue As String)
  mSubtitulo = vNewValue
End Property

Public Property Get RazonSocial() As String
  RazonSocial = mRazonSocial
End Property

Public Property Let RazonSocial(ByRef vNewValue As String)
  mRazonSocial = vNewValue
End Property


' constantes de exportacion

Public Property Get vsExpTodo() As Byte
  vsExpTodo = 0
End Property

Public Property Get vsExpConMarca() As Byte
  vsExpConMarca = 1
End Property

Public Property Get vsExpSinMarca() As Byte
  vsExpSinMarca = 2
End Property


' Funciones para VSFlexGrid

Private Sub Inicializar()
  With grilla
    ' dejar encabezado en negrita
    .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True

    ' tipo de refresco
    .Redraw = flexRDBuffered
  End With
End Sub


Public Sub NegritaFila(nroFila As Long, Valor As Boolean)
  With grilla
    .Cell(flexcpFontBold, nroFila, 0, nroFila, .Cols - 1) = Valor
  End With
End Sub

Public Sub ItalicaFila(nroFila As Long, Valor As Boolean)
  With grilla
    .Cell(flexcpFontItalic, nroFila, 0, nroFila, .Cols - 1) = Valor
  End With
End Sub

Public Sub ResaltarFila(nroFila As Long, Valor As Boolean, Optional color As Long = vbYellow)
  With grilla
    If .FixedCols > 0 Then
      .Cell(flexcpBackColor, nroFila, 0, nroFila, .FixedCols - 1) = IIf(Valor, color, 0)
    End If
    If .FixedCols < .Cols Then
      .Cell(flexcpBackColor, nroFila, .FixedCols, nroFila, .Cols - 1) = IIf(Valor, color, 0)
    End If
  End With
End Sub

Public Function FilaResaltada(nroFila As Long) As Boolean
  With grilla
    FilaResaltada = (.Cell(flexcpBackColor, nroFila, 0, nroFila, 0) <> 0)
  End With
End Function

Public Sub MarcarFila(nroFila As Long, Valor As Boolean, Optional Col As Integer = 0)
  grilla.Cell(flexcpChecked, nroFila, Col, nroFila, Col) = IIf(Valor, flexChecked, flexUnchecked)
End Sub

Public Function FilaMarcada(nroFila As Long, Optional Col As Integer = 0) As Boolean
  Select Case grilla.Cell(flexcpChecked, nroFila, Col, nroFila, Col)
  Case flexChecked
    FilaMarcada = True
  Case flexUnchecked
    FilaMarcada = False
  Case Else
    ' no tiene para marcar
    FilaMarcada = True
  End Select
End Function

Public Function ColumnaOculta(nroCol As Integer) As Boolean
  ColumnaOculta = (grilla.ColHidden(nroCol) Or (grilla.ColWidth(nroCol) <= 0))
End Function

Public Function FilaOculta(nroFila As Long) As Boolean
  FilaOculta = (grilla.RowHidden(nroFila) Or (grilla.RowHeight(nroFila) <= 0))
End Function


Private Function MostrarFila(nroFila As Long, filasAMostrar As Byte) As Boolean
  If FilaOculta(nroFila) Then
    ' nunca mostrar filas ocultas
    MostrarFila = False
    Exit Function
  End If

  If nroFila < grilla.FixedCols Then
    ' siempre mostrar encabezados
    MostrarFila = True
    Exit Function
  End If

  Select Case filasAMostrar
  Case vsExpTodo
    MostrarFila = True
  Case vsExpConMarca
    MostrarFila = FilaMarcada(nroFila)
  Case vsExpSinMarca
    MostrarFila = (Not FilaMarcada(nroFila))
  Case Else
    MostrarFila = True
  End Select
End Function

Public Sub ExportarCsv(nombreArch As String, Optional filasAMostrar As Byte = 0, Optional renglonFinal As String = vbNullString)
  Const SEP = ";"

  Dim nArch As Integer
  Dim renglon() As String
  Dim nroFila As Long
  Dim nroCol As Integer
  Dim nroCols As Integer
  Dim c As Integer
  Dim s As String

  If grilla.Rows < 1 Then Exit Sub
  If nombreArch = vbNullString Then Exit Sub

  On Error GoTo ErrExp
  If Dir(nombreArch) <> vbNullString Then
    If MsgBox("El archivo ya existe, ¿reemplazarlo?", vbYesNo + vbQuestion, Titulo) = vbYes Then
      Kill nombreArch
    Else
      Exit Sub
    End If
  End If

  nArch = FreeFile
  Open nombreArch For Output As #nArch
  On Error GoTo 0

  Screen.MousePointer = vbHourglass

  Print #nArch, "sep=" & SEP
  If mRazonSocial <> vbNullString Then Print #nArch, mRazonSocial
  If mTitulo <> vbNullString Then Print #nArch, mTitulo
  If mSubtitulo <> vbNullString Then Print #nArch, mSubtitulo
  Print #nArch, vbNullString

  With grilla
    ' contar columnas visibles
    nroCols = 0
    For nroCol = 0 To .Cols - 1
      If Not ColumnaOculta(nroCol) Then nroCols = nroCols + 1
    Next

    ' inicializar vector
    ReDim renglon(nroCols)

    For nroFila = 0 To .Rows - 1
      If MostrarFila(nroFila, filasAMostrar) Then
        c = 0
        For nroCol = 0 To .Cols - 1
          If Not ColumnaOculta(nroCol) Then
            ' cargar vector, reemplazando SEP y "
            renglon(c) = Replace(.TextMatrix(nroFila, nroCol), SEP, " ")
            renglon(c) = Replace(renglon(c), """", " ")
            c = c + 1
          End If
        Next
        ' exportar renglon
        Print #nArch, Join(renglon, SEP)
      End If
    Next
  End With

  If renglonFinal <> vbNullString Then
    Print #nArch, vbNullString
    Print #nArch, renglonFinal
  End If

  Close #nArch
  Screen.MousePointer = vbDefault

  MsgBox "Archivo " & nombreArch & " exportado exitosamente.", vbOKOnly, Titulo
  Shell "explorer.exe /select," & nombreArch, vbNormalFocus
  Exit Sub

ErrExp:
  If Err.Number = 70 Then
    s = "Permiso denegado. Verifique que el archivo no se encuentre abierto."
  Else
    s = "Error " & Err.Number & " - " & Err.Description
  End If
  MsgBox s, vbOKOnly + vbExclamation, Titulo
  On Error GoTo 0
End Sub


' *** Impresion ***

Private Sub CalcularAnchosImp(margenIzq As Integer, tamano As Integer)
  Dim anchoTwips() As Long
  Dim nroCol As Integer
  Dim cantCols As Integer
  Dim anchoTot As Currency
  Dim maxChar As Integer

  anchoTot = 0
  cantCols = 0

  With grilla
    ReDim anchoCols(.Cols)
    ReDim anchoTwips(.Cols)

    ' guardar anchos en twips y calcular total
    For nroCol = 0 To .Cols - 1
      If ColumnaOculta(nroCol) Then
        anchoTwips(nroCol) = 0
      Else
        anchoTot = anchoTot + .ColWidth(nroCol)
        anchoTwips(nroCol) = .ColWidth(nroCol)
        cantCols = cantCols + 1
      End If
    Next

    ' cantidad de caracteres que entran con este tamano
    Printer.FontSize = tamano
    maxChar = Int((Printer.ScaleWidth - margenIzq * 2) / Printer.TextWidth("O")) + 1

    ' dejarlo proporcional al total, en caracteres
    For nroCol = 0 To .Cols - 1
      anchoCols(nroCol) = (anchoTwips(nroCol) / anchoTot) * (maxChar - cantCols + 1)
    Next
  End With
End Sub

Private Function Celda2StringImp(nroFila As Long, nroCol As Integer) As String
  With grilla
    If .Cell(flexcpAlignment, nroFila, nroCol, nroFila, nroCol) = flexAlignRightCenter Then
      Celda2StringImp = Der(.TextMatrix(nroFila, nroCol), anchoCols(nroCol))
    Else
      Celda2StringImp = Izq(.TextMatrix(nroFila, nroCol), anchoCols(nroCol))
    End If
  End With
End Function

Private Function RenglonImp(nroFila As Long) As String
  Dim renglon As String
  Dim nroCol As Integer

  renglon = vbNullString

  With grilla
    For nroCol = 0 To .Cols - 1
      If Not ColumnaOculta(nroCol) Then
        renglon = renglon & Celda2StringImp(nroFila, nroCol) & " "
      End If
    Next
  End With

  RenglonImp = renglon
End Function

Public Sub Imprimir(Optional filasAMostrar As Byte = 0, Optional orientacion As Byte = vbPRORPortrait, Optional tamano As Integer = 8, Optional renglonFinal As String = vbNullString)
  Dim list As New Listado
  Dim nroFila As Long
  Dim orientAnt As Byte

  Screen.MousePointer = vbHourglass
  orientAnt = Printer.Orientation
  Printer.Orientation = orientacion

  With list
    .empresa = mRazonSocial
    .Titulo = mTitulo
    .Subtitulo = mSubtitulo

    .Tipo = 0  ' con encabezamiento
    .margenIzq = 300
    .MargenSup = 300
    .LetraListado = tamano

    Call CalcularAnchosImp(.margenIzq, tamano)
    .Columnas = RenglonImp(0)

    .Iniciar

    For nroFila = 1 To grilla.Rows - 1
      If MostrarFila(nroFila, filasAMostrar) Then
        .renglon RenglonImp(nroFila)
      End If
    Next

    If renglonFinal <> vbNullString Then
      .renglon "-"
      Printer.FontBold = True
      .renglon renglonFinal
      Printer.FontBold = False
    End If

    .Fin
  End With

  Printer.Orientation = orientAnt
  Screen.MousePointer = vbDefault
End Sub
