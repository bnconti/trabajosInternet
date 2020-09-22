VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsInf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim fSuscrip As New LinePrinter
Private xr As Single 'posicion x del ultimo renglon
Private XY As Single 'posicion y de la columna
Dim K As Integer
Dim stUsua As String
Dim StServ As String
Dim StNroFac As String
Dim StImporte As String
Dim CONT As Integer
Dim SumaPar As Currency
Dim Saldo As Currency
Dim IMPLET(1 To 2) As String
'Variables para textos de posicion fija en Suscripcion

     Private mCOOP As String
     Private mCoop_RazonSocial As String
     Private mCoop_Direccion As String
     Private MPERIODO As String
     


Public Property Get COOP() As String
  COOP = mCOOP
End Property

Public Property Let COOP(ByVal vNewValue As String)
   mCOOP = vNewValue
End Property

Public Property Get Coop_RazonSocial() As String
  Coop_RazonSocial = mCoop_RazonSocial
End Property

Public Property Let Coop_RazonSocial(ByVal vNewValue As String)
   mCoop_RazonSocial = vNewValue
End Property

Public Property Get Coop_Direccion() As String
  Coop_Direccion = mCoop_Direccion
End Property

Public Property Let Coop_Direccion(ByVal vNewValue As String)
   mCoop_Direccion = vNewValue
End Property



'* * * * * * * * * * * * * * * * * * * * * * * * *
'       M E T O D O S   D E   L A   C L A S E
'* * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub Init(Optional pict)
 Dim TipoDocFinal As String
 Dim EnLetras(1 To 2) As String
  With fSuscrip
    
 '   .Thin = 10
'    .Fat = 30
    .Thin = 10
    .Fat = 16
    .Init 2, 0, pict  ', vbPRPSLetter

 '   If NuevaImpresion Then
 '       .DoDocument , "2mm", "0mm", 0, 0
 '   Else
 '       .NewPage
 '   End If
    
    'Call Membrete(pv, 0, "RENTAS")
'     .Bmp 25, 5, "\\Servidor\D\Compu\Sfs2000\rentas.bmp", True, 162, 130
'    .Bmp 177, 10, Sys.FuenteDeDatos & "\rentasnro.bmp", True, 0, 0
    '*****************************************************************************
    'CUADROS Y LINEAS
    '*****************************************************************************
    .LineH 25, 20, 165
    
    .LineH 25, 35, 165
 
 
    .LineV 158, 47, 175
    '*****************************************************************************
    'TEXTOS
    '*****************************************************************************
    .Text 25, 40, main.ini.GetVar("empresa", "razonsocial"), 12, True, "Times New Roman"
    
    .Text 25, 23, "Liquidación de Cobranza de Tasa de Alumbrado Público y compensación con suministros y servicios ", 11, True, "Times New Roman"
    .Text 25, 28, "prestados por la Cooperativa, según Convenio de fecha 15 de Marzo de 2002.-", 11, True, "Times New Roman"
      
      MPERIODO = Format(Month(main.pERIODO), "00") & "/" & Year(main.pERIODO)

    'MPERIODO = "MARZO 2002"
    .Text 130, 37, "PERÍODO DE FACTURACIÓN ", 12, True, "Times New Roman"
    .Text 130, 42, MPERIODO, 12, True, "Times New Roman"
    
    
    
    
    .Text 130, 47, "Municipalidad", 10, True, "Times New Roman"
    .Text 130, 51, "de Pergamino", 10, True, "Times New Roman"
    
    .Text 160, 47, "Cooperativa de", 10, True, "Times New Roman"
    .Text 160, 51, "El Socorro Ltda", 10, True, "Times New Roman"
    
    
    .Text 25, 60, " NRO.US                                  SERVICIO                                                    FACT.Nº        IMPORTE   ", 7, True, "Times New Roman"
    CONT = 0
   For K = 1 To Val(main.cTotalMov)
    main.Lista.Row = K
    main.Lista.col = 0
    If main.Lista.Text = "X" Then
        CONT = CONT + 1
        
        main.Lista.col = 2
        stUsua = main.Lista.Text
       
        main.Lista.col = 7
        StServ = main.Lista.Text
        
        main.Lista.col = 3
        StNroFac = main.Lista.Text
        main.Lista.col = 6
        StImporte = main.Lista.Text
        .Text 25, 60 + (CONT * 3.5), Der(stUsua, 8) & "    " & Izq(StServ, 37), 7, False, "Times New Roman"
        .Text 85, 60 + (CONT * 3.5), Der(StNroFac, 15), 7, False, "Times New Roman"
        
        .Text 105, 60 + (CONT * 3.5), Der(StImporte, 15), 7, False, "Times New Roman"
    End If
  Next K

'Importe Ley
.Text 25, 140, "Total Facturas:", 12, True, "Times New Roman"
.Text 140, 140, Der(main.tmSALDO, 27), 14, True, "Times New Roman"


.Text 25, 150, "Percepción Art. 72 Ter. De la Ley 11.769", 12, True, "Times New Roman"
.Text 140, 150, Der("(" & Trim(main.TxtImpLey) & ")", 27), 14, True, "Times New Roman"

.Text 25, 160, "Impuesto Ley 13.404", 12, True, "Times New Roman"
.Text 140, 160, Der("(" & Trim(main.TxtAgua) & ")", 27), 14, True, "Times New Roman"



.Text 25, 170, "Tasa Alumbrado Público:", 12, True, "Times New Roman"

.Text 35, 180, "   CANTIDAD DE           ", 12, False, "Times New Roman"
.Text 35, 185, "CONTRIBUYENTES      ", 12, False, "Times New Roman"

.Text 42, 192, main.TxtCantTasa, 14, True, "Times New Roman"

'.Text 79, 192, "$" & Trim(TxtValorTasa), 14, True, "Times New Roman"

.Text 115, 192, Der(main.TxtTasa, 27), 14, True, "Times New Roman"


.Text 115, 202, Der(main.TxtTasa, 27), 14, True, "Times New Roman"
SumaPar = Val(main.tmSALDO) - (Val(main.TxtImpLey) + Val(main.TxtAgua))

.Text 140, 202, Der(Format(SumaPar, "fixed"), 27), 14, True, "Times New Roman"
Saldo = Val(main.TxtTasa) - SumaPar

.Text 110, 215, "SALDO", 15, True, "Times New Roman", True
If Saldo > 0 Then
     .Text 112, 215, Der(Format(Saldo, "fixed"), 27), 15, True, "Times New Roman", True
    Else
     .Text 140, 215, Der(Format(Abs(Saldo), "fixed"), 27), 15, True, "Times New Roman", True
End If
Call NumALet(Abs(CDbl(Saldo)), 74, IMPLET())

      

.Text 25, 225, "Son " & IMPLET(1) & IMPLET(2), 12, False, "Times New Roman"

.Text 25, 230, "Documentación adjunta:", 12, False, "Times New Roman"
.Text 25, 235, "Facturas por suministros y servicios, listado de facturación Tasa Alumbrado Público y listado", 12, False, "Times New Roman"
.Text 25, 240, "percepción Ley 11769 Art. 72 ter,  correspondientes al período liquidado", 12, False, "Times New Roman"
'
.Text 25, 250, "Pergamino, " & Date, 12, False, "Times New Roman"

 End With
 Printer.EndDoc
End Sub



Public Sub SendToPrinter()
  fSuscrip.SendToPrinter
End Sub

Public Sub Text(Optional ByVal X, Optional ByVal Y, Optional ByVal texto, Optional size, Optional bold, Optional font)
  fSuscrip.Text X, Y, texto, size, bold, font
End Sub

Public Sub Obs(ByVal RenglonNro As Byte, ByVal texto As String)
  fSuscrip.Text 21, 193 + ((RenglonNro - 1) * 4), texto, 9
End Sub

Public Sub Consumos(ByVal RenglonNro As Byte, ByVal texto As String)
  fSuscrip.Text 135, 51 + ((RenglonNro - 1) * 3.2), texto, 8, , "Courier New"
End Sub

Public Sub Renglon(ByVal nro As Byte, ByVal texto As String, Optional cantidad, Optional precio, Optional Importe)
Dim aux As Single
  xr = 103 + ((nro - 1) * 4)
  aux = IIf(mGrandes, 116, 21)
  fSuscrip.Text aux, xr, Izq(texto, 24), 9, , "Courier New"
  If Not IsMissing(cantidad) Then fSuscrip.Text aux + 47, xr, Der(Format(cantidad, "0"), 7), 9, , "Courier New"
  If Not IsMissing(precio) Then fSuscrip.Text aux + 63, xr, Der(Format(precio, "0.0000"), 7), 9, , "Courier New"
  If Not IsMissing(Importe) Then fSuscrip.Text aux + 74, xr, Der(Format(Importe, "0.00"), 10), 9, , "Courier New"
End Sub

Public Sub Renglon2(ByVal nro As Byte, ByVal texto As String, Optional cantidad, Optional precio, Optional Importe)
Dim aux As Single
  xr = 103 + ((nro - 1) * 4)
  aux = IIf(mGrandes, 21, 116)
  fSuscrip.Text aux, xr, Izq(texto, 24), 9, , "Courier New"
  If Not IsMissing(cantidad) Then fSuscrip.Text aux + 47, xr, Der(Format(cantidad, "0"), 7), 9, , "Courier New"
  If Not IsMissing(precio) Then fSuscrip.Text aux + 63, xr, Der(Format(precio, "0.0000"), 7), 9, , "Courier New"
  If Not IsMissing(Importe) Then fSuscrip.Text aux + 74, xr, Der(Format(Importe, "0.00"), 10), 9, , "Courier New"
End Sub

