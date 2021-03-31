Attribute VB_Name = "moduloFunciones"
Option Explicit

Private pdf As PDFCreator.clsPDFCreator
Private opt As clsPDFCreatorOptions

Private Const directorioOrdenes = "C:\OrdenesInternet\"

Public Declare Function SetDefaultPrinter Lib "winspool.drv" Alias "SetDefaultPrinterA" (ByVal pszPrinter As String) As Boolean

Public Function prepararPDFOrden(idTrabajo As Long) As String

  Call crearDirectorio(directorioOrdenes)
  
  On Error GoTo instalarPDFCreator
  
  Set pdf = New clsPDFCreator
  Set opt = New clsPDFCreatorOptions
  
  With pdf
    .cVisible = True
    If .cStart("/NoProcessingAtStartup") = False Then
      If .cStart("/NoProcessingAtStartup", True) = False Then
        Exit Function
      End If
      .cVisible = True
    End If

    Set opt = .cOptions
    .cClearCache
  End With
  
  Dim nombreOrden As String
  nombreOrden = "Orden" & idTrabajo & Trim(getNombreCuadrilla(idTrabajo)) & "-" & Format(DateTime.Now, "ddMMhhmmss") & ".pdf"
  
  With opt
    .AutosaveDirectory = directorioOrdenes
    .AutosaveFilename = nombreOrden
    .DisableUpdateCheck = True
    .StandardTitle = nombreOrden
    .UseAutosave = 1
    .UseAutosaveDirectory = 1
    .AutosaveFormat = 0  ' PDF
  End With

  Set pdf.cOptions = opt
  Set Printer = Printers(PrinterIndex("PDFCreator"))

  Call imprimirOrden(idTrabajo)
  
  pdf.cPrinterStop = False
  Sleep 1000
  
  pdf.cClose
  Set pdf = Nothing
  Set opt = Nothing
  DoEvents
  
  prepararPDFOrden = directorioOrdenes & nombreOrden
  Exit Function
  
instalarPDFCreator:
  MsgBox "No podemos generar el PDF porque en este equipo no está instalado el programa PDFCreator. Por favor, contáctenos para que lo ayudemos.", vbCritical, "No se puede generar el .pdf"
  Exit Function
  
End Function

Public Sub cambiarNoFacturar(nroOrden As Long, estadoNuevo As String)
' estadoNuevo deberá ser 0 para activado y 1 para activado

  Dim estado As Byte

  If estadoNuevo = "NOFACTURAR" Then
    estado = 1
  ElseIf estadoNuevo = "SIFACTURAR" Then
    estado = 0
  End If

  With main.VOrdenes
    .IndexNumber = 0
    .FieldValue("nroOrden") = nroOrden

    If .GetEqual = 0 Then
      .FieldValue("NOFACTURAR") = estado
      .Update
    End If
  End With
End Sub

Public Sub cargarComboPrioridad(cmbPrioridad As ComboBox)
  cmbPrioridad.AddItem ("BAJA")
  cmbPrioridad.ItemData(0) = 3

  cmbPrioridad.AddItem ("MEDIA")
  cmbPrioridad.ItemData(1) = 2

  cmbPrioridad.AddItem ("ALTA")
  cmbPrioridad.ItemData(2) = 1

  cmbPrioridad.ListIndex = 0
End Sub

Public Function getCodAlumbrado(nroOrden As Long) As Long
  With main
    .VOrdenes.IndexNumber = 0
    .VOrdenes.FieldValue("NroOrden") = nroOrden

    If .VOrdenes.GetEqual = 0 Then
      getCodAlumbrado = .VOrdenes.FieldValue("CodAlumbrado")
    End If

  End With
End Function

Public Sub actualizarTarifa(nroOrden As Long, idTarifa As Long)
' Cambia la tarifa a la seleccionada por el operador

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

Public Function getIdTarifa(idTrabajo As Long) As Long
  With main
    .vTrabInternet.IndexNumber = 0
    .vTrabInternet.FieldValue("id_trabajo") = idTrabajo

    If .vTrabInternet.GetEqual = 0 Then
      ' Si el trabajo no tiene ancho de banda definido, le asigno por defecto el de la tarifa 1002 (20 MB FTTH).
      getIdTarifa = IIf(IsNull(.vTrabInternet.FieldValue("ancho_banda")), 1002, .vTrabInternet.FieldValue("ancho_banda"))
    End If
  End With
End Function

Public Function getDescripTarifa(idTrabajo As Long) As String
  With main
    .vTrabInternet.IndexNumber = 0
    .VOrdenes.IndexNumber = 0

    .vTrabInternet.FieldValue("id_trabajo") = idTrabajo
    If .vTrabInternet.GetEqual = 0 Then

      .VOrdenes.FieldValue("NroOrden") = .vTrabInternet.FieldValue("NroOrden")
      If .VOrdenes.GetEqual = 0 Then
        .VAsumAlum.FieldValue("CodAlumbrado") = .VOrdenes.FieldValue("CodAlumbrado")
        If .VAsumAlum.GetEqual = 0 Then
          .VTarifas.FieldValue("Id_Tarifa") = .VAsumAlum.FieldValue("Id_Tarifa")
          If .VTarifas.GetEqual = 0 Then
            getDescripTarifa = .VTarifas.FieldValue("descrip")
          End If
        End If
      End If

    End If
  End With
End Function

Public Sub seleccionarPorItemData(id As Long, combo As ComboBox)
  Dim i As Long
  For i = 0 To combo.ListCount - 1
    If combo.ItemData(i) = id Then
      combo.ListIndex = i
      Exit For
    End If
  Next
End Sub

Public Sub cargarTarifasFTTH(cmbTarifas As ComboBox)
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
      st = .GetNext
    Loop

  End With
End Sub

Public Sub cargarTarifasTodas(cmbTarifas As ComboBox)
  With main.VTarifas
    Dim st As Integer

    .IndexNumber = 2
    .FieldValue("Id_Servicio") = 3
    .FieldValue("Id_Tipo") = 0

    st = .GetGreaterOrEqual

    Do While st = 0 And .FieldValue("Id_Servicio") = 3
      cmbTarifas.AddItem Format(.FieldValue("Id_Tarifa"), "0000") & " - " & (UCase(.FieldValue("descrip")))
      cmbTarifas.ItemData(cmbTarifas.NewIndex) = .FieldValue("Id_Tarifa")
      st = .GetNext
    Loop

  End With
End Sub

Public Function getNombreCuadrilla(idTrabajo) As String
  With main
    .vTrabInternet.IndexNumber = 0
    .vTrabInternet.FieldValue("id_trabajo") = idTrabajo
    If .vTrabInternet.GetEqual = 0 Then
      .VCuadrillas.IndexNumber = 0
      .VCuadrillas.FieldValue("idcuadrilla") = .vTrabInternet.FieldValue("idcuadrilla")
      If .VCuadrillas.GetEqual = 0 Then
        getNombreCuadrilla = .VCuadrillas.FieldValue("miembros")
      End If
    End If
  End With
End Function

Public Function getNroTfno(CODCLI) As String
  With main.VAClientes
    .IndexNumber = 0
    .FieldValue("CodCli") = CODCLI

    If .GetEqual = 0 Then
      getNroTfno = .FieldValue("reserva")
    Else
      getNroTfno = "-"
    End If

  End With
End Function

Public Function Izq(Texto As String, largo As Integer) As String
  If largo <= 0 Then
    Izq = vbNullString
  Else
    Izq = left$(LTrim$(Texto), largo)
    Izq = Izq & String$(largo - Len(Izq), " ")
  End If
End Function

Public Function Der(Texto As String, largo As Integer) As String
  Der = Right$(RTrim$(Texto), largo)
  Der = String$(largo - Len(Der), " ") & Der
End Function

Public Sub ponerEncabezadoEnNegrita(tabla As VSFlexGrid)
  tabla.Cell(flexcpFontBold, 0, 0, 0, tabla.Cols - 1) = True
End Sub

Public Sub imprimirOrden(idTrabajo As Long)

' =======================================================
' Variables para guardar los datos de la orden de trabajo

  Dim nroUsuario As String
  Dim nroOrden As String
  Dim tipoConexion As String
  Dim fechaInstalacion As String
  Dim horaInstalacion As String
  Dim ruta_ub As String

  Dim obs As String
  Dim Nombre As String
  Dim dni As String
  Dim domicilioFacturacion As String
  Dim domicilioConexion As String
  Dim telefono As String
  Dim celular As String
  Dim email As String
  Dim nombreUsuario As String
  Dim anchoDeBanda As String
  Dim iva As String
  Dim cuadrilla As String
  Dim ancho_banda As String

  ' =================================
  ' Cargar los datos en las variables

  With main
    .vTrabInternet.IndexNumber = 0
    .vTrabInternet.FieldValue("id_trabajo") = idTrabajo
    .VOrdenes.GetEqual

    .VOrdenes.IndexNumber = 0
    .VAClientes.IndexNumber = 0
    .VAsumAlumInte.IndexNumber = 0
    .VAsumAlum.IndexNumber = 0
    .VCuadrillas.IndexNumber = 0
    .VTarifas.IndexNumber = 0

    .VOrdenes.FieldValue("NroOrden") = .vTrabInternet.FieldValue("Nroorden")
    .VOrdenes.GetEqual

    .VAClientes.FieldValue("CodCli") = .VOrdenes.FieldValue("CodCli")
    .VAClientes.GetEqual

    .VAsumAlumInte.FieldValue("CodAlumbrado") = .VOrdenes.FieldValue("CodAlumbrado")
    .VAsumAlumInte.GetEqual

    .VAsumAlum.FieldValue("CodAlumbrado") = .VOrdenes.FieldValue("CodAlumbrado")
    .VAsumAlum.GetEqual

    .VCuadrillas.FieldValue("idcuadrilla") = .vTrabInternet.FieldValue("idcuadrilla")
    .VCuadrillas.GetEqual

    .VTarifas.FieldValue("id_tarifa") = .vTrabInternet.FieldValue("ancho_banda")
    .VTarifas.GetEqual

    If .VOrdenes.status = 0 And _
       .VAClientes.status = 0 And _
       .VAsumAlumInte.status = 0 And _
       .VAsumAlum.status = 0 And _
       .VCuadrillas.status = 0 Then
      
      ruta_ub = Format$(.VOrdenes.FieldValue("Ruta"), String$(2, "0")) & "-" & Format$(.VOrdenes.FieldValue("SubRuta"), String$(6, "0"))
      
      nroUsuario = .VAsumAlumInte.FieldValue("CodAlumbrado")
      nroOrden = .VOrdenes.FieldValue("NroOrden")
      tipoConexion = frmTrabajo.cmbTipoConexion.Text
      fechaInstalacion = .vTrabInternet.FieldValue("fecha_inst")
      horaInstalacion = .vTrabInternet.FieldValue("hora_inst")
      obs = .vTrabInternet.FieldValue("obs")

      Nombre = .VAClientes.FieldValue("nombre") & " " & .VAClientes.FieldValue("apellido")
      dni = .VAClientes.FieldValue("NroDocIde") & vbNullString
      domicilioFacturacion = .VOrdenes.FieldValue("domicilio") & vbNullString
      domicilioConexion = .VAsumAlum.FieldValue("cuenta") & vbNullString

      telefono = .VAClientes.FieldValue("reserva") & vbNullString
      nombreUsuario = .VAsumAlumInte.FieldValue("UsInt") & vbNullString
      anchoDeBanda = IIf(.VTarifas.status = 0, .VTarifas.FieldValue("descrip") & " (" & .VTarifas.FieldValue("id_tarifa") & ")", vbNullString)
      email = IIf(.VAsumAlum.FieldValue("direelec") = vbNullString, "-", .VAsumAlum.FieldValue("direelec"))
      iva = CIVADescrip(.VOrdenes.FieldValue("civa"))

      cuadrilla = .VCuadrillas.FieldValue("miembros")
    End If
  End With

  ' ===================
  ' Comienza a imprimir
  
  Dim R As Single ' renglon, para que no sea estático - tiene que ser de tipo single o no funca
  R = 20
  
  Dim pdfOrden As New LinePrinter
  With pdfOrden

    .Thin = 10
    .Fat = 16
    .Init 2, 0

    .LineH 25, R, 165
    
    .Text 25, incRenglon(R), "Cooperativa Eléctrica Integral de Provisión de Servicios Públicos", 14, True, "Arial"
    .Text 25, incRenglon(R), "y Sociales de Todd Ltda.", 14, True, "Arial"
    
    .LineH 25, incRenglon(R, 10), 165

    .Text 25, incRenglon(R), "Todd Net - Nueva conexión", 12, True, "Arial"
    .Text 25, incRenglon(R, 10), "Observaciones: " & obs, 11, True, "Arial"
    
    .Text 25, incRenglon(R), "Ruta y ubicación: " & ruta_ub, 11, False, "Arial"
    .Text 100, R, "N.º de usuario: " & nroUsuario, 11, False, "Arial"
    
    
    .Text 25, incRenglon(R), "Fecha de inst. programada: " & Format$(fechaInstalacion, "dd/MM/yyyy"), 11, False, "Arial"
    .Text 100, R, "Hora de inst. programada: " & Format$(horaInstalacion, "hh:mm AMPM"), 11, False, "Arial"
    
    .LineH 25, incRenglon(R, 10), 165

    .Text 25, incRenglon(R), "Apellido y nombre: " & Nombre, 11, False, "Arial"
    .Text 25, incRenglon(R), "DNI/CUIT: " & dni, 11, False, "Arial"
    .Text 25, incRenglon(R), "Condición IVA: " & iva, 11, False, "Arial"
    .Text 25, incRenglon(R), "Domicilio de facturación: " & domicilioFacturacion, 11, False, "Arial"
    .Text 25, incRenglon(R), "Domicilio de conexión: " & domicilioConexion, 11, False, "Arial"
    .Text 25, incRenglon(R), "Teléfono: " & telefono, 11, False, "Arial"
    .Text 25, incRenglon(R), "Correo eléctronico: " & email, 11, False, "Arial"
    .Text 25, incRenglon(R), "Nombre de usuario de Internet: " & nombreUsuario, 11, False, "Arial"
    .Text 25, incRenglon(R), "Ancho de Banda a instalar: " & anchoDeBanda, 11, False, "Arial"

    If tipoConexion = "CAMBIO A FTTH" Then
      .Text 25, incRenglon(R), "Ancho de Banda anterior: " & getDescripTarifa(idTrabajo), 11, False, "Arial"
    End If

    .LineH 25, incRenglon(R, 10), 165

    .Text 25, incRenglon(R), "Datos de la instalación", 12, True, "Arial"

    .Text 25, incRenglon(R, 10), "Cuadrilla: " & cuadrilla, 11, True, "Arial"
    .Text 100, R, "Tipo de trabajo: " & tipoConexion, 11, True, "Arial"
    .Text 25, incRenglon(R), "Fecha de inst.: ______________", 11, True, "Arial"
    .Text 100, R, "Hora de inst.: __________", 11, True, "Arial"
    
    .Text 25, incRenglon(R), "Retiro equipo: [  ]", 11, False, "Arial"
    .Text 100, R, "Retiro fuente: [  ]", 11, False, "Arial"
    
    .Text 25, incRenglon(R), "Dir. MAC:       :      :      :      :      :      ", 11, False, "Arial"
    .Text 25, incRenglon(R), "N.º de fibra: ____________", 11, False, "Arial"
    .Text 25, incRenglon(R), "Nombre y ubicación de la caja: ____________", 11, False, "Arial"
    .Text 25, incRenglon(R), "Cant. de cables (mts.): ____________", 11, False, "Arial"
    .Text 25, incRenglon(R), "Cant. de conectores: ____________", 11, False, "Arial"
    
    If tipoConexion = "CAMBIO A FTTH" And TieneIpPublica(nroOrden) Then
        .Text 100, R, "¡¡¡EL USUARIO TIENE IP PÚBLICA!!!", 12, True, "Arial"
    End If
    
    .Text 25, incRenglon(R), "Cant. de mordazas: ____________", 11, False, "Arial"
    .Text 25, incRenglon(R), "Cant. de cadenas: ____________", 11, False, "Arial"
    .Text 25, incRenglon(R), "Patchcord: ____________", 11, False, "Arial"
    .Text 25, incRenglon(R), "Potencia en abonado: ____________", 11, False, "Arial"
    .Text 25, incRenglon(R), "IP: ____________", 11, False, "Arial"
    .Text 25, incRenglon(R), "Cant. de cable UTP: ____________", 11, False, "Arial"
    .Text 25, incRenglon(R), "Cant. de fichas RJ45: ____________", 11, False, "Arial"
    
    .Text 25, incRenglon(R), "¿Antena retirada? [  ]", 11, False, "Arial"
    .Text 100, R, "Dir. MAC de la antena:       :      :      :      :      :      ", 11, False, "Arial"

    .LineH 25, incRenglon(R, 10), 165

    .Text 25, incRenglon(R, 10), "Conformidad del usuario", 11, False, "Arial"
    .Text 150, R, "Nombre del instalador", 11, False, "Arial"
    .Text 25, incRenglon(R, 10), "____________________", 11, False, "Arial"
    .Text 150, R, "____________________", 11, False, "Arial"

    .Text 25, incRenglon(R, 20), "Acepto bases y condiciones del servicio de Internet indicadas en la página web www.todd.com.ar", 8, False, "Arial"

    .SendToPrinter
  End With

End Sub

Public Function incRenglon(ByRef renglon As Single, Optional cant As Single = 5) As Single
  renglon = renglon + cant
  incRenglon = renglon
End Function

Public Function TieneIpPublica(nroOrden As String) As Boolean
    With main.VSVarios
        .IndexNumber = 0
        .FieldValue("NroOrden") = nroOrden
        .FieldValue("ID_CTO") = 6023 ' nro. cto. IP pública
        
        If .GetEqual = 0 Then
            If .FieldValue("Importe") > 0 Then
                TieneIpPublica = True
                Exit Function
            End If
        End If
        
        TieneIpPublica = False
    End With
End Function

Public Function crearDirectorio(rutaDirectorio As String)
  If Dir(rutaDirectorio, vbDirectory) = "" Then
    MkDir rutaDirectorio
  End If
End Function

' Establecer impresora
Public Function PrinterIndex(PrinterName As String) As Long
  Dim i As Long

  For i = 0 To Printers.Count - 1
    If UCase$(Printers(i).DeviceName) = UCase$(PrinterName) Then
      PrinterIndex = i
      Exit For
    End If
  Next
End Function


Public Function EsCorreoValido(correo As String) As Boolean
  Dim pos As Integer
  Dim nCorreo As Integer

  ' posicion del arroba
  pos = InStr(2, correo, "@")
  If (pos < 1) Or (pos > (Len(correo) - 5)) Then
    ' no tiene arroba a partir del 2do caracter, o tiene un sufijo menor a cuatro caracteres
    EsCorreoValido = False
  Else
    EsCorreoValido = True
  End If
End Function

' valida la lista de correos separada por punto y coma
Public Function ValidarCorreos(listaCorreos As String) As Boolean
  Dim correos() As String
  Dim nCorreo As Byte
  Dim valido As Boolean

  ' sacar espacios del textbox
  listaCorreos = Replace(listaCorreos, " ", vbNullString)
  If listaCorreos = vbNullString Then
    ' no puso nada, salir nomasss
    ValidarCorreos = True
    Exit Function
  End If

  correos = Split(listaCorreos, ";")
  For nCorreo = LBound(correos) To UBound(correos)
    valido = EsCorreoValido(correos(nCorreo))
    If Not valido Then
      Exit For
    End If
  Next

  ValidarCorreos = valido
End Function


Public Function numeroOrden(idTrabajo As Long) As Long
  With main.vTrabInternet
    .IndexNumber = 0
    .FieldValue("idtrabajo") = idTrabajo
    .GetEqual

    If .status = 0 Then
      numeroOrden = .FieldValue("nroorden")
    Else
      numeroOrden = 0
    End If
  End With
End Function

Public Function idTipoConexion(idTrabajo As Long) As Byte
  With main.vTrabInternet
    .IndexNumber = 0
    .FieldValue("tipo_conexion") = idTrabajo
    .GetEqual

    If .status = 0 Then
      idTipoConexion = .FieldValue("nroorden")
    Else
      idTipoConexion = 0
    End If
  End With
End Function

Public Function stringTipoConexion(idTipoCon As Byte) As String
  stringTipoConexion = IIf(idTipoCon < 1, "Algo salió mal", main.arrConexiones(idTipoCon - 1))
End Function

Public Sub definirImpresora(nombreImpresora As String)
  Dim Impresora As Printer
  If Printers.Count > 0 Then
    For Each Impresora In Printers
      If Impresora.DeviceName = nombreImpresora Then
        Set Printer = Impresora
        Exit For
      End If
    Next Impresora
  End If
End Sub

