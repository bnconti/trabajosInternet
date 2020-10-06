Attribute VB_Name = "moduloFunciones"
Option Explicit

Public Declare Function SetDefaultPrinter Lib "winspool.drv" Alias "SetDefaultPrinterA" (ByVal pszPrinter As String) As Boolean

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
    Dim tipoConexion As String
    Dim fechaInstalacion As String
    Dim horaInstalacion As String
    
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
        
        .VTarifas.FieldValue("id_tarifa") = .VAsumAlum.FieldValue("id_tarifa")
        .VTarifas.GetEqual
        
        If .VOrdenes.status = 0 And _
            .VAClientes.status = 0 And _
            .VAsumAlumInte.status = 0 And _
            .VAsumAlum.status = 0 And _
            .VCuadrillas.status = 0 Then
            
            nroUsuario = .VAsumAlumInte.FieldValue("CodAlumbrado")
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
            anchoDeBanda = IIf(.VTarifas.status = 0, .VTarifas.FieldValue("descrip"), vbNullString)
            email = .VAsumAlum.FieldValue("direelec")
            iva = CIVADescrip(.VOrdenes.FieldValue("civa"))
            
            cuadrilla = .VCuadrillas.FieldValue("miembros")
        End If
    End With
        
    ' ===================
    ' Comienza a imprimir
    
    Dim pdfOrden As New LinePrinter
    With pdfOrden
    
        .Thin = 10
        .Fat = 16
        .Init 2, 0
        
        .LineH 25, 20, 165
        .Text 25, 25, "Cooperativa Eléctrica Integral de Provisión de Servicios Públicos", 12, True, "Arial"
        .Text 25, 30, "y Sociales de Todd Ltda.", 12, True, "Arial"
        .LineH 25, 40, 165
        
        .Text 25, 45, "Todd Net - Nueva conexión", 10, True, "Arial"
        .Text 25, 50, "Observaciones: " & obs, 9, True, "Arial"
        
        .Text 25, 55, "N.º de usuario: " & nroUsuario, 9, False, "Arial"
        .Text 70, 55, "Fecha de inst. programada: " & Format$(fechaInstalacion, "dd/MM/yyyy"), 9, False, "Arial"
        .Text 135, 55, "Hora de inst. programada: " & Format$(horaInstalacion, "hh:mm AMPM"), 9, False, "Arial"
        .LineH 25, 65, 165
        
        .Text 25, 70, "Apellido y nombre: " & Nombre, 9, False, "Arial"
        .Text 25, 75, "DNI/CUIT: " & dni, 9, False, "Arial"
        .Text 25, 80, "Condición IVA: " & iva, 9, False, "Arial"
        .Text 25, 85, "Domicilio de facturación: " & domicilioFacturacion, 9, False, "Arial"
        .Text 25, 90, "Domicilio de conexión: " & domicilioConexion, 9, False, "Arial"
        .Text 25, 95, "Teléfono: " & telefono, 9, False, "Arial"
        .Text 25, 100, "Correo eléctronico: " & email, 9, False, "Arial"
        .Text 25, 105, "Nombre de usuario de Internet: " & nombreUsuario, 9, False, "Arial"
        .Text 25, 110, "Ancho de Banda: " & anchoDeBanda, 9, False, "Arial"
        .LineH 25, 120, 165
        
        .Text 25, 125, "Datos de la instalación", 10, True, "Arial"
        
        .Text 25, 130, "Cuadrilla: " & cuadrilla, 9, True, "Arial"
        .Text 100, 130, "Tipo de trabajo: " & tipoConexion, 9, True, "Arial"
        .Text 25, 135, "Fecha de inst.: ______________", 9, True, "Arial"
        .Text 100, 135, "Hora de inst.: __________", 9, True, "Arial"
        
        .Text 25, 140, "N.º de fibra: ____________", 9, False, "Arial"
        .Text 25, 145, "Nombre y ubicación de la caja: ____________", 9, False, "Arial"
        .Text 25, 150, "Cant. de cables (mts.): ____________", 9, False, "Arial"
        .Text 25, 155, "Cant. de conectores: ____________", 9, False, "Arial"
        .Text 25, 160, "Cant. de anilla de distr.: ____________", 9, False, "Arial"
        .Text 25, 165, "Cant. de anilla de paso: ____________", 9, False, "Arial"
        .Text 25, 170, "Cant. de mordazas: ____________", 9, False, "Arial"
        .Text 25, 175, "Cant. de cadenas: ____________", 9, False, "Arial"
        .Text 25, 180, "Patchcord: ____________", 9, False, "Arial"
        .Text 25, 185, "Potencia en abonado: ____________", 9, False, "Arial"
        .Text 25, 190, "IP: ____________", 9, False, "Arial"
        .Text 25, 195, "Cant. de cable UTP: ____________", 9, False, "Arial"
        .Text 25, 200, "Cant. de fichas RJ45: ____________", 9, False, "Arial"
        
        .LineH 25, 210, 165
        
        .Text 25, 220, "Conformidad del usuario", 9, False, "Arial"
        .Text 150, 220, "Nombre del instalador", 9, False, "Arial"
        .Text 25, 230, "____________________", 9, False, "Arial"
        .Text 150, 230, "____________________", 9, False, "Arial"
        
        .Text 25, 245, "Acepto bases y condiciones del servicio de Internet indicadas en la página web www.todd.com.ar", 7, False, "Arial"
        
        .SendToPrinter
    End With
    
End Sub

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
