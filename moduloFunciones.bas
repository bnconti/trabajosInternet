Attribute VB_Name = "moduloFunciones"
Option Explicit

Public Declare Function SetDefaultPrinter Lib "Winspool.drv" Alias "SetDefaultPrinterA" (ByVal pszPrinter As String) As Boolean

Public Function Izq(Texto As String, largo As Integer) As String
  Izq = left$(LTrim$(Texto), largo)
  Izq = Izq & String$(largo - Len(Izq), " ")
End Function

Public Function Der(Texto As String, largo As Integer) As String
  Der = Right$(RTrim$(Texto), largo)
  Der = String$(largo - Len(Der), " ") & Der
End Function

Public Sub ponerEncabezadoEnNegrita(tabla As VSFlexGrid)
    tabla.Cell(flexcpFontBold, 0, 0, 0, tabla.Cols - 1) = True
End Sub

Public Sub imprimirOrden(idTrabajo As Long)
    Dim nroUsuario As String
    Dim tipoConexion As String
    Dim fechaInstalacion As String
    Dim horaInstalacion As String
    
    Dim Obs As String
    Dim nombre As String
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
    
    ' Traer datos
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
            .VCuadrillas.status = 0 And _
            .VTarifas.status = 0 Then
            
            nroUsuario = .VAsumAlumInte.FieldValue("CodAlumbrado")
            tipoConexion = frmTrabajo.cmbTipoConexion.Text
            fechaInstalacion = .vTrabInternet.FieldValue("fecha_inst")
            horaInstalacion = .vTrabInternet.FieldValue("hora_inst")
            Obs = .vTrabInternet.FieldValue("obs")
            
            nombre = .VAClientes.FieldValue("nombre") & " " & .VAClientes.FieldValue("apellido")
            dni = .VAClientes.FieldValue("NroDocIde") & vbNullString
            domicilioFacturacion = .VOrdenes.FieldValue("domicilio") & vbNullString
            domicilioConexion = .VAsumAlum.FieldValue("cuenta") & vbNullString
            
            telefono = .VAClientes.FieldValue("reserva") & vbNullString
            nombreUsuario = .VAsumAlumInte.FieldValue("UsInt") & vbNullString
            anchoDeBanda = .VTarifas.FieldValue("descrip")
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
        .Text 25, 25, "Cooperativa El�ctrica Integral de Provisi�n de Servicios P�blicos", 12, True, "Arial"
        .Text 25, 30, "y Sociales de Todd Ltda.", 12, True, "Arial"
        .LineH 25, 40, 165
        
        .Text 25, 45, "Todd Net - Nueva conexi�n", 10, True, "Arial"
        .Text 25, 50, "Observaciones: " & Obs, 9, True, "Arial"
        
        .Text 25, 55, "N.� de usuario: " & nroUsuario, 9, False, "Arial"
        .Text 70, 55, "Fecha de inst. programada: " & fechaInstalacion, 9, False, "Arial"
        .Text 135, 55, "Hora de inst. programada: " & horaInstalacion, 9, False, "Arial"
        .LineH 25, 65, 165
        
        .Text 25, 70, "Apellido y nombre: " & nombre, 9, False, "Arial"
        .Text 25, 75, "DNI/CUIT: " & dni, 9, False, "Arial"
        .Text 25, 80, "Condici�n IVA: " & iva, 9, False, "Arial"
        .Text 25, 85, "Domicilio de facturaci�n: " & domicilioFacturacion, 9, False, "Arial"
        .Text 25, 90, "Domicilio de conexi�n: " & domicilioConexion, 9, False, "Arial"
        .Text 25, 95, "Tel�fono: " & telefono, 9, False, "Arial"
        .Text 25, 100, "Correo el�ctronico: " & email, 9, False, "Arial"
        .Text 25, 105, "Nombre de usuario de Internet: " & nombreUsuario, 9, False, "Arial"
        .Text 25, 110, "Ancho de Banda: " & anchoDeBanda, 9, False, "Arial"
        .LineH 25, 120, 165
        
        .Text 25, 125, "Datos de la instalaci�n", 10, True, "Arial"
        
        .Text 25, 130, "Cuadrilla: " & cuadrilla, 9, True, "Arial"
        .Text 100, 130, "Tipo de trabajo: " & tipoConexion, 9, True, "Arial"
        .Text 25, 135, "Fecha de inst.: ______________", 9, True, "Arial"
        .Text 100, 135, "Hora de inst.: __________", 9, True, "Arial"
        
        .Text 25, 140, "N.� de fibra: ____________", 9, False, "Arial"
        .Text 25, 145, "Nombre y ubicaci�n de la caja: ____________", 9, False, "Arial"
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
        
        .Text 25, 245, "Acepto bases y condiciones del servicio de Internet indicadas en la p�gina web www.todd.com.ar", 7, False, "Arial"
        
        .SendToPrinter
    End With
    
End Sub

