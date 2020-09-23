Attribute VB_Name = "moduloCorreo"
Option Explicit

Private pdf As PDFCreator.clsPDFCreator
Private opt As clsPDFCreatorOptions

Public Type tCorreo
    direccion As String
    contrasenia As String
    servidor As String
    puerto As String
    seguridad As Boolean
    autenticacion As Boolean
    adjunto As String
End Type

Dim correo As tCorreo

Public Sub prepararCorreo(idTrabajo As Long)

    Dim correosDestino As String

    Call prepararPDF
    Call cargarDatosCorreo
    
    ' hacer for por si es más de un correo
    correosDestino = "bruno.soportecoop@gmail.com"
    
    ' Genera la orden temporalmente - desp borrar
    Call imprimirOrden
    Call enviarCorreo
End Sub

Private Function traerCorreoCuadrilla(idTrabajo As Long) As String
    With main
        .vTrabInternet.IndexNumber = 0
        .vTrabInternet.FieldValue ("id_trabajo")
        .vTrabInternet.GetEqual
        
        If .vTrabInternet.status = 0 Then
            .VCuadrillas.IndexNumber = 0
            .VCuadrillas.FieldValue("idcuadrilla") = .vTrabInternet.FieldValue("idcuadrilla")
            .VCuadrillas.GetEqual
            
            If .VCuadrillas.status = 0 Then
                traerCorreoCuadrilla = .VCuadrillas.FieldValue("email") & vbNullString
            End If
        End If
    End With
End Function


Private Sub prepararPDF()
    Set pdf = New clsPDFCreator
    Set opt = New clsPDFCreatorOptions
    
    With pdf
        .cVisible = True
        If .cStart("/NoProcessingAtStartup") = False Then
            If .cStart("/NoProcessingAtStartup", True) = False Then
                Exit Sub
            End If
            .cVisible = True
        End If
        
        Set opt = .cOptions
        .cClearCache
    End With
    
    With opt
        ' Desp cambiar la ubicacion
        .AutosaveDirectory = "C:\"
        .AutosaveFilename = "ORDENCORREO.pdf"
        .DisableUpdateCheck = True
        .StandardTitle = "Orden de trabajo"
        .UseAutosave = 1
        .UseAutosaveDirectory = 1
        .AutosaveFormat = 0 ' PDF
    End With
    
    Set pdf.cOptions = opt
    Set Printer = Printers(PrinterIndex("PDFCreator"))
    
    
End Sub


Private Sub cargarDatosCorreo()
    correo.direccion = empresa.GetVar("empresa", "emailrte")
    correo.contrasenia = empresa.GetVar("empresa", "contraseniaEmail")
    correo.puerto = empresa.GetVar("empresa", "puertoSmtp")
    correo.servidor = empresa.GetVar("empresa", "servidorsmtp")
    correo.adjunto = "C:\ORDENCORREO.pdf"
    correo.seguridad = IIf(empresa.GetVar("empresa", "seguridadEmail") = "true", True, False)
    correo.autenticacion = IIf(empresa.GetVar("empresa", "autenticacionEmail") = "true", True, False)
End Sub

Private Sub enviarCorreo(destino As String, rutaAdjunto As String)
    Set cdoCorreo = CreateObject("CDO.Message")
    
    With cdoCorreo.Configuration.Fields
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2  'Send the message using the network (SMTP over the network).
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = correo.servidor
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = correo.puerto ' 25 si es sin contraseña, 465 si es gmail
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = correo.seguridad 'True si es con seguridad, sino False
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 15
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = correo.autenticacion 'basic (clear-text) authentication
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "bruno.soportecoop@gmail.com" ' correo.direccion
        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "elchiqui20" ' correo.contrasenia
        .Update
    End With
  
    cdoCorreo.To = Trim(destino)
    cdoCorreo.From = correo.direccion
    cdoCorreo.Subject = "Orden de trabajo" & "" ' agregar algún dato para que quede mejor
    cdoCorreo.Sender = "Todd Net"
    cdoCorreo.AddAttachment rutaAdjunto

End Sub
