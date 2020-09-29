Attribute VB_Name = "moduloCorreo"
Option Explicit

Private pdf As PDFCreator.clsPDFCreator
Private opt As clsPDFCreatorOptions

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Type tCorreo
    Direccion As String
    Contrasenia As String
    Servidor As String
    Puerto As String
    Seguridad As Boolean
    Autenticacion As Boolean
    Adjunto As String
End Type

Dim correo As tCorreo

' Definen dónde se guarda el .pdf
Private Const DIRECTORIO As String = "C:\PDFTEMP\"
Private Nombre As String

Public Sub prepararCorreo(idTrabajo As Long)
    Dim correosDestino As String

    Nombre = "ORDEN " & Format(DateTime.Now, "dd-MM hh-mm-ss") & ".pdf"
    
    Screen.MousePointer = vbHourglass
    
    Call prepararPDF(idTrabajo)
    Call cargarDatosCorreo
    
    ' el correo destino se saca de la tabla CUADRILLAS
    ' hacer for por si es más de un correo
    ' traerCorreoCuadrilla(idTrabajo)
    correosDestino = "bruno.soportecoop@gmail.com"
    
    Call enviarCorreo(correosDestino)
    
    On Error GoTo NoFunco
    Kill correo.Adjunto
    
NoFunco:
    On Error GoTo 0
End Sub

Private Function traerCorreoCuadrilla(idTrabajo As Long) As String
    With main
        .vTrabInternet.IndexNumber = 0
        .vTrabInternet.FieldValue("id_trabajo") = idTrabajo
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


Private Sub prepararPDF(idTrabajo As Long)
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
        .AutosaveDirectory = DIRECTORIO
        .AutosaveFilename = Nombre
        .DisableUpdateCheck = True
        .StandardTitle = "Orden de trabajo"
        .UseAutosave = 1
        .UseAutosaveDirectory = 1
        .AutosaveFormat = 0 ' PDF
    End With
    
    Set pdf.cOptions = opt
    Set Printer = Printers(PrinterIndex("PDFCreator"))
    
    Call imprimirOrden(idTrabajo)
    
    pdf.cPrinterStop = False
    
    Sleep 1000
End Sub


Private Sub cargarDatosCorreo()
    With main.ini
        correo.Direccion = .GetVar("empresa", "emailrte")
        correo.Contrasenia = .GetVar("empresa", "contraseniaEmail")
        correo.Puerto = .GetVar("empresa", "puertoSmtp")
        correo.Servidor = .GetVar("empresa", "servidorsmtp")
        correo.Adjunto = DIRECTORIO & Nombre
        correo.Seguridad = IIf(.GetVar("empresa", "seguridadEmail") = "true", True, False)
        correo.Autenticacion = IIf(.GetVar("empresa", "autenticacionEmail") = "true", True, False)
    End With
End Sub

Private Sub enviarCorreo(destino As String)
    Dim cdoCorreo As New CDO.Message
    
    With cdoCorreo.Configuration.Fields
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2  'Send the message using the network (SMTP over the network).
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = correo.Servidor
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = correo.Puerto ' 25 si es sin contraseña, 465 si es gmail
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = correo.Seguridad 'True si es con seguridad, sino False
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 15
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = correo.Autenticacion 'basic (clear-text) authentication
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "bruno.soportecoop@gmail.com" ' correo.direccion
        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "elchiqui20" ' correo.contrasenia
        .Update
    End With
  
    cdoCorreo.To = Trim(destino)
    cdoCorreo.From = correo.Direccion
    cdoCorreo.Subject = "Orden de trabajo" & "" ' agregar algún dato para que quede mejor
    cdoCorreo.Sender = "Todd Net"
    cdoCorreo.AddAttachment correo.Adjunto
    
    cdoCorreo.HTMLBody = "<div>" & "Nueva orden de trabajo" & "</div>"
    cdoCorreo.TextBodyPart.Charset = "utf-8"
        
    On Error GoTo ErrorAlEnviar
    cdoCorreo.Send
    
    Screen.MousePointer = vbDefault
    
    MsgBox "Correo enviado exitosamente", vbInformation + vbOKOnly, "Éxito"
    Exit Sub
    
ErrorAlEnviar:
    MsgBox "Hubo un problema al enviar el correo", vbCritical + vbOKOnly, "Error"
    On Error GoTo 0

End Sub
