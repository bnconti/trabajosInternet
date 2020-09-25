VERSION 5.00
Begin VB.Form frmCuadrillaNueva 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nueva cuadrilla"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   Icon            =   "frmCuadrillaNueva.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   4905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnGuardarCuadrilla 
      Caption         =   "&Guardar"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton btnVolver 
      Cancel          =   -1  'True
      Caption         =   "&Volver"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox txtCorreoCuadrilla 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   4095
   End
   Begin VB.TextBox txtMiembros 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   4095
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   4560
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label lblCorreo 
      Caption         =   "Correo electrónico (separar direcciones por punto y coma)"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Label lblMiembros 
      Caption         =   "Miembros o nombre de la cuadrilla"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "frmCuadrillaNueva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnGuardarCuadrilla_Click()
    main.VCuadrillas.Clear ' Para borrar el ID
    
    If txtMiembros.Text <> vbNullString Then
        main.VCuadrillas.FieldValue("miembros") = txtMiembros
        main.VCuadrillas.FieldValue("email") = txtCorreoCuadrilla
        main.VCuadrillas.FieldValue("habilitado") = True
        main.VCuadrillas.Insert
        
        Unload Me
    Else
        ' Mostrar mensaje si no se cargo porque faltan datos
        Call MsgBox("Debe completar todos los datos", vbOKOnly + vbInformation, Me.Caption)
        txtMiembros.SetFocus
    End If
End Sub

Private Sub btnVolver_Click()
    Unload Me
End Sub

Private Sub txtCorreoCuadrilla_Validate(Cancel As Boolean)
    Dim correos() As String
    Dim nCorreo As Byte
    Dim pos As Integer
    
    Cancel = False
    
    ' sacar espacios iniciales y finales del textbox
    txtCorreoCuadrilla.Text = Trim$(txtCorreoCuadrilla.Text)
    If txtCorreoCuadrilla.Text = vbNullString Then
        ' no puso nada, salir nomasss
        Exit Sub
    End If
    
    'txtCorreoCuadrilla.Text = Trim$(txtCorreoCuadrilla.Text)
    correos = Split(txtCorreoCuadrilla.Text, ";")
    For nCorreo = LBound(correos) To UBound(correos)
        ' posicion del arroba
        pos = InStr(2, correos(nCorreo), "@")
        If (pos < 1) Or (pos > (Len(correos(nCorreo)) - 4)) Then
            ' no tiene arroba a partir del 2do caracter, o tiene un sufijo menor a cuatro caracteres
            Call MsgBox("Ingrese una dirección válida")
            Cancel = True
            Exit Sub
        End If
    Next
End Sub
