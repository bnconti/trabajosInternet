VERSION 5.00
Begin VB.Form frmCuadrillaNueva 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nueva cuadrilla"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   Icon            =   "frmCuadrillaNueva.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   6435
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnGuardarCuadrilla 
      Caption         =   "&Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1222
      TabIndex        =   4
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton btnVolver 
      Cancel          =   -1  'True
      Caption         =   "&Volver"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox txtCorreoCuadrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   330
      TabIndex        =   3
      Top             =   1440
      Width           =   5775
   End
   Begin VB.TextBox txtMiembros 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   330
      TabIndex        =   1
      Top             =   480
      Width           =   5775
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   6120
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label lblCorreo 
      Caption         =   "Correo electrónico (separar direcciones por punto y coma)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   5535
   End
   Begin VB.Label lblMiembros 
      Caption         =   "Miembros o nombre de la cuadrilla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   120
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
  main.VCuadrillas.Clear  ' Para borrar el ID

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
  Cancel = False

  ' sacar espacios del textbox
  txtCorreoCuadrilla.Text = Replace(txtCorreoCuadrilla.Text, " ", vbNullString)
  If Not ValidarCorreos(txtCorreoCuadrilla.Text) Then
    Call MsgBox("Ingrese una dirección válida", vbOKOnly + vbInformation, Me.Caption)
    Cancel = True
  End If
End Sub
