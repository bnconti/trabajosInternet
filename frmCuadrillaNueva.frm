VERSION 5.00
Begin VB.Form frmCuadrillaNueva 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nueva cuadrilla"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnGuardarCuadrilla 
      Caption         =   "&Guardar"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton btnVolver 
      Cancel          =   -1  'True
      Caption         =   "&Volver"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txtCorreoCuadrilla 
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txtMiembros 
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "(separar direcciones por punto y coma)"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblCorreo 
      Caption         =   "Correo electrónico"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblMiembros 
      Caption         =   "Miembros"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1695
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
    
    If txtMiembros <> vbNullString And txtCorreoCuadrilla <> vbNullString Then
        main.VCuadrillas.FieldValue("miembros") = txtMiembros
        main.VCuadrillas.FieldValue("email") = txtCorreoCuadrilla
        main.VCuadrillas.FieldValue("habilitado") = True
        main.VCuadrillas.Insert
    End If
    
    Unload Me
End Sub

Private Sub btnVolver_Click()
    Unload Me
End Sub

