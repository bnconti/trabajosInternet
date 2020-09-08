VERSION 5.00
Begin VB.Form frmCuadrillaNueva 
   Caption         =   "Nueva cuadrilla"
   ClientHeight    =   1950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1950
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnGuardarCuadrilla 
      Caption         =   "&Guardar"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton btnVolver 
      Cancel          =   -1  'True
      Caption         =   "&Volver"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1440
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
   Begin VB.Label Label2 
      Caption         =   "Correo electrónico"
      Height          =   375
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
    If txtMiembros <> vbNullString And txtCorreoCuadrilla <> vbNullString Then
        main.VCuadrillas.FieldValue("miembros") = txtMiembros
        main.VCuadrillas.FieldValue("email") = txtCorreoCuadrilla
        main.VCuadrillas.FieldValue("habilitado") = True
        main.VCuadrillas.Insert
    End If
End Sub
