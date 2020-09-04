Attribute VB_Name = "moduloFunciones"
Option Explicit

Private Sub cargarCuadrillas(control)
    Dim status As Integer
    With control
        status = main.VCuadrillas.GetFirst
        While status = 0
            .AddItem (VCuadrillas.FieldValue("MIEMBROS"))
            status = VCuadrillas.GetNext
        Wend
    End With
End Sub
