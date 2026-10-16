Attribute VB_Name = "ModificarOeliminar"
Public Sub AbrirFormularioParaModificar()
    ' Pasa el argumento "Modificar" al formulario
    Buscar.Tag = "Modificar"
    Buscar.Show
End Sub

Public Sub AbrirFormularioParaEliminar()
    ' Pasa el argumento "Eliminar" al formulario
    Buscar.Tag = "Eliminar"
    Buscar.Show
End Sub
