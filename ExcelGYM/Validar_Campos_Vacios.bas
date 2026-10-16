Attribute VB_Name = "Validar_Campos_Vacios"
Public contador

Sub Validar_Vacios()

contador = 0

'Validar Nombre
If Registrar.Box_Name.Value = "" Then
Registrar.Box_Name.BackColor = VBA.RGB(250, 211, 151)
contador = contador + 1
Else
Registrar.Box_Name.BackColor = VBA.vbWhite
End If

'Validar Nombre del plan
If Registrar.Box_Plan.Value = "" Then
Registrar.Box_Plan.BackColor = VBA.RGB(250, 211, 151)
contador = contador + 1
Else
Registrar.Box_Plan.BackColor = VBA.vbWhite
End If

'Validar Condicion
If Registrar.ComboBox1.Value = "" Then
Registrar.ComboBox1.BackColor = VBA.RGB(250, 211, 151)
contador = contador + 1
Else
Registrar.ComboBox1.BackColor = VBA.vbWhite
End If

'Validar Fecha
If Registrar.Box_Fecha.Value = "" Then
Registrar.Box_Fecha.BackColor = VBA.RGB(250, 211, 151)
contador = contador + 1
Else
Registrar.Box_Fecha.BackColor = VBA.vbWhite
End If

'Validar Total
If Registrar.Box_Total.Value = "" Then
Registrar.Box_Total.BackColor = VBA.RGB(250, 211, 151)
contador = contador + 1
Else
Registrar.Box_Total.BackColor = VBA.vbWhite
End If

'Validar Metodo
If Registrar.ComboBox2.Value = "" Then
Registrar.ComboBox2.BackColor = VBA.RGB(250, 211, 151)
contador = contador + 1
Else
Registrar.ComboBox2.BackColor = VBA.vbWhite
End If

End Sub
