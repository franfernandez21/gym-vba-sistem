VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Eliminar 
   Caption         =   "¿Está seguro que quiere elminar el registro: ?"
   ClientHeight    =   2460
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5085
   OleObjectBlob   =   "Eliminar.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Eliminar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Unload Me
End Sub

Private Sub CommandButton2_Click()

'Capturar codigo seleccionado del lisbox1
Codigo_Registro = Buscar.ListBox1.List(Buscar.ListBox1.ListIndex, 0)
Nombre = Buscar.ListBox1.List(Buscar.ListBox1.ListIndex, 1)

'MsgBox Codigo_Registro
'Identificar en que fila se encuentra el codigo seleccionado

Set fila = Sheets("Base de datos").Range("B:B").Find(Codigo_Registro, Lookat:=xlWhole)
Linea = fila.Row

'MsgBox Linea

If TextBox1.Value = "12345" Then

Range("B" & Linea).EntireRow.Delete
MsgBox "Se elimino el Cliente: " & Nombre
Unload Me

'se presiona el boton buscar
Buscar.CommandButton1_Click

Else

MsgBox "Contraseña Incorrecta"
Unload Me

End If



End Sub

Private Sub UserForm_Activate()

Nombre = Buscar.ListBox1.List(Buscar.ListBox1.ListIndex, 1)

Label1.Caption = Nombre



End Sub

Private Sub UserForm_Click()

End Sub
