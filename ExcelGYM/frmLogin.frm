VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin 
   Caption         =   "INICIAR SESIÓN"
   ClientHeight    =   2850
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4740
   OleObjectBlob   =   "frmLogin.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Pega este código en el módulo del UserForm 'frmLogin'
Option Explicit
Public LoginExitoso As Boolean

Private Sub cmdIngresar_Click()
    
    ' CAMBIA "tu_contrasena" por la que quieras usar
    Const CONTRASENA_MAESTRA As String = "3333"
    
    ' Verifica si el campo de contraseña está vacío
    If txtContrasena.Text = "" Then
        MsgBox "Por favor, ingrese la contraseña.", vbExclamation
        txtContrasena.SetFocus
        Exit Sub
    End If
    
    If txtContrasena.Text = CONTRASENA_MAESTRA Then
        LoginExitoso = True ' Le dice al código principal que el login fue exitoso
        Me.Hide
    Else
        ' Si la contraseña es incorrecta, avisa y prepara el cierre
        MsgBox "Contraseña incorrecta. El archivo se cerrará.", vbCritical
        LoginExitoso = False ' Le dice al código principal que el login no fue exitoso
        Me.Hide
    End If
    
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Si el usuario intenta cerrar con la "X", le pedimos la contraseña
    If CloseMode = vbFormControlMenu Then
        MsgBox "Debe ingresar la contraseña para continuar.", vbCritical
        Cancel = 1
    End If
    
    ' Cuando el formulario se oculta (Me.Hide) para cerrar,
    ' nos aseguramos de que se cierre completamente
    If CloseMode = vbFormControlMenu Then
        ' No hace nada, el código de arriba ya lo maneja
    ElseIf CloseMode = vbFormCode Then
        ' Establece LoginExitoso en False si el formulario se cierra por código
        LoginExitoso = False
    End If
End Sub

Private Sub txtContrasena_Change()
End Sub
