VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Pago 
   Caption         =   "Agregar Pago"
   ClientHeight    =   2655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5535
   OleObjectBlob   =   "Pago.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Pago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()
    If Me.ComboBox1.ListIndex <> -1 Then
        Me.ComboBox1.BackColor = vbWhite
    End If
End Sub

Private Sub CommandButton1_Click()
    
    Dim Monto As Double
    Dim Medio As String
    Dim fila As Variant
    Dim nombreCliente As String
    Dim wsBase As Worksheet
    Dim wsHistorial As Worksheet
    Dim UltFila As Long
    Dim hayError As Boolean
    
    Set wsBase = ThisWorkbook.Sheets("Base de datos")
    Set wsHistorial = ThisWorkbook.Sheets("Historial de Pagos") ' Asegurate que exista
    
    hayError = False
    
    ' Validar monto en TextBox1
    If Trim(Me.TextBox1.Value) = "" Then
        Me.TextBox1.BackColor = VBA.RGB(250, 211, 151)
        hayError = True
    Else
        Me.TextBox1.BackColor = vbWhite
    End If
    
    ' Validar selección en ComboBox1
    If Me.ComboBox1.ListIndex = -1 Then
        Me.ComboBox1.BackColor = VBA.RGB(250, 211, 151)
        hayError = True
    Else
        Me.ComboBox1.BackColor = vbWhite
    End If
    
    If hayError Then
        MsgBox "Por favor, complete todos los campos resaltados.", vbExclamation
        Exit Sub
    End If
    
    ' Validar que el monto sea numérico y positivo
    If Not IsNumeric(Me.TextBox1.Value) Then
        MsgBox "Ingrese un monto válido.", vbExclamation
        Exit Sub
    End If
    
    Monto = CDbl(Me.TextBox1.Value)
    
    If Monto <= 0 Then
        MsgBox "El monto debe ser mayor a cero.", vbExclamation
        Exit Sub
    End If
    
    Medio = Me.ComboBox1.Value
    
     ' Buscar fila cliente en Base de datos usando variable global
    On Error Resume Next
    fila = Application.WorksheetFunction.Match(gCodigoCliente, wsBase.Range("B:B"), 0)
    If Err.Number <> 0 Then
        MsgBox "No se encontró el cliente en la Base de datos.", vbCritical
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Obtener nombre cliente y sumar monto a columna "G" (Pago)
    nombreCliente = wsBase.Cells(fila, "C").Value
    wsBase.Cells(fila, "G").Value = Val(wsBase.Cells(fila, "G").Value) + Monto  ' En esta linea
    
    ' Obtenemos el plan del cliente de la columna "O" (PLAN)
    PlanCliente = wsBase.Cells(fila, "O").Value
    
    ' Insertar datos una fila arriba en Historial de Pagos
    UltFila = wsHistorial.Cells(wsHistorial.Rows.Count, "A").End(xlUp).Row + 1
    If UltFila < 2 Then UltFila = 2 ' Para no pisar encabezados
    
    wsHistorial.Rows(UltFila).Insert Shift:=xlDown
    
    wsHistorial.Cells(UltFila, 1).Value = gCodigoCliente
    wsHistorial.Cells(UltFila, 2).Value = nombreCliente
    wsHistorial.Cells(UltFila, 3).Value = Date
    wsHistorial.Cells(UltFila, 4).Value = PlanCliente
    wsHistorial.Cells(UltFila, 5).Value = Medio
    wsHistorial.Cells(UltFila, 6).Value = Monto
    
    MsgBox "Pago registrado correctamente.", vbInformation
    
    Dim uf As Object
    For Each uf In VBA.UserForms
        If TypeName(uf) = "Buscar" Then
            uf.CargarTodosLosClientes
            Exit For
        End If
    Next uf
    Unload Me
End Sub

Private Sub TextBox1_Change()
    If Trim(Me.TextBox1.Value) <> "" Then
        Me.TextBox1.BackColor = vbWhite ' o VBA.RGB(255, 255, 255)
    End If
End Sub

Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii >= 46 And KeyAscii <= 57) Then
    KeyAscii = KeyAscii
  Else
    KeyAscii = 0
  End If
End Sub
Private Sub UserForm_Initialize()
    Me.ComboBox1.Clear
    Me.ComboBox1.AddItem "EFECTIVO"
    Me.ComboBox1.AddItem "TRANSFERENCIA"
End Sub
