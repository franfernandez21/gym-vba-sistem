VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Buscar 
   Caption         =   "Buscador"
   ClientHeight    =   6690
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9960.001
   OleObjectBlob   =   "Buscar.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Buscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Se ejecuta al iniciar el formulario
Private Sub UserForm_Initialize()
    ' Ocultamos todos los botones para que el evento Activate los muestre
    Me.CommandButton4.Visible = False
    Me.CommandButton5.Visible = False
    Me.CommandButton3.Visible = False ' Nuevo boton de Eliminar
    
    Call CargarTodosLosClientes
End Sub

' Se ejecuta al activar el formulario (cuando lo abrimos)
Private Sub UserForm_Activate()
    ' Verifica el valor de la propiedad Tag para mostrar los botones
    If Me.Tag = "Eliminar" Then
        Me.CommandButton3.Visible = True
        Me.CommandButton5.Visible = False
        Me.CommandButton4.Visible = False
        Me.Label1.Caption = "Eliminar Cliente" ' Opcional: Cambia el titulo del formulario
    Else
        Me.CommandButton5.Visible = True
        Me.CommandButton4.Visible = True
        Me.CommandButton3.Visible = False
        Me.Label1.Caption = "Buscar Cliente"
    End If
End Sub

' Se ejecuta cuando cambiás el texto del TextBox (búsqueda en vivo)
Private Sub TextBox1_Change()
    Call FiltrarClientes(Me.TextBox1.Value)
End Sub

' Botón para agregar pago a un cliente (ahora CommandButton4)
Private Sub CommandButton4_Click()
    Dim Codigo_Cliente As String
    Dim Estado As String
    Dim fila As Variant
    
    If Me.ListBox1.ListIndex = -1 Then
        MsgBox "Seleccione un cliente", vbExclamation
        Exit Sub
    End If
    
    Codigo_Cliente = Me.ListBox1.List(Me.ListBox1.ListIndex, 0)
    
    With Hoja1
        On Error Resume Next
        fila = Application.WorksheetFunction.Match(Codigo_Cliente, .Range("B:B"), 0)
        If Err.Number <> 0 Then
            MsgBox "No se encontró el cliente en la Base de datos", vbCritical
            Err.Clear
            Exit Sub
        End If
        On Error GoTo 0
        
        Estado = .Cells(fila, "K").Value
    End With
    
    If LCase(Trim(Estado)) = "cancelado" Then
        MsgBox "No se puede agregar pago porque el cliente está CANCELADO.", vbCritical
        Exit Sub
    End If
    
    gCodigoCliente = Codigo_Cliente
    
    Pago.Show
End Sub

' Botón para modificar el cliente (ahora CommandButton5)
Private Sub CommandButton5_Click()
    Dim Codigo_Cliente As String
    
    If Me.ListBox1.ListIndex = -1 Then
        MsgBox "Seleccione un cliente para modificar.", vbExclamation
        Exit Sub
    End If
    
    Codigo_Cliente = Me.ListBox1.List(Me.ListBox1.ListIndex, 0)
    
    gCodigoCliente = Codigo_Cliente
    
    Registrar.Show
    
    Unload Me
End Sub

'BOTON DE ELIMINAR
Private Sub CommandButton3_Click()
    Dim wsBD As Worksheet
    Dim IDaEliminar As String
    Dim FilaaEliminar As Variant
    Dim Msg As VbMsgBoxResult
    Dim Contrasena As String
    
    ' Obtiene la hoja de la Base de Datos
    Set wsBD = ThisWorkbook.Sheets("Base de datos")
    
    ' Verifica si se seleccionó un cliente en la lista
    If Me.ListBox1.ListIndex = -1 Then
        MsgBox "Por favor, seleccione un cliente para eliminar.", vbExclamation
        Exit Sub
    End If
    
    ' Obtiene el ID del cliente seleccionado
    IDaEliminar = Me.ListBox1.List(Me.ListBox1.ListIndex, 0)
    
    ' Pide confirmacion y contrasena
    Msg = MsgBox("¿Estás seguro de que deseas eliminar este cliente? Esta acción no se puede deshacer.", vbYesNo + vbExclamation, "Confirmar eliminación")
    
    If Msg = vbYes Then
        Contrasena = InputBox("Ingrese la contraseña de seguridad para confirmar la eliminación:", "Confirmar")
        
        If Contrasena = "3333" Then ' <--- REEMPLAZA "PUEBLOFITNESS" CON TU CONTRASEÑA REAL
            On Error Resume Next
            ' Busca la fila del cliente por su ID
            FilaaEliminar = Application.Match(IDaEliminar, wsBD.Range("B:B"), 0)
            On Error GoTo 0
            
            If Not IsError(FilaaEliminar) Then
                ' Elimina la fila completa
                wsBD.Rows(FilaaEliminar).Delete
                MsgBox "Cliente eliminado correctamente.", vbInformation
                Me.Hide ' Oculta el formulario después de la eliminación
            Else
                MsgBox "No se encontró el cliente.", vbExclamation
            End If
        Else
            MsgBox "Contraseña incorrecta. La eliminación ha sido cancelada.", vbCritical
        End If
    End If
End Sub

' Se ejecuta al hacer doble clic en un cliente
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim Codigo_Registro As String
    Codigo_Registro = Me.ListBox1.List(Me.ListBox1.ListIndex, 0)
    MsgBox Codigo_Registro
End Sub

' Carga todos los clientes (sin filtros)
Public Sub CargarTodosLosClientes()
    Dim wsBD As Worksheet
    Dim ultimaFila As Long
    Dim i As Long, j As Long
    Dim DataArray() As Variant
    Dim arrColumnas As Variant
    Dim k As Long
    
    Set wsBD = ThisWorkbook.Sheets("Base de datos")
    ultimaFila = wsBD.Range("B" & wsBD.Rows.Count).End(xlUp).Row
    
    ' Definimos las columnas que queremos mostrar: B, C, D, K, O
    arrColumnas = Array("B", "C", "D", "K", "O")
    
    Me.ListBox1.Clear
    Me.ListBox1.ColumnCount = UBound(arrColumnas) + 1
    Me.ListBox1.ColumnWidths = "50;120;70;70;70" ' Ancho de las 5 columnas
    
    If ultimaFila < 7 Then Exit Sub ' No hay datos
    
    ReDim DataArray(1 To ultimaFila - 6, 1 To UBound(arrColumnas) + 1)
    
    For i = 7 To ultimaFila
        k = 0
        For j = LBound(arrColumnas) To UBound(arrColumnas)
            DataArray(i - 6, k + 1) = wsBD.Cells(i, arrColumnas(j)).Value
            k = k + 1
        Next j
    Next i
    
    Me.ListBox1.List = DataArray
End Sub

' Filtra clientes según lo que escribís en el TextBox
Private Sub FiltrarClientes(Filtro As String)
    Dim wsBD As Worksheet
    Dim ultimaFila As Long
    Dim fila As Long
    Dim Y As Long
    Dim Encontrado As Boolean
    
    Set wsBD = ThisWorkbook.Sheets("Base de datos")
    ultimaFila = wsBD.Range("B" & wsBD.Rows.Count).End(xlUp).Row
    
    Me.ListBox1.Clear
    Me.ListBox1.ColumnCount = 5
    Me.ListBox1.ColumnWidths = "50;120;70;70;70"
    
    Y = 0
    
    For fila = 7 To ultimaFila
        Encontrado = False
        ' El filtro busca en las columnas B (Código), C (Nombre), D (Condición), K (Estado) y O (Plan)
        If LCase(wsBD.Cells(fila, "B").Value) Like "*" & LCase(Filtro) & "*" Or _
           LCase(wsBD.Cells(fila, "C").Value) Like "*" & LCase(Filtro) & "*" Or _
           LCase(wsBD.Cells(fila, "D").Value) Like "*" & LCase(Filtro) & "*" Or _
           LCase(wsBD.Cells(fila, "K").Value) Like "*" & LCase(Filtro) & "*" Or _
           LCase(wsBD.Cells(fila, "O").Value) Like "*" & LCase(Filtro) & "*" Then
            
            Encontrado = True
        End If
        
        If Encontrado Then
            Me.ListBox1.AddItem
            Me.ListBox1.List(Y, 0) = wsBD.Cells(fila, "B").Value
            Me.ListBox1.List(Y, 1) = wsBD.Cells(fila, "C").Value
            Me.ListBox1.List(Y, 2) = wsBD.Cells(fila, "D").Value
            Me.ListBox1.List(Y, 3) = wsBD.Cells(fila, "K").Value
            Me.ListBox1.List(Y, 4) = wsBD.Cells(fila, "O").Value
            
            Y = Y + 1
        End If
    Next fila
End Sub

