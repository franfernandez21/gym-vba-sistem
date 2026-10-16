VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Registrar 
   Caption         =   "Datos del Cliente"
   ClientHeight    =   6660
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7530
   OleObjectBlob   =   "Registrar.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Registrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Box_Telefono_Change()

End Sub

Private Sub Frame1_Click()

End Sub

' Se ejecuta cuando el formulario se activa.
Private Sub UserForm_Activate()
    ' Declaraci\u00f3n de variables para el \u00e1mbito de toda la subrutina
    Dim wsBD As Worksheet
    Dim FilaEncontrada As Variant
    Dim ultimaFila As Long
    Dim UltimoNumeroID As Long
    Dim celda As Range

    With Me.ComboBox1
        .AddItem "2 SEMANAS"
        .AddItem "1 MES"
    End With
    
    With Me.ComboBox2
        .AddItem "EFECTIVO"
        .AddItem "TRANSFERENCIA"
    End With
    
    ' Comprobamos si la variable global tiene un valor para saber si es un cliente nuevo o existente
    If gCodigoCliente <> "" Then
        ' Modo de modificaci\u00f3n: cargamos los datos del cliente
        Set wsBD = ThisWorkbook.Sheets("Base de datos")
        
        On Error GoTo ErrorHandler
        
        ultimaFila = wsBD.Cells(wsBD.Rows.Count, "B").End(xlUp).Row
        
        ' Buscar la fila del cliente en la Base de datos
        FilaEncontrada = Application.Match(gCodigoCliente, wsBD.Range("B7:B" & ultimaFila), 0)
        
        If Not IsError(FilaEncontrada) Then
            FilaEncontrada = FilaEncontrada + 6
            
            ' Rellenar los campos del formulario con los datos del cliente
            Me.Box_ID.Value = wsBD.Cells(FilaEncontrada, "B").Value
            Me.Box_Name.Value = wsBD.Cells(FilaEncontrada, "C").Value
            Me.ComboBox1.Value = wsBD.Cells(FilaEncontrada, "D").Value
            Me.Box_Fecha.Value = wsBD.Cells(FilaEncontrada, "E").Value
            Me.Box_Total.Value = wsBD.Cells(FilaEncontrada, "F").Value
            Me.Box_Abona.Value = wsBD.Cells(FilaEncontrada, "G").Value
            Me.ComboBox2.Value = wsBD.Cells(FilaEncontrada, "H").Value
            Me.Box_Plan.Value = wsBD.Cells(FilaEncontrada, "O").Value
            Me.Box_Telefono.Value = wsBD.Cells(FilaEncontrada, "P").Value
            
            ' Deshabilita el Box_ID para que no se pueda modificar
            Me.Box_ID.Enabled = False
            
        Else
            MsgBox "No se pudieron cargar los datos del cliente seleccionado.", vbExclamation
        End If
        
        ' Reiniciamos la variable global para que no interfiera en el futuro
        gCodigoCliente = ""
        
    Else
        ' Modo de nuevo cliente: generamos un ID
        Me.Box_ID.Enabled = True
        
        Set wsBD = ThisWorkbook.Sheets("Base de datos")
        
        ' Buscamos la ultima fila con datos en la columna B
        ultimaFila = wsBD.Cells(wsBD.Rows.Count, "B").End(xlUp).Row
        
        UltimoNumeroID = 0
        
        ' Iteramos sobre todos los IDs para encontrar el maximo numero
        If ultimaFila >= 7 Then
            For Each celda In wsBD.Range("B7:B" & ultimaFila)
                ' Verificamos que la celda no este vacia y contenga el prefijo "Cli_"
                If InStr(1, celda.Value, "Cli_", vbTextCompare) > 0 Then
                    Dim numero As Long
                    ' Extraemos solo el numero despues del prefijo "Cli_"
                    On Error Resume Next ' Manejamos errores si el valor no es un numero
                    numero = CLng(Replace(celda.Value, "Cli_", "", 1, 1, vbTextCompare))
                    On Error GoTo 0
                    
                    If numero > UltimoNumeroID Then
                        UltimoNumeroID = numero
                    End If
                End If
            Next celda
        End If
        
        ' Formateamos el nuevo ID con el prefijo "Cli_" y los ceros iniciales
        Me.Box_ID.Value = "Cli_" & Format(UltimoNumeroID + 1, "00000")
        
        ' Deshabilitamos el campo ID para que no se pueda modificar
        Me.Box_ID.Enabled = False
        
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "Ocurrio un error al intentar cargar los datos del cliente: " & Err.Description, vbCritical
End Sub

' Manejo de eventos de los controles del formulario
Private Sub Box_Abona_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii >= 46 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Box_Fecha_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii >= 46 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Box_ID_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii >= 46 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub Box_Name_Change()
    Me.Box_Name.Value = ValidarTexto(Me.Box_Name)
End Sub

Private Sub Box_Total_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If (KeyAscii >= 46 And KeyAscii <= 57) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub CommandButton1_Click()
    Me.Box_Fecha.Value = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub CommandButton2_Click()
    On Error GoTo ErrorHandler
    ' Asumiendo que la subrutina "Validar_Vacios" este en un Modulo Estandar
    ' y que la variable "Contador" es una variable publica en ese mismo modulo.
    Call Validar_Vacios
    
    If contador = 0 Then
        Dim wsBD As Worksheet
        Dim ultimaFilaBD As Long
        Dim wsHistorial As Worksheet
        Dim UltimaFilaHistorial As Long
        Dim FilaExistente As Variant
        
        If Me.Box_ID.Value = "" Or Me.Box_Name.Value = "" Then
            MsgBox "Por favor, complete los campos de ID y Nombre.", vbExclamation
            Exit Sub
        End If
        
        Set wsBD = ThisWorkbook.Sheets("Base de datos")
        Set wsHistorial = ThisWorkbook.Sheets("Historial de Pagos")
        
        On Error Resume Next
        FilaExistente = Application.Match(Me.Box_ID.Value, wsBD.Range("B:B"), 0)
        On Error GoTo 0
        
        If Not IsError(FilaExistente) Then
            ' Modo de modificacion
            wsBD.Cells(FilaExistente, "C").Value = Me.Box_Name.Value
            wsBD.Cells(FilaExistente, "D").Value = Me.ComboBox1.Value
            wsBD.Cells(FilaExistente, "E").Value = Me.Box_Fecha.Value
            wsBD.Cells(FilaExistente, "F").Value = Me.Box_Total.Value
            wsBD.Cells(FilaExistente, "G").Value = Me.Box_Abona.Value
            wsBD.Cells(FilaExistente, "H").Value = Me.ComboBox2.Value
            wsBD.Cells(FilaExistente, "O").Value = Me.Box_Plan.Value
            wsBD.Cells(FilaExistente, "P").Value = Me.Box_Telefono.Value
            
            ' Agregamos el registro de la modificacion al historial de pagos
            If wsHistorial.Range("A2").Value = "" Then
                UltimaFilaHistorial = 2
            Else
                UltimaFilaHistorial = wsHistorial.Cells(wsHistorial.Rows.Count, "A").End(xlUp).Row + 1
            End If
            
            wsHistorial.Cells(UltimaFilaHistorial, "A").Value = Me.Box_ID.Value
            wsHistorial.Cells(UltimaFilaHistorial, "B").Value = Me.Box_Name.Value
            wsHistorial.Cells(UltimaFilaHistorial, "C").Value = Me.Box_Fecha.Value
            wsHistorial.Cells(UltimaFilaHistorial, "D").Value = Me.Box_Plan.Value
            wsHistorial.Cells(UltimaFilaHistorial, "E").Value = Me.ComboBox2.Value
            wsHistorial.Cells(UltimaFilaHistorial, "F").Value = Me.Box_Abona.Value
            
            MsgBox "Cliente modificado correctamente y pago registrado en el historial.", vbInformation
        Else
            ' Modo de nuevo cliente
            If wsBD.Range("B7").Value = "" Then
                ultimaFilaBD = 7
            Else
                ultimaFilaBD = wsBD.Cells(wsBD.Rows.Count, "B").End(xlUp).Row + 1
            End If
            
            wsBD.Cells(ultimaFilaBD, "B").Value = Me.Box_ID.Value
            wsBD.Cells(ultimaFilaBD, "C").Value = Me.Box_Name.Value
            wsBD.Cells(ultimaFilaBD, "D").Value = Me.ComboBox1.Value
            wsBD.Cells(ultimaFilaBD, "E").Value = Me.Box_Fecha.Value
            wsBD.Cells(ultimaFilaBD, "F").Value = Me.Box_Total.Value
            wsBD.Cells(ultimaFilaBD, "G").Value = Me.Box_Abona.Value
            wsBD.Cells(ultimaFilaBD, "H").Value = Me.ComboBox2.Value
            wsBD.Cells(ultimaFilaBD, "O").Value = Me.Box_Plan.Value
            wsBD.Cells(ultimaFilaBD, "P").Value = Me.Box_Telefono.Value
            
            ' Registrar el pago en el historial
            If wsHistorial.Range("A2").Value = "" Then
                UltimaFilaHistorial = 2
            Else
                UltimaFilaHistorial = wsHistorial.Cells(wsHistorial.Rows.Count, "A").End(xlUp).Row + 1
            End If
            
            wsHistorial.Cells(UltimaFilaHistorial, "A").Value = Me.Box_ID.Value
            wsHistorial.Cells(UltimaFilaHistorial, "B").Value = Me.Box_Name.Value
            wsHistorial.Cells(UltimaFilaHistorial, "C").Value = Me.Box_Fecha.Value
            wsHistorial.Cells(UltimaFilaHistorial, "D").Value = Me.Box_Plan.Value
            wsHistorial.Cells(UltimaFilaHistorial, "E").Value = Me.ComboBox2.Value
            wsHistorial.Cells(UltimaFilaHistorial, "F").Value = Me.Box_Abona.Value
            
            MsgBox "Nuevo cliente agregado correctamente.", vbInformation
        End If
        
        Unload Me
    Else
        MsgBox "Por favor, complete los campos vacíos.", vbExclamation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ha ocurrido un error inesperado: " & Err.Description, vbCritical
End Sub
