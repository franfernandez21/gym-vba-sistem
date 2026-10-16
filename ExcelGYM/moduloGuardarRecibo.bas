Attribute VB_Name = "moduloGuardarRecibo"
Sub guardarRecibo()
    ' Asume que la hoja de la plantilla de recibo se llama "Recibo"
    ' y la hoja de la base de datos se llama "Base de Datos de Recibos".
    Dim wsRecibo As Worksheet
    Dim wsBD As Worksheet
    
    Set wsRecibo = ThisWorkbook.Sheets("Recibo")
    Set wsBD = ThisWorkbook.Sheets("Base de Datos de Recibos")

    ' --- Define las celdas con los datos del recibo en tu plantilla ---
    Dim idRecibo As Variant
    idRecibo = wsRecibo.Range("I5").Value
    
    Dim idCliente As String
    idCliente = wsRecibo.Range("I8").Value
    
    Dim nombreCliente As String
    nombreCliente = wsRecibo.Range("D9").Value
    
    Dim descripcion As String
    descripcion = wsRecibo.Range("E16").Value
    
    Dim total As Double
    total = wsRecibo.Range("I20").Value

    ' --- Encuentra la primera fila vacía en la base de datos ---
    Dim ultimaFila As Long
    ' Primero, verifica si la celda A2 está vacía (es la primera fila de datos)
    If wsBD.Range("A2").Value = "" Then
        ultimaFila = 2
    Else
        ultimaFila = wsBD.Cells(wsBD.Rows.Count, "A").End(xlUp).Row + 1
    End If

    ' --- Copia y pega los datos en la base de datos ---
    wsBD.Cells(ultimaFila, 1).Value = idRecibo
    wsBD.Cells(ultimaFila, 2).Value = idCliente
    wsBD.Cells(ultimaFila, 3).Value = nombreCliente
    wsBD.Cells(ultimaFila, 4).Value = Date
    wsBD.Cells(ultimaFila, 5).Value = descripcion
    wsBD.Cells(ultimaFila, 6).Value = total

    ' --- Lógica de contador que reinicia diariamente ---
    Dim contador As Long
    Dim fechaID As String
    Dim ultimoID As String
    Dim fechaUltimoID As String
    
    ' Obtiene la fecha actual formateada (ej. 20250902)
    fechaID = Format(Date, "YYYYMMDD")
    
    ' Encuentra la última fila de la base de datos para obtener el último ID
    Dim ultimaFilaParaID As Long
    ultimaFilaParaID = wsBD.Cells(wsBD.Rows.Count, "A").End(xlUp).Row
    
    ' Verifica si hay algún recibo guardado previamente
    If ultimaFilaParaID > 1 Then ' > 1 porque la fila 1 es el encabezado
        ' Obtiene el último ID y la fecha del mismo
        ultimoID = wsBD.Cells(ultimaFilaParaID, "A").Value
        fechaUltimoID = Left(ultimoID, 8) ' Extrae la fecha (ej. 20250902)
        
        ' Compara la fecha de hoy con la fecha del último ID
        If fechaID = fechaUltimoID Then
            ' Si las fechas coinciden, extrae el numero y lo incrementa
            contador = Right(ultimoID, 3) + 1
        Else
            ' Si las fechas son diferentes, reinicia el contador
            contador = 1
        End If
    Else
        ' Si es el primer recibo en la base de datos, el contador es 1
        contador = 1
    End If
    
    ' Formatea el contador para que tenga 3 dígitos (ej. 001)
    Dim nuevoContador As String
    nuevoContador = Format(contador, "000")
    
    ' Crea el nuevo ID completo y lo asigna en la celda I5 del recibo
    wsRecibo.Range("I5").Value = fechaID & "-" & nuevoContador
    
End Sub
