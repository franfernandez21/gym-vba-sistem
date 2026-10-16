Attribute VB_Name = "Generar_Codigo"
Public Codigo_Cliente As String
Public gModoEdicion As Boolean
' Generador de Código de Cliente (Versión final)
' Maneja errores y respeta el formato "Cli_XXXXX"
Sub Genera_Codigo()
    Dim wsBD As Worksheet
    Dim i As Long
    Dim j As Long
    Dim Dato As String
    Dim Posicion As Long
    Dim numero As Long
    Dim Maximo As Long
    Dim miArray() As Long
    
    ' Definir la hoja de la base de datos
    Set wsBD = ThisWorkbook.Sheets("Base de datos")
    
    On Error GoTo ErrorHandler
    
    ' Encontrar la última fila con datos en la columna B
    Dim ultimaFila As Long
    ultimaFila = wsBD.Cells(wsBD.Rows.Count, "B").End(xlUp).Row
    
    ' Si la base de datos solo tiene encabezados (o no tiene datos), el primer ID es Cli_00001.
    If ultimaFila <= 6 Then ' <-- Se ajusta para que el primer cliente esté en la fila 7.
        Codigo_Cliente = "Cli_00001"
        Exit Sub
    End If
    
    j = 0
    ' La principal mejora está aquí: se ajusta el inicio del bucle a la fila 7.
    For i = 7 To ultimaFila
        Dato = wsBD.Cells(i, "B").Value
        
        If Dato <> "" And InStr(Dato, "_") > 0 Then
            Posicion = InStr(Dato, "_")
            numero = CLng(Mid(Dato, Posicion + 1))
            
            ReDim Preserve miArray(0 To j)
            miArray(j) = numero
            
            j = j + 1
        End If
    Next i
    
    ' Se valida si el array tiene elementos antes de usarlo.
    If j = 0 Then
        ' Si el array está vacío, significa que no se encontraron códigos válidos.
        Codigo_Cliente = "Cli_00001"
    Else
        ' Si se encontraron códigos, se encuentra el máximo y se le suma 1.
        Maximo = WorksheetFunction.Max(miArray) + 1
        Codigo_Cliente = "Cli_" & Format(Maximo, "00000")
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ha ocurrido un error inesperado al generar el código. Se asignará el código inicial.", vbExclamation
    Codigo_Cliente = "Cli_00001"
End Sub
