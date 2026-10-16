Attribute VB_Name = "Módulo1"
Public gCodigoCliente As String
Sub Llamar_Formulario_Registrar()
Registrar.Show
End Sub
Sub Llamar_Buscar()
Buscar.Show
End Sub
Sub ActualizarNumeroDeFactura()
    
    Dim wsFactura As Worksheet
    Dim wsHojaPagos As Worksheet
    Dim loTablaPagos As ListObject
    Dim proximoNumero As Long
    
    ' --- CONFIGURACIÓN ---
    ' La hoja donde quieres que aparezca el número.
    ' La celda donde quieres que aparezca el nuevo número.
    Const NOMBRE_HOJA_FACTURA As String = "Factura"
    Const NOMBRE_HOJA_PAGOS As String = "Historial de Pagos"
    Const CELDA_DESTINO As String = "i5"
    
    On Error GoTo ErrorHandler
    
    ' Referencia a las hojas y la tabla
    Set wsFactura = ThisWorkbook.Sheets(NOMBRE_HOJA_FACTURA)
    Set wsHojaPagos = ThisWorkbook.Sheets(NOMBRE_HOJA_PAGOS)
    Set loTablaPagos = wsHojaPagos.ListObjects("Tabla_Pagos")
    
    ' Obtiene la cantidad de filas en la tabla y le suma 1
    proximoNumero = loTablaPagos.ListRows.Count + 1
    
    ' Pone el resultado en la celda destino
    wsFactura.Range(CELDA_DESTINO).Value = proximoNumero
    Exit Sub
ErrorHandler:
    ' No muestra mensaje al usuario, simplemente sale
    Exit Sub
End Sub

Sub ActualizarFechaDeHoy()
    
    Dim wsFactura As Worksheet
    
    ' --- CONFIGURACIÓN ---
    ' La hoja donde quieres que aparezca la fecha.
    ' La celda donde quieres que aparezca la fecha.
    Const NOMBRE_HOJA_FACTURA As String = "Recibo"
    Const CELDA_DESTINO_FECHA As String = "i9"
    
    On Error GoTo ErrorHandler
    
    Set wsFactura = ThisWorkbook.Sheets(NOMBRE_HOJA_FACTURA)
    
    ' Pone la fecha actual en la celda destino.
    wsFactura.Range(CELDA_DESTINO_FECHA).Value = Date
    
    Exit Sub
ErrorHandler:
    Exit Sub
End Sub
