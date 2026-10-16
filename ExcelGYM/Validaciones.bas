Attribute VB_Name = "Validaciones"
Function ValidarTexto(Texto As Variant)

    Dim Caracter As Variant
    Dim Largo As String
    Dim i As Integer
    On Error Resume Next
    Largo = Len(Texto)
    
    For i = 1 To Largo
        Caracter = CInt(Mid(Texto, i, 1))
        
        If Caracter <> "" Then
            If Not Application.WorksheetFunction.IsText(Caracter) Then
                Texto = Replace(Texto, Caracter, "")
                ValidarTexto = Texto
            Else
            End If
        End If
        
    Next i
    
    ValidarTexto = Texto
    On Error GoTo 0
    
End Function
