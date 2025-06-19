Option Explicit

' Valida que un valor sea numérico y mayor que cero
Public Function ValidarValorNumerico(nombreCampo As String, valor As Variant) As Boolean
    If Not IsNumeric(valor) Then
        MsgBox "Error: El campo '" & nombreCampo & "' debe ser un valor numérico.", vbExclamation, "Error de Validación"
        ValidarValorNumerico = False
        Exit Function
    ElseIf CDbl(valor) <= 0 Then
        MsgBox "Error: El campo '" & nombreCampo & "' debe ser un número mayor que cero.", vbExclamation, "Error de Validación"
        ValidarValorNumerico = False
        Exit Function
    Else
        ValidarValorNumerico = True
    End If
End Function

' Valida que un valor sea texto no vacío
Public Function ValidarTexto(nombreCampo As String, texto As Variant) As Boolean
    If IsNumeric(texto) Then
        MsgBox "Error: El campo '" & nombreCampo & "' no debe ser un número.", vbExclamation, "Error de Validación"
        ValidarTexto = False
        Exit Function
    ElseIf Not EsCadena(texto) Then
        MsgBox "Error: El campo '" & nombreCampo & "' debe ser un texto.", vbExclamation, "Error de Validación"
        ValidarTexto = False
        Exit Function
    ElseIf Trim(texto) = "" Then
        MsgBox "Error: El campo '" & nombreCampo & "' no puede estar vacío.", vbExclamation, "Error de Validación"
        ValidarTexto = False
        Exit Function
    Else
        ValidarTexto = True
    End If
End Function

' Valida que un valor esté dentro de un rango definido
Public Function ValidarEnRango(nombreCampo As String, valor As Variant, minimo As Double, maximo As Double) As Boolean
    If Not IsNumeric(valor) Then
        MsgBox "Error: El campo '" & nombreCampo & "' debe ser numérico.", vbExclamation, "Error de Validación"
        ValidarEnRango = False
        Exit Function
    ElseIf CDbl(valor) < minimo Or CDbl(valor) > maximo Then
        MsgBox "Error: El campo '" & nombreCampo & "' debe estar entre " & minimo & " y " & maximo & ".", vbExclamation, "Error de Validación"
        ValidarEnRango = False
        Exit Function
    Else
        ValidarEnRango = True
    End If
End Function

' Valida que un campo exista en la hoja de entrada
Public Function ValidarExistenciaCampo(nombreCampo As String, hoja As Worksheet) As Boolean
    Dim i As Long
    For i = 2 To hoja.Cells(hoja.Rows.Count, 1).End(xlUp).Row
        If Trim(hoja.Cells(i, 1).Value) = nombreCampo Then
            ValidarExistenciaCampo = True
            Exit Function
        End If
    Next i
    MsgBox "Error: El campo '" & nombreCampo & "' no se encuentra en la hoja '" & hoja.Name & "'.", vbCritical, "Campo faltante"
    ValidarExistenciaCampo = False
End Function

' Valida fila horizontal (si se usa formato de fila)
Public Function ValidarFilaDatos(hoja As Worksheet, fila As Long) As Boolean
    Dim servicios As Variant
    Dim precio As Variant
    servicios = hoja.Cells(fila, 4).Value
    precio = hoja.Cells(fila, 5).Value

    If Not IsNumeric(servicios) Then
        MsgBox "Advertencia: Fila " & fila & " tiene número de servicios inválido.", vbExclamation
        ValidarFilaDatos = False
        Exit Function
    ElseIf CDbl(servicios) <= 0 Then
        MsgBox "Advertencia: Fila " & fila & " debe tener servicios > 0.", vbExclamation
        ValidarFilaDatos = False
        Exit Function
    ElseIf Not IsNumeric(precio) Then
        MsgBox "Advertencia: Fila " & fila & " tiene precio por servicio inválido.", vbExclamation
        ValidarFilaDatos = False
        Exit Function
    ElseIf CDbl(precio) <= 0 Then
        MsgBox "Advertencia: Fila " & fila & " debe tener precio > 0.", vbExclamation
        ValidarFilaDatos = False
        Exit Function
    Else
        ValidarFilaDatos = True
    End If
End Function

' Valida que un valor sea numérico
Public Function ValidarEsNumero(nombreCampo As String, valor As Variant) As Boolean
    If Not IsNumeric(valor) Then
        MsgBox "Error: El campo '" & nombreCampo & "' debe ser numérico.", vbExclamation, "Error de Validación"
        ValidarEsNumero = False
        Exit Function
    Else
        ValidarEsNumero = True
    End If
End Function

' Valida que un valor sea string
Public Function EsCadena(v As Variant) As Boolean
    EsCadena = (VarType(v) = vbString)
End Function

' Convierte un valor de porcentaje flexible (20 o 0.2) a decimal (0.2)
Public Function NormalizarPorcentaje(nombreCampo As String, valorOriginal As Variant) As Double
    Dim valor As Double

    If Not IsNumeric(valorOriginal) Then
        MsgBox "Error: El campo '" & nombreCampo & "' debe ser numérico (ej. 20 o 0.2).", vbExclamation, "Error de Validación"
        Exit Function
    Else
        valor = CDbl(valorOriginal)
    End If

    If valor > 1 And valor <= 100 Then
        valor = valor / 100
    End If

    If valor < 0 Or valor > 1 Then
        MsgBox "Error: El campo '" & nombreCampo & "' debe estar entre 0 y 1 (o entre 0% y 100%).", vbExclamation, "Error de Validación"
        Exit Function
    End If

    NormalizarPorcentaje = valor
End Function