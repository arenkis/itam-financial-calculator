Option Explicit

Public Function ProcesarDatos() As Boolean
    ' Manejar errores
    On Error GoTo ErrHandler
    
    ' Leer las hojas
    Dim hojaDatos As Worksheet
    Dim hojaResultados As Worksheet
    Set hojaDatos = ThisWorkbook.Sheets("DatosEntrada")
    Set hojaResultados = ThisWorkbook.Sheets("Resultados")

    ' Borrar resultados anteriores
    hojaResultados.Range("E2:E14").ClearContents

    ' Variables de entrada
    Dim nombreNegocio As String
    Dim serviciosRealizados As Double, precioServicio As Double, costoServicio As Double
    Dim salarioHora As Double, horasPorServicio As Double, numeroTrabajadores As Double
    Dim tasaComision As Double, cac As Double

    ' Tabla de campos y tipos de validación
    Dim campos() As Variant
    campos = Array( _
        Array("Nombre del Negocio", "texto"), _
        Array("Servicios Realizados", "mayor0"), _
        Array("Precio por Servicio", "mayor0"), _
        Array("Costo por Servicio", "mayor0"), _
        Array("Salario por Hora", "mayor0"), _
        Array("Horas por Servicio", "mayor0"), _
        Array("Número de Trabajadores", "mayor0"), _
        Array("Tasa de Comisión", "porcentaje"), _
        Array("CAC", "esNumero") _
    )

    Dim v As Variant
    Dim texto As String
    Dim i As Integer
    Dim fieldName As String

    ' Leer y validar campos
    For i = 0 To UBound(campos)
        fieldName = campos(i)(0)
        Select Case campos(i)(1)

            Case "texto"
                texto = BuscarValor(fieldName, hojaDatos)
                If Not ValidarTexto(fieldName, texto) Then Exit Function
                If fieldName = "Nombre del Negocio" Then nombreNegocio = texto

            Case "mayor0"
                v = BuscarValor(fieldName, hojaDatos)
                If Not ValidarValorNumerico(fieldName, v) Then Exit Function
                Select Case fieldName
                    Case "Servicios Realizados": serviciosRealizados = v
                    Case "Precio por Servicio": precioServicio = v
                    Case "Costo por Servicio": costoServicio = v
                    Case "Salario por Hora": salarioHora = v
                    Case "Horas por Servicio": horasPorServicio = v
                    Case "Número de Trabajadores": numeroTrabajadores = v
                End Select

            Case "porcentaje"
                v = BuscarValor(fieldName, hojaDatos)
                tasaComision = NormalizarPorcentaje(fieldName, v)

            Case "esNumero"
                v = BuscarValor(fieldName, hojaDatos)
                If Not ValidarEsNumero(fieldName, v) Then Exit Function
                If fieldName = "CAC" Then cac = v

        End Select
    Next i

    ' Calcular resultados
    Dim resultados(1 To 12) As Double
    For i = 1 To 12
        Select Case i
            Case 1
                resultados(i) = CalcularIngresoBruto(serviciosRealizados, precioServicio)
            Case 2
                resultados(i) = CalcularIngresoNeto(resultados(1), tasaComision)
            Case 3
                resultados(i) = CalcularHorasPorTrabajador(horasPorServicio, numeroTrabajadores)
            Case 4
                resultados(i) = CalcularCostoManoObra(resultados(3), salarioHora)
            Case 5
                resultados(i) = CalcularCostoVariableTotalServicio(costoServicio, resultados(4))
            Case 6
                resultados(i) = CalcularCostoVariableGlobal(resultados(5), serviciosRealizados)
            Case 7
                resultados(i) = CalcularCostoFijoAsignado()
            Case 8
                resultados(i) = CalcularUtilidadBruta(resultados(2), resultados(7), resultados(6))
            Case 9
                resultados(i) = CalcularUtilidadAntesImpuestos(resultados(8), cac)
            Case 10
                resultados(i) = CalcularImpuestos(resultados(9))
            Case 11
                resultados(i) = CalcularUtilidadNeta(resultados(9), resultados(10))
            Case 12
                resultados(i) = CalcularROI(resultados(11), resultados(7) + resultados(6) + cac)
        End Select
    Next i

    ' Escribir resultados
    hojaResultados.Range("E2").Value = nombreNegocio
    Dim ultimaFila As Long
    ultimaFila = hojaResultados.Cells(hojaResultados.Rows.Count, 1).End(xlUp).Row

    For i = 3 To ultimaFila
        hojaResultados.Range("E" & i).Value = resultados(i - 2)
    Next i
    
    ' Manejar errores
    ProcesarDatos = True
    Exit Function
ErrHandler:
    MsgBox "Ocurrió un error interno: " & Err.Description, vbExclamation
    ProcesarDatos = False
End Function

' Buscar los datos de entrada
Public Function BuscarValor(nombreCampo As String, hoja As Worksheet) As Variant
    If Not ValidarExistenciaCampo(nombreCampo, hoja) Then
        MsgBox "No se puede continuar sin el campo: " & nombreCampo, vbCritical, "Campo requerido"
        End
    End If

    Dim i As Long
    For i = 2 To hoja.Cells(hoja.Rows.Count, 1).End(xlUp).Row
        If Trim(hoja.Cells(i, 1).Value) = nombreCampo Then
            BuscarValor = hoja.Cells(i, 5).Value
            Exit Function
        End If
    Next i
End Function