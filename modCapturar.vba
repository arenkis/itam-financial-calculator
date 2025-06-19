Option Explicit

' Variables globales para usar en otros módulos
Public CostosFijosMensuales As Double
Public NumeroNegociosActivos As Long
Public TasaImpuestoCorporativo As Double

' Solicita y valida los datos generales necesarios para los cálculos
Public Function CapturarDatos() As Boolean
    ' Manejar errores
    On Error GoTo ErrHandler
    
    Dim entrada As Variant
    Dim msg As String

    ' Costos Fijos Mensuales
    entrada = InputBox("Ingrese el monto de Costos Fijos Mensuales (USD):", "Costos Fijos")
    If Not ValidarValorNumerico("Costos Fijos Mensuales", entrada) Then
        CapturarDatos = False
        Exit Function
    End If
    CostosFijosMensuales = CDbl(entrada)

    ' Número de Negocios Activos
    entrada = InputBox("Ingrese el número de negocios o proyectos activos:", "Negocios Activos")
    If Not ValidarValorNumerico("Número de Negocios Activos", entrada) Then
        CapturarDatos = False
        Exit Function
    End If
    NumeroNegociosActivos = CLng(entrada)

    ' Tasa de Impuesto Corporativo
    entrada = InputBox("Ingrese la tasa de impuesto corporativo (%):", "Tasa de Impuestos")
    If Not ValidarEnRango("Tasa de Impuestos", entrada, 0, 100) Then
        CapturarDatos = False
        Exit Function
    End If
    TasaImpuestoCorporativo = NormalizarPorcentaje("Tasa de Impuestos", entrada)

    ' Confirmación final de datos capturados correctamente
    msg = "Datos capturados correctamente." & vbCrLf & _
           "Costos Fijos Mensuales: " & FormatCurrency(CostosFijosMensuales) & vbCrLf & _
           "Negocios Activos: " & NumeroNegociosActivos & vbCrLf & _
           "Tasa de Impuestos: " & FormatPercent(TasaImpuestoCorporativo)

    MsgBox msg, vbInformation, "Confirmación"
    CapturarDatos = True
    Exit Function

ErrHandler:
    MsgBox "Ocurrió un error interno: " & Err.Description, vbExclamation
    CapturarDatos = False
End Function