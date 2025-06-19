Option Explicit

' Calcula el retorno sobre inversión (ROI) en porcentaje
Public Function CalcularROI(utilidadNeta As Double, costosTotales As Double) As Double
    If costosTotales <= 0 Then
        CalcularROI = 0 ' Evita división por cero
    Else
        CalcularROI = (utilidadNeta / costosTotales) * 100
    End If
End Function

' Calcula el ingreso bruto total: servicios realizados * precio unitario
Public Function CalcularIngresoBruto(servicios As Double, precio As Double) As Double
    CalcularIngresoBruto = servicios * precio
End Function

' Calcula el ingreso neto después de aplicar la tasa de comisión.
Public Function CalcularIngresoNeto(ingresoBruto As Double, tasaComision As Double) As Double
    CalcularIngresoNeto = ingresoBruto * (1 - tasaComision)
End Function

' Calcula las horas de trabajo por trabajador en un servicio
Public Function CalcularHorasPorTrabajador(horas As Double, trabajadores As Double) As Double
    If trabajadores > 0 Then
        CalcularHorasPorTrabajador = horas / trabajadores
    Else
        CalcularHorasPorTrabajador = horas
    End If
End Function

' Calcula el costo de mano de obra por servicio
Public Function CalcularCostoManoObra(horasPorTrabajador As Double, salarioHora As Double) As Double
    CalcularCostoManoObra = horasPorTrabajador * salarioHora
End Function

' Suma el costo logístico con el costo de mano de obra por servicio
Public Function CalcularCostoVariableTotalServicio(costoFijoPorServicio As Double, costoManoObra As Double) As Double
    CalcularCostoVariableTotalServicio = costoFijoPorServicio + costoManoObra
End Function

' Calcula el costo variable total para todos los servicios realizados
Public Function CalcularCostoVariableGlobal(costoVariableUnitario As Double, servicios As Double) As Double
    CalcularCostoVariableGlobal = costoVariableUnitario * servicios
End Function

' Calcula el costo fijo mensual asignado a un solo cliente/proyecto
Public Function CalcularCostoFijoAsignado() As Double
    If NumeroNegociosActivos > 0 Then
        CalcularCostoFijoAsignado = CostosFijosMensuales / NumeroNegociosActivos
    Else
        CalcularCostoFijoAsignado = 0
    End If
End Function

' Calcula la utilidad bruta antes de CAC e impuestos
Public Function CalcularUtilidadBruta(ingresoNeto As Double, costoFijo As Double, costoVariableGlobal As Double) As Double
    CalcularUtilidadBruta = ingresoNeto - (costoFijo + costoVariableGlobal)
End Function

' Resta el CAC de la utilidad bruta para obtener utilidad antes de impuestos
Public Function CalcularUtilidadAntesImpuestos(utilidadBruta As Double, cac As Double) As Double
    CalcularUtilidadAntesImpuestos = utilidadBruta - cac
End Function

' Calcula el monto de impuestos corporativos sobre la utilidad antes de impuestos
Public Function CalcularImpuestos(utilidadAntesImpuestos As Double) As Double
    If utilidadAntesImpuestos > 0 Then
        CalcularImpuestos = utilidadAntesImpuestos * TasaImpuestoCorporativo
    Else
        CalcularImpuestos = 0
    End If
End Function

' Calcula la utilidad neta después de impuestos
Public Function CalcularUtilidadNeta(utilidadAntesImpuestos As Double, impuestos As Double) As Double
    CalcularUtilidadNeta = utilidadAntesImpuestos - impuestos
End Function