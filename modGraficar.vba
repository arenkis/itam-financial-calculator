Option Explicit

' Graficar las métricas financieras
Public Function GraficarDatos() As Boolean
    ' Manejar errores
    On Error GoTo ErrHandler
    
    Dim hojaDatos     As Worksheet
    Dim hojaGraficas  As Worksheet
    Dim ultimaFila    As Long
    Dim fila          As Long
    Dim rngEtiquetas  As Range
    Dim rngValores    As Range
    Dim graficaObjeto As ChartObject

    Set hojaDatos = ThisWorkbook.Sheets("Resultados")
    Set hojaGraficas = ThisWorkbook.Sheets("Gráficas")

    ' Eliminar gráficos existentes
    For Each graficaObjeto In hojaGraficas.ChartObjects
        graficaObjeto.Delete
    Next graficaObjeto

    ' Determinar última fila con datos en columna E
    ultimaFila = hojaDatos.Cells(hojaDatos.Rows.Count, "E").End(xlUp).Row

    ' Construir rangos sólo con unidad "USD"
    For fila = 2 To ultimaFila
        If UCase(hojaDatos.Cells(fila, "B").Value) = "USD" Then
            If rngEtiquetas Is Nothing Then
                Set rngEtiquetas = hojaDatos.Cells(fila, "A")
                Set rngValores = hojaDatos.Cells(fila, "E")
            Else
                Set rngEtiquetas = Union(rngEtiquetas, hojaDatos.Cells(fila, "A"))
                Set rngValores = Union(rngValores, hojaDatos.Cells(fila, "E"))
            End If
        End If
    Next fila

    If rngEtiquetas Is Nothing Then
        MsgBox "No se encontraron métricas en USD para graficar.", vbExclamation
        Exit Function
    End If

    ' Crear y configurar la gráfica
    Set graficaObjeto = hojaGraficas.ChartObjects.Add(Left:=0, Top:=0, Width:=1000, Height:=500)

    With graficaObjeto.Chart
        .ChartType = xlColumnClustered
        .HasLegend = False
        .SeriesCollection.NewSeries
        With .SeriesCollection(1)
            .XValues = rngEtiquetas
            .Values = rngValores
        End With
        .HasTitle = True
        .ChartTitle.Text = "Métricas Financieras (USD)"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Campo"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Valor (USD)"
        .Axes(xlValue).HasMajorGridlines = True
    End With
    
    ' Manejar errores
    GraficarDatos = True
    Exit Function
ErrHandler:
    MsgBox "Ocurrió un error interno: " & Err.Description, vbExclamation
    GraficarDatos = False
End Function