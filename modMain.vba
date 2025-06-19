Option Explicit

' Ejecutar el proyecto
Public Sub EjecutarProyecto()
    ' Paso 1: Captura de datos generales
    If Not CapturarDatos() Then Exit Sub
    
    ' Paso 2: Procesamiento de datos
    If Not ProcesarDatos() Then Exit Sub
    
    ' Paso 3: Generar gráfica
    If Not GraficarDatos() Then Exit Sub
    
    ' Paso 4: Confirmación de éxito
    MsgBox "El proyecto se ejecutó correctamente.", vbInformation, "Finalizado"
End Sub