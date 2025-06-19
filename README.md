# Calculadora Financiera VBA para Excel

Este proyecto de **VBA** implementa un conjunto completo de funciones y macros para evaluar la rentabilidad de un negocio o proyecto dentro de una hoja de cálculo de Excel.
Permite capturar los datos clave de la operación, validar su consistencia, calcular las principales métricas financieras —ingresos, costos, utilidades, impuestos y ROI— y generar un gráfico con los resultados.

---

## Estructura del libro

| Hoja             | Propósito                                                                       | Comentarios                                                      |
| ---------------- | ------------------------------------------------------------------------------- | ---------------------------------------------------------------- |
| **DatosEntrada** | Captura de variables de entrada (columna A: nombre del campo, columna E: valor) | Se valida automáticamente la existencia y el tipo de cada campo. |
| **Resultados**   | Presenta los cálculos finales (columna E) y sirve de fuente para la gráfica     | Las filas 2 – 14 se corresponden con cada métrica calculada.     |
| **Gráficas**     | Hoja de salida donde se dibuja el gráfico de columnas con las métricas en USD   | Se limpia en cada ejecución para evitar duplicados.              |

---

## Principales módulos y procedimientos

| Módulo                | Procedimiento / Función                                                                                                   | Descripción                                                                                           |
| --------------------- | ------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------- |
| **Finanzas.bas**      | `CalcularIngresoBruto`, `CalcularIngresoNeto`, `CalcularCostoVariableGlobal`, `CalcularUtilidadNeta`, `CalcularROI`, etc. | Funciones puramente aritméticas que devuelven cada métrica financiera.                                |
| **GlobVars.bas**      | Variables públicas (`CostosFijosMensuales`, `NumeroNegociosActivos`, `TasaImpuestoCorporativo`)                           | Se inicializan una sola vez mediante `CapturarDatos()`.                                               |
| **Captura.bas**       | `CapturarDatos`, `ValidarValorNumerico`, `NormalizarPorcentaje`, etc.                                                     | Solicitan datos al usuario con `InputBox`, verifican tipos y rangos y guardan las variables globales. |
| **Procesamiento.bas** | `ProcesarDatos`, `BuscarValor`, `ValidarFilaDatos`                                                                        | Lee las hojas, ejecuta las funciones de cálculo en orden lógico y escribe los resultados.             |
| **Graficas.bas**      | `GraficarDatos`                                                                                                           | Crea una gráfica de columnas (métricas en USD) en la hoja **Gráficas**.                               |
| **Main.bas**          | `EjecutarProyecto`                                                                                                        | Orquestador: llama a captura, procesamiento, graficado y muestra mensajes de éxito o error.           |

---

## Flujo de uso

1. **Abrir el libro** y habilitar macros.
2. Rellenar la hoja **DatosEntrada** con los campos requeridos
   (ver lista en `ProcesarDatos`, por ejemplo *“Servicios Realizados”*, *“Precio por Servicio”*, *“CAC”*, etc.).
3. Ejecutar la macro `EjecutarProyecto` (desde el Explorador de macros o asignándola a un botón).
4. Revisar:

   * Hoja **Resultados**: valores numéricos calculados.
   * Hoja **Gráficas**: gráfico de barras “Métricas Financieras (USD)”.
   * Ventanas emergentes de confirmación o advertencia.

---

## Personalización

* **Agregar nuevos indicadores**
  Añadir una fila descriptiva en **Resultados** y ampliar el `Select Case` dentro de `ProcesarDatos` para incluir la fórmula correspondiente.
* **Cambiar validaciones**
  Modificar las funciones de la sección *Validación* para ajustar rangos o tipos permitidos.
* **Adaptar la gráfica**
  Editar `GraficarDatos`: cambiar tipo de gráfico (`ChartType`), títulos o ejes.

---

## Requisitos

* Microsoft Excel con soporte VBA (Windows o macOS con Excel 365 / 2019+).
* Permisos para habilitar macros.
* No se necesitan complementos externos: todo el código está en módulos `.bas`.

---

## Ejecución rápida desde el Editor VBA

```vb
Sub Demo()
    ' 1) Captura variables globales
    If Not CapturarDatos Then Exit Sub
    
    ' 2) Procesa datos y escribe resultados
    If Not ProcesarDatos Then Exit Sub
    
    ' 3) Dibuja la gráfica
    Call GraficarDatos
End Sub
```

---

## Licencia

Este proyecto se distribuye bajo la licencia MIT. Consulte el archivo `LICENSE` para más detalles.