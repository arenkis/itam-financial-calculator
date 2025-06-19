# Calculadora Financiera VBA para Excel

Este proyecto implementa, totalmente en **VBA**, un flujo de trabajo para evaluar la rentabilidad de un negocio o proyecto **dentro de un libro de Excel**.  
Incluye:

* Captura interactiva y validación de datos clave (ingresos, costos, CAC, etc.).
* Cálculo de las principales métricas financieras (ingreso neto, utilidad bruta, impuestos, ROI, …).
* Generación automática de una **gráfica de columnas** con los resultados en USD.
* Código modular y comentado para que puedas extenderlo o adaptarlo.

---

## Estructura del repositorio

```

itam-financial-calculator/
├─ .gitignore                    # ignora binarios y archivos temporales
├─ financial-calculator.xlsm     # libro Excel con las macros listas para usar
├─ LICENSE                       # MIT
├─ modCalcular.vba               # funciones financieras puras
├─ modCapturar.vba               # captura y validación básica
├─ modGraficar.vba               # creación de la gráfica
├─ modMain.vba                   # orquestador: EjecutarProyecto
├─ modProcesar.vba               # lógica de cálculo / escritura de resultados
├─ modValidar.vba                # validaciones generales (numérico, porcentaje, etc.)
└─ README.md                     # este archivo

```

---

## Hojas del libro

| Hoja             | Propósito                                                           |
| ---------------- | ------------------------------------------------------------------- |
| **DatosEntrada** | Captura manual de las variables (col. A = nombre, col. E = valor).  |
| **Resultados**   | Salida numérica: filas 2-14 corresponden a cada métrica calculada.  |
| **Gráficas**     | Se limpia en cada ejecución y alberga un gráfico de columnas (USD). |

---

## Módulos y procedimientos clave

| Módulo `.vba` | Rutinas destacadas                                              | Rol principal                                          |
| ------------- | --------------------------------------------------------------- | ------------------------------------------------------ |
| `modCalcular` | `CalcularIngresoBruto`, `CalcularROI`, etc.                     | Fórmulas financieras puras.                            |
| `modCapturar` | `CapturarDatos`, `ValidarValorNumerico`, `NormalizarPorcentaje` | Diálogo con el usuario; guarda variables globales.     |
| `modProcesar` | `ProcesarDatos`, `BuscarValor`, `ValidarFilaDatos`              | Orquestación de cálculos y escritura en **Resultados** |
| `modGraficar` | `GraficarDatos`                                                 | Dibuja el gráfico en **Gráficas**.                     |
| `modMain`     | `EjecutarProyecto`                                              | Llama a captura → proceso → gráfica → mensajes.        |
| `modValidar`  | Helpers de validación re-usados por varios módulos.             |                                                        |

---

## Uso rápido

1. **Clona o descarga** el repositorio, habilita macros en Excel y abre `financial-calculator.xlsm`.

   > Si prefieres importar solo el código, abre cualquier libro nuevo y **VBA > Import File…** con los módulos `.vba`.
2. En la hoja **DatosEntrada** introduce los valores necesarios (ver lista dentro de `modProcesar`: “Servicios Realizados”, “Precio por Servicio”, “CAC”, etc.).
3. Ejecuta la macro **`EjecutarProyecto`**:

   * **Alt + F8** → selecciona `EjecutarProyecto` → **Run**
   * o asigna la macro a un botón de la cinta.
4. Revisa:

   * Hoja **Resultados**: métricas numéricas finales.
   * Hoja **Gráficas**: gráfico “Métricas Financieras (USD)”.
   * Ventanas emergentes de confirmación o advertencia.

---

## Personalización

* **Agregar métricas**:

  * Añade una fila descriptiva en **Resultados**.
  * Extiende el `Select Case` en `modProcesar` para incluir la nueva fórmula.

* **Ajustar validaciones**:

  * Modifica las funciones de `modValidar` para cambiar rangos o tipos permitidos.

* **Cambiar la gráfica**:

  * Edita `modGraficar`: `ChartType`, títulos, ejes, colores, etc.

---

## Requisitos

* **Microsoft Excel** con soporte VBA (Windows o macOS — Excel 365 / 2019+).
* Permisos para habilitar macros.
* No se necesitan complementos externos: todo el proyecto es puro VBA.

---

## Ejemplo de ejecución desde el Editor VBA

```vb
Sub Demo()
    If Not CapturarDatos     Then Exit Sub    ' Paso 1
    If Not ProcesarDatos     Then Exit Sub    ' Paso 2
    Call GraficarDatos                       ' Paso 3
End Sub
```

---

## Licencia

Distribuido bajo la licencia **MIT**.
Consulta el archivo `LICENSE` para más información.