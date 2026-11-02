# Limpiador de planillas

Toma un Excel desordenado y entrega uno limpio. Además deja constancia de **todo lo que eliminó**, para que nada se pierda sin que lo sepas.

**Qué corrige:** espacios sobrantes, mayúsculas y minúsculas, emails inválidos, fechas en distintos formatos, montos con puntos de miles, filas duplicadas y filas sin nombre.

---

## 1. Preparación (se hace una sola vez)

1. Instala Python desde https://www.python.org/downloads/
   En Windows, marca la casilla **"Add Python to PATH"** al instalar.
2. Descarga esta carpeta y ábrela con Visual Studio Code (Archivo > Abrir carpeta).
3. Abre la terminal de VS Code (menú Terminal > Nuevo terminal) y escribe:

```
pip install pandas openpyxl
```

Espera a que termine. Si aparece un error, revisa la sección "Si algo falla".

## 2. Probar con datos de ejemplo

En la terminal escribe, una línea a la vez:

```
python crear_datos.py
python limpiar.py
```

Se creará el archivo `ventas_limpias.xlsx`.

## 3. Usar con tu propio archivo

1. Copia tu Excel a esta misma carpeta (por ejemplo `mis_datos.xlsx`).
2. Cierra el archivo en Excel.
3. En la terminal escribe:

```
python limpiar.py mis_datos.xlsx resultado.xlsx
```

Tu archivo debe tener estas columnas en la **fila 1**, escritas exactamente así: `Nombre`, `Email`, `Fecha`, `Monto`.

## 4. Qué obtienes

El archivo de resultado tiene 3 hojas:

| Hoja | Contenido |
|------|-----------|
| Datos limpios | Tu planilla corregida |
| Resumen | Cuántas filas había y cuántas se eliminaron o vaciaron, por motivo |
| Eliminados | Cada dato eliminado, con su fila original, para poder recuperarlo |

La misma lista aparece en la pantalla al ejecutar, pero **desaparece al cerrar la terminal**. Por eso también queda guardada en la hoja "Eliminados".

## 5. Si algo falla

| Mensaje | Qué significa | Solución |
|---------|---------------|----------|
| `python no se reconoce...` | Python no está en el PATH | Reinstala Python marcando "Add Python to PATH". En Mac prueba `python3` |
| `No module named 'pandas'` | Falta instalar las librerías | Ejecuta `pip install pandas openpyxl` (en Mac, `pip3`) |
| `no encuentro el archivo` | Nombre mal escrito o carpeta distinta | Revisa el nombre y que el Excel esté junto a `limpiar.py` |
| `faltan estas columnas` | Los títulos no coinciden | Renombra las columnas a Nombre, Email, Fecha, Monto |
| `no pude guardar` | El archivo de resultado está abierto | Ciérralo en Excel y repite |

## 6. Importante

- Revisa siempre la hoja "Eliminados" antes de entregar el resultado.
- Trabaja con **copias** de tus archivos, nunca con el único original.
- El script asume que el punto en los montos es separador de miles (`20.500` = 20500) y que las fechas van en orden día/mes/año.
