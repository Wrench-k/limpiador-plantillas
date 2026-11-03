"""Limpiador de planillas: limpia un Excel y deja constancia de todo lo eliminado."""
import sys
from collections import Counter

import pandas as pd

# Evita que las tildes se vean como "?" en algunas terminales de Windows
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

COLUMNAS = ["Nombre", "Email", "Fecha", "Monto"]
PATRON_EMAIL = r"^[\w\.\+-]+@[\w\.-]+\.\w+$"
FORMATOS_FECHA = ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y")  # día primero (formato chileno)


def parsear_fechas(serie):
    resultado = pd.Series(pd.NaT, index=serie.index, dtype="datetime64[ns]")
    for formato in FORMATOS_FECHA:
        intento = pd.to_datetime(serie, format=formato, errors="coerce")
        resultado = resultado.fillna(intento)
    return resultado


def limpiar(ruta_entrada, ruta_salida):
    try:
        original = pd.read_excel(ruta_entrada, dtype=str)
    except FileNotFoundError:
        print(f"ERROR: no encuentro el archivo '{ruta_entrada}'.")
        print("Revisa el nombre y que esté en la misma carpeta que limpiar.py.")
        return

    faltantes = [c for c in COLUMNAS if c not in original.columns]
    if faltantes:
        print(f"ERROR: faltan estas columnas en el Excel: {faltantes}")
        print(f"Columnas encontradas: {list(original.columns)}")
        return

    df = original.copy()
    eliminados = []

    def registrar(motivo, indice, columna, valor):
        eliminados.append({
            "Motivo": motivo,
            "Fila en Excel": indice + 2,  # +2: encabezado y conteo desde 0
            "Columna": columna,
            "Valor eliminado": valor,
        })

    def fila_texto(indice):
        valores = ["(vacío)" if pd.isna(v) else v for v in original.loc[indice]]
        return " | ".join(valores)

    # 1. Quitar espacios sobrantes
    for col in COLUMNAS:
        df[col] = df[col].str.strip().str.replace(r"\s+", " ", regex=True)

    # 2. Normalizar nombres y emails
    df["Nombre"] = df["Nombre"].str.title()
    df["Email"] = df["Email"].str.lower()

    # 3. Emails inválidos
    valido = df["Email"].str.contains(PATRON_EMAIL, na=False)
    for idx in df.index[~valido & df["Email"].notna()]:
        registrar("Email inválido (quedó vacío)", idx, "Email", original.at[idx, "Email"])
    df.loc[~valido, "Email"] = None

    # 4. Fechas
    fechas = parsear_fechas(df["Fecha"])
    for idx in df.index[fechas.isna() & df["Fecha"].notna()]:
        registrar("Fecha no reconocida (quedó vacía)", idx, "Fecha", original.at[idx, "Fecha"])
    df["Fecha"] = fechas.dt.strftime("%Y-%m-%d")

    # 5. Montos (se asume que el punto es separador de miles)
    montos = pd.to_numeric(df["Monto"].str.replace(".", "", regex=False), errors="coerce")
    for idx in df.index[montos.isna() & df["Monto"].notna()]:
        registrar("Monto no numérico (quedó vacío)", idx, "Monto", original.at[idx, "Monto"])
    df["Monto"] = montos

    # 6. Filas sin nombre
    sin_nombre = df["Nombre"].isna()
    for idx in df.index[sin_nombre]:
        registrar("Fila sin nombre (fila eliminada)", idx, "(fila completa)", fila_texto(idx))
    df = df[~sin_nombre]

    # 7. Filas duplicadas
    duplicada = df.duplicated()
    for idx in df.index[duplicada]:
        registrar("Fila duplicada (fila eliminada)", idx, "(fila completa)", fila_texto(idx))
    df = df[~duplicada]

    # 8. Mostrar por pantalla TODO lo eliminado
    print("\n=== DETALLE DE LO ELIMINADO (respaldo temporal) ===")
    if eliminados:
        for e in eliminados:
            print(f"- Fila {e['Fila en Excel']} | {e['Columna']} | {e['Motivo']} | Valor: {e['Valor eliminado']!r}")
    else:
        print("No se eliminó nada.")

    conteo = Counter(e["Motivo"] for e in eliminados)
    resumen = [f"Filas iniciales: {len(original)}"]
    resumen += [f"{motivo}: {cantidad}" for motivo, cantidad in conteo.items()]
    resumen.append(f"Filas finales: {len(df)}")

    print("\n=== RESUMEN ===")
    print("\n".join(resumen))

    # 9. Guardar (el respaldo también queda en el Excel, porque la pantalla se borra al cerrar)
    try:
        with pd.ExcelWriter(ruta_salida) as writer:
            df.to_excel(writer, sheet_name="Datos limpios", index=False)
            pd.DataFrame({"Resumen de limpieza": resumen}).to_excel(writer, sheet_name="Resumen", index=False)
            pd.DataFrame(eliminados, columns=["Motivo", "Fila en Excel", "Columna", "Valor eliminado"]).to_excel(
                writer, sheet_name="Eliminados", index=False)
    except PermissionError:
        print(f"\nERROR: no pude guardar '{ruta_salida}'. Cierra ese archivo en Excel y vuelve a ejecutar.")
        return

    print(f"\nListo. Archivo creado: {ruta_salida}")


if __name__ == "__main__":
    entrada = sys.argv[1] if len(sys.argv) > 1 else "ventas_sucias.xlsx"
    salida = sys.argv[2] if len(sys.argv) > 2 else "ventas_limpias.xlsx"
    limpiar(entrada, salida)