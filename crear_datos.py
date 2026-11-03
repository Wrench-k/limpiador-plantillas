import pandas as pd

datos = {
    "Nombre": ["  Ana Pérez", "luis GONZÁLEZ", "Ana Pérez", "María  Soto", None, "pedro rojas"],
    "Email": ["ANA@MAIL.COM", "luis@mail.com ", "ana@mail.com", "maria@mail", "x@mail.com", "pedro+ventas@mail.com"],
    "Fecha": ["2026-10-01", "01/10/2026", "2026-10-01", "15-10-2026", "2026-10-03", "no registrada"],
    "Monto": ["15000", "20.500", "15000", "8000", "12000", None],
}

pd.DataFrame(datos).to_excel("ventas_sucias.xlsx", index=False)
print("Archivo ventas_sucias.xlsx creado")