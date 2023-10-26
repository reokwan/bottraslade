import pandas as pd
from openpyxl import load_workbook

# Leer el archivo "cuentas_activas.xlsx" y extraer la columna "Numero"
df_cuentas_activas = pd.read_excel("cuentas_activas.xlsx")
columna_Numero = df_cuentas_activas["Numero"]

# Leer el archivo "Plantilla Informe Mensual.xlsx"
df_plantilla_informe = pd.read_excel("Plantilla Informe Mensual.xlsx")

# Agregar una nueva columna con los datos extraídos
df_plantilla_informe["Nueva Columna"] = columna_Numero

# Guardar el archivo "Plantilla Informe Mensual.xlsx" con los datos agregados
with pd.ExcelWriter("Plantilla Informe Mensual.xlsx", engine='openpyxl', mode='a') as writer:
    df_plantilla_informe.to_excel(writer, sheet_name='Sheet1', index=False)

# Cargar el archivo Excel nuevamente para agregar un estilo predeterminado
book = load_workbook("Plantilla Informe Mensual.xlsx")
writer = pd.ExcelWriter("Plantilla Informe Mensual.xlsx", engine='openpyxl') 
writer.book = book

# Agregar un estilo predeterminado
default_style = NamedStyle(name="default")
default_style.alignment = openpyxl.styles.Alignment(horizontal="left")
default_style.font = openpyxl.styles.Font(name="Calibri", size=11)
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
for sheetname in writer.sheets:
    writer.sheets[sheetname].default_style = default_style

# Guardar el archivo nuevamente con el estilo predeterminado
df_plantilla_informe.to_excel(writer, "Sheet1", index=False)
writer.save()
