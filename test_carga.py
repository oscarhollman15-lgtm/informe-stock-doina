import pandas as pd

ARCHIVO = "stock doina.xlsx"
HOJA_STOCK_PROCESO = "STOCK EN PROCESO"

# Test 1: leer hojas
xl = pd.ExcelFile(ARCHIVO)
print("Hojas disponibles:", xl.sheet_names)

# Test 2: leer RESUMEN STOCK
try:
    raw = pd.read_excel(ARCHIVO, sheet_name="RESUMEN STOCK", header=None)
    print("\nRESUMEN STOCK fila 0:", raw.iloc[0].tolist())
    print("RESUMEN STOCK fila 4:", raw.iloc[4].tolist())
    df = raw.iloc[4:].copy()
    df.columns = ["Producto", "Exp_Unidades", "Exp_KG", "Prod_Unidades", "Prod_KG", "Proc_Unidades", "Proc_KG"]
    df = df.dropna(subset=["Producto"])
    print("Resumen stock OK:", len(df), "filas")
except Exception as e:
    print("ERROR RESUMEN STOCK:", e)

# Test 3: leer STOCK EN PROCESO
try:
    df_raw = pd.read_excel(ARCHIVO, sheet_name=HOJA_STOCK_PROCESO, skiprows=3)
    print("\nSTOCK EN PROCESO columnas:", df_raw.columns.tolist())
    print("Total filas raw:", len(df_raw))
    df_raw.columns = df_raw.columns.astype(str).str.strip()
    df_clean = df_raw[df_raw["ETAPA"].notna() & ~df_raw["ETAPA"].astype(str).str.startswith("TOTAL")]
    print("Filas limpias:", len(df_clean))
    print("Primeras filas:\n", df_clean.head(3).to_string())
except Exception as e:
    import traceback
    print("ERROR STOCK EN PROCESO:", e)
    traceback.print_exc()
