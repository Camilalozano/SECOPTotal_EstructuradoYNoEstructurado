# ============================================
# LEFT JOIN ENTRE:
# SECOP_ATENEA_TOTAL  +  SECOP_NoEstructurado
# ============================================

import pandas as pd
import os

# =========================================================
# 1. SOLICITAR RUTAS DE LOS ARCHIVOS
# =========================================================

print("========================================")
print("MERGE SECOP_ATENEA_TOTAL + NoEstructurado")
print("========================================\n")

ruta_secop_total = input(
    "Ingrese la ruta del archivo SECOP_ATENEA_TOTAL (.xlsx): "
).strip().replace('"', '')

ruta_no_estructurado = input(
    "Ingrese la ruta del archivo SECOP_NoEstructurado (.xlsx): "
).strip().replace('"', '')

ruta_output = input(
    "Ingrese la carpeta donde desea guardar el archivo resultado: "
).strip().replace('"', '')

# =========================================================
# 2. VALIDAR EXISTENCIA DE ARCHIVOS
# =========================================================

if not os.path.exists(ruta_secop_total):
    raise FileNotFoundError(
        f"No se encontró el archivo:\n{ruta_secop_total}"
    )

if not os.path.exists(ruta_no_estructurado):
    raise FileNotFoundError(
        f"No se encontró el archivo:\n{ruta_no_estructurado}"
    )

if not os.path.exists(ruta_output):
    raise FileNotFoundError(
        f"No existe la carpeta de salida:\n{ruta_output}"
    )

# =========================================================
# 3. CARGAR ARCHIVOS EXCEL
# =========================================================

print("\n📥 Cargando archivos Excel...")

df_secop_total = pd.read_excel(ruta_secop_total)
df_no_estructurado = pd.read_excel(ruta_no_estructurado)

print("✅ Archivos cargados correctamente")

# =========================================================
# 4. VALIDAR COLUMNAS NECESARIAS
# =========================================================

columna_left = "urlproceso (contratos_electronicos)"
columna_right = "url"

if columna_left not in df_secop_total.columns:
    raise ValueError(
        f"La columna '{columna_left}' no existe en SECOP_ATENEA_TOTAL"
    )

if columna_right not in df_no_estructurado.columns:
    raise ValueError(
        f"La columna '{columna_right}' no existe en SECOP_NoEstructurado"
    )

# =========================================================
# 5. LIMPIEZA BÁSICA DE VARIABLES DE UNIÓN
# =========================================================

print("\n🧹 Limpiando variables de unión...")

df_secop_total[columna_left] = (
    df_secop_total[columna_left]
    .astype(str)
    .str.strip()
)

df_no_estructurado[columna_right] = (
    df_no_estructurado[columna_right]
    .astype(str)
    .str.strip()
)

# =========================================================
# 6. REALIZAR LEFT JOIN
# =========================================================

print("\n🔗 Realizando LEFT JOIN...")

df_merge = pd.merge(
    df_secop_total,
    df_no_estructurado,
    how="left",
    left_on=columna_left,
    right_on=columna_right,
    suffixes=("", "_NoEstructurado")
)

print("✅ Merge realizado correctamente")

# =========================================================
# 7. EXPORTAR RESULTADO
# =========================================================

nombre_output = "SECOP_ATENEA_TOTAL_Merge_NoEstructurado.xlsx"

ruta_archivo_salida = os.path.join(
    ruta_output,
    nombre_output
)

print("\n💾 Exportando archivo Excel...")

df_merge.to_excel(
    ruta_archivo_salida,
    index=False
)

# =========================================================
# 8. MENSAJE FINAL
# =========================================================

print("\n========================================")
print("✅ PROCESO FINALIZADO")
print("========================================")
print(f"📁 Archivo guardado en:\n{ruta_archivo_salida}")
print(f"\n📊 Filas resultado: {len(df_merge):,}")
print("========================================")
