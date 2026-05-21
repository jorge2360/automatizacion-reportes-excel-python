import pandas as pd
from pathlib import Path

RUTA_ENTRADA = Path("entrada")
RUTA_SALIDA = Path("salida")

def generar_reporte():
    RUTA_SALIDA.mkdir(exist_ok=True)

    ventas = pd.read_csv(RUTA_ENTRADA / "ventas.csv")
    inventario = pd.read_csv(RUTA_ENTRADA / "inventario.csv")

    ventas["total"] = ventas["cantidad"] * ventas["precio_unitario"]
    inventario["valor_inventario"] = inventario["stock"] * inventario["precio_compra"]

    ventas_por_categoria = ventas.groupby("categoria")["total"].sum().reset_index()
    ventas_por_producto = ventas.groupby("producto")["total"].sum().reset_index()
    inventario_valorizado = inventario[["producto", "categoria", "stock", "valor_inventario"]]

    ruta_reporte = RUTA_SALIDA / "reporte_empresarial.xlsx"

    with pd.ExcelWriter(ruta_reporte, engine="openpyxl") as writer:
        ventas.to_excel(writer, sheet_name="Ventas", index=False)
        inventario.to_excel(writer, sheet_name="Inventario", index=False)
        ventas_por_categoria.to_excel(writer, sheet_name="Ventas por categoria", index=False)
        ventas_por_producto.to_excel(writer, sheet_name="Ventas por producto", index=False)
        inventario_valorizado.to_excel(writer, sheet_name="Inventario valorizado", index=False)

    print(f"Reporte generado correctamente: {ruta_reporte}")

if __name__ == "__main__":
    generar_reporte()