import json
import os
from pathlib import Path

import pandas as pd


CARPETA = Path(__file__).resolve().parent


def calcular_precio_final(precio_neto: float, iva: float) -> float:
    """Devuelve el precio neto aumentado por la tasa de IVA indicada."""
    return precio_neto * (1 + iva)


def main() -> None:
    """Carga el inventario, calcula totales y exporta un archivo Excel."""
    try:
        inventario = pd.read_csv(CARPETA / "inventario.csv")
        with (CARPETA / "config.json").open(encoding="utf-8") as archivo_config:
            configuracion = json.load(archivo_config)
    except FileNotFoundError as error:
        print(f"No se encontró uno de los archivos requeridos: {error.filename}")
        return

    iva = configuracion["iva"]
    totales_por_categoria: dict[str, float] = {}

    for _, producto in inventario.iterrows():
        categoria = producto["categoria"]
        valor_producto = float(producto["precio_neto"] * producto["stock"])
        totales_por_categoria[categoria] = (
            totales_por_categoria.get(categoria, 0.0) + valor_producto
        )

    inventario["precio_final_iva"] = inventario["precio_neto"].apply(
        calcular_precio_final, iva=iva
    )
    ruta_salida = CARPETA / "inventario_final.xlsx"
    inventario.to_excel(ruta_salida, index=False)

    print("Valor total del inventario por categoría:")
    for categoria, total in totales_por_categoria.items():
        print(f"- {categoria}: {total:,.0f} {configuracion['moneda']}")
    print(f"Archivo generado: {ruta_salida}")

    try:
        os.startfile(str(ruta_salida))
    except OSError as error:
        print(f"No se pudo abrir el Excel automáticamente: {error}")


if __name__ == "__main__":
    main()