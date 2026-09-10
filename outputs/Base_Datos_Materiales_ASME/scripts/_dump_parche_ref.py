"""Un solo uso: vuelca Parche_PCC2_Art212 del .xlsm real a parche_art212_ref.json.
El JSON es el *oracle* del desanclado total — la fuente de las formulas de las
secciones 1-6, que la Regla n.1 prohibe reescribir de memoria. Regenerarlo solo
si se repone la hoja desde un Rev validado en Excel."""
import json, os
from pathlib import Path
import openpyxl

AQUI = Path(__file__).resolve().parent
# scripts -> Base_Datos_Materiales_ASME -> outputs, donde vive el entregable
# (mismo criterio que POR_DEFECTO en test_dashboard.py).
RUTA = Path(os.environ.get(
    "MOTOR_XLSM",
    AQUI.parents[1] / "Motor_de_Calculo_ASME_PCC_Rev4.xlsm"))
SALIDA = Path(__file__).resolve().parent / "parche_art212_ref.json"
HOJA = "Parche_PCC2_Art212"


def main():
    wb = openpyxl.load_workbook(RUTA, keep_vba=True)
    ws = wb[HOJA]
    formulas = {}
    for fila in ws.iter_rows(min_row=1, max_row=129, min_col=1, max_col=7):
        for c in fila:
            if c.value is not None:
                formulas[c.coordinate] = c.value
    validaciones = sorted(
        ({"sqref": str(dv.sqref), "formula1": dv.formula1} for dv in ws.data_validations.dataValidation),
        key=lambda d: d["sqref"])
    fusionados = sorted(str(r) for r in ws.merged_cells.ranges)
    anchos = {col: dim.width for col, dim in ws.column_dimensions.items() if dim.width}
    SALIDA.write_text(json.dumps(
        {"formulas": formulas, "validaciones": validaciones,
         "fusionados": fusionados, "anchos": anchos},
        ensure_ascii=False, indent=1, sort_keys=True), encoding="utf-8")
    print(f"{len(formulas)} celdas, {len(validaciones)} validaciones -> {SALIDA}")


if __name__ == "__main__":
    main()
