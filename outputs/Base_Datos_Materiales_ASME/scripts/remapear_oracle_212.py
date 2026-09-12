"""Un solo uso (Fase 3): traslada el *oracle* del Art. 212 al layout nuevo.

La Fase 3 sube la Seccion de Material justo detras de la Seccion 1, y con ella
se mueven las direcciones de casi toda la hoja. El *oracle*
(`parche_art212_ref.json`) es la hoja Rev0 validada en Excel, volcada una vez
con `_dump_parche_ref.py`: es el unico control INDEPENDIENTE del builder que
tiene el Art. 212.

**Por eso no se vuelve a volcar desde la hoja reconstruida.** Hacerlo lo
convertiria en un espejo del builder —pasaria a decir "el builder produce lo
que el builder produjo" y no detectaria un error del builder—. Aqui se le
aplica EL MISMO mapa de filas que el builder aplica a la hoja
(`build_db_materiales.MAPA_FILAS_212`), a dos sitios:

  * la CLAVE de cada celda            D64  -> D92
  * las FILAS de las referencias      =D63/($D$42*$D$41) -> =D91/($D$70*$D$69)
    dentro de cada formula, reusando `remapear_referencias` del builder, que es
    la misma funcion que mueve la hoja (no una copia).

Ninguna formula cambia de ESTRUCTURA: solo se mueven numeros de fila. El diff
del JSON tiene que poder leerse asi, y es la unica garantia que se ofrece.

    python remapear_oracle_212.py [--dry-run]
"""
import argparse
import json
from pathlib import Path

import build_db_materiales as B

ORACLE = Path(__file__).resolve().parent / "parche_art212_ref.json"


def remapear(datos, mapa):
    def celda(ref):
        return f"{ref[0]}{mapa.get(int(ref[1:]), int(ref[1:]))}"

    return {
        "formulas": {celda(k): B.remapear_referencias(v, mapa)
                     for k, v in datos["formulas"].items()},
        "validaciones": sorted(
            ({"sqref": B._remapear_rango(v["sqref"], mapa),
              "formula1": v["formula1"]} for v in datos["validaciones"]),
            key=lambda d: d["sqref"]),
        "fusionados": sorted(B._remapear_rango(r, mapa)
                             for r in datos["fusionados"]),
        # Los anchos son por COLUMNA: no los toca un mapa de filas.
        "anchos": datos["anchos"],
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    viejo = json.loads(ORACLE.read_text(encoding="utf-8"))
    nuevo = remapear(viejo, B.MAPA_FILAS_212)

    # Un mapa inyectivo no puede perder ni fabricar celdas. Si el conteo no
    # cuadra, dos claves distintas cayeron en la misma y el oracle quedaria
    # silenciosamente incompleto.
    for clave in ("formulas", "validaciones", "fusionados"):
        assert len(nuevo[clave]) == len(viejo[clave]), (
            f"{clave}: {len(viejo[clave])} -> {len(nuevo[clave])}; el mapa "
            f"no es inyectivo sobre las celdas del oracle")

    print(f"{len(nuevo['formulas'])} celdas, "
          f"{len(nuevo['validaciones'])} validaciones, "
          f"{len(nuevo['fusionados'])} fusionados")
    if args.dry_run:
        print("(dry-run: no se escribe nada)")
        return
    ORACLE.write_text(json.dumps(nuevo, ensure_ascii=False, indent=1,
                                 sort_keys=True), encoding="utf-8")
    print(f"-> {ORACLE}")


if __name__ == "__main__":
    main()
