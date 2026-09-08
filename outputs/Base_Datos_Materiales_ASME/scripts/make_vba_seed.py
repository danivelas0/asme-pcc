"""Siembra el maestro con macros del Motor de Calculo ASME PCC.

Paso unico y previo al builder. Toma el respaldo Rev. 0 (3 hojas, sin VBA),
le inyecta los dos componentes VBA versionados en `vba/` y lo guarda como
`templates/maestro_con_macros.xlsm`.

Desde ahi, `build_db_materiales.py` lo abre con `keep_vba=True` y openpyxl
arrastra `vbaProject.bin` intacto hasta el entregable. Se hace asi -y no con
cirugia sobre el ZIP- porque openpyxl emite por su cuenta los content-types y
la relacion del binario cuando el proyecto VBA viene ya en el archivo de
entrada; construirlos a mano es fragil y silenciosamente corruptible.

Requiere Excel instalado y, momentaneamente, el ajuste del Centro de confianza
"Confiar en el acceso al modelo de objetos de proyectos de VBA". Este script lo
activa, siembra, y lo devuelve a su valor original en un `finally` -incluso si
la siembra falla-. La ventana de exposicion dura segundos y solo afecta a
HKEY_CURRENT_USER.

La restauracion se verifica y se reintenta: Excel vuelca los ajustes del Centro
de confianza al registro mientras se cierra, y ese volcado asincrono puede
reescribir el 1 despues de haberlo borrado. Si aun asi no se puede restaurar, el
script lo dice por stderr; **compruebelo** antes de dar por terminado el paso.
Si el proceso se mata a la fuerza, la restauracion no llega a correr y el ajuste
queda encendido.

    python make_vba_seed.py
"""
from __future__ import annotations

import argparse
import re
import sys
import time
import winreg
from pathlib import Path

# xlOpenXMLWorkbookMacroEnabled. Es el unico formato que conserva vbaProject.bin.
XL_XLSM = 52

# Componentes VBA a sembrar: (archivo fuente, nombre del componente en el proyecto).
# ThisWorkbook es un modulo de documento: no se puede importar ni sustituir, solo
# rellenar. modNav si se crea desde cero.
SRC_THIS_WORKBOOK = "this_workbook.vba"
SRC_MOD_NAV = "mod_nav.vba"
NOMBRE_MOD_NAV = "modNav"

VBEXT_CT_STD_MODULE = 1  # vbext_ct_StdModule

CLAVE_SEGURIDAD = r"Software\Microsoft\Office\16.0\Excel\Security"
VALOR_VBOM = "AccessVBOM"


# ---------------------------------------------------------------------------
# Ajuste temporal del Centro de confianza
# ---------------------------------------------------------------------------
def _leer_vbom() -> int | None:
    """Valor actual de AccessVBOM, o None si no existe la clave o el valor."""
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, CLAVE_SEGURIDAD) as k:
            return int(winreg.QueryValueEx(k, VALOR_VBOM)[0])
    except FileNotFoundError:
        return None


def _escribir_vbom(valor: int | None) -> None:
    """Fija AccessVBOM, o lo borra si `valor` es None (estado 'ausente')."""
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, CLAVE_SEGURIDAD) as k:
        if valor is None:
            try:
                winreg.DeleteValue(k, VALOR_VBOM)
            except FileNotFoundError:
                pass
        else:
            winreg.SetValueEx(k, VALOR_VBOM, 0, winreg.REG_DWORD, valor)


# ---------------------------------------------------------------------------
# Siembra
# ---------------------------------------------------------------------------
def _restaurar_vbom(previo: int | None, intentos: int = 6) -> int | None:
    """Devuelve AccessVBOM a `previo`, insistiendo hasta que quede escrito.

    Excel mantiene los ajustes del Centro de confianza en memoria y los vuelca
    al registro durante su cierre, que es asincrono: si se restaura el valor
    justo despues de Quit, el proceso moribundo puede volver a escribir el 1
    que leyo al arrancar. Dejar encendido "Confiar en el acceso al modelo de
    objetos de VBA" es dejar abierta una via para que cualquier macro reescriba
    macros, asi que se comprueba y se reintenta hasta confirmarlo.
    """
    import time
    for i in range(intentos):
        _escribir_vbom(previo)
        time.sleep(0.5)
        if _leer_vbom() == previo:
            # Segunda lectura tras una pausa: el vuelco de Excel puede llegar
            # despues de la primera confirmacion.
            time.sleep(1.0)
            if _leer_vbom() == previo:
                return previo
        print(f"  reintentando restaurar AccessVBOM ({i + 1}/{intentos})...",
              file=sys.stderr)
    return _leer_vbom()


# ---------------------------------------------------------------------------
# Lint previo
# ---------------------------------------------------------------------------
# Excel compila el VBA al guardar. Si no compila, abre un cuadro de dialogo
# modal que la automatizacion no puede cerrar: SaveAs se queda colgado hasta
# que la llamada RPC caduca y deja un Excel huerfano bloqueando el archivo.
# Vale la pena atrapar en Python los errores estructurales mas probables antes
# de tocar COM.
_RE_PROC = re.compile(
    r"^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?"
    r"(Sub|Function|Property\s+(?:Get|Let|Set))\b", re.IGNORECASE)
_RE_FIN_PROC = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.IGNORECASE)
_RE_DECL = re.compile(
    r"^\s*(?:Public\s+|Private\s+|Global\s+)?(Const|Dim|Type|Declare|Enum)\b",
    re.IGNORECASE)

# VBA no admite mas de 25 continuaciones de linea ('_') en una misma linea
# logica. Pasarse no da un error de compilacion legible: AddFromString falla en
# el acto con "Demasiadas continuaciones de linea", el modulo se queda VACIO, y
# lo que el usuario ve despues es un "No se ha definido Sub o Function" al
# guardar -porque el otro modulo llama a rutinas que ya no existen- dentro de
# un dialogo modal que cuelga la automatizacion y deja un Excel huerfano.
# Ya se pago una vez, con un Array() de 31 hojas navegables.
MAX_CONTINUACIONES = 25


def lint_vba(texto: str, nombre: str) -> list[str]:
    """Comprueba las reglas estructurales que Excel solo delata al compilar."""
    fallos: list[str] = []
    dentro = False          # dentro de una rutina
    primera_proc: int | None = None
    continuacion = False
    n_cont, ini_logica, ya_avisado = 0, 1, False

    for i, cruda in enumerate(texto.splitlines(), 1):
        linea = cruda.split("'")[0] if not cruda.lstrip().startswith("'") else ""
        if continuacion:
            n_cont += 1
            if n_cont > MAX_CONTINUACIONES and not ya_avisado:
                fallos.append(
                    f"{nombre}:{ini_logica}: la linea logica encadena {n_cont} o mas "
                    f"continuaciones ('_'); VBA admite {MAX_CONTINUACIONES}. "
                    f"AddFromString la rechaza y deja el modulo vacio.")
                ya_avisado = True
            continuacion = linea.rstrip().endswith("_")
            continue
        continuacion = linea.rstrip().endswith("_")
        n_cont, ini_logica, ya_avisado = 0, i, False

        if _RE_FIN_PROC.match(linea):
            dentro = False
            continue
        if _RE_PROC.match(linea):
            dentro = True
            if primera_proc is None:
                primera_proc = i
            continue
        # Una declaracion de modulo (fuera de rutina) despues de la primera
        # rutina no falla en su linea: VBA reporta "Variable not defined" en
        # cada uso, que es mucho mas dificil de rastrear.
        if not dentro and primera_proc is not None and _RE_DECL.match(linea):
            fallos.append(
                f"{nombre}:{i}: declaracion de modulo '{linea.strip()[:50]}' despues de la "
                f"primera rutina (linea {primera_proc}). VBA exige que Const/Dim/Type/"
                f"Declare/Enum precedan a toda rutina.")
    if dentro:
        fallos.append(f"{nombre}: falta un 'End Sub'/'End Function' al final del modulo.")
    return fallos


def _pid_de(excel) -> int | None:
    """PID de la instancia de Excel, via la ventana principal."""
    try:
        import win32process
        return win32process.GetWindowThreadProcessId(excel.Hwnd)[1]
    except Exception:  # noqa: BLE001
        return None


def _cerrar(excel, wb, pid: int | None) -> None:
    """Cierra Excel sin dejar huerfanos y sin enmascarar el error original.

    Cada paso va en su propio try: si la siembra fallo a medias, es normal que
    Close o Quit tambien fallen, y esa segunda excepcion no debe sustituir a la
    primera -que es la que dice que paso-. Si Quit no basta (COM puede dejar el
    proceso vivo con el editor VBA abierto, bloqueando el archivo y la
    siguiente ejecucion), se termina el proceso por PID.
    """
    for accion in (lambda: wb.Close(SaveChanges=False) if wb is not None else None,
                   excel.Quit):
        try:
            accion()
        except Exception as e:  # noqa: BLE001
            print(f"  aviso al cerrar Excel: {e}", file=sys.stderr)
    if pid:
        try:
            import win32api
            import win32con
            h = win32api.OpenProcess(win32con.PROCESS_TERMINATE | win32con.SYNCHRONIZE,
                                     False, pid)
            import win32event
            if win32event.WaitForSingleObject(h, 5000) != win32event.WAIT_OBJECT_0:
                print(f"  Excel (PID {pid}) sigue vivo tras Quit; se termina.",
                      file=sys.stderr)
                win32api.TerminateProcess(h, 0)
            win32api.CloseHandle(h)
        except Exception:  # noqa: BLE001
            pass


def sembrar(entrada: Path, salida: Path, dir_vba: Path) -> dict:
    import win32com.client

    cod_wb = (dir_vba / SRC_THIS_WORKBOOK).read_text(encoding="utf-8")
    cod_nav = (dir_vba / SRC_MOD_NAV).read_text(encoding="utf-8")

    fallos = (lint_vba(cod_wb, SRC_THIS_WORKBOOK) + lint_vba(cod_nav, SRC_MOD_NAV))
    if fallos:
        raise ValueError("el VBA no pasa el lint estructural:\n  " + "\n  ".join(fallos))

    salida.parent.mkdir(parents=True, exist_ok=True)
    if salida.exists():
        salida.unlink()  # SaveAs sobre un archivo existente puede pedir confirmacion

    excel = win32com.client.DispatchEx("Excel.Application")  # instancia propia
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    pid = _pid_de(excel)
    wb = None
    try:
        wb = excel.Workbooks.Open(str(entrada.resolve()))
        proyecto = wb.VBProject  # falla con -2147352567 si AccessVBOM sigue en 0

        # ThisWorkbook es un modulo de documento: no se puede importar ni
        # sustituir, solo vaciar y rellenar. AddFromString sobre un modulo con
        # codigo previo lo duplicaria, de ahi el DeleteLines.
        cm = proyecto.VBComponents("ThisWorkbook").CodeModule
        if cm.CountOfLines:
            cm.DeleteLines(1, cm.CountOfLines)
        cm.AddFromString(cod_wb)

        # modNav: se rellena si ya existe en vez de borrarlo y recrearlo.
        # VBComponents.Remove es asincrono y deja el editor VBA abierto, lo que
        # cuelga a Excel en modo automatizacion.
        existentes = {c.Name: c for c in proyecto.VBComponents}
        if NOMBRE_MOD_NAV in existentes:
            mod = existentes[NOMBRE_MOD_NAV]
            if mod.CodeModule.CountOfLines:
                mod.CodeModule.DeleteLines(1, mod.CodeModule.CountOfLines)
        else:
            mod = proyecto.VBComponents.Add(VBEXT_CT_STD_MODULE)
            mod.Name = NOMBRE_MOD_NAV
        mod.CodeModule.AddFromString(cod_nav)

        wb.SaveAs(str(salida.resolve()), FileFormat=XL_XLSM)

        info = {
            "hojas": [ws.Name for ws in wb.Worksheets],
            "lineas_this_workbook": proyecto.VBComponents("ThisWorkbook").CodeModule.CountOfLines,
            "lineas_mod_nav": proyecto.VBComponents(NOMBRE_MOD_NAV).CodeModule.CountOfLines,
        }
    finally:
        _cerrar(excel, wb, pid)
    return info


def main(argv=None) -> int:
    aqui = Path(__file__).resolve().parent
    raiz = aqui.parents[2]  # scripts -> Base_Datos_Materiales_ASME -> outputs -> repo

    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--in", dest="inp", type=Path,
                    default=aqui.parent / "Motor_de_Calculo_ASME_PCC_Rev0_respaldo.xlsx")
    ap.add_argument("--out", dest="out", type=Path,
                    default=raiz / "templates" / "maestro_con_macros.xlsm")
    ap.add_argument("--vba", dest="vba", type=Path, default=aqui / "vba")
    a = ap.parse_args(argv)

    if not a.inp.exists():
        print(f"ERROR: no existe la entrada {a.inp}", file=sys.stderr)
        return 2
    for src in (SRC_THIS_WORKBOOK, SRC_MOD_NAV):
        if not (a.vba / src).exists():
            print(f"ERROR: falta la fuente VBA {a.vba / src}", file=sys.stderr)
            return 2

    previo = _leer_vbom()
    print(f"AccessVBOM previo: {'(ausente)' if previo is None else previo}", flush=True)
    if previo == 1:
        print("  AVISO: ya estaba activado. O el Centro de confianza lo tiene puesto a "
              "proposito, o una ejecucion anterior murio sin restaurarlo. Al terminar se "
              "dejara como esta ahora.", file=sys.stderr)

    try:
        _escribir_vbom(1)
        info = sembrar(a.inp, a.out, a.vba)
    except Exception:  # noqa: BLE001 - se reporta y se restaura el registro
        import traceback
        print("ERROR durante la siembra:", file=sys.stderr)
        traceback.print_exc()
        return 1
    finally:
        actual = _restaurar_vbom(previo)
        print(f"AccessVBOM restaurado a: {'(ausente)' if actual is None else actual}")
        if actual != previo:
            print("AVISO: NO se pudo restaurar AccessVBOM a su valor original "
                  f"({'ausente' if previo is None else previo}). Revise el Centro de "
                  "confianza -> Configuracion de macros -> 'Confiar en el acceso al "
                  "modelo de objetos de proyectos de VBA' y desactivelo a mano.",
                  file=sys.stderr)

    esperadas = ["Instrucciones", "Datos_Ref", "Parche_PCC2_Art212"]
    ok = (sorted(info["hojas"]) == sorted(esperadas)
          and info["lineas_this_workbook"] > 0
          and info["lineas_mod_nav"] > 0)

    print(f"Hojas del maestro ({len(info['hojas'])}): {', '.join(info['hojas'])}")
    print(f"ThisWorkbook: {info['lineas_this_workbook']} lineas de codigo")
    print(f"{NOMBRE_MOD_NAV}: {info['lineas_mod_nav']} lineas de codigo")
    print(f"Escrito: {a.out}")
    if not ok:
        print("ERROR: la verificacion posterior a la siembra no cuadra.", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
