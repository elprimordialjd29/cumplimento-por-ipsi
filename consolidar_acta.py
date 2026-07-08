#!/usr/bin/env python3
"""
CLI de validacion y consolidacion de Actas PYM/PMT/DNT - DUSAKAWI.
La logica vive en core.py (compartida con la app web en webapp/app.py).

Uso:
  1) Acta en Excel (con hoja tipo "erev", encabezado ACTA No / EMPRESA / etc.):
       python3 consolidar_acta.py --xlsx "ruta/al/acta.xlsx" --periodo "OTRO SI"

  2) Acta en PDF (o cualquier archivo sin estructura de celdas): se entrega un
     JSON con los datos ya transcritos (ver entradas/ejemplo.json) y se corre:
       python3 consolidar_acta.py --json "ruta/al/extraido.json" --periodo "OTRO SI"
"""
import argparse
import json
import sys

import core


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--xlsx", help="Ruta al acta en Excel (extraccion automatica)")
    ap.add_argument("--json", help="Ruta a un JSON con los datos ya transcritos (para actas en PDF)")
    ap.add_argument("--periodo", required=True, choices=list(core.PERIODO_MESES.keys()),
                     help="Periodo contractual del acta")
    args = ap.parse_args()

    if not args.xlsx and not args.json:
        ap.error("Debes indicar --xlsx o --json")

    if args.xlsx:
        acta = core.parse_acta_xlsx(args.xlsx)
    else:
        with open(args.json, encoding="utf-8") as f:
            acta = json.load(f)

    checks = core.validar(acta)
    master = core.consolidar(acta, args.periodo, checks)

    n_err = sum(1 for c in checks if c["resultado"] == "ERROR")
    print(f"Acta: {acta.get('municipio')} | {acta.get('contrato')} | {acta.get('acta_no')} | periodo={args.periodo}")
    print(f"Chequeos: {len(checks)} total, {n_err} con ERROR")
    for c in checks:
        marca = "OK " if c["resultado"] == "OK" else "ERR"
        print(f"  [{marca}] {c['chequeo']}: {c['detalle']}")
    print(f"\nMaestro actualizado: {master}")
    if n_err:
        sys.exit(1)


if __name__ == "__main__":
    main()
