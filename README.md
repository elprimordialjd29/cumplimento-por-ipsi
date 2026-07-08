# Cumplimiento por IPSI — DUSAKAWI

Herramienta de validación y consolidación de Actas de Evaluación de Servicios
(PYM / PMT / DNT) de DUSAKAWI IPSI. Extrae los datos de cada acta (Excel o
PDF), valida su consistencia interna y consolida el % de cumplimiento por
actividad y periodo en un Excel maestro.

## Qué valida por cada acta
- Que el Nº de Acta empiece con el Nº de Contrato correspondiente.
- Que la Empresa sea la esperada (DUSAKAWI IPSI).
- Que por cada actividad: `Vr Exigido − Descuento = Vr Reconocido`.
- Que la suma de todas las actividades cuadre con la fila `TOTAL EJECUCION`.

## Estructura
- `core.py` — extracción, validación y consolidación (lógica compartida).
- `consolidar_acta.py` — CLI.
- `webapp/` — app web (Flask) para cargar actas desde el navegador.
- `entradas/` — respaldo en JSON de actas transcritas manualmente (PDFs sin texto).
- `CONSOLIDADO_ACTAS_PYM_DUSAKAWI.xlsx` — maestro consolidado (Detalle,
  Validaciones, Resumen_Cumplimiento).

## Uso

### CLI
```bash
# Acta en Excel (extracción automática)
python3 consolidar_acta.py --xlsx "ruta/al/acta.xlsx" --periodo "OTRO SI"

# Acta en PDF (JSON ya transcrito, ver entradas/)
python3 consolidar_acta.py --json "ruta/al/extraido.json" --periodo "OTRO SI"
```

### App web
```bash
pip install flask openpyxl
python3 webapp/app.py
```
Abrir `http://localhost:5001`.

Periodos soportados: `OTRO SI`, `1-TRIM`, `2-TRIM`, `3-TRIM`, `4-TRIM`.
