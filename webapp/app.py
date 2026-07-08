#!/usr/bin/env python3
"""App web (Flask) para cargar, validar y consolidar Actas PYM/PMT/DNT - DUSAKAWI."""
import sys
import tempfile
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import core  # noqa: E402

app = Flask(__name__)
UPLOAD_DIR = Path(__file__).resolve().parent / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)


@app.route("/")
def index():
    return render_template(
        "index.html",
        periodos=[{"value": k, "label": f"{k} ({v})"} for k, v in core.PERIODO_MESES.items()],
    )


@app.route("/api/upload-xlsx", methods=["POST"])
def upload_xlsx():
    file = request.files.get("file")
    periodo = request.form.get("periodo")
    if not file or not periodo:
        return jsonify({"error": "Falta el archivo o el periodo"}), 400
    if periodo not in core.PERIODO_MESES:
        return jsonify({"error": f"Periodo invalido: {periodo}"}), 400

    dest = UPLOAD_DIR / file.filename
    file.save(dest)
    try:
        acta = core.parse_acta_xlsx(dest)
    except ValueError as e:
        return jsonify({"error": str(e)}), 400

    checks = core.validar(acta)
    core.consolidar(acta, periodo, checks)
    return jsonify({"acta": acta, "checks": checks, "periodo": periodo})


@app.route("/api/upload-manual", methods=["POST"])
def upload_manual():
    payload = request.get_json(force=True)
    periodo = payload.get("periodo")
    acta = payload.get("acta")
    if not periodo or not acta:
        return jsonify({"error": "Falta 'periodo' o 'acta'"}), 400
    if periodo not in core.PERIODO_MESES:
        return jsonify({"error": f"Periodo invalido: {periodo}"}), 400

    checks = core.validar(acta)
    core.consolidar(acta, periodo, checks)
    return jsonify({"acta": acta, "checks": checks, "periodo": periodo})


@app.route("/api/resumen")
def api_resumen():
    rows = core.read_sheet_as_dicts("Resumen_Cumplimiento")
    periodos = core.periodos_presentes()
    return jsonify({"rows": rows, "periodos": periodos})


@app.route("/api/validaciones")
def api_validaciones():
    return jsonify(core.read_sheet_as_dicts("Validaciones"))


@app.route("/api/detalle")
def api_detalle():
    return jsonify(core.read_sheet_as_dicts("Detalle"))


@app.route("/api/download")
def api_download():
    if not core.MASTER_PATH.exists():
        return jsonify({"error": "Aun no hay datos consolidados"}), 404
    return send_file(core.MASTER_PATH, as_attachment=True, download_name=core.MASTER_PATH.name)


if __name__ == "__main__":
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 5001
    app.run(debug=True, port=port)
