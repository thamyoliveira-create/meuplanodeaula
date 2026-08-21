# -*- coding: utf-8 -*-
"""
Flask Web Application for Lesson Plan Generator
Vercel Serverless entrypoint: api/index.py
"""

import io
import os
import pathlib
import sys
from flask import Flask, render_template, request, send_file, jsonify

# Ensure local imports work in both Vercel and local environment
BASE_DIR = pathlib.Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from lesson_plan_generator.plan_generator import (
    load_spreadsheet,
    get_available_weeks,
    filter_by_week,
    render_docx,
    get_plan_preview,
    DEFAULT_TEMPLATE_PATH
)

app = Flask(
    __name__,
    template_folder=str(BASE_DIR / "templates"),
    static_folder=str(BASE_DIR / "static")
)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB max upload


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/api/weeks", methods=["POST"])
def detect_weeks():
    """Detect available weeks from uploaded spreadsheet."""
    if "spreadsheet" not in request.files:
        return jsonify({"success": False, "error": "Nenhum arquivo enviado."}), 400

    file = request.files["spreadsheet"]
    if not file.filename:
        return jsonify({"success": False, "error": "Arquivo vazio."}), 400

    try:
        content = io.BytesIO(file.read())
        df = load_spreadsheet(content, filename=file.filename)
        weeks = get_available_weeks(df)
        return jsonify({"success": True, "weeks": weeks, "columns": list(df.columns)})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 400


@app.route("/api/preview", methods=["POST"])
def preview_plan():
    """Return JSON preview of lesson plan for selected week."""
    if "spreadsheet" not in request.files:
        return jsonify({"success": False, "error": "Planilha não fornecida."}), 400

    week = request.form.get("week", "").strip()
    if not week:
        return jsonify({"success": False, "error": "Semana não informada."}), 400

    file = request.files["spreadsheet"]
    try:
        content = io.BytesIO(file.read())
        df = load_spreadsheet(content, filename=file.filename)
        filtered = filter_by_week(df, week)
        preview_data = get_plan_preview(filtered)
        return jsonify({"success": True, "data": preview_data, "week": week})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 400


@app.route("/generate", methods=["POST"])
def generate():
    """Generate and download filled DOCX."""
    if "spreadsheet" not in request.files:
        return "Erro: Planilha não enviada.", 400

    week = request.form.get("week", "").strip()
    if not week:
        return "Erro: Semana não informada.", 400

    sheet_file = request.files["spreadsheet"]
    if not sheet_file.filename:
        return "Erro: Arquivo de planilha inválido.", 400

    try:
        sheet_bytes = io.BytesIO(sheet_file.read())
        df = load_spreadsheet(sheet_bytes, filename=sheet_file.filename)
        filtered = filter_by_week(df, week)

        # Check for optional custom template
        template_bytes = None
        if "template" in request.files:
            tpl_file = request.files["template"]
            if tpl_file and tpl_file.filename:
                template_bytes = io.BytesIO(tpl_file.read())

        docx_buffer = render_docx(filtered, template_source=template_bytes or DEFAULT_TEMPLATE_PATH)

        download_name = f"plano_de_aula_semana_{week}.docx"
        return send_file(
            docx_buffer,
            as_attachment=True,
            download_name=download_name,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        return f"Erro ao gerar plano de aula: {str(e)}", 400


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
