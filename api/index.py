# -*- coding: utf-8 -*-
"""
Flask Web Application for Lesson Plan Generator
Entrypoint for Vercel Serverless Function: api/index.py
"""

import io
import os
import pathlib
import sys
from typing import Union, List, Dict, Any

from flask import Flask, request, send_file, jsonify, render_template_string
import pandas as pd
from docx import Document

BASE_DIR = pathlib.Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB limit


# --- Core Helper Functions ---

def get_default_template_path() -> pathlib.Path:
    candidates = [
        BASE_DIR / "template" / "template.docx",
        BASE_DIR / "lesson_plan_generator" / "template" / "template.docx",
        pathlib.Path(__file__).parent.parent / "template" / "template.docx",
        pathlib.Path(__file__).parent / "template.docx"
    ]
    for p in candidates:
        if p.exists():
            return p
    return candidates[0]


def load_spreadsheet(file_or_path: Union[str, pathlib.Path, io.BytesIO], filename: str = "") -> pd.DataFrame:
    if isinstance(file_or_path, (str, pathlib.Path)):
        path = pathlib.Path(file_or_path)
        if not path.exists():
            raise FileNotFoundError(f"Planilha não encontrada: {path}")
        if path.suffix.lower() in {".xlsx", ".xls"}:
            df = pd.read_excel(path)
        else:
            df = pd.read_csv(path, delimiter=None, engine="python")
    else:
        # In-memory BytesIO
        if filename.lower().endswith((".xlsx", ".xls")):
            df = pd.read_excel(file_or_path)
        else:
            try:
                df = pd.read_csv(file_or_path, delimiter=",")
            except Exception:
                file_or_path.seek(0)
                df = pd.read_csv(file_or_path, delimiter=";")

    df.columns = [str(col).strip() for col in df.columns]
    return df


def find_week_column(df: pd.DataFrame) -> str:
    for col in df.columns:
        if "semana" in col.lower():
            return col
    raise KeyError("Não foi encontrada a coluna 'Semana' na planilha enviada.")


def get_available_weeks(df: pd.DataFrame) -> List[str]:
    week_col = find_week_column(df)
    weeks = df[week_col].dropna().astype(str).str.strip().unique().tolist()
    try:
        weeks.sort(key=lambda x: int(x) if x.isdigit() else x)
    except Exception:
        pass
    return [w for w in weeks if w]


def filter_by_week(df: pd.DataFrame, week: str) -> pd.DataFrame:
    week_col = find_week_column(df)
    filtered = df[df[week_col].astype(str).str.strip() == str(week).strip()]
    if filtered.empty:
        raise ValueError(f"Nenhum registro encontrado para a semana '{week}'.")
    return filtered


def _replace_in_paragraph(paragraph, placeholder_map: Dict[str, str]):
    full_text = "".join(run.text for run in paragraph.runs)
    has_match = any(f"{{{{{key}}}}}" in full_text for key in placeholder_map)
    if not has_match:
        return

    for run in paragraph.runs:
        for key, val in placeholder_map.items():
            token = f"{{{{{key}}}}}"
            if token in run.text:
                run.text = run.text.replace(token, val)

    full_text = "".join(run.text for run in paragraph.runs)
    if any(f"{{{{{key}}}}}" in full_text for key in placeholder_map):
        for key, val in placeholder_map.items():
            token = f"{{{{{key}}}}}"
            full_text = full_text.replace(token, val)
        if paragraph.runs:
            paragraph.runs[0].text = full_text
            for run in paragraph.runs[1:]:
                run.text = ""


def render_docx(df: pd.DataFrame, template_source: Union[str, pathlib.Path, io.BytesIO] = None) -> io.BytesIO:
    if template_source is None:
        template_source = get_default_template_path()

    if isinstance(template_source, (str, pathlib.Path)):
        template_path = pathlib.Path(template_source)
        if not template_path.exists():
            raise FileNotFoundError(f"Arquivo de modelo não encontrado: {template_path}")
        doc = Document(str(template_path))
    else:
        doc = Document(template_source)

    placeholder_map = {}
    for col in df.columns:
        values = df[col].dropna().astype(str).tolist()
        placeholder_map[col] = "\n".join(values)

    for paragraph in doc.paragraphs:
        _replace_in_paragraph(paragraph, placeholder_map)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    _replace_in_paragraph(paragraph, placeholder_map)

    out_stream = io.BytesIO()
    doc.save(out_stream)
    out_stream.seek(0)
    return out_stream


def get_plan_preview(df: pd.DataFrame) -> Dict[str, Any]:
    data = {}
    for col in df.columns:
        values = df[col].dropna().astype(str).tolist()
        data[col] = values
    return data


# --- HTML Template (Embedded for 100% reliability on Vercel) ---
HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Gerador de Plano de Aula</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css">
    <style>
        :root {
            --primary-gradient: linear-gradient(135deg, #4f46e5 0%, #7c3aed 100%);
            --card-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.08), 0 8px 10px -6px rgba(0, 0, 0, 0.05);
        }
        body {
            background-color: #f8fafc;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
            color: #1e293b;
            min-height: 100vh;
        }
        .navbar-custom {
            background: var(--primary-gradient);
            padding: 1.2rem 0;
            box-shadow: 0 4px 15px rgba(79, 70, 229, 0.2);
        }
        .hero-section {
            background: white;
            border-radius: 1.25rem;
            box-shadow: var(--card-shadow);
            padding: 2.5rem;
            margin-top: -2rem;
            border: 1px solid #e2e8f0;
        }
        .btn-gradient {
            background: var(--primary-gradient);
            border: none;
            color: white;
            padding: 0.75rem 1.5rem;
            font-weight: 600;
            border-radius: 0.75rem;
            transition: all 0.2s ease;
        }
        .btn-gradient:hover {
            transform: translateY(-1px);
            box-shadow: 0 6px 20px rgba(79, 70, 229, 0.35);
            color: white;
        }
        .card-custom {
            border: 1px solid #e2e8f0;
            border-radius: 1rem;
            box-shadow: var(--card-shadow);
            background: white;
            overflow: hidden;
        }
        .upload-box {
            border: 2px dashed #cbd5e1;
            border-radius: 0.85rem;
            padding: 1.5rem;
            text-align: center;
            background: #f8fafc;
            cursor: pointer;
            transition: border-color 0.2s ease;
        }
        .upload-box:hover {
            border-color: #6366f1;
            background: #f1f5f9;
        }
        .preview-card {
            background: #ffffff;
            border-left: 4px solid #4f46e5;
            border-radius: 0.5rem;
            padding: 1rem 1.25rem;
            margin-bottom: 0.75rem;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        }
        .preview-label {
            font-weight: 700;
            font-size: 0.8rem;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            color: #64748b;
            margin-bottom: 0.25rem;
        }
        .preview-value {
            font-size: 0.95rem;
            color: #0f172a;
            white-space: pre-wrap;
        }
        @media print {
            .no-print { display: none !important; }
            .hero-section { box-shadow: none; border: none; margin: 0; padding: 0; }
            body { background: white; }
        }
    </style>
</head>
<body>

    <nav class="navbar navbar-dark navbar-custom no-print">
        <div class="container">
            <a class="navbar-brand d-flex align-items-center gap-2 fw-bold fs-4" href="#">
                <i class="bi bi-journal-bookmark-fill fs-3"></i>
                Gerador de Plano de Aula
            </a>
            <span class="navbar-text text-white-50 d-none d-sm-inline">
                Crie planos de aula automaticamente
            </span>
        </div>
    </nav>

    <div class="container my-4">
        <div class="hero-section mb-4 no-print">
            <h2 class="fw-bold mb-2 text-dark">Gerar Plano de Aula Semanal</h2>
            <p class="text-muted mb-4">
                Envie a planilha do curso técnico (Excel ou CSV), selecione a semana desejada e gere o plano de aula formatado.
            </p>

            <form id="planForm" action="/generate" method="POST" enctype="multipart/form-data">
                <div class="row g-4">
                    <div class="col-md-6">
                        <label class="form-label fw-bold"><i class="bi bi-file-earmark-spreadsheet text-success me-1"></i> Planilha do Curso (.xlsx ou .csv)</label>
                        <div class="upload-box" onclick="document.getElementById('spreadsheet').click()">
                            <i class="bi bi-cloud-arrow-up text-primary fs-2 d-block mb-1"></i>
                            <span id="spreadsheetName" class="fw-semibold text-secondary">Clique para selecionar a planilha</span>
                            <input type="file" class="d-none" id="spreadsheet" name="spreadsheet" accept=".xlsx, .xls, .csv" required>
                        </div>
                        <small class="text-muted mt-1 d-block">Planilha contendo a coluna <code>Semana</code>, <code>Título da aula</code>, etc.</small>
                    </div>

                    <div class="col-md-6">
                        <label class="form-label fw-bold"><i class="bi bi-file-earmark-word text-primary me-1"></i> Modelo Word (.docx) <span class="badge bg-light text-secondary border">Opcional</span></label>
                        <div class="upload-box" onclick="document.getElementById('template').click()">
                            <i class="bi bi-file-earmark-richtext text-indigo fs-2 d-block mb-1"></i>
                            <span id="templateName" class="fw-semibold text-secondary">Usar modelo padrão do curso</span>
                            <input type="file" class="d-none" id="template" name="template" accept=".docx">
                        </div>
                        <small class="text-muted mt-1 d-block">Deixe em branco para usar o modelo padrão com placeholders.</small>
                    </div>

                    <div class="col-md-6">
                        <label for="weekInput" class="form-label fw-bold"><i class="bi bi-calendar3 me-1"></i> Semana da Aula</label>
                        <div class="input-group">
                            <span class="input-group-text bg-white"><i class="bi bi-calendar-event"></i></span>
                            <input type="text" class="form-control form-control-lg" id="weekInput" name="week" placeholder="Ex: 1, 2, 3..." required>
                            <select class="form-select form-control-lg d-none" id="weekSelect" onchange="document.getElementById('weekInput').value = this.value">
                                <option value="">Selecione a semana...</option>
                            </select>
                        </div>
                        <small class="text-muted mt-1 d-block" id="weekHint">Informe o número ou identificador da semana.</small>
                    </div>

                    <div class="col-md-6 d-flex align-items-end gap-2">
                        <button type="submit" class="btn btn-gradient flex-grow-1 py-3" id="btnDownload">
                            <i class="bi bi-file-earmark-arrow-down me-1 fs-5"></i> Baixar Word (.docx)
                        </button>
                        <button type="button" class="btn btn-outline-secondary py-3 px-3" id="btnPreview" title="Visualizar antes de baixar">
                            <i class="bi bi-eye me-1"></i> Pré-visualizar
                        </button>
                    </div>
                </div>
            </form>
        </div>

        <div id="alertBox" class="alert d-none no-print" role="alert"></div>

        <div id="previewSection" class="card-custom p-4 d-none">
            <div class="d-flex justify-content-between align-items-center mb-3 pb-2 border-bottom">
                <h4 class="fw-bold mb-0 text-primary">
                    <i class="bi bi-card-checklist me-2"></i>Pré-visualização do Plano de Aula
                </h4>
                <div class="no-print d-flex gap-2">
                    <button class="btn btn-sm btn-outline-primary" onclick="window.print()">
                        <i class="bi bi-printer me-1"></i> Imprimir / Salvar PDF
                    </button>
                </div>
            </div>

            <div id="previewContent" class="row"></div>
        </div>
    </div>

    <footer class="text-center text-muted py-4 mt-5 no-print border-top">
        <div class="container">
            <small>Gerador de Plano de Aula &bull; Desenvolvido para Cursos Técnicos</small>
        </div>
    </footer>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        const spreadsheetInput = document.getElementById('spreadsheet');
        const spreadsheetName = document.getElementById('spreadsheetName');
        const templateInput = document.getElementById('template');
        const templateName = document.getElementById('templateName');
        const weekInput = document.getElementById('weekInput');
        const weekSelect = document.getElementById('weekSelect');
        const weekHint = document.getElementById('weekHint');
        const btnPreview = document.getElementById('btnPreview');
        const previewSection = document.getElementById('previewSection');
        const previewContent = document.getElementById('previewContent');
        const alertBox = document.getElementById('alertBox');

        function showAlert(msg, type = 'danger') {
            alertBox.className = `alert alert-${type} alert-dismissible fade show no-print`;
            alertBox.innerHTML = `<span>${msg}</span><button type="button" class="btn-close" data-bs-dismiss="alert"></button>`;
            alertBox.classList.remove('d-none');
            alertBox.scrollIntoView({ behavior: 'smooth' });
        }

        spreadsheetInput.addEventListener('change', async function() {
            if (this.files && this.files[0]) {
                const file = this.files[0];
                spreadsheetName.innerHTML = `<span class="text-success"><i class="bi bi-check-circle-fill me-1"></i>${file.name}</span>`;

                const formData = new FormData();
                formData.append('spreadsheet', file);

                try {
                    const res = await fetch('/api/weeks', { method: 'POST', body: formData });
                    const data = await res.json();
                    if (data.success && data.weeks && data.weeks.length > 0) {
                        weekSelect.innerHTML = '<option value="">Selecione uma semana detectada...</option>';
                        data.weeks.forEach(w => {
                            const opt = document.createElement('option');
                            opt.value = w;
                            opt.textContent = `Semana ${w}`;
                            weekSelect.appendChild(opt);
                        });
                        weekSelect.classList.remove('d-none');
                        weekInput.value = data.weeks[0];
                        weekSelect.value = data.weeks[0];
                        weekHint.innerHTML = `<span class="text-success"><i class="bi bi-stars me-1"></i>${data.weeks.length} semanas detectadas automaticamente!</span>`;
                    }
                } catch (e) {
                    console.log('Detect weeks failed:', e);
                }
            }
        });

        templateInput.addEventListener('change', function() {
            if (this.files && this.files[0]) {
                templateName.innerHTML = `<span class="text-primary"><i class="bi bi-check-circle-fill me-1"></i>${this.files[0].name}</span>`;
            }
        });

        btnPreview.addEventListener('click', async function() {
            if (!spreadsheetInput.files || !spreadsheetInput.files[0]) {
                showAlert('Por favor, selecione primeiro a planilha do curso.');
                return;
            }
            const week = weekInput.value.trim();
            if (!week) {
                showAlert('Por favor, informe ou selecione a semana desejada.');
                return;
            }

            btnPreview.disabled = true;
            btnPreview.innerHTML = `<span class="spinner-border spinner-border-sm me-1"></span> Carregando...`;

            const formData = new FormData();
            formData.append('spreadsheet', spreadsheetInput.files[0]);
            formData.append('week', week);

            try {
                const res = await fetch('/api/preview', { method: 'POST', body: formData });
                const data = await res.json();

                if (!data.success) {
                    showAlert(data.error || 'Não foi possível carregar os dados.');
                } else {
                    alertBox.classList.add('d-none');
                    previewContent.innerHTML = '';
                    
                    for (const [colName, valList] of Object.entries(data.data)) {
                        const valText = Array.isArray(valList) ? valList.join('\\n') : String(valList);
                        if (!valText.trim()) continue;

                        const colDiv = document.createElement('div');
                        colDiv.className = 'col-md-6 mb-3';
                        colDiv.innerHTML = `
                            <div class="preview-card h-100">
                                <div class="preview-label">${colName}</div>
                                <div class="preview-value">${valText}</div>
                            </div>
                        `;
                        previewContent.appendChild(colDiv);
                    }
                    previewSection.classList.remove('d-none');
                    previewSection.scrollIntoView({ behavior: 'smooth' });
                }
            } catch (err) {
                showAlert('Erro ao processar a pré-visualização.');
            } finally {
                btnPreview.disabled = false;
                btnPreview.innerHTML = `<i class="bi bi-eye me-1"></i> Pré-visualizar`;
            }
        });
    </script>
</body>
</html>"""


# --- Flask Routes ---

@app.route("/", methods=["GET"])
def index():
    return render_template_string(HTML_TEMPLATE)


@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "app": "Gerador de Plano de Aula"})


@app.route("/api/weeks", methods=["POST"])
def detect_weeks():
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

        template_bytes = None
        if "template" in request.files:
            tpl_file = request.files["template"]
            if tpl_file and tpl_file.filename:
                template_bytes = io.BytesIO(tpl_file.read())

        docx_buffer = render_docx(filtered, template_source=template_bytes)

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
    app.run(debug=True, host="127.0.0.1", port=5000)
