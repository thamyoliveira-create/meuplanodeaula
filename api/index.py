# -*- coding: utf-8 -*-
"""
Flask Web Application for Lesson Plan Generator
Entrypoint for Vercel Serverless Function: api/index.py
Supports 1º Ano and 2º Ano technical courses (Administração e Vendas).
"""

import io
import os
import csv
import pathlib
import sys
import traceback
from typing import Union, List, Dict, Any

from flask import Flask, request, send_file, jsonify, render_template_string
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

BASE_DIR = pathlib.Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB limit


# --- File Path Resolvers ---

def get_schedule_paths() -> List[pathlib.Path]:
    """Find schedule files for 1º and 2º ano."""
    files = []
    candidates = [
        pathlib.Path(__file__).parent / "schedule.csv",
        BASE_DIR / "data" / "schedule.csv",
        pathlib.Path(__file__).parent / "schedule_2ano.csv",
        BASE_DIR / "data" / "schedule_2ano.csv"
    ]
    for p in candidates:
        if p.exists() and p not in files:
            files.append(p)
    return files


def get_default_template_path() -> pathlib.Path:
    candidates = [
        pathlib.Path(__file__).parent / "template.docx",
        BASE_DIR / "template" / "template.docx",
        BASE_DIR / "lesson_plan_generator" / "template" / "template.docx",
        BASE_DIR / "api" / "template.docx"
    ]
    for p in candidates:
        if p.exists():
            return p
    return candidates[0]


# --- Data Parsing & Normalization ---

def normalize_row_dict(row: Dict[str, str]) -> Dict[str, str]:
    """Clean keys and values for robust column lookup."""
    clean = {}
    for k, v in row.items():
        if k:
            norm_key = k.replace("\n", " ").strip()
            clean[norm_key] = (v or "").strip()
    return clean


def parse_schedule_rows(source: Union[str, pathlib.Path, io.BytesIO], filename: str = "") -> List[Dict[str, str]]:
    """Parse CSV and return list of normalized dict rows."""
    if isinstance(source, (str, pathlib.Path)):
        p = pathlib.Path(source)
        if not p.exists():
            return []
        with open(str(p), "r", encoding="utf-8", errors="replace") as f:
            reader = csv.DictReader(f)
            return [normalize_row_dict(row) for row in reader]
    else:
        text = source.read().decode("utf-8", errors="replace")
        delimiter = ";" if ";" in text.split("\n")[0] else ","
        reader = csv.DictReader(io.StringIO(text), delimiter=delimiter)
        return [normalize_row_dict(row) for row in reader]


def load_all_schedules() -> List[Dict[str, str]]:
    """Load rows from all available schedule CSVs."""
    all_rows = []
    for p in get_schedule_paths():
        all_rows.extend(parse_schedule_rows(p))
    return all_rows


def get_organized_metadata(rows: List[Dict[str, str]]) -> Dict[str, Any]:
    """Organize disciplines and weeks by 2º Ano and 3º Ano."""
    year_2_disciplines = [
        "Introdução à Administração, Legislação e Pessoas",
        "Matemática Aplicada à Administração"
    ]
    year_3_disciplines = [
        "Marketing Estratégico",
        "Comunicação Empresarial",
        "Gestão Financeira e Contabilidade",
        "Gestão de Operações",
        "Empreendedorismo e Desenvolvimento de Negócios"
    ]

    # Dynamically find any extra disciplines
    all_disc_found = set()
    discipline_weeks = {}

    for r in rows:
        disc = r.get("Nome do componente", "").strip()
        week = r.get("Semana", "").strip()
        if disc:
            all_disc_found.add(disc)
            if disc not in discipline_weeks:
                discipline_weeks[disc] = set()
            if week:
                discipline_weeks[disc].add(week)

    def sort_weeks(w_list):
        try:
            return sorted(list(w_list), key=lambda x: int(x) if x.isdigit() else 999)
        except Exception:
            return sorted(list(w_list))

    y2_final = [d for d in year_2_disciplines if d in all_disc_found] or year_2_disciplines
    y3_final = [d for d in year_3_disciplines if d in all_disc_found] or year_3_disciplines

    # Include any remaining
    for d in sorted(list(all_disc_found)):
        if d not in y2_final and d not in y3_final:
            y3_final.append(d)

    sorted_disc_weeks = {d: sort_weeks(discipline_weeks.get(d, [str(i) for i in range(1, 29)])) for d in all_disc_found}

    return {
        "years": {
            "2º Ano (MTEC)": y2_final,
            "3º Ano (MTEC)": y3_final
        },
        "discipline_weeks": sorted_disc_weeks,
        "all_weeks": [str(i) for i in range(1, 29)]
    }


def filter_rows(rows: List[Dict[str, str]], week: str, discipline: str = "") -> List[Dict[str, str]]:
    results = []
    for r in rows:
        r_week = str(r.get("Semana", "")).strip()
        r_disc = str(r.get("Nome do componente", "")).strip()

        if r_week == str(week).strip():
            if not discipline or r_disc.lower() == discipline.strip().lower():
                results.append(r)
    return results


def compile_lesson_data(rows: List[Dict[str, str]], teacher: str, period: str, classroom: str, discipline: str, week: str) -> Dict[str, str]:
    if not rows:
        return {
            "professor": teacher or "Professor(a)",
            "disciplina": discipline or "Componente Curricular",
            "turma": classroom or "Turma",
            "periodo": period or f"Semana {week}",
            "semana": week or "1",
            "habilidades": "Desenvolvimento das competências técnicas e socioemocionais da semana.",
            "objetos": "Objetos de conhecimento previstos para o componente curricular.",
            "conteudo": f"Aulas e temas correspondentes à Semana {week}.",
            "objetivos": "Compreender, analisar e aplicar os conceitos abordados.",
            "estrategias": "Aulas expositivas dialogadas, estudos de caso práticos e atividades em grupo.",
            "recursos": "Lousa, projetor multimídia, computadores do laboratório e material de apoio.",
            "avaliacao": "Participação ativa, resolução de exercícios práticos e autoavaliação.",
            "referencias": "Currículo Paulista / Diretrizes da Educação Profissional Técnica - SEDUC-SP."
        }

    aulas_titles = []
    temas = set()
    hab_tecnicas = set()
    hab_socio = set()
    obj_conhecimento = set()
    objetivos = []

    for idx, r in enumerate(rows, 1):
        tema = r.get("Tema da semana", "").strip()
        if tema: temas.add(tema)

        titulo = r.get("Título da aula", "").strip()
        if titulo: aulas_titles.append(f"• {titulo}")

        ht = r.get("Habilidades técnicas", "").strip() or r.get("Habilidade técnica", "").strip()
        if ht: hab_tecnicas.add(f"• {ht}")

        hs = r.get("Habildades socioemocionais", "").strip() or r.get("Competências Socioemocionais", "").strip()
        if hs: hab_socio.add(f"• {hs}")

        oc = r.get("Objeto de conhecimento", "").strip()
        if oc: obj_conhecimento.add(f"• {oc}")

        obj = r.get("Objetivo da aula", "").strip() or r.get("Objetivos da aula", "").strip()
        if obj: objetivos.append(f"• {obj}")

    tema_str = " | ".join(temas) if temas else f"Semana {week}"
    conteudo_str = f"Tema: {tema_str}\n" + "\n".join(aulas_titles) if aulas_titles else tema_str
    habilidades_str = "Habilidades Técnicas:\n" + "\n".join(hab_tecnicas) if hab_tecnicas else "Desenvolver competências técnicas da área."
    if hab_socio:
        habilidades_str += "\n\nHabilidades Socioemocionais:\n" + "\n".join(hab_socio)

    return {
        "professor": teacher,
        "disciplina": discipline or (rows[0].get("Nome do componente", "") if rows else "Técnico"),
        "turma": classroom,
        "periodo": period,
        "semana": week,
        "habilidades": habilidades_str,
        "objetos": "\n".join(obj_conhecimento) if obj_conhecimento else "Conceitos fundamentais da disciplina.",
        "conteudo": conteudo_str,
        "objetivos": "\n".join(objetivos) if objetivos else "Desenvolver os conhecimentos previstos.",
        "estrategias": "Aulas expositivas dialogadas, estudos de casos práticos, dinâmicas e resolução de problemas em grupo.",
        "recursos": "Quadro, projetor multimídia, computadores do laboratório, material de apoio impresso/digital.",
        "avaliacao": "Avaliação contínua, participação ativa nas discussões, entrega de exercícios e produções práticas.",
        "referencias": "Diretrizes Curriculares da Educação Profissional Técnica - SEDUC-SP e bibliografia recomendada."
    }


def create_clean_document_from_scratch(lesson_data: Dict[str, str]) -> Document:
    """Build official school lesson plan document from scratch."""
    doc = Document()

    # School Header
    p_header = doc.add_paragraph()
    p_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p_header.add_run("GOVERNO DO ESTADO DE SÃO PAULO – SECRETARIA DA EDUCAÇÃO\n")
    r1.bold = True
    r1.font.size = Pt(10)
    r2 = p_header.add_run("DIRETORIA DE ENSINO - REGIÃO DE ARARAQUARA\nEE. PROF. DINORA MARCONDES GOMES\n")
    r2.bold = True
    r2.font.size = Pt(9.5)
    r3 = p_header.add_run("RUA EMILIA GALLI, 549 - CENTRO - AMÉRICO BRASILIENSE - SP | TEL: (16) 3392-1335\n")
    r3.font.size = Pt(8.5)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run(f"PLANO DE AULA - 2026 - SEMANA {lesson_data['semana']}\n")
    r_title.bold = True
    r_title.font.size = Pt(12)

    # Main Table
    items = [
        ("Professor(a)", lesson_data['professor']),
        ("Componente curricular", lesson_data['disciplina']),
        ("3. Ano/Série/Turma", lesson_data['turma']),
        ("4. Período de realização", lesson_data['periodo']),
        ("5. Habilidades", lesson_data['habilidades']),
        ("6. Objetos de Conhecimento", lesson_data['objetos']),
        ("7. Conteúdo", lesson_data['conteudo']),
        ("8. Objetivos", lesson_data['objetivos']),
        ("9. Estratégias", lesson_data['estrategias']),
        ("10. Recursos didáticos", lesson_data['recursos']),
        ("11. Critérios de Avaliação", lesson_data['avaliacao']),
        ("12. Referências", lesson_data['referencias']),
    ]

    table = doc.add_table(rows=len(items), cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for idx, (label, val) in enumerate(items):
        cell = table.cell(idx, 0)
        p = cell.paragraphs[0]
        r_lbl = p.add_run(f"{label}: ")
        r_lbl.bold = True
        r_lbl.font.size = Pt(10)
        
        if "\n" in val or len(val) > 40:
            p.add_run("\n")
        r_val = p.add_run(val)
        r_val.font.size = Pt(10)

    return doc


def generate_filled_docx(lesson_data: Dict[str, str], custom_template: io.BytesIO = None) -> io.BytesIO:
    doc = None
    if custom_template:
        try:
            custom_template.seek(0)
            doc = Document(custom_template)
        except Exception:
            pass

    if doc is None:
        template_path = get_default_template_path()
        if template_path.exists():
            try:
                doc = Document(str(template_path))
            except Exception:
                pass

    if doc is None:
        doc = create_clean_document_from_scratch(lesson_data)
    else:
        placeholders = {
            "Professor": lesson_data["professor"],
            "Nome do componente": lesson_data["disciplina"],
            "Disciplina": lesson_data["disciplina"],
            "Componente curricular": lesson_data["disciplina"],
            "Ano/Série/Turma": lesson_data["turma"],
            "Turma": lesson_data["turma"],
            "Período de realização": lesson_data["periodo"],
            "Data": lesson_data["periodo"],
            "Semana": lesson_data["semana"],
            "Habilidades": lesson_data["habilidades"],
            "Objetos de Conhecimento": lesson_data["objetos"],
            "Conteúdo": lesson_data["conteudo"],
            "Objetivos": lesson_data["objetivos"],
            "Estratégias": lesson_data["estrategias"],
            "Recursos didáticos": lesson_data["recursos"],
            "Critérios de Avaliação": lesson_data["avaliacao"],
            "Referências": lesson_data["referencias"]
        }

        for p in doc.paragraphs:
            for k, v in placeholders.items():
                token = f"{{{{{k}}}}}"
                if token in p.text:
                    p.text = p.text.replace(token, v)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for k, v in placeholders.items():
                        token = f"{{{{{k}}}}}"
                        if token in cell.text:
                            cell.text = cell.text.replace(token, v)

                    c_text = cell.text.strip()
                    if c_text == "Professor":
                        cell.text = f"Professor: {lesson_data['professor']}"
                    elif c_text.startswith("Componente curricular:"):
                        cell.text = f"Componente curricular: {lesson_data['disciplina']}"
                    elif "Ano/Série/Turma" in c_text:
                        cell.text = f"3. Ano/Série/Turma: {lesson_data['turma']}"
                    elif "Período de realização" in c_text:
                        cell.text = f"4. Período de realização: {lesson_data['periodo']}"
                    elif c_text.startswith("5. Habilidades") or c_text.startswith("5 . Habilidades"):
                        cell.text = f"5. Habilidades:\n{lesson_data['habilidades']}"
                    elif "Objetos de Conhecimento" in c_text:
                        cell.text = f"6. Objetos de Conhecimento:\n{lesson_data['objetos']}"
                    elif "Conteúdo" in c_text:
                        cell.text = f"7. Conteúdo:\n{lesson_data['conteudo']}"
                    elif "Objetivos" in c_text and not "Objetos" in c_text:
                        cell.text = f"8. Objetivos:\n{lesson_data['objetivos']}"
                    elif "Estratégias" in c_text:
                        cell.text = f"9. Estratégias:\n{lesson_data['estrategias']}"
                    elif "Recursos didáticos" in c_text:
                        cell.text = f"10. Recursos didáticos:\n{lesson_data['recursos']}"
                    elif "Critérios de Avaliação" in c_text:
                        cell.text = f"11. Critérios de Avaliação:\n{lesson_data['avaliacao']}"
                    elif "Referências" in c_text:
                        cell.text = f"12. Referências:\n{lesson_data['referencias']}"

    out_stream = io.BytesIO()
    doc.save(out_stream)
    out_stream.seek(0)
    return out_stream


# --- Flask Routes ---

@app.route("/", methods=["GET"])
@app.route("/api", methods=["GET"])
def index():
    static_html = BASE_DIR / "public" / "index.html"
    if static_html.exists():
        with open(str(static_html), "r", encoding="utf-8") as f:
            return f.read()
    return jsonify({"status": "ok", "app": "Gerador de Plano de Aula"})


@app.route("/api/health", methods=["GET"])
@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "app": "Gerador de Plano de Aula"})


@app.route("/api/initial-data", methods=["GET"])
@app.route("/initial-data", methods=["GET"])
def initial_data():
    try:
        rows = load_all_schedules()
        meta = get_organized_metadata(rows)
        return jsonify({"success": True, **meta})
    except Exception as e:
        return jsonify({"success": True, "years": {
            "2º Ano (MTEC)": [
                "Introdução à Administração, Legislação e Pessoas",
                "Matemática Aplicada à Administração"
            ],
            "3º Ano (MTEC)": [
                "Marketing Estratégico",
                "Comunicação Empresarial",
                "Gestão Financeira e Contabilidade",
                "Gestão de Operações",
                "Empreendedorismo e Desenvolvimento de Negócios"
            ]
        }, "all_weeks": [str(i) for i in range(1, 29)]})


@app.route("/api/preview-plan", methods=["POST"])
@app.route("/preview-plan", methods=["POST"])
def preview_plan():
    try:
        teacher = request.form.get("teacher", "").strip()
        period = request.form.get("period", "").strip()
        classroom = request.form.get("classroom", "").strip()
        discipline = request.form.get("discipline", "").strip()
        week = request.form.get("week", "1").strip()

        rows = []
        if "custom_spreadsheet" in request.files and request.files["custom_spreadsheet"].filename:
            f = request.files["custom_spreadsheet"]
            rows = parse_schedule_rows(io.BytesIO(f.read()), filename=f.filename)
        else:
            rows = load_all_schedules()

        filtered = filter_rows(rows, week, discipline)
        lesson_data = compile_lesson_data(filtered, teacher, period, classroom, discipline, week)
        return jsonify({"success": True, "data": lesson_data})
    except Exception as e:
        return jsonify({"success": False, "error": f"Erro interno: {str(e)}"}), 500


@app.route("/api/generate", methods=["POST"])
@app.route("/generate", methods=["POST"])
def generate():
    try:
        teacher = request.form.get("teacher", "").strip() or "Professor"
        period = request.form.get("period", "").strip() or "2026"
        classroom = request.form.get("classroom", "").strip() or "Turma"
        discipline = request.form.get("discipline", "").strip() or "Componente"
        week = request.form.get("week", "1").strip()

        rows = []
        if "custom_spreadsheet" in request.files and request.files["custom_spreadsheet"].filename:
            f = request.files["custom_spreadsheet"]
            rows = parse_schedule_rows(io.BytesIO(f.read()), filename=f.filename)
        else:
            rows = load_all_schedules()

        custom_tpl = None
        if "custom_template" in request.files and request.files["custom_template"].filename:
            t = request.files["custom_template"]
            custom_tpl = io.BytesIO(t.read())

        filtered = filter_rows(rows, week, discipline)
        lesson_data = compile_lesson_data(filtered, teacher, period, classroom, discipline, week)
        docx_buffer = generate_filled_docx(lesson_data, custom_template=custom_tpl)

        clean_disc = "".join(c for c in discipline if c.isalnum() or c in " _-").strip().replace(" ", "_")
        filename = f"Plano_de_Aula_Semana_{week}_{clean_disc or 'Tecnico'}.docx"

        docx_buffer.seek(0)
        return send_file(
            docx_buffer,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        err_msg = traceback.format_exc()
        return f"Erro ao gerar documento Word: {str(e)}\n\nDetalhes:\n{err_msg}", 500


if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=5000)
