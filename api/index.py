# -*- coding: utf-8 -*-
"""
Flask Web Application for Lesson Plan Generator
Entrypoint for Vercel Serverless Function: api/index.py
Includes Intelligent Pedagogical Engine with Doug Lemov Techniques and Active Methodologies.
"""

import io
import os
import csv
import pathlib
import sys
import traceback
from typing import Union, List, Dict, Any

from flask import Flask, request, send_file, jsonify
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

BASE_DIR = pathlib.Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

try:
    from api.curriculum_data import CURRICULUM_DATA
except Exception:
    try:
        from curriculum_data import CURRICULUM_DATA
    except Exception:
        CURRICULUM_DATA = []

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB limit


# --- Pedagogical Intelligence: Lemov Techniques & Active Methodologies ---

def select_lemov_and_active_methodology(discipline: str, titles: List[str], objectives: List[str]) -> Dict[str, str]:
    """
    Intelligently select and construct Doug Lemov techniques and active methodologies
    tailored to the technical subject and lesson objectives.
    """
    disc_lower = discipline.lower()
    content_text = " ".join(titles + objectives).lower()

    # Case 1: Mathematical / Financial / Quantitative
    if "matemática" in disc_lower or "financeira" in disc_lower or "contabilidade" in disc_lower or "cálculo" in content_text:
        return {
            "estrategia_principal": "Modelagem Gradual (Eu Faço, Nós Fazemos, Você Faz) e Aprendizagem Baseada em Problemas (PBL)",
            "lemov_tecnicas": "• Faça Agora (Do Now): Exercício rápido de aquecimento de 3 a 5 min.\n• Modelagem Gradual (I Do, We Do, You Do): Demonstração passo a passo do cálculo na lousa, resolução conjunta e prática autônoma individual.\n• Padronize o Formato: Organização estruturada do raciocínio matemático no caderno.\n• Circule pela Sala: Monitoramento ativo das etapas de cálculo com intervenções pontuais.\n• Bilhete de Saída (Exit Ticket): Mini-desafio final de 3 minutos para checagem imediata de retenção.",
            "recursos_lemov": "Lousa interativa/quadro, folhas de atividades práticas dirigidas, calculadoras e formulários de exercícios estruturados.",
            "avaliacao_lemov": "Avaliação formativa processual com aplicação do 'Certo é Certo' (correção de etapas) e conferência do 'Bilhete de Saída'."
        }

    # Case 2: Marketing, Sales, Strategy & Business Communication
    elif "marketing" in disc_lower or "vendas" in disc_lower or "comunicação" in disc_lower or "estratégico" in disc_lower or "negócios" in disc_lower:
        return {
            "estrategia_principal": "Estudo de Casos Empresariais Reais (PBL), Aprendizagem por Pares e Design de Estratégias",
            "lemov_tecnicas": "• Gancho Inicial (Hook): Apresentação de um case ou campanha de mercado instigante nos primeiros minutos.\n• Vire e Converse (Turn and Talk): Debate rápido em duplas para gerar hipóteses de mercado antes da discussão plenária.\n• Puxe Mais (Stretch It): Perguntas de aprofundamento aos alunos que respondem corretamente para expandir a visão crítica.\n• Mostre e Conte (Show Call): Projeção de soluções e pitches dos alunos para análise construtiva coletiva.\n• Formato Importa: Estímulo à comunicação oral e escrita com vocabulário técnico de marketing e negócios.",
            "recursos_lemov": "Projetor multimídia, computadores do laboratório com acesso a estudos de casos reais, material de apoio e ambiente virtual.",
            "avaliacao_lemov": "Avaliação contínua por rubricas: análise crítica nos estudos de caso, qualidade da argumentação técnica e participação cooperativa."
        }

    # Case 3: Management, Operations, People & Legal (RH)
    elif "pessoas" in disc_lower or "legislação" in disc_lower or "administração" in disc_lower or "operações" in disc_lower or "recursos humanos" in content_text:
        return {
            "estrategia_principal": "Simulação de Rotinas Organizacionais, Análise Crítica de Cenários e Sala de Aula Invertida",
            "lemov_tecnicas": "• Faça Agora (Do Now): Situação-problema inicial sobre dilemas da rotina empresarial/jurídica.\n• Todos Escrevem (Everybody Writes): Tempo individual cronometrado para os alunos registrarem soluções antes do debate.\n• Sem Escapatória (No Opt Out): Garantir que todo aluno participe ativamente e reformule a resposta até atingir o rigor técnico.\n• Chamada Fria (Cold Call): Perguntas dirigidas de forma acolhedora para manter o foco e engajamento uniforme de toda a turma.\n• Certo é Certo (Right is Right): Exigência de conformidade às normas legais (ex: CLT, LGPD) e procedimentos organizacionais.",
            "recursos_lemov": "Lousa digital, cenários simulados de departamento de RH e gestão, legislação comentada e formulários administrativos.",
            "avaliacao_lemov": "Acompanhamento formativo diário, verificação das produções escritas individuais e assertividade na aplicação das rotinas administrativas."
        }

    # General technical subjects fallback
    else:
        return {
            "estrategia_principal": "Metodologias Ativas com Modelagem Gradual e Resolução Prática de Problemas",
            "lemov_tecnicas": "• Faça Agora (Do Now): Atividade de retomada nos primeiros 5 minutos.\n• Modelagem Gradual (Eu Faço, Nós Fazemos, Você Faz): Explanação conceitual, aplicação orientada e execução individual.\n• Vire e Converse (Turn and Talk): Discussão em pares para fixação dos conceitos técnicos.\n• Circule pela Sala: Feedback formativo imediato durante as atividades.\n• Bilhete de Saída (Exit Ticket): Fechamento com verificação de aprendizagem.",
            "recursos_lemov": "Lousa, projetor multimídia, material impresso e ambiente de aprendizagem.",
            "avaliacao_lemov": "Avaliação formativa e contínua, observação da prática orientada e entrega dos bilhetes de saída."
        }


# --- Curriculum Rows & Compiling ---

def get_curriculum_rows() -> List[Dict[str, str]]:
    if CURRICULUM_DATA:
        return CURRICULUM_DATA

    rows = []
    for fname in ["schedule_2ano.csv", "schedule.csv"]:
        for parent in [pathlib.Path(__file__).parent, BASE_DIR / "data", BASE_DIR]:
            p = parent / fname
            if p.exists():
                with open(str(p), "r", encoding="utf-8", errors="replace") as f:
                    reader = csv.DictReader(f)
                    for r in reader:
                        clean_r = {k.replace("\n", " ").strip(): (v or "").strip() for k, v in r.items() if k}
                        rows.append(clean_r)
    return rows


def filter_rows(rows: List[Dict[str, str]], week: str, discipline: str = "") -> List[Dict[str, str]]:
    results = []
    w_target = str(week).strip()
    d_target = discipline.strip().lower()

    for r in rows:
        r_week = str(r.get("Semana", "")).strip()
        r_disc = str(r.get("Nome do componente", "")).strip().lower()

        if r_week == w_target:
            if not d_target or d_target in r_disc or r_disc in d_target:
                results.append(r)
    return results


def compile_lesson_data(rows: List[Dict[str, str]], teacher: str, period: str, classroom: str, discipline: str, week: str) -> Dict[str, str]:
    aulas_titles = []
    temas = set()
    hab_tecnicas = set()
    hab_socio = set()
    obj_conhecimento = set()
    objetivos = []

    for r in rows:
        tema = r.get("Tema da semana", "").strip()
        if tema: temas.add(tema)

        titulo = r.get("Título da aula", "").strip()
        if titulo: aulas_titles.append(f"• {titulo}")

        ht = r.get("Habilidades técnicas", "").strip() or r.get("Habilidade técnica", "").strip()
        if ht: hab_tecnicas.add(f"• {ht}")

        hs = r.get("Habildades socioemocionais", "").strip() or r.get("Competências Socioemocionais ", "").strip() or r.get("Competências Socioemocionais", "").strip()
        if hs: hab_socio.add(f"• {hs}")

        oc = r.get("Objeto de conhecimento", "").strip()
        if oc: obj_conhecimento.add(f"• {oc}")

        obj = r.get("Objetivo da aula", "").strip() or r.get("Objetivos da aula", "").strip()
        if obj: objetivos.append(f"• {obj}")

    tema_str = " | ".join(temas) if temas else f"Semana {week}"
    conteudo_str = f"Tema da Semana: {tema_str}\n\nAulas Previstas:\n" + "\n".join(aulas_titles) if aulas_titles else tema_str
    habilidades_str = "Habilidades Técnicas:\n" + "\n".join(hab_tecnicas) if hab_tecnicas else "Desenvolver competências técnicas da área profissional."
    if hab_socio:
        habilidades_str += "\n\nHabilidades Socioemocionais:\n" + "\n".join(hab_socio)

    # Dynamic Lemov & Active Methodologies Assignment
    lemov_info = select_lemov_and_active_methodology(discipline, aulas_titles, objetivos)

    estrategias_completas = (
        f"Metodologia Ativa Central:\n{lemov_info['estrategia_principal']}\n\n"
        f"Técnicas de Ensino Estruturadas (Doug Lemov - Aula Nota 10):\n"
        f"{lemov_info['lemov_tecnicas']}"
    )

    return {
        "professor": teacher or "Professor(a)",
        "disciplina": discipline or (rows[0].get("Nome do componente", "") if rows else "Técnico"),
        "turma": classroom or "Turma MTEC",
        "periodo": period or f"Semana {week}",
        "semana": week or "1",
        "habilidades": habilidades_str,
        "objetos": "\n".join(obj_conhecimento) if obj_conhecimento else "Conceitos fundamentais da disciplina.",
        "conteudo": conteudo_str,
        "objetivos": "\n".join(objetivos) if objetivos else "Desenvolver os conhecimentos previstos.",
        "estrategias": estrategias_completas,
        "recursos": lemov_info["recursos_lemov"],
        "avaliacao": lemov_info["avaliacao_lemov"],
        "referencias": "Diretrizes Curriculares da Educação Profissional Técnica - SEDUC-SP; LEMOV, Doug. Aula Nota 10 3.0: 63 técnicas para melhorar a gestão da sala de aula."
    }


def create_clean_document(lesson_data: Dict[str, str]) -> Document:
    """Build pristine school lesson plan document with Lemov methodologies."""
    doc = Document()

    # Header
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

    items = [
        ("Professor(a)", lesson_data['professor']),
        ("Componente curricular", lesson_data['disciplina']),
        ("3. Ano/Série/Turma", lesson_data['turma']),
        ("4. Período de realização", lesson_data['periodo']),
        ("5. Habilidades", lesson_data['habilidades']),
        ("6. Objetos de Conhecimento", lesson_data['objetos']),
        ("7. Conteúdo", lesson_data['conteudo']),
        ("8. Objetivos", lesson_data['objetivos']),
        ("9. Estratégias (Metodologias Ativas & Lemov)", lesson_data['estrategias']),
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
    return jsonify({"status": "ok", "app": "Gerador de Plano de Aula", "rows": len(get_curriculum_rows())})


@app.route("/api/initial-data", methods=["GET"])
@app.route("/initial-data", methods=["GET"])
def initial_data():
    return jsonify({
        "success": True,
        "years": {
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
        },
        "all_weeks": [str(i) for i in range(1, 29)]
    })


@app.route("/api/preview-plan", methods=["POST"])
@app.route("/preview-plan", methods=["POST"])
def preview_plan():
    try:
        teacher = request.form.get("teacher", "").strip()
        period = request.form.get("period", "").strip()
        classroom = request.form.get("classroom", "").strip()
        discipline = request.form.get("discipline", "").strip()
        week = request.form.get("week", "1").strip()

        rows = get_curriculum_rows()
        filtered = filter_rows(rows, week, discipline)
        lesson_data = compile_lesson_data(filtered, teacher, period, classroom, discipline, week)
        return jsonify({"success": True, "data": lesson_data})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


@app.route("/api/generate", methods=["POST"])
@app.route("/generate", methods=["POST"])
def generate():
    try:
        teacher = request.form.get("teacher", "").strip() or "Professor(a)"
        period = request.form.get("period", "").strip() or "Semana de Aula"
        classroom = request.form.get("classroom", "").strip() or "Turma MTEC"
        discipline = request.form.get("discipline", "").strip() or "Administração"
        week = request.form.get("week", "1").strip()

        rows = get_curriculum_rows()
        filtered = filter_rows(rows, week, discipline)
        lesson_data = compile_lesson_data(filtered, teacher, period, classroom, discipline, week)

        doc = create_clean_document(lesson_data)

        docx_buffer = io.BytesIO()
        doc.save(docx_buffer)
        docx_buffer.seek(0)

        clean_disc = "".join(c for c in discipline if c.isalnum() or c in " _-").strip().replace(" ", "_")
        filename = f"Plano_de_Aula_Semana_{week}_{clean_disc or 'Tecnico'}.docx"

        return send_file(
            docx_buffer,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        err_msg = traceback.format_exc()
        return f"Erro ao gerar documento: {str(e)}\n\nDetalhes:\n{err_msg}", 500


if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=5000)
