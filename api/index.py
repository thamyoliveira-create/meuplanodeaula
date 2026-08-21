# -*- coding: utf-8 -*-
"""
Flask Web Application for Lesson Plan Generator
Entrypoint for Vercel Serverless Function: api/index.py
Matches exact EE. Prof. Dinora Marcondes Gomes school template format, 50-min periods,
real school infrastructure and customizable lesson styles (interativa, pratica, tradicional, revisao).
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
from docx.shared import Pt
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


# --- Pedagogical Engine: 50-min lesson, Lesson Styles, Lemov & Real School Resources ---

def build_coherent_lesson_development(discipline: str, titles: List[str], objectives: List[str], tema: str, lesson_style: str = "interativa", with_inclusion: bool = False) -> Dict[str, str]:
    """
    Generate coherent 50-minute lesson development customized by Lesson Style:
    - 'interativa' (Metodologias Ativas, Lemov Hook / Turn & Talk, Debates e Cases)
    - 'pratica' (Mão na Massa, Sala Técnica, Netbooks, Planilhas e Simulações Reais)
    - 'tradicional' (Expositiva Dialogada, Modelagem Gradual, Quadro Branco e Exercícios Dirigidos)
    - 'revisao' (Diagnóstica, Técnica Ache o Erro, Ticket de Saída e Feedback)
    """
    disc_lower = (discipline or "").lower()
    text_content = " ".join(titles + objectives).lower()
    tema_clean = tema or "o conteúdo previsto"

    inclusion_text = ""
    inclusion_recurso = ""
    if with_inclusion:
        inclusion_text = "\n\n• Adaptações Curriculares e Acessibilidade (Inclusão/PEI): Mediação orientada com apoio de pares, fragmentação de comandos em etapas curtas, disponibilização prévia de materiais digitais com fonte ampliada e flexibilização de tempo para realização das atividades nos netbooks e sala especializada."
        inclusion_recurso = ", materiais com fonte ampliada/leitura facilitada e suporte individualizado"

    # 1. Aula Prática & Mão na Massa
    if lesson_style == "pratica":
        estrategias = (
            f"• Momento Inicial (5 a 10 min): Acolhimento e apresentação rápida da missão prática na TV/quadro branco, contextualizando uma situação real de mercado sobre {tema_clean}.\n\n"
            f"• Desenvolvimento (Mediação e Prática - 30 a 35 min): Na sala especializada do curso técnico e utilizando os netbooks, os alunos realizam uma simulação prática das rotinas corporativas (como elaboração de propostas, planilhas técnicas ou documentos administrativos). O professor circula orientando individualmente e aplicando a técnica de 'Prática Guiada'.\n\n"
            f"• Fechamento (5 a 10 min): Validação das entregas práticas, correção de detalhes técnicos recorrentes e síntese dos procedimentos realizados.{inclusion_text}"
        )
        recursos = f"Sala especializada do curso técnico, netbooks escolares para atividades práticas e planilhas, TV para projeção de instruções, modelos de documentos empresariais e quadro branco{inclusion_recurso}."
        avaliacao = "Avaliação processual e de desempenho prático, considerando a precisão na execução das tarefas, a organização nos netbooks e a postura profissional."

    # 2. Aula Tradicional & Estruturada
    elif lesson_style == "tradicional":
        estrategias = (
            f"• Momento Inicial (5 a 10 min): Acolhimento e ativação prévia com pergunta conceitual projetada na TV/quadro branco para situar a turma em relação a {tema_clean}.\n\n"
            f"• Desenvolvimento (Mediação e Prática - 30 a 35 min): Aula expositiva dialogada com apoio da TV e registros no quadro branco. Demonstração detalhada dos conceitos através da 'Modelagem Gradual' (Eu Faço, Nós Fazemos, Você Faz). Em seguida, os alunos realizam individualmente exercícios dirigidos de fixação e registro no caderno/netbooks.\n\n"
            f"• Fechamento (5 a 10 min): Correção coletiva dos principais pontos na lousa, esclarecimento pontual de dúvidas conceituais e sistematização final.{inclusion_text}"
        )
        recursos = f"Quadro branco, TV para exibição de esquemas e conceitos teóricos, apostilas, cadernos e folhas de exercícios dirigidos{inclusion_recurso}."
        avaliacao = "Avaliação formativa e contínua, observando a atenção, a resolução correta dos exercícios dirigidos e a compreensão dos conceitos fundamentais."

    # 3. Aula de Revisão, Fixação & Feedback
    elif lesson_style == "revisao":
        estrategias = (
            f"• Momento Inicial (5 a 10 min): Acolhimento e aplicação da técnica 'Faça Agora' (Do Now) com uma questão diagnóstica rápida no quadro/TV sobre os tópicos de {tema_clean}.\n\n"
            f"• Desenvolvimento (Mediação e Prática - 30 a 35 min): Aplicação da técnica Lemov 'Ache o Erro' (Find the Mistake) projetada na TV, na qual os alunos identificam e corrigem falhas em cenários organizacionais simulados. Atendimento individualizado pelo professor para esclarecimento de dúvidas pontuais com apoio dos netbooks.\n\n"
            f"• Fechamento (5 a 10 min): Aplicação de 'Ticket de Saída' (Exit Ticket) de 3 minutos para checagem rápida de retenção e diagnóstico da aprendizagem da turma.{inclusion_text}"
        )
        recursos = f"Quadro branco, TV para projeção de exercícios diagnósticos, netbooks para consulta de resumos e fichas de avaliação formativa{inclusion_recurso}."
        avaliacao = "Avaliação diagnóstica e formativa, com foco na superação de dúvidas recorrentes, na identificação de erros conceituais e na evolução individual."

    # 4. Aula Interativa & Metodologias Ativas (Default)
    else:
        estrategias = (
            f"• Momento Inicial (5 a 10 min): Acolhimento e apresentação de um 'Gancho Inicial' (Hook) na TV com um vídeo/case de mercado relacionado a {tema_clean}, despertando o interesse da turma.\n\n"
            f"• Desenvolvimento (Mediação e Prática - 30 a 35 min): Explanação dialogada pelo professor com suporte da TV. Em seguida, na sala especializada do curso técnico ou com uso de netbooks, aplicação da técnica 'Vire e Converse' (Turn and Talk) em duplas para debaterem a estratégia e formularem soluções práticas.\n\n"
            f"• Fechamento (5 a 10 min): Compartilhamento das propostas das duplas, síntese do professor no quadro branco e alinhamento com as práticas reais do mercado (utilizando o auditório quando houver apresentação de projetos).{inclusion_text}"
        )
        recursos = f"Quadro branco, TV para exibição de cases e campanhas, netbooks para pesquisas e propostas, sala especializada do curso técnico para dinâmicas e auditório para apresentações{inclusion_recurso}."
        avaliacao = "Avaliação processual baseada na participação ativa nas discussões, na clareza da argumentação técnica durante os debates e na produção em equipe."

    return {
        "estrategias": estrategias,
        "recursos": recursos,
        "avaliacao": avaliacao
    }


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
        r_disc = str(r.get("Nome do componente", "") or r.get("Unidade curricular", "")).strip().lower()

        if r_week == w_target:
            if not d_target or d_target in r_disc or r_disc in d_target:
                results.append(r)
    return results


def compile_lesson_data(rows: List[Dict[str, str]], teacher: str, period: str, classroom: str, discipline: str, week: str, lesson_style: str = "interativa", with_inclusion: bool = False) -> Dict[str, str]:
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

        hs = r.get("Habildades socioemocionais", "").strip() or r.get("Competências Socioemocionais ", "").strip() or r.get("Competências Socioemocionais", "").strip() or r.get("Habilidades socioemocionais", "").strip()
        if hs: hab_socio.add(f"• {hs}")

        oc = r.get("Objeto de conhecimento", "").strip() or r.get("Objeto de conhecimento – macro", "").strip()
        if oc: obj_conhecimento.add(f"• {oc}")

        obj = r.get("Objetivo da aula", "").strip() or r.get("Objetivos da aula", "").strip()
        if obj: objetivos.append(f"• {obj}")

    tema_str = " | ".join(temas) if temas else f"Semana {week}"
    conteudo_str = f"Tema: {tema_str}\n\n" + "\n".join(aulas_titles) if aulas_titles else tema_str
    
    habilidades_str = "Habilidades Técnicas:\n" + "\n".join(hab_tecnicas) if hab_tecnicas else "Desenvolver as competências técnicas previstas no currículo do curso técnico."
    if hab_socio:
        habilidades_str += "\n\nHabilidades Socioemocionais:\n" + "\n".join(hab_socio)

    pedagogy = build_coherent_lesson_development(discipline, aulas_titles, objetivos, tema_str, lesson_style, with_inclusion)

    return {
        "professor": teacher or "Prof. Thamy Oliveira",
        "disciplina": discipline or (rows[0].get("Nome do componente", "") if rows else "Administração"),
        "turma": classroom or "2º MTEC - Administração",
        "periodo": period or f"Semana de Aulas",
        "semana": week or "1",
        "habilidades": habilidades_str,
        "objetos": "\n".join(obj_conhecimento) if obj_conhecimento else "Conceitos fundamentais e aplicados da disciplina.",
        "conteudo": conteudo_str,
        "objetivos": "\n".join(objetivos) if objetivos else "Compreender e aplicar os conhecimentos técnicos abordados na prática.",
        "estrategias": pedagogy["estrategias"],
        "recursos": pedagogy["recursos"],
        "avaliacao": pedagogy["avaliacao"],
        "referencias": "Currículo Paulista / Diretrizes Curriculares da Educação Profissional Técnica - SEDUC-SP."
    }


def create_exact_school_document(lesson_data: Dict[str, str]) -> Document:
    """Build exact school template document."""
    doc = Document()

    # Exact School Header matching template.docx
    p_header = doc.add_paragraph()
    p_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p_header.add_run("GOVERNO DE ESTADO DE SÃO PAULO – SECRETÁRIA DA EDUCAÇÃO\n")
    r1.bold = True
    r1.font.size = Pt(10)
    r2 = p_header.add_run("DIRETORIA DE ENSINO-REGIÃO DE ARARAQUARA\nEE. PROF.DINORA MARCODES GOMES\n")
    r2.bold = True
    r2.font.size = Pt(9.5)
    r3 = p_header.add_run("RUA EMILIA GALLI, 549 - CENTRO-AMERICO BRASILIENSE - SP. TELEFONE: (16) 33921335\nCEP: 14.820.015 | E-mail: e021830a@educacao.sp.gov.br\n")
    r3.font.size = Pt(8.5)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run(f"Plano de Aula - 2026 - SEMANA {lesson_data['semana']}\n")
    r_title.bold = True
    r_title.font.size = Pt(11.5)

    items = [
        ("Professor", lesson_data['professor']),
        ("Componente curricular", lesson_data['disciplina']),
        ("3 . Ano/Série/Turma", lesson_data['turma']),
        ("4 . Período de realização", lesson_data['periodo']),
        ("5 . Habilidades", lesson_data['habilidades']),
        ("6 . Objetos de Conhecimento", lesson_data['objetos']),
        ("7 . Conteúdo", lesson_data['conteudo']),
        ("8 . Objetivos", lesson_data['objetivos']),
        ("9 . Estratégias", lesson_data['estrategias']),
        ("1 0 . Recursos didáticos", lesson_data['recursos']),
        ("1 1 . Critérios de Avaliação", lesson_data['avaliacao']),
        ("12 . Referências", lesson_data['referencias']),
    ]

    table = doc.add_table(rows=len(items), cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for idx, (label, val) in enumerate(items):
        cell = table.cell(idx, 0)
        p = cell.paragraphs[0]
        r_lbl = p.add_run(f"{label} : ")
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


@app.route("/api/preview-plan", methods=["POST"])
@app.route("/preview-plan", methods=["POST"])
def preview_plan():
    try:
        teacher = request.form.get("teacher", "").strip()
        period = request.form.get("period", "").strip()
        classroom = request.form.get("classroom", "").strip()
        discipline = request.form.get("discipline", "").strip()
        week = request.form.get("week", "1").strip()
        lesson_style = request.form.get("lessonStyle", "interativa").strip()
        with_inclusion = request.form.get("withInclusion", "false").lower() == "true"

        rows = get_curriculum_rows()
        filtered = filter_rows(rows, week, discipline)
        lesson_data = compile_lesson_data(filtered, teacher, period, classroom, discipline, week, lesson_style, with_inclusion)
        return jsonify({"success": True, "data": lesson_data})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


@app.route("/api/generate", methods=["POST"])
@app.route("/generate", methods=["POST"])
def generate():
    try:
        teacher = request.form.get("teacher", "").strip() or "Professor"
        period = request.form.get("period", "").strip() or "2026"
        classroom = request.form.get("classroom", "").strip() or "Turma MTEC"
        discipline = request.form.get("discipline", "").strip() or "Administração"
        week = request.form.get("week", "1").strip()
        lesson_style = request.form.get("lessonStyle", "interativa").strip()
        with_inclusion = request.form.get("withInclusion", "false").lower() == "true"

        rows = get_curriculum_rows()
        filtered = filter_rows(rows, week, discipline)
        lesson_data = compile_lesson_data(filtered, teacher, period, classroom, discipline, week, lesson_style, with_inclusion)

        doc = create_exact_school_document(lesson_data)

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
