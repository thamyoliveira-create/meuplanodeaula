#!/usr/bin/env python3
"""Build the compact regular-education curriculum used by the web app.

Input files are plain text extracted from the official SEDUC-SP 2026 guides.
The generated JSON is the only artifact shipped to the browser.
"""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path


SOURCES = {
    "ef1_portugues.txt": ("fundamental_1", "Língua Portuguesa"),
    "ef1_matematica.txt": ("fundamental_1", "Matemática"),
    "ef2_portugues.txt": ("fundamental_2", "Língua Portuguesa"),
    "ef2_matematica.txt": ("fundamental_2", "Matemática"),
    "ef2_ciencias.txt": ("fundamental_2", "Ciências"),
    "ef2_historia.txt": ("fundamental_2", "História"),
    "ef2_geografia.txt": ("fundamental_2", "Geografia"),
    "ef2_ingles.txt": ("fundamental_2", "Língua Inglesa"),
    "ef2_edfisica.txt": ("fundamental_2", "Educação Física"),
    "ef2_arte.txt": ("fundamental_2", "Arte"),
    "em_portugues.txt": ("ensino_medio", "Língua Portuguesa"),
    "em_matematica.txt": ("ensino_medio", "Matemática"),
    "em_biologia.txt": ("ensino_medio", "Biologia"),
    "em_fisica.txt": ("ensino_medio", "Física"),
    "em_quimica.txt": ("ensino_medio", "Química"),
    "em_historia.txt": ("ensino_medio", "História"),
    "em_geografia.txt": ("ensino_medio", "Geografia"),
    "em_filosofia.txt": ("ensino_medio", "Filosofia"),
    "em_sociologia.txt": ("ensino_medio", "Sociologia"),
    "em_ingles.txt": ("ensino_medio", "Língua Inglesa"),
    "em_edfisica.txt": ("ensino_medio", "Educação Física"),
    "em_arte.txt": ("ensino_medio", "Arte"),
}

HEADER_RE = re.compile(
    r"Aula\s+Conteúdo\s+Objetivos\s+de\s+aprendizagem\s+Habilidades\s+Aprendizagem\s+Essencial",
    re.IGNORECASE,
)
GRADE_RE = re.compile(
    r"(?P<grade>[1-9])[ºª]\s*(?:Ano|Série)\s*[–—-]*\s*(?P<term>[1-4])[ºª]\s*Bimestre",
    re.IGNORECASE,
)
GRADE_ONLY_RE = re.compile(r"(?P<grade>[1-9])[ºª]\s*(?:Ano|Série)", re.IGNORECASE)
ROW_RE = re.compile(r"^\s*(?P<number>\d{1,2})(?:\s+(?P<rest>.*\S))?\s*$")
SKILL_RE = re.compile(r"\b(?:EF\d{2}[A-Z]{2}\d{2}[A-Z]?|EM13[A-Z]{3}\d{3}|EMIF[A-Z]{2,}\d{2})\b")
AE_RE = re.compile(r"\bAE\d+\s*[-–:]\s*", re.IGNORECASE)

OBJECTIVE_VERBS = (
    "analisar", "aplicar", "apreciar", "argumentar", "avaliar", "calcular",
    "caracterizar", "classificar", "comparar", "compreender", "conhecer",
    "construir", "criar", "descrever", "desenvolver", "diferenciar",
    "discutir", "elaborar", "estabelecer", "experimentar", "explorar",
    "expressar", "identificar", "inferir", "interpretar", "investigar",
    "justificar", "ler", "localizar", "modelar", "nomear", "observar",
    "ordenar", "planejar", "produzir", "reconhecer", "registrar", "relacionar",
    "representar", "resolver", "retomar", "selecionar", "utilizar", "verificar",
)


def clean(value: str) -> str:
    value = value.replace("\u0002", "-").replace("\uf0b7", "•")
    value = re.sub(r"\s+", " ", value).strip(" •-\t")
    return value


def normalize_text(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("Escopo-SequênciaEscopo", "Escopo-Sequência\nEscopo")
    text = re.sub(r"(?<!\n)(Escopo\s*-\s*Sequência)", r"\n\1", text)
    return text


def nearest_grade(text: str, header_start: int) -> tuple[int, int] | None:
    context_start = max(0, header_start - 350)
    context_end = min(len(text), header_start + 350)
    context = text[context_start:context_end]
    relative_header = header_start - context_start

    def distance(match: re.Match[str]) -> int:
        if match.end() <= relative_header:
            return relative_header - match.end()
        return match.start() - relative_header

    matches = list(GRADE_RE.finditer(context))
    if matches:
        match = min(matches, key=distance)
        return int(match.group("grade")), int(match.group("term"))
    grade_matches = list(GRADE_ONLY_RE.finditer(context))
    if not grade_matches:
        return None
    match = min(grade_matches, key=distance)
    return int(match.group("grade")), 0


def is_objective(item: str) -> bool:
    first = clean(item).lower().split(" ", 1)[0]
    return any(first.startswith(verb) for verb in OBJECTIVE_VERBS)


def parse_row(number: int, lines: list[str]) -> dict | None:
    normalized = [line.rstrip() for line in lines if line.strip()]
    if not normalized:
        return None

    first_marker = next(
        (i for i, line in enumerate(normalized) if line.lstrip().startswith(("•", "-"))),
        len(normalized),
    )
    body = "\n".join(normalized[first_marker:]) if first_marker < len(normalized) else "\n".join(normalized)
    skills = list(dict.fromkeys(SKILL_RE.findall(body)))

    ae_match = AE_RE.search(body)
    essential = clean(body[ae_match.start():]) if ae_match else ""
    if essential:
        essential = re.split(r"\n\s*(?:Escopo|Aula\s+Conteúdo)", essential, maxsplit=1)[0]
        essential = clean(essential)

    if first_marker < len(normalized):
        title = clean(" ".join(normalized[:first_marker]))
        before_ae = body[:ae_match.start()] if ae_match else body
        before_ae = SKILL_RE.sub("", before_ae)
        bullet_parts = re.split(r"(?:^|\n)\s*[•-]\s*", before_ae)
        items = [clean(part) for part in bullet_parts if clean(part)]
        objective_index = next((i for i, item in enumerate(items) if is_objective(item)), len(items))
        contents = items[:objective_index]
        objectives = items[objective_index:]
    else:
        # Some guides use plain table cells rather than bullets. Recover the
        # title and columns from line wrapping so the title does not absorb the
        # entire row.
        column_end = next(
            (i for i, line in enumerate(normalized) if SKILL_RE.search(line) or AE_RE.search(line)),
            len(normalized),
        )
        columns = normalized[:column_end]
        objective_index = next(
            (i for i, line in enumerate(columns) if is_objective(line)),
            len(columns),
        )

        title_end = 0
        title_parts: list[str] = []
        for i, line in enumerate(columns[:objective_index]):
            candidate = clean(" ".join(title_parts + [line]))
            if i > 0 and title_parts and (
                len(candidate) > 80
                or re.search(r"[.;!?]$", title_parts[-1].strip())
                or (re.search(r"[.;]$", line.strip()) and len(clean(" ".join(title_parts))) >= 20)
            ):
                break
            title_parts.append(line)
            title_end = i + 1

        if not title_parts and columns:
            title_parts = [columns[0]]
            title_end = 1
        title = clean(" ".join(title_parts))
        content_text = clean(" ".join(columns[title_end:objective_index]))
        contents = [content_text] if content_text else []

        objectives = []
        for line in columns[objective_index:]:
            value = clean(line)
            if not value:
                continue
            if is_objective(value) or not objectives:
                objectives.append(value)
            else:
                objectives[-1] = clean(f"{objectives[-1]} {value}")

    if not title:
        title = f"Aula {number}"

    if not objectives and essential:
        objectives = [re.sub(r"^AE\d+\s*[-–:]\s*", "", essential, flags=re.IGNORECASE)]

    return {
        "aula": number,
        "titulo": title,
        "conteudos": contents,
        "objetivos": objectives,
        "habilidades": skills,
        "aprendizagem_essencial": essential,
    }


def parse_block(block: str) -> list[dict]:
    rows: list[dict] = []
    current_number: int | None = None
    current_lines: list[str] = []

    def flush() -> None:
        nonlocal current_number, current_lines
        if current_number is not None:
            row = parse_row(current_number, current_lines)
            if row:
                rows.append(row)
        current_number = None
        current_lines = []

    for raw_line in block.splitlines():
        line = raw_line.strip()
        if not line or line.startswith("Escopo-Sequência") or line.startswith("Escopo - Sequência"):
            continue
        match = ROW_RE.match(line)
        if match:
            number = int(match.group("number"))
            rest = match.group("rest") or ""
            if 1 <= number <= 40:
                flush()
                current_number = number
                if rest:
                    current_lines.append(rest)
                continue
        if current_number is not None:
            current_lines.append(line)
    flush()
    return rows


def parse_source(path: Path, stage: str, subject: str) -> list[dict]:
    text = normalize_text(path.read_text(encoding="utf-8", errors="replace"))
    headers = list(HEADER_RE.finditer(text))
    parsed: list[dict] = []
    implicit_terms: dict[int, int] = {}
    last_lesson: dict[int, int] = {}

    for index, header in enumerate(headers):
        grade_term = nearest_grade(text, header.start())
        if not grade_term:
            continue
        grade, term = grade_term
        end = headers[index + 1].start() if index + 1 < len(headers) else len(text)
        block = text[header.end():end]
        # The guides append assessment matrices after the lesson sequence. In
        # extracted text those matrices also contain numbered lines, which can
        # otherwise be mistaken for lessons in the final bimestre.
        block = re.split(r"\n\s*Matriz\s+Prova\s+Paulista\b", block, maxsplit=1, flags=re.IGNORECASE)[0]
        block_rows = parse_block(block)
        if term == 0 and block_rows:
            current_term = implicit_terms.get(grade, 1)
            first_lesson = min(row["aula"] for row in block_rows)
            previous_lesson = last_lesson.get(grade, 0)
            if previous_lesson and first_lesson <= previous_lesson:
                current_term = min(4, current_term + 1)
            implicit_terms[grade] = current_term
            last_lesson[grade] = max(row["aula"] for row in block_rows)
            term = current_term
        elif term:
            implicit_terms[grade] = term
            if block_rows:
                last_lesson[grade] = max(row["aula"] for row in block_rows)

        for row in block_rows:
            row.update({
                "etapa": stage,
                "ano": grade,
                "disciplina": subject,
                "bimestre": term,
                "fonte": "Guia de Aprendizagens Essenciais SEDUC-SP 2026",
            })
            parsed.append(row)

    # Final fallback for unusual text layers without repeated page headers.
    maxima: dict[int, int] = {}
    for row in parsed:
        maxima[row["ano"]] = max(maxima.get(row["ano"], 0), row["aula"])
    for row in parsed:
        if row["bimestre"] == 0:
            per_term = max(1, (maxima[row["ano"]] + 3) // 4)
            row["bimestre"] = min(4, ((row["aula"] - 1) // per_term) + 1)

    # Page headers repeat, so deduplicate the same lesson in the same term.
    unique: dict[tuple, dict] = {}
    for row in parsed:
        key = (row["etapa"], row["ano"], row["disciplina"], row["bimestre"], row["aula"])
        old = unique.get(key)
        if old is None or len(json.dumps(row, ensure_ascii=False)) > len(json.dumps(old, ensure_ascii=False)):
            unique[key] = row
    return list(unique.values())


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("input_dir", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()

    rows: list[dict] = []
    missing: list[str] = []
    for filename, (stage, subject) in SOURCES.items():
        path = args.input_dir / filename
        if not path.exists():
            missing.append(filename)
            continue
        rows.extend(parse_source(path, stage, subject))

    if missing:
        raise SystemExit(f"Missing source files: {', '.join(missing)}")

    rows.sort(key=lambda row: (
        row["etapa"], row["ano"], row["disciplina"], row["bimestre"], row["aula"]
    ))
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(rows, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")

    summary: dict[str, int] = {}
    for row in rows:
        key = f'{row["etapa"]} | {row["ano"]} | {row["disciplina"]}'
        summary[key] = summary.get(key, 0) + 1
    print(json.dumps({"rows": len(rows), "groups": summary}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
