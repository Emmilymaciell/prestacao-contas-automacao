from __future__ import annotations

import argparse
import calendar
import dataclasses
import logging
import os
import re
import shutil
import subprocess
import sys
import tempfile
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime, date
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from pypdf import PdfReader, PdfWriter


# =========================
# CONFIGURAÇÃO PRINCIPAL
# =========================
_AQUI = Path(__file__).resolve().parent
# Pastas configuráveis por variável de ambiente (padrão: ./examples/engenharia e ./saida)
BASE_ENGINEERING_PATH = Path(os.environ.get("PRESTACAO_BASE_DIR", _AQUI / "examples" / "engenharia"))

OUTPUT_BASE_PATH = Path(os.environ.get("PRESTACAO_SAIDA_DIR", _AQUI / "saida"))
OUTPUT_WORD_PATH = OUTPUT_BASE_PATH / "Word"
LOGS_PATH = OUTPUT_BASE_PATH / "_logs"

ALLOWED_LIST_EXTENSIONS = {".pdf", ".doc", ".docx"}
MERGEABLE_PDF_EXTENSION = ".pdf"

# Pastas normalmente irrelevantes para localizar o arquivo principal do relatório.
IGNORED_FOLDER_NAMES = {
    "anexos",
    "arquivos auxiliares",
    "auxiliares",
    "backup",
    "backups",
    "old",
    "antigos",
    "rascunhos",
    "draft",
    "drafts",
}

TEMP_FILE_PREFIXES = ("~$",)
TEMP_FILE_SUFFIXES = (".tmp", ".temp", ".part")

MONTH_NAMES_PT = {
    1: "JAN",
    2: "FEV",
    3: "MAR",
    4: "ABR",
    5: "MAI",
    6: "JUN",
    7: "JUL",
    8: "AGO",
    9: "SET",
    10: "OUT",
    11: "NOV",
    12: "DEZ",
}

MONTH_NAME_TO_NUMBER = {
    "jan": 1,
    "janeiro": 1,
    "fev": 2,
    "fevereiro": 2,
    "mar": 3,
    "marco": 3,
    "março": 3,
    "abr": 4,
    "abril": 4,
    "mai": 5,
    "maio": 5,
    "jun": 6,
    "junho": 6,
    "jul": 7,
    "julho": 7,
    "ago": 8,
    "agosto": 8,
    "set": 9,
    "setembro": 9,
    "sep": 9,
    "sept": 9,
    "out": 10,
    "outubro": 10,
    "oct": 10,
    "nov": 11,
    "novembro": 11,
    "dez": 12,
    "dezembro": 12,
    "dec": 12,
}


@dataclass(frozen=True)
class ContractConfig:
    road: str
    start_year: int
    start_month: int
    monthly_root: Path
    weekly_root: Path
    output_root: Path


CONTRACTS: Dict[str, ContractConfig] = {
    "RD101": ContractConfig(
        road="RD101",
        start_year=2025,
        start_month=9,
        monthly_root=BASE_ENGINEERING_PATH / "RD101" / "3_Gestao Tecnica de Obra" / "5. Relatório Mensal da Obra",
        weekly_root=BASE_ENGINEERING_PATH / "RD101" / "3_Gestao Tecnica de Obra" / "4. Relatório Semanal Obras" / "5. Relatório Semanal de Acompanhamento de Obras",
        output_root=OUTPUT_BASE_PATH / "RD101",
    ),
    "RD102A": ContractConfig(
        road="RD102A",
        start_year=2025,
        start_month=9,
        monthly_root=BASE_ENGINEERING_PATH / "RD102A" / "3_Gestao Tecnica de Obra" / "5. Relatório Mensal da Obra",
        weekly_root=BASE_ENGINEERING_PATH / "RD102A" / "3_Gestao Tecnica de Obra" / "4. Relatório Semanal Obras" / "5. Relatório Semanal de Acompanhamento de Obras",
        output_root=OUTPUT_BASE_PATH / "RD102A",
    ),
    "RD103": ContractConfig(
        road="RD103",
        start_year=2025,
        start_month=9,
        monthly_root=BASE_ENGINEERING_PATH / "RD103" / "3_Gestao Tecnica de Obra" / "5. Relatório Mensal da Obra",
        weekly_root=BASE_ENGINEERING_PATH / "RD103" / "3_Gestao Tecnica de Obra" / "4. Relatório Semanal Obras" / "5. Relatório Semanal de Acompanhamento de Obras",
        output_root=OUTPUT_BASE_PATH / "RD103",
    ),
    "RD104": ContractConfig(
        road="RD104",
        start_year=2025,
        start_month=10,
        monthly_root=BASE_ENGINEERING_PATH / "RD104" / "3_Gestao Tecnica de Obra" / "5. Relatório Mensal da Obra",
        weekly_root=BASE_ENGINEERING_PATH / "RD104" / "3_Gestao Tecnica de Obra" / "4. Relatório Semanal Obras" / "5. Relatório Semanal de Acompanhamento de Obras",
        output_root=OUTPUT_BASE_PATH / "RD104",
    ),
}


@dataclass
class DocumentItem:
    road: str
    source_kind: str  # monthly_root | weekly_root
    doc_type: str  # RM | RO
    path: Path
    extension: str
    modified_at: datetime
    month_key: Tuple[int, int]
    inferred_by: str
    sort_date: datetime
    display_name: str = field(init=False)
    month_label: str = field(init=False)
    quarter_index: int = field(init=False)

    def __post_init__(self) -> None:
        self.display_name = self.path.name
        self.month_label = month_label(self.month_key)
        self.quarter_index = -1


@dataclass(frozen=True)
class MergeSource:
    path: Path
    label: str
    origin: str  # rm | anexo | ro


def setup_logging(log_level: str = "INFO") -> None:
    LOGS_PATH.mkdir(parents=True, exist_ok=True)
    log_file = LOGS_PATH / f"prestacao_contas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

    root_logger = logging.getLogger()
    root_logger.setLevel(getattr(logging, log_level.upper(), logging.INFO))
    root_logger.handlers.clear()

    formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")

    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setFormatter(formatter)
    root_logger.addHandler(file_handler)

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    root_logger.addHandler(console_handler)

    logging.info("Log gravado em: %s", log_file)


# =========================
# UTILIDADES DE DATA
# =========================
def month_key_from_date(dt: date | datetime) -> Tuple[int, int]:
    return (dt.year, dt.month)


def add_months(year: int, month: int, delta: int) -> Tuple[int, int]:
    total = (year * 12 + (month - 1)) + delta
    new_year = total // 12
    new_month = (total % 12) + 1
    return new_year, new_month


def months_between(start_key: Tuple[int, int], end_key: Tuple[int, int]) -> int:
    sy, sm = start_key
    ey, em = end_key
    return (ey - sy) * 12 + (em - sm)


def month_label(month_key: Tuple[int, int]) -> str:
    year, month = month_key
    return f"{MONTH_NAMES_PT[month]} {year}"


def month_folder_label(month_key: Tuple[int, int]) -> str:
    year, month = month_key
    return f"{MONTH_NAMES_PT[month]} - {year}"


def quarter_label(contract: ContractConfig, quarter_index: int) -> str:
    start_key = add_months(contract.start_year, contract.start_month, quarter_index * 3)
    end_key = add_months(contract.start_year, contract.start_month, quarter_index * 3 + 2)

    sy, sm = start_key
    ey, em = end_key

    if sy == ey:
        return f"{MONTH_NAMES_PT[sm]} - {MONTH_NAMES_PT[em]} {sy}"
    return f"{MONTH_NAMES_PT[sm]} {sy} - {MONTH_NAMES_PT[em]} {ey}"


def quarter_index_for_month(contract: ContractConfig, month_key: Tuple[int, int]) -> int:
    offset = months_between((contract.start_year, contract.start_month), month_key)
    if offset < 0:
        raise ValueError(f"Mês {month_key} é anterior ao início do contrato {contract.road}.")
    return offset // 3


# =========================
# EXTRAÇÃO DE PADRÕES
# =========================
def normalize_text(value: str) -> str:
    return value.lower().replace("_", " ").replace("-", " ")


def normalized_file_key(path: Path) -> str:
    key = normalize_text(path.stem)
    key = re.sub(r"[^a-z0-9]+", "", key)
    return key


def is_temp_like_file(path: Path) -> bool:
    name_lower = path.name.lower()
    if any(path.name.startswith(prefix) for prefix in TEMP_FILE_PREFIXES):
        return True
    if name_lower.endswith(TEMP_FILE_SUFFIXES):
        return True
    return False


def is_annex_folder_name(name: str) -> bool:
    normalized = normalize_text(name).strip()
    return "anexo" in normalized


def is_ignored_path(path: Path) -> bool:
    parts_lower = {part.lower().strip() for part in path.parts}
    if parts_lower & IGNORED_FOLDER_NAMES:
        return True
    if is_temp_like_file(path):
        return True
    return False


def looks_like_candidate_report(file_path: Path) -> bool:
    if not file_path.is_file():
        return False
    if file_path.suffix.lower() not in ALLOWED_LIST_EXTENSIONS:
        return False
    if is_ignored_path(file_path):
        return False

    name = file_path.stem.lower()
    return bool(re.search(r"\b(rm|ro)[\s._-]*\d+", name)) or bool(re.search(r"\b(rm|ro)\b", name))


def detect_doc_type_from_name(file_path: Path) -> Optional[str]:
    name = file_path.stem.lower()
    if re.search(r"\brm[\s._-]*\d+", name) or re.search(r"\brm\b", name):
        return "RM"
    if re.search(r"\bro[\s._-]*\d+", name) or re.search(r"\bro\b", name):
        return "RO"
    return None


def detect_doc_type(file_path: Path, root_kind: str) -> Optional[str]:
    by_name = detect_doc_type_from_name(file_path)
    if by_name:
        return by_name

    # Fallback pelo contexto da pasta apenas se o nome não ajudou.
    if root_kind == "monthly_root":
        return "RM"
    if root_kind == "weekly_root":
        return "RO"
    return None


def extract_month_year_from_text(text: str) -> Optional[Tuple[int, int, str]]:
    lower = text.lower()

    # Primeiro tenta datas completas para evitar leituras erradas como 05-09 -> ano 2009.
    full_date_patterns = [
        re.compile(r"\b(0?[1-9]|[12]\d|3[01])[\s._/-](0?[1-9]|1[0-2])[\s._/-]((?:20)?\d{2})\b"),
        re.compile(r"\b((?:20)?\d{2})[\s._/-](0?[1-9]|1[0-2])[\s._/-](0?[1-9]|[12]\d|3[01])\b"),
    ]
    for idx, pattern in enumerate(full_date_patterns):
        match = pattern.search(lower)
        if not match:
            continue
        if idx == 0:
            month = int(match.group(2))
            year = normalize_year(match.group(3))
        else:
            year = normalize_year(match.group(1))
            month = int(match.group(2))
        return year, month, "data_completa_no_texto"

    patterns = [
        re.compile(r"(?<!\d)(0?[1-9]|1[0-2])[\s._/-]+((?:20)?\d{2})(?!\d)"),
        re.compile(r"(?<!\d)((?:20)?\d{2})[\s._/-]+(0?[1-9]|1[0-2])(?!\d)"),
        re.compile(r"\b(jan|janeiro|fev|fevereiro|mar|marco|março|abr|abril|mai|maio|jun|junho|jul|julho|ago|agosto|set|setembro|sep|sept|out|outubro|oct|nov|novembro|dez|dezembro|dec)[\s._/-]+((?:20)?\d{2})\b"),
        re.compile(r"\b((?:20)?\d{2})[\s._/-]+(jan|janeiro|fev|fevereiro|mar|marco|março|abr|abril|mai|maio|jun|junho|jul|julho|ago|agosto|set|setembro|sep|sept|out|outubro|oct|nov|novembro|dez|dezembro|dec)\b"),
    ]

    for idx, pattern in enumerate(patterns):
        match = pattern.search(lower)
        if not match:
            continue

        if idx == 0:
            month = int(match.group(1))
            year = normalize_year(match.group(2))
        elif idx == 1:
            year = normalize_year(match.group(1))
            month = int(match.group(2))
        elif idx == 2:
            month = MONTH_NAME_TO_NUMBER[match.group(1)]
            year = normalize_year(match.group(2))
        else:
            year = normalize_year(match.group(1))
            month = MONTH_NAME_TO_NUMBER[match.group(2)]

        return year, month, "texto_nome_ou_pasta"
    return None


def normalize_year(raw: str) -> int:
    year = int(raw)
    if year < 100:
        return 2000 + year
    return year


def extract_report_sequence(text: str, doc_type: str) -> Optional[int]:
    pattern = re.compile(rf"\b{doc_type.lower()}[\s._-]*(\d{{1,3}})(?:\.(\d+))?\b", re.IGNORECASE)
    match = pattern.search(text)
    if not match:
        return None
    return int(match.group(1))


def infer_month_for_item(
    file_path: Path,
    contract: ContractConfig,
    doc_type: str,
    modified_at: datetime,
) -> Tuple[Tuple[int, int], str, datetime]:
    # 1) Tenta achar mês/ano no nome do arquivo ou em pastas próximas.
    candidate_texts = [file_path.stem, file_path.parent.name]
    if file_path.parent.parent != file_path.parent:
        candidate_texts.append(file_path.parent.parent.name)

    for text in candidate_texts:
        result = extract_month_year_from_text(text)
        if result:
            year, month, source = result
            inferred_date = modified_at.replace(year=year, month=month)
            return (year, month), source, inferred_date

    # 2) Se for RM, tenta pelo sequencial do relatório a partir do mês inicial da obra.
    if doc_type == "RM":
        seq = None
        for text in candidate_texts:
            seq = extract_report_sequence(text, doc_type)
            if seq is not None:
                break
        if seq is not None:
            year, month = add_months(contract.start_year, contract.start_month, seq - 1)
            inferred_date = modified_at.replace(year=year, month=month)
            return (year, month), "sequencial_rm", inferred_date

    # 3) Fallback: usa mês da modificação local do arquivo.
    return month_key_from_date(modified_at), "data_modificacao", modified_at


def extract_best_sort_date(file_path: Path, default_dt: datetime) -> datetime:
    text = f"{file_path.stem} {file_path.parent.name}"

    full_date_patterns = [
        re.compile(r"\b(0?[1-9]|[12]\d|3[01])[\s._/-](0?[1-9]|1[0-2])[\s._/-]((?:20)?\d{2})\b"),
        re.compile(r"\b((?:20)?\d{2})[\s._/-](0?[1-9]|1[0-2])[\s._/-](0?[1-9]|[12]\d|3[01])\b"),
    ]

    for idx, pattern in enumerate(full_date_patterns):
        match = pattern.search(text)
        if not match:
            continue
        try:
            if idx == 0:
                day = int(match.group(1))
                month = int(match.group(2))
                year = normalize_year(match.group(3))
            else:
                year = normalize_year(match.group(1))
                month = int(match.group(2))
                day = int(match.group(3))
            return datetime(year, month, day, default_dt.hour, default_dt.minute, default_dt.second)
        except ValueError:
            pass

    return default_dt


# =========================
# ONE DRIVE / SHAREPOINT
# =========================
def release_onedrive_space(path: Path) -> None:
    """
    Tenta marcar o arquivo novamente como online-only para liberar espaço local.
    Não exclui nada da nuvem. Em máquinas sem OneDrive/SharePoint ou sem esse recurso,
    apenas registra o aviso e segue.
    """
    if os.name != "nt":
        return

    try:
        subprocess.run(
            ["attrib", "+U", "-P", str(path)],
            check=False,
            capture_output=True,
            text=True,
        )
    except Exception as exc:
        logging.debug("Não foi possível liberar espaço do arquivo %s: %s", path, exc)


# =========================
# SCAN
# =========================
def iter_candidate_files(root: Path) -> Iterable[Path]:
    if not root.exists():
        logging.warning("Pasta não encontrada: %s", root)
        return

    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if d.lower() not in IGNORED_FOLDER_NAMES]
        current_dir = Path(dirpath)
        for filename in filenames:
            file_path = current_dir / filename
            try:
                if looks_like_candidate_report(file_path):
                    yield file_path
            except OSError as exc:
                logging.warning("Falha ao inspecionar %s: %s", file_path, exc)


def scan_contract_documents(contract: ContractConfig) -> List[DocumentItem]:
    items: List[DocumentItem] = []
    unique_paths: set[str] = set()

    effective_weekly_root = resolve_weekly_root(contract)

    for root_kind, root_path in (("monthly_root", contract.monthly_root), ("weekly_root", effective_weekly_root)):
        for file_path in iter_candidate_files(root_path):
            normalized_path = str(file_path).lower()
            if normalized_path in unique_paths:
                continue
            unique_paths.add(normalized_path)

            doc_type = detect_doc_type(file_path, root_kind)
            if doc_type is None:
                logging.info("Arquivo ignorado por não identificar RM/RO: %s", file_path)
                continue

            try:
                stat_result = file_path.stat()
                modified_at = datetime.fromtimestamp(stat_result.st_mtime)
            except OSError as exc:
                logging.warning("Não foi possível ler metadados de %s: %s", file_path, exc)
                continue

            month_key, inferred_by, inferred_date = infer_month_for_item(
                file_path=file_path,
                contract=contract,
                doc_type=doc_type,
                modified_at=modified_at,
            )

            sort_date = extract_best_sort_date(file_path, inferred_date)

            item = DocumentItem(
                road=contract.road,
                source_kind=root_kind,
                doc_type=doc_type,
                path=file_path,
                extension=file_path.suffix.lower(),
                modified_at=modified_at,
                month_key=month_key,
                inferred_by=inferred_by,
                sort_date=sort_date,
            )

            try:
                item.quarter_index = quarter_index_for_month(contract, item.month_key)
            except ValueError:
                logging.warning(
                    "Arquivo %s caiu em mês anterior ao início do contrato (%s). Será ignorado.",
                    file_path,
                    contract.road,
                )
                continue

            items.append(item)

    items.sort(key=lambda x: (x.month_key[0], x.month_key[1], x.sort_date, x.doc_type, x.display_name.lower()))
    logging.info("%s: %s arquivos elegíveis encontrados.", contract.road, len(items))
    return items


# =========================
# PDF
# =========================
def safe_filename(value: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*]+', "-", value)
    cleaned = re.sub(r"\s+", " ", cleaned).strip().rstrip(".")
    return cleaned


def merge_pdfs(pdf_sources: List[MergeSource], output_pdf: Path, description: str) -> None:
    if not pdf_sources:
        logging.info("Nenhum PDF para juntar em %s", description)
        return

    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    writer = PdfWriter()
    total_pages = 0

    logging.info("Iniciando junção: %s -> %s arquivo(s)", description, len(pdf_sources))

    for source in pdf_sources:
        if source.path.suffix.lower() != MERGEABLE_PDF_EXTENSION:
            continue

        logging.info("Lendo PDF (%s): %s", source.origin, source.path)
        try:
            with open(source.path, "rb") as fh:
                reader = PdfReader(fh)
                if reader.is_encrypted:
                    try:
                        reader.decrypt("")
                    except Exception:
                        raise RuntimeError("PDF criptografado e não foi possível abrir sem senha.")

                for page in reader.pages:
                    writer.add_page(page)
                    total_pages += 1
        except Exception as exc:
            logging.error("Falha ao incluir PDF %s: %s", source.path, exc)
            continue
        finally:
            release_onedrive_space(source.path)

    if total_pages == 0:
        logging.warning("Nenhuma página foi adicionada ao PDF final de %s", description)
        return

    temp_output = output_pdf.with_suffix(".tmp.pdf")
    with open(temp_output, "wb") as out_fh:
        writer.write(out_fh)

    if output_pdf.exists():
        output_pdf.unlink()
    temp_output.replace(output_pdf)
    logging.info("PDF criado: %s (%s páginas)", output_pdf, total_pages)


# =========================
# WORD
# =========================
def set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def add_document_header(doc: Document, title: str, subtitle: str) -> None:
    section = doc.sections[0]
    section.top_margin = Pt(42)
    section.bottom_margin = Pt(42)
    section.left_margin = Pt(50)
    section.right_margin = Pt(50)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(title)
    run.bold = True
    run.font.size = Pt(18)

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = p2.add_run(subtitle)
    run2.italic = True
    run2.font.size = Pt(10)

    doc.add_paragraph()


def style_normal(doc: Document) -> None:
    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(10.5)


def add_road_section(doc: Document, road: str, items: List[DocumentItem]) -> None:
    heading = doc.add_paragraph()
    heading.style = doc.styles["Heading 1"]
    heading_run = heading.add_run(road)
    heading_run.bold = True

    if not items:
        p = doc.add_paragraph("Nenhum relatório RM/RO localizado para esta rodovia neste mês.")
        p.style = doc.styles["Normal"]
        return

    grouped: Dict[str, List[DocumentItem]] = defaultdict(list)
    for item in items:
        grouped[item.doc_type].append(item)

    for doc_type in ("RO", "RM"):
        subset = grouped.get(doc_type, [])
        subheading = doc.add_paragraph()
        subheading.style = doc.styles["Heading 2"]
        subheading.add_run(f"{doc_type} ({len(subset)})")

        if not subset:
            doc.add_paragraph("Nenhum arquivo encontrado.")
            continue

        table = doc.add_table(rows=1, cols=4)
        table.style = "Table Grid"
        hdr = table.rows[0].cells
        hdr[0].text = "Data"
        hdr[1].text = "Arquivo"
        hdr[2].text = "Extensão"
        hdr[3].text = "Origem"

        for cell in hdr:
            set_cell_shading(cell, "D9EAF7")
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.bold = True

        subset.sort(key=lambda x: (x.sort_date, x.display_name.lower()))
        for item in subset:
            row = table.add_row().cells
            row[0].text = item.sort_date.strftime("%d/%m/%Y")
            row[1].text = item.display_name
            row[2].text = item.extension.upper().replace(".", "")
            row[3].text = "Mensal" if item.source_kind == "monthly_root" else "Semanal"

        doc.add_paragraph()


def create_monthly_word(month_key: Tuple[int, int], items_by_road: Dict[str, List[DocumentItem]], output_docx: Path) -> None:
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    doc = Document()
    style_normal(doc)

    add_document_header(
        doc,
        title=f"Prestação de contas - {month_label(month_key)}",
        subtitle="Relação consolidada de relatórios entregues por rodovia (RO e RM)",
    )

    generated = doc.add_paragraph()
    generated.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    generated.add_run(f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}")

    intro = doc.add_paragraph()
    intro.add_run("Observação: ").bold = True
    intro.add_run(
        "foram listados os arquivos principais de relatórios encontrados nas pastas mensal e semanal, "
        "agrupados pelo mês de competência identificado no nome, na pasta ou, em último caso, pela data de modificação."
    )

    doc.add_paragraph()

    for idx, road in enumerate(sorted(CONTRACTS.keys())):
        if idx > 0:
            doc.add_page_break()
        add_road_section(doc, road, items_by_road.get(road, []))

    temp_output = output_docx.with_suffix(".tmp.docx")
    doc.save(temp_output)
    if output_docx.exists():
        output_docx.unlink()
    temp_output.replace(output_docx)
    logging.info("Word mensal criado: %s", output_docx)


# =========================
# AGRUPAMENTO
# =========================
def group_by_month(items: Iterable[DocumentItem]) -> Dict[Tuple[int, int], List[DocumentItem]]:
    grouped: Dict[Tuple[int, int], List[DocumentItem]] = defaultdict(list)
    for item in items:
        grouped[item.month_key].append(item)
    for key in grouped:
        grouped[key].sort(key=lambda x: (x.sort_date, x.doc_type, x.display_name.lower()))
    return dict(grouped)


def group_by_quarter(items: Iterable[DocumentItem]) -> Dict[int, List[DocumentItem]]:
    grouped: Dict[int, List[DocumentItem]] = defaultdict(list)
    for item in items:
        grouped[item.quarter_index].append(item)
    for key in grouped:
        grouped[key].sort(key=lambda x: (x.month_key, x.sort_date, x.doc_type, x.display_name.lower()))
    return dict(grouped)


# =========================
# ANEXOS E PACOTES MENSAIS
# =========================
def candidate_annex_dirs_for_rm(rm_path: Path) -> List[Path]:
    candidates: List[Path] = []
    month_dir = rm_path.parent

    try:
        for child in month_dir.iterdir():
            if child.is_dir() and is_annex_folder_name(child.name):
                candidates.append(child)
    except OSError as exc:
        logging.warning("Não foi possível inspecionar anexos de %s: %s", rm_path, exc)

    candidates.sort(key=lambda p: p.name.lower())
    return candidates


def iter_annex_pdfs(root: Path) -> Iterable[Path]:
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if not d.startswith('~$')]
        current_dir = Path(dirpath)
        for filename in sorted(filenames):
            file_path = current_dir / filename
            if is_temp_like_file(file_path):
                continue
            if file_path.suffix.lower() != MERGEABLE_PDF_EXTENSION:
                continue
            yield file_path


def sort_key_for_path(path: Path) -> Tuple[datetime, str]:
    try:
        modified_at = datetime.fromtimestamp(path.stat().st_mtime)
    except OSError:
        modified_at = datetime.now()
    return (extract_best_sort_date(path, modified_at), path.name.lower())


def looks_like_ro_pdf(path: Path) -> bool:
    return detect_doc_type_from_name(path) == "RO"


def merge_source_from_path(path: Path, origin: str) -> MergeSource:
    return MergeSource(path=path, label=path.name, origin=origin)


def collect_annex_sources_for_rm(rm_path: Path) -> Tuple[List[MergeSource], List[MergeSource]]:
    non_ro_sources: List[MergeSource] = []
    ro_sources: List[MergeSource] = []
    seen: set[str] = set()

    for annex_dir in candidate_annex_dirs_for_rm(rm_path):
        for pdf_path in sorted(iter_annex_pdfs(annex_dir), key=sort_key_for_path):
            key = str(pdf_path).lower()
            if key in seen:
                continue
            seen.add(key)

            source = merge_source_from_path(pdf_path, "ro" if looks_like_ro_pdf(pdf_path) else "anexo")
            if source.origin == "ro":
                ro_sources.append(source)
            else:
                non_ro_sources.append(source)

    return non_ro_sources, ro_sources


def expand_month_pdf_sources(month_items: List[DocumentItem]) -> List[MergeSource]:
    month_items = [item for item in month_items if item.extension == MERGEABLE_PDF_EXTENSION]
    month_items.sort(key=lambda x: (x.sort_date, x.doc_type, x.display_name.lower()))

    rms = [item for item in month_items if item.doc_type == "RM"]
    ros_by_date = [item for item in month_items if item.doc_type == "RO"]

    if not rms:
        return [merge_source_from_path(item.path, "ro") for item in ros_by_date]

    expanded: List[MergeSource] = []
    fallback_ros = [merge_source_from_path(item.path, "ro") for item in ros_by_date]
    fallback_appended = False

    for idx, rm in enumerate(rms):
        expanded.append(merge_source_from_path(rm.path, "rm"))

        annex_non_ro, annex_ro = collect_annex_sources_for_rm(rm.path)
        expanded.extend(annex_non_ro)

        if annex_ro:
            expanded.extend(annex_ro)
            logging.info(
                "%s | %s: ROs definidos pelos anexos do RM %s (%s arquivo(s)).",
                rm.road,
                rm.month_label,
                rm.display_name,
                len(annex_ro),
            )
            continue

        if not fallback_ros or fallback_appended:
            continue

        if len(rms) == 1:
            expanded.extend(fallback_ros)
            fallback_appended = True
            logging.info(
                "%s | %s: sem RO em anexo no RM %s, usando fallback por data (%s arquivo(s)).",
                rm.road,
                rm.month_label,
                rm.display_name,
                len(fallback_ros),
            )
        elif idx == len(rms) - 1:
            expanded.extend(fallback_ros)
            fallback_appended = True
            logging.warning(
                "%s | %s: existem múltiplos RMs sem RO em anexo; fallback por data foi anexado uma única vez ao final do último RM.",
                rm.road,
                rm.month_label,
            )

    return expanded


def resolve_weekly_root(contract: ContractConfig) -> Path:
    if contract.weekly_root.exists():
        return contract.weekly_root

    base = BASE_ENGINEERING_PATH / contract.road / "3_Gestao Tecnica de Obra" / "4. Relatório Semanal Obras"
    if not base.exists():
        return contract.weekly_root

    candidates: List[Path] = []
    for dirpath, dirnames, _filenames in os.walk(base):
        current = Path(dirpath)
        normalized = normalize_text(current.name)
        if "relatorio semanal" in normalized and "obra" in normalized:
            candidates.append(current)

    if not candidates:
        return contract.weekly_root

    candidates.sort(key=lambda p: (-len(p.parts), p.as_posix().lower()))
    resolved = candidates[0]
    logging.info("%s: pasta semanal alternativa localizada automaticamente: %s", contract.road, resolved)
    return resolved


# =========================
# PIPELINE PRINCIPAL
# =========================
def build_outputs_for_contract(contract: ContractConfig, items: List[DocumentItem]) -> None:
    contract.output_root.mkdir(parents=True, exist_ok=True)

    pdf_items = [i for i in items if i.extension == MERGEABLE_PDF_EXTENSION]
    monthly_groups = group_by_month(pdf_items)
    expanded_month_sources: Dict[Tuple[int, int], List[MergeSource]] = {}

    # PDFs mensais por rodovia, incluindo anexos do RM logo após o RM correspondente.
    for month_key in sorted(monthly_groups.keys()):
        month_items = monthly_groups[month_key]
        expanded_sources = expand_month_pdf_sources(month_items)
        expanded_month_sources[month_key] = expanded_sources

        quarter_idx = quarter_index_for_month(contract, month_key)
        quarter_dir = contract.output_root / safe_filename(quarter_label(contract, quarter_idx))
        month_dir = quarter_dir / safe_filename(month_folder_label(month_key))
        output_pdf = month_dir / safe_filename(f"Prestação de contas - {contract.road} - {month_label(month_key)}.pdf")
        merge_pdfs(
            pdf_sources=expanded_sources,
            output_pdf=output_pdf,
            description=f"{contract.road} | {month_label(month_key)}",
        )

    # PDFs trimestrais por rodovia, reaproveitando a mesma ordem dos meses e incluindo anexos.
    quarter_to_sources: Dict[int, List[MergeSource]] = defaultdict(list)
    for month_key in sorted(expanded_month_sources.keys()):
        quarter_idx = quarter_index_for_month(contract, month_key)
        quarter_to_sources[quarter_idx].extend(expanded_month_sources[month_key])

    for quarter_idx in sorted(quarter_to_sources.keys()):
        quarter_dir = contract.output_root / safe_filename(quarter_label(contract, quarter_idx))
        output_pdf = quarter_dir / safe_filename(
            f"Prestação de contas trimestral - {contract.road} - {quarter_label(contract, quarter_idx)}.pdf"
        )
        merge_pdfs(
            pdf_sources=quarter_to_sources[quarter_idx],
            output_pdf=output_pdf,
            description=f"{contract.road} | {quarter_label(contract, quarter_idx)}",
        )


def build_monthly_words(all_items: Dict[str, List[DocumentItem]]) -> None:
    OUTPUT_WORD_PATH.mkdir(parents=True, exist_ok=True)

    month_index: Dict[Tuple[int, int], Dict[str, List[DocumentItem]]] = defaultdict(lambda: defaultdict(list))
    for road, items in all_items.items():
        for item in items:
            month_index[item.month_key][road].append(item)

    for month_key in sorted(month_index.keys()):
        per_road = month_index[month_key]
        output_docx = OUTPUT_WORD_PATH / safe_filename(f"Prestação de contas - {month_label(month_key)}.docx")
        create_monthly_word(month_key, per_road, output_docx)


def validate_paths(selected_contracts: List[ContractConfig]) -> None:
    OUTPUT_BASE_PATH.mkdir(parents=True, exist_ok=True)
    OUTPUT_WORD_PATH.mkdir(parents=True, exist_ok=True)
    for contract in selected_contracts:
        if not contract.monthly_root.exists():
            logging.warning("Caminho mensal não encontrado para %s: %s", contract.road, contract.monthly_root)
        resolved_weekly = resolve_weekly_root(contract)
        if not resolved_weekly.exists():
            logging.warning("Caminho semanal não encontrado para %s: %s", contract.road, contract.weekly_root)


def run(selected_roads: Optional[List[str]] = None) -> None:
    if selected_roads:
        selected_contracts = [CONTRACTS[road] for road in selected_roads]
    else:
        selected_contracts = [CONTRACTS[key] for key in sorted(CONTRACTS.keys())]

    validate_paths(selected_contracts)

    all_items: Dict[str, List[DocumentItem]] = {}
    for contract in selected_contracts:
        items = scan_contract_documents(contract)
        all_items[contract.road] = items
        build_outputs_for_contract(contract, items)

    build_monthly_words(all_items)

    logging.info("Processamento concluído com sucesso.")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Gera Word consolidado por mês e PDFs mensal/trimestral por rodovia com relatórios RM e RO."
    )
    parser.add_argument(
        "--roads",
        nargs="*",
        choices=sorted(CONTRACTS.keys()),
        help="Processa apenas as rodovias informadas. Se omitido, processa todas.",
    )
    parser.add_argument(
        "--log-level",
        default="INFO",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        help="Nível de log.",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    setup_logging(args.log_level)
    try:
        run(args.roads)
    except KeyboardInterrupt:
        logging.error("Execução interrompida pelo usuário.")
        raise
    except Exception:
        logging.exception("Erro fatal na execução.")
        raise
