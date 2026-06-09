#!/usr/bin/env python3
"""script_dp1268"""
from __future__ import annotations

import hashlib
import json
import re
import shutil
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt

BASE = Path(__file__).resolve().parent.parent
PRODUTOS = BASE / "produtos_finais"
WORKBOOK_SRC = PRODUTOS / "WORKBOOK_DP_1268_DICIONARIO_CONSOLIDADO.xlsx"
README_MD = PRODUTOS / "README_ENTREGA_DP_1268.md"
REPORT_MD = PRODUTOS / "RELATORIO_FINAL_CONSOLIDADO_DP1268.md"
README_TXT = PRODUTOS / "README_ENTREGA_DP_1268.txt"
REPORT_DOCX = PRODUTOS / "RELATORIO_FINAL_CONSOLIDADO_DP1268.docx"
ZIP_PATH = PRODUTOS / "DP_1268.zip"
VALIDACAO_MD = PRODUTOS / "VALIDACAO_EMPACOTAMENTO_FINAL_DP1268.md"
VALIDACAO_JSON = PRODUTOS / "VALIDACAO_EMPACOTAMENTO_FINAL_DP1268.json"

EXPECTED_FILES = {
    "WORKBOOK_DP_1268_DICIONARIO_CONSOLIDADO.xlsx",
    "README_ENTREGA_DP_1268.txt",
    "RELATORIO_FINAL_CONSOLIDADO_DP1268.docx",
}


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest().upper()


def add_formatted_run(paragraph, text: str) -> None:
    """Parse **bold**, *italic*, `code` inline markers."""
    pattern = re.compile(r"(\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`)")
    pos = 0
    for m in pattern.finditer(text):
        if m.start() > pos:
            paragraph.add_run(text[pos : m.start()])
        token = m.group(0)
        if token.startswith("**"):
            run = paragraph.add_run(token[2:-2])
            run.bold = True
        elif token.startswith("*"):
            run = paragraph.add_run(token[1:-1])
            run.italic = True
        elif token.startswith("`"):
            run = paragraph.add_run(token[1:-1])
            run.font.name = "Consolas"
            run._element.rPr.rFonts.set(qn("w:eastAsia"), "Consolas")
        pos = m.end()
    if pos < len(text):
        paragraph.add_run(text[pos:])


def set_cell_shading(cell, fill: str = "D9E2F3") -> None:
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    shd.set(qn("w:val"), "clear")
    tc_pr.append(shd)


def parse_table_row(line: str) -> list[str]:
    line = line.strip()
    if line.startswith("|"):
        line = line[1:]
    if line.endswith("|"):
        line = line[:-1]
    return [c.strip() for c in line.split("|")]


def is_table_separator(line: str) -> bool:
    cells = parse_table_row(line)
    return bool(cells) and all(re.fullmatch(r":?-{3,}:?", c or "---") for c in cells)


def heading_level(line: str) -> int | None:
    m = re.match(r"^(#{1,6})\s+(.*)$", line)
    if m:
        return len(m.group(1)), m.group(2).strip()
    return None


def md_to_docx(md_path: Path, docx_path: Path) -> None:
    lines = md_path.read_text(encoding="utf-8").splitlines()
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    i = 0
    in_ul = False
    in_ol = False

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if stripped == "---":
            p = doc.add_paragraph()
            p.add_run("—" * 40)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            in_ul = in_ol = False
            i += 1
            continue

        if stripped.startswith("|") and i + 1 < len(lines) and is_table_separator(lines[i + 1]):
            headers = parse_table_row(line)
            i += 2
            rows: list[list[str]] = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                rows.append(parse_table_row(lines[i]))
                i += 1
            ncols = len(headers)
            table = doc.add_table(rows=1 + len(rows), cols=ncols)
            table.style = "Table Grid"
            for j, h in enumerate(headers):
                cell = table.rows[0].cells[j]
                cell.text = ""
                p = cell.paragraphs[0]
                add_formatted_run(p, h)
                for run in p.runs:
                    run.bold = True
                set_cell_shading(cell)
            for r_idx, row in enumerate(rows):
                for c_idx in range(ncols):
                    val = row[c_idx] if c_idx < len(row) else ""
                    cell = table.rows[r_idx + 1].cells[c_idx]
                    cell.text = ""
                    add_formatted_run(cell.paragraphs[0], val)
            doc.add_paragraph()
            in_ul = in_ol = False
            continue

        hl = heading_level(line)
        if hl:
            level, text = hl
            doc.add_heading(text, level=min(level, 4))
            in_ul = in_ol = False
            i += 1
            continue

        if re.match(r"^[-*]\s+", stripped):
            text = re.sub(r"^[-*]\s+", "", stripped)
            text = re.sub(r"^\[x\]\s*", "☑ ", text)
            text = re.sub(r"^\[ \]\s*", "☐ ", text)
            p = doc.add_paragraph(style="List Bullet")
            add_formatted_run(p, text)
            in_ul = True
            in_ol = False
            i += 1
            continue

        if re.match(r"^\d+\.\s+", stripped):
            text = re.sub(r"^\d+\.\s+", "", stripped)
            p = doc.add_paragraph(style="List Number")
            add_formatted_run(p, text)
            in_ol = True
            in_ul = False
            i += 1
            continue

        if stripped == "":
            in_ul = in_ol = False
            i += 1
            continue

        p = doc.add_paragraph()
        add_formatted_run(p, stripped)
        in_ul = in_ol = False
        i += 1

    doc.save(docx_path)


def files_identical(a: Path, b: Path) -> bool:
    return sha256_file(a) == sha256_file(b)


def text_content_equal(a: Path, b: Path) -> bool:
    return a.read_text(encoding="utf-8") == b.read_text(encoding="utf-8")


def main() -> int:
    timestamp = datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")
    checks: dict[str, str] = {}

    source_paths = {
        "workbook": WORKBOOK_SRC,
        "readme_md": README_MD,
        "report_md": REPORT_MD,
    }

    hashes_before = {k: sha256_file(p) for k, p in source_paths.items()}

    readme_md_bytes = README_MD.read_bytes()
    README_TXT.write_bytes(readme_md_bytes)

    md_to_docx(REPORT_MD, REPORT_DOCX)

    hashes_after_sources = {k: sha256_file(p) for k, p in source_paths.items()}

    source_workbook_unchanged = hashes_before["workbook"] == hashes_after_sources["workbook"]
    source_readme_md_unchanged = hashes_before["readme_md"] == hashes_after_sources["readme_md"]
    source_report_md_unchanged = hashes_before["report_md"] == hashes_after_sources["report_md"]

    checks["source_workbook_unchanged"] = "PASS" if source_workbook_unchanged else "FAIL"
    checks["source_readme_md_unchanged"] = "PASS" if source_readme_md_unchanged else "FAIL"
    checks["source_report_md_unchanged"] = "PASS" if source_report_md_unchanged else "FAIL"

    hashes_generated = {
        "readme_txt": sha256_file(README_TXT),
        "report_docx": sha256_file(REPORT_DOCX),
    }

    checks["readme_txt_matches_md"] = (
        "PASS" if text_content_equal(README_TXT, README_MD) else "FAIL"
    )

    report_text = REPORT_MD.read_text(encoding="utf-8")
    has_item_13 = "## 13. Próximos passos recomendados" in report_text
    docx_exists = REPORT_DOCX.exists() and REPORT_DOCX.stat().st_size > 0
    checks["docx_created_from_current_md"] = (
        "PASS" if (docx_exists and has_item_13) else "FAIL"
    )

    with tempfile.TemporaryDirectory(prefix="dp1268_pack_") as tmpdir:
        pack_root = Path(tmpdir) / "DP_1268"
        pack_root.mkdir()
        shutil.copy2(WORKBOOK_SRC, pack_root / "WORKBOOK_DP_1268_DICIONARIO_CONSOLIDADO.xlsx")
        shutil.copy2(README_TXT, pack_root / "README_ENTREGA_DP_1268.txt")
        shutil.copy2(REPORT_DOCX, pack_root / "RELATORIO_FINAL_CONSOLIDADO_DP1268.docx")

        pack_files = sorted(p.name for p in pack_root.iterdir() if p.is_file())
        checks["pack_folder_file_count"] = "PASS" if len(pack_files) == 3 else "FAIL"

        if ZIP_PATH.exists():
            ZIP_PATH.unlink()
        with zipfile.ZipFile(ZIP_PATH, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for f in pack_root.iterdir():
                if f.is_file():
                    arcname = f"DP_1268/{f.name}"
                    zf.write(f, arcname)

        hashes_generated["zip"] = sha256_file(ZIP_PATH)

        extract_dir = Path(tmpdir) / "extract_validate"
        extract_dir.mkdir()
        with zipfile.ZipFile(ZIP_PATH, "r") as zf:
            zf.extractall(extract_dir)
            zip_listing = sorted(zf.namelist())

        root_entries = list(extract_dir.iterdir())
        zip_root_ok = len(root_entries) == 1 and root_entries[0].name == "DP_1268"
        checks["zip_root_ok"] = "PASS" if zip_root_ok else "FAIL"

        inner = root_entries[0] if zip_root_ok else None
        inner_files = sorted(p.name for p in inner.iterdir() if p.is_file()) if inner else []
        zip_file_count = len(inner_files)
        checks["zip_file_count"] = "PASS" if zip_file_count == 3 else "FAIL"

        zip_has_only_expected = set(inner_files) == EXPECTED_FILES
        checks["zip_has_only_expected_files"] = "PASS" if zip_has_only_expected else "FAIL"

        wb_in_zip = inner / "WORKBOOK_DP_1268_DICIONARIO_CONSOLIDADO.xlsx"
        readme_in_zip = inner / "README_ENTREGA_DP_1268.txt"
        docx_in_zip = inner / "RELATORIO_FINAL_CONSOLIDADO_DP1268.docx"

        workbook_inside_zip_matches_source = files_identical(wb_in_zip, WORKBOOK_SRC)
        readme_inside_zip_matches_txt = files_identical(readme_in_zip, README_TXT)
        docx_inside_zip_matches_docx = files_identical(docx_in_zip, REPORT_DOCX)

        checks["workbook_inside_zip_matches_source"] = (
            "PASS" if workbook_inside_zip_matches_source else "FAIL"
        )
        checks["readme_inside_zip_matches_txt"] = (
            "PASS" if readme_inside_zip_matches_txt else "FAIL"
        )
        checks["docx_inside_zip_matches_docx"] = (
            "PASS" if docx_inside_zip_matches_docx else "FAIL"
        )

        extraction_result = {
            "zip_listing": zip_listing,
            "root_folder": root_entries[0].name if root_entries else None,
            "inner_files": inner_files,
            "inner_file_count": zip_file_count,
        }

    checks_pass = sum(1 for v in checks.values() if v == "PASS")
    checks_fail = sum(1 for v in checks.values() if v == "FAIL")
    all_pass = checks_fail == 0

    validation = {
        "timestamp": timestamp,
        "activity": "DP_1268",
        "hashes_before_sources": hashes_before,
        "hashes_after_sources": hashes_after_sources,
        "hashes_generated": hashes_generated,
        "source_files_unchanged": {
            "workbook": source_workbook_unchanged,
            "readme_md": source_readme_md_unchanged,
            "report_md": source_report_md_unchanged,
        },
        "zip_path": str(ZIP_PATH),
        "zip_listing": extraction_result["zip_listing"],
        "extraction_validation": extraction_result,
        "checks": checks,
        "checks_pass": checks_pass,
        "checks_fail": checks_fail,
        "overall": "PASS" if all_pass else "FAIL",
    }

    VALIDACAO_JSON.write_text(
        json.dumps(validation, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )

    md_lines = [
        "# Validação de empacotamento final — DP_1268",
        "",
        f"**Timestamp:** {timestamp}",
        "",
        "## Hashes dos arquivos fonte (antes)",
        "",
        f"- `WORKBOOK_DP_1268_DICIONARIO_CONSOLIDADO.xlsx`: `{hashes_before['workbook']}`",
        f"- `README_ENTREGA_DP_1268.md`: `{hashes_before['readme_md']}`",
        f"- `RELATORIO_FINAL_CONSOLIDADO_DP1268.md`: `{hashes_before['report_md']}`",
        "",
        "## Hashes dos arquivos fonte (depois)",
        "",
        f"- `WORKBOOK_DP_1268_DICIONARIO_CONSOLIDADO.xlsx`: `{hashes_after_sources['workbook']}`",
        f"- `README_ENTREGA_DP_1268.md`: `{hashes_after_sources['readme_md']}`",
        f"- `RELATORIO_FINAL_CONSOLIDADO_DP1268.md`: `{hashes_after_sources['report_md']}`",
        "",
        "## Confirmação: arquivos fonte não alterados",
        "",
        f"- Workbook: **{'SIM' if source_workbook_unchanged else 'NÃO'}**",
        f"- README MD: **{'SIM' if source_readme_md_unchanged else 'NÃO'}**",
        f"- Relatório MD: **{'SIM' if source_report_md_unchanged else 'NÃO'}**",
        "",
        "## Hashes dos arquivos gerados",
        "",
        f"- `README_ENTREGA_DP_1268.txt`: `{hashes_generated['readme_txt']}`",
        f"- `RELATORIO_FINAL_CONSOLIDADO_DP1268.docx`: `{hashes_generated['report_docx']}`",
        f"- `DP_1268.zip`: `{hashes_generated['zip']}`",
        "",
        "## Conteúdo do ZIP",
        "",
    ]
    for entry in extraction_result["zip_listing"]:
        md_lines.append(f"- `{entry}`")
    md_lines.extend(
        [
            "",
            "## Resultado da extração temporária",
            "",
            f"- Pasta raiz: `{extraction_result['root_folder']}`",
            f"- Arquivos internos ({extraction_result['inner_file_count']}): {', '.join(f'`{f}`' for f in extraction_result['inner_files'])}",
            "",
            "## Checks",
            "",
            "| Check | Resultado |",
            "|-------|-----------|",
        ]
    )
    for name, result in checks.items():
        md_lines.append(f"| `{name}` | **{result}** |")
    md_lines.extend(
        [
            "",
            f"**Total PASS:** {checks_pass} | **Total FAIL:** {checks_fail}",
            "",
            f"**Resultado geral:** {'EMPACOTAMENTO_FINAL_DP1268_CONCLUIDO_PASS' if all_pass else 'EMPACOTAMENTO_FINAL_DP1268_REQUER_AJUSTE'}",
            "",
        ]
    )
    VALIDACAO_MD.write_text("\n".join(md_lines), encoding="utf-8")

    if all_pass:
        print("EMPACOTAMENTO_FINAL_DP1268_CONCLUIDO_PASS")
    else:
        print("EMPACOTAMENTO_FINAL_DP1268_REQUER_AJUSTE")
        fail_list = [k for k, v in checks.items() if v == "FAIL"]
        for f in fail_list:
            print(f"FAIL: {f}")

    print(f"zip_path={ZIP_PATH}")
    print(f"docx_path={REPORT_DOCX}")
    print(f"readme_txt_path={README_TXT}")
    print(f"checks_pass={checks_pass}")
    print(f"checks_fail={checks_fail}")
    print(f"source_workbook_unchanged={str(source_workbook_unchanged).lower()}")
    print(f"source_readme_md_unchanged={str(source_readme_md_unchanged).lower()}")
    print(f"source_report_md_unchanged={str(source_report_md_unchanged).lower()}")
    print(f"zip_root_ok={str(zip_root_ok).lower()}")
    print(f"zip_file_count={zip_file_count}")
    print(f"zip_has_only_expected_files={str(zip_has_only_expected).lower()}")
    print(f"readme_txt_matches_md={str(text_content_equal(README_TXT, README_MD)).lower()}")
    print(f"docx_created_from_current_md={str(docx_exists and has_item_13).lower()}")
    print(f"workbook_inside_zip_matches_source={str(workbook_inside_zip_matches_source).lower()}")
    print(f"readme_inside_zip_matches_txt={str(readme_inside_zip_matches_txt).lower()}")
    print(f"docx_inside_zip_matches_docx={str(docx_inside_zip_matches_docx).lower()}")
    print(f"sha256_zip={hashes_generated['zip']}")

    return 0 if all_pass else 1


if __name__ == "__main__":
    raise SystemExit(main())
