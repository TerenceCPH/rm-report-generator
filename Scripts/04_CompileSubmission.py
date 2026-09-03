"""04 Compile Submission: insert generated appendix PDFs after matching header pages."""

import argparse
import glob
import os
import re

import fitz

from env import OUTPUT_DIR, SUBMISSION_DIR, output_subdir

# Match header-page text -> PDFs to insert after that page.
INSERT_RULES = [
    {
        "name": "Appendix A – Report Builder",
        "page_re": re.compile(r"Appendix\s*A.{0,80}Report Builder", re.I | re.S),
        "globs": [
            os.path.join(OUTPUT_DIR, "01_ReportBuilder", "*.pdf"),
        ],
    },
    {
        "name": "Appendix B – Changes Agreed",
        "page_re": re.compile(
            r"Appendix\s*B.{0,120}(Agreed|Artifact Types)", re.I | re.S
        ),
        "globs": [
            os.path.join(OUTPUT_DIR, "02_ComplianceSummary", "Appendix B*.pdf"),
        ],
    },
    {
        "name": "Appendix C – Proposed Changes",
        "page_re": re.compile(
            r"Appendix\s*C.{0,120}Proposed Changes", re.I | re.S
        ),
        "globs": [
            os.path.join(OUTPUT_DIR, "02_ComplianceSummary", "Appendix C*.pdf"),
        ],
    },
    {
        "name": "Appendix D – RVT Attachment",
        "page_re": re.compile(r"Appendix\s*D.{0,80}RVT", re.I | re.S),
        "globs": [
            os.path.join(OUTPUT_DIR, "03_RVTGenerator", "*.pdf"),
        ],
    },
]


def list_pdfs(patterns):
    files = []
    for pattern in patterns:
        files.extend(glob.glob(pattern))
    files = [os.path.abspath(p) for p in files if os.path.isfile(p)]
    return sorted(set(files), key=lambda p: os.path.basename(p).lower())


def page_text(page):
    return page.get_text("text") or ""


def find_header_page(doc, pattern):
    """Return 1-based page number of the appendix divider page, or None.

    Skips the table of contents and prefers short title-only pages so body
    mentions do not win over the dedicated header sheet.
    """
    candidates = []
    for i in range(doc.page_count):
        text = page_text(doc.load_page(i))
        if re.search(r"table\s+of\s+contents", text, re.I):
            continue
        if not pattern.search(text):
            continue
        candidates.append((i + 1, len(text)))
    if not candidates:
        return None
    short = [item for item in candidates if item[1] < 800]
    pool = short or candidates
    return pool[-1][0]


def merge_pdfs(paths):
    merged = fitz.open()
    for path in paths:
        src = fitz.open(path)
        merged.insert_pdf(src)
        src.close()
    return merged


def compile_submission(input_pdf, output_pdf):
    if not os.path.exists(input_pdf):
        raise SystemExit(f"Input PDF not found: {input_pdf}")

    stub = fitz.open(input_pdf)
    print(f"Loaded stub PDF ({stub.page_count} pages): {input_pdf}")

    jobs = []
    for rule in INSERT_RULES:
        page_no = find_header_page(stub, rule["page_re"])
        pdfs = list_pdfs(rule["globs"])
        if page_no is None:
            print(f"[skip] No header page matched '{rule['name']}'.")
            continue
        if not pdfs:
            print(
                f"[skip] Header on p.{page_no} for '{rule['name']}', "
                f"but no PDF files found."
            )
            continue
        print(
            f"[insert] '{rule['name']}' after p.{page_no} "
            f"({len(pdfs)} file(s)): {', '.join(os.path.basename(p) for p in pdfs)}"
        )
        jobs.append((page_no, pdfs, rule["name"]))

    stub.close()

    if not jobs:
        raise SystemExit("Nothing to insert. Generate appendix PDFs first, then retry.")

    # Insert from the last header page so earlier page numbers stay valid.
    jobs.sort(key=lambda item: item[0], reverse=True)

    compiled = fitz.open(input_pdf)
    for page_no, pdfs, name in jobs:
        block = merge_pdfs(pdfs)
        # After 1-based page P, insert before 0-based index P.
        compiled.insert_pdf(block, start_at=page_no)
        print(f"  inserted {block.page_count} page(s) for {name} after p.{page_no}")
        block.close()

    out_dir = os.path.dirname(output_pdf)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    compiled.save(output_pdf)
    print(f"Success! Saved '{output_pdf}' ({compiled.page_count} pages).")
    compiled.close()
    return output_pdf


def default_output_path(input_pdf):
    stem = os.path.splitext(os.path.basename(input_pdf))[0]
    return os.path.join(output_subdir("04_CompileSubmission"), f"{stem} - compiled.pdf")


def parse_args():
    parser = argparse.ArgumentParser(
        description=(
            "Insert generated appendix PDFs after matching header pages in the "
            "submission PDF (e.g. Appendix A after its header page)."
        )
    )
    parser.add_argument(
        "input_pdf",
        nargs="?",
        default=os.path.join(SUBMISSION_DIR, "1263-W-000-SOB-190-000014 (Rev A).pdf"),
        help="Path to the stub submission PDF.",
    )
    parser.add_argument(
        "--output",
        default=None,
        help="Compiled PDF path. Default: Output/04_CompileSubmission/<stem> - compiled.pdf",
    )
    return parser.parse_args()


def main():
    args = parse_args()
    output_pdf = args.output or default_output_path(args.input_pdf)
    compile_submission(args.input_pdf, output_pdf)


if __name__ == "__main__":
    main()
