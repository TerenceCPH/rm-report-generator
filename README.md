# RM Automation

Python scripts for generating DOORS Next appendices, exporting Excel files to PDF via Microsoft Excel, and compiling the final submission PDF.

## Project Layout

```text
RM/
  Inputs/
    Submission/
      RM_Report.docx
      RM_Report.pdf
    Template/
      Requirement_Compliance_Summary.xlsx
    XML/
      Compliance Summary_AppB_AppC.xml
      Module List.xml
      Report Builder.xml
      Requirements Checker.xml
  Output/
    01_ReportBuilder/
    02_ComplianceSummary/
    03_RVTGenerator/
    04_CompileSubmission/
  Scripts/
    00_EnvironmentVariables.py
    01_ReportBuilder.py
    02_ComplianceSummary.py
    03_RVTGenerator.py
    04_CompileSubmission.py
  requirements.txt
  README.md
```

## Requirements

- Python 3.11+
- Microsoft Excel installed on Windows
- Edge browser installed for `03_RVTGenerator.py`

Install Python packages:

```bash
pip install -r requirements.txt
python -m playwright install msedge
```

## Scripts

### `Scripts/01_ReportBuilder.py`

Builds Appendix A from the Report Builder XML and exports:

- `Output/01_ReportBuilder/Appendix A – Report Builder.xlsx`
- `Output/01_ReportBuilder/Appendix A – Report Builder.pdf`

Example:

```bash
python Scripts/01_ReportBuilder.py
python Scripts/01_ReportBuilder.py --xml "Inputs\XML\Report Builder.xml"
```

### `Scripts/02_ComplianceSummary.py`

Fills the compliance summary template, generates Appendix B/C workbooks, and exports Appendix B/C PDFs only.

Outputs:

- `Output/02_ComplianceSummary/Requirement_Compliance_Summary.xlsx`
- `Output/02_ComplianceSummary/Appendix B – List Of Changes Agreed By MTR In Artifact Types.xlsx`
- `Output/02_ComplianceSummary/Appendix B – List Of Changes Agreed By MTR In Artifact Types.pdf`
- `Output/02_ComplianceSummary/Appendix C – List Of Proposed Changes In Artifact Types.xlsx` / `.pdf` when matching Proposed rows exist

Example:

```bash
python Scripts/02_ComplianceSummary.py
python Scripts/02_ComplianceSummary.py --xml "Inputs\XML\Compliance Summary_AppB_AppC.xml"
python Scripts/02_ComplianceSummary.py --skip-pdf
```

### `Scripts/03_RVTGenerator.py`

Downloads RVT attachment workbooks from DOORS Next, formats them, and exports PDFs.

Default CSV lookup order:

1. `Inputs/AppD/RM Module List.csv`
2. `RM/AppD/RM Module List.csv`
3. `RM/RM Module List.csv`
4. legacy Terence path

Example:

```bash
python Scripts/03_RVTGenerator.py --csv "C:\path\to\RM Module List.csv"
python Scripts/03_RVTGenerator.py --skip-pdf
```

### `Scripts/04_CompileSubmission.py`

Takes the stub submission PDF, finds the appendix divider pages by text, and inserts generated appendix PDFs after them.

Default input:

- `Inputs/Submission/RM_Report.pdf`

Output:

- `Output/04_CompileSubmission/<submission name> - compiled.pdf`

Example:

```bash
python Scripts/04_CompileSubmission.py
python Scripts/04_CompileSubmission.py "Inputs\Submission\RM_Report.pdf"
```

## Typical Workflow

```bash
python Scripts/01_ReportBuilder.py
python Scripts/02_ComplianceSummary.py
python Scripts/03_RVTGenerator.py --csv "C:\path\to\RM Module List.csv"
python Scripts/04_CompileSubmission.py
```

## Notes

- XML fetching is online-first with fallback to the cached files in `Inputs/XML/`.
- PDF export uses Excel COM through `pywin32`. Adobe PDF is not required.
- `02_ComplianceSummary.py` does not generate a PDF for `Requirement_Compliance_Summary.xlsx` because that sheet is not used in the final compiled submission.
