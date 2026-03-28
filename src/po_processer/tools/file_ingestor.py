from crewai.tools import BaseTool
from pydantic import BaseModel, Field
from pathlib import Path


class FileIngestorInput(BaseModel):
    file_path: str = Field(..., description="Absolute path to the PO file to ingest.")


class FileIngestorTool(BaseTool):
    name: str = "file_ingestor"
    description: str = (
        "Reads a PO file (PDF, Excel, or plain text) and returns its full content as "
        "structured text. For PDFs, tables are extracted with column alignment preserved. "
        "For Excel files, all sheets are returned as tab-separated text. "
        "For plain text files, content is returned as-is."
    )
    args_schema: type[BaseModel] = FileIngestorInput

    def _run(self, file_path: str) -> str:
        path = Path(file_path)
        if not path.exists():
            return f"[FileIngestorTool] File not found: {file_path}"

        suffix = path.suffix.lower()

        if suffix == ".pdf":
            return self._ingest_pdf(path)
        elif suffix in (".xlsx", ".xls"):
            return self._ingest_excel(path)
        elif suffix in (".txt", ".csv"):
            return self._ingest_text(path)
        else:
            return f"[FileIngestorTool] Unsupported file type: {suffix}"

    # ------------------------------------------------------------------
    # PDF — pdfplumber with table-aware extraction
    # ------------------------------------------------------------------

    def _ingest_pdf(self, path: Path) -> str:
        try:
            import pdfplumber
        except ImportError:
            return "[FileIngestorTool] pdfplumber is not installed. Run: pip install pdfplumber"

        output_parts = []

        with pdfplumber.open(str(path)) as pdf:
            for page_num, page in enumerate(pdf.pages, start=1):
                output_parts.append(f"--- Page {page_num} ---")

                # Extract tables first — pdfplumber detects column boundaries
                # and returns rows as lists, preserving alignment.
                tables = page.extract_tables()

                if tables:
                    for table_idx, table in enumerate(tables, start=1):
                        if len(tables) > 1:
                            output_parts.append(f"[Table {table_idx}]")
                        output_parts.append(self._format_table(table))
                else:
                    # No tables detected — fall back to plain text for this page.
                    # words=True preserves reading order better than the default.
                    text = page.extract_text(x_tolerance=3, y_tolerance=3)
                    if text:
                        output_parts.append(text.strip())

        return "\n".join(output_parts)

    def _format_table(self, table: list) -> str:
        """
        Convert a pdfplumber table (list of rows, each row a list of cell strings)
        into a clean tab-separated string.

        pdfplumber returns None for empty cells — we replace those with empty strings
        so the column count stays consistent across rows.
        """
        rows = []
        for row in table:
            # Replace None cells with "" and strip whitespace from each cell
            cleaned = [
                (cell.strip().replace("\n", " ") if cell else "")
                for cell in row
            ]
            rows.append("\t".join(cleaned))
        return "\n".join(rows)

    # ------------------------------------------------------------------
    # Excel — openpyxl, sheet by sheet
    # ------------------------------------------------------------------

    def _ingest_excel(self, path: Path) -> str:
        try:
            import openpyxl
        except ImportError:
            return "[FileIngestorTool] openpyxl is not installed. Run: pip install openpyxl"

        wb = openpyxl.load_workbook(str(path), data_only=True)
        output_parts = []

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            output_parts.append(f"--- Sheet: {sheet_name} ---")
            for row in ws.iter_rows(values_only=True):
                cleaned = [str(cell) if cell is not None else "" for cell in row]
                output_parts.append("\t".join(cleaned))

        return "\n".join(output_parts)

    # ------------------------------------------------------------------
    # Plain text / CSV
    # ------------------------------------------------------------------

    def _ingest_text(self, path: Path) -> str:
        try:
            return path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            return path.read_text(encoding="latin-1")
