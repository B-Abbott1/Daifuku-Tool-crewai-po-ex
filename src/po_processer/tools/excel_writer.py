import json
from crewai.tools import BaseTool
from pydantic import BaseModel, Field
from typing import Optional
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


class ExcelWriterInput(BaseModel):
    normalized_data: str = Field(
        ...,
        description="JSON string of the extracted PO data (header + line_items).",
    )
    anomaly_report: Optional[str] = Field(
        default=None,
        description="Optional JSON string of an anomaly report. Ignored in the current 3-agent workflow.",
    )
    output_path: str = Field(
        default="po_output.xlsx",
        description="File path for the output Excel workbook.",
    )


class ExcelWriterTool(BaseTool):
    name: str = "Excel Writer"
    description: str = (
        "Writes extracted PO line items to a formatted Excel workbook (.xlsx). "
        "Creates a 'Line Items' sheet with one row per line item, bold headers, "
        "frozen top row, and autofilter. Accepts normalized_data as a JSON string "
        "containing a 'header' object and a 'line_items' array."
    )
    args_schema: type[BaseModel] = ExcelWriterInput

    # Color constants
    DARK_BLUE: str = "1F3864"
    WHITE: str = "FFFFFF"

    def _extract_first_json_object(self, text: str) -> dict:
        """
        Robustly extract the first complete JSON object from a string.
        Handles cases where the LLM appends extra text or a second JSON object
        after the main payload (which causes json.loads to raise 'Extra data').
        Also strips markdown code fences if present.
        """
        # Strip markdown fences
        text = text.strip()
        if text.startswith("```"):
            text = text.split("```")[1]
            if text.startswith("json"):
                text = text[4:]
            text = text.strip()

        # Use a decoder that stops after the first complete object
        decoder = json.JSONDecoder()
        obj, _ = decoder.raw_decode(text)
        return obj

    def _run(
        self,
        normalized_data: str,
        anomaly_report: Optional[str] = None,
        output_path: str = "po_output.xlsx",
    ) -> str:
        try:
            po_data = self._extract_first_json_object(normalized_data)
        except Exception as e:
            return f"[ExcelWriterTool] Failed to parse input JSON: {e}"

        try:
            wb = Workbook()
            self._write_line_items_sheet(wb, po_data)

            # Remove the default blank sheet if a named sheet was created
            if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
                del wb["Sheet"]

            wb.save(output_path)

            line_count = len(po_data.get("line_items", []))
            return (
                f"Excel workbook written to '{output_path}' — "
                f"Sheet: 'Line Items' — {line_count} line item(s) written."
            )
        except Exception as e:
            return f"[ExcelWriterTool] Failed to write Excel file: {e}"

    def _apply_header(self, cell, value: str) -> None:
        cell.value = value
        cell.font = Font(name="Arial", bold=True, size=11, color=self.WHITE)
        cell.fill = PatternFill("solid", fgColor=self.DARK_BLUE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def _autosize_columns(self, ws, min_width: int = 12, max_width: int = 60) -> None:
        for col_cells in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col_cells[0].column)
            for cell in col_cells:
                try:
                    max_len = max(max_len, len(str(cell.value or "")))
                except Exception:
                    pass
            ws.column_dimensions[col_letter].width = max(min_width, min(max_len + 2, max_width))

    def _write_line_items_sheet(self, wb: Workbook, po_data: dict) -> None:
        ws = wb.active
        ws.title = "Line Items"

        columns = [
            "line_number",
            "part_number",
            "description",
            "quantity",
            "unit_of_measure",
            "unit_price",
            "total_price",
            "supplier_part_number",
            "notes",
        ]

        # Header row
        for col_idx, col_name in enumerate(columns, start=1):
            self._apply_header(
                ws.cell(row=1, column=col_idx),
                col_name.replace("_", " ").title(),
            )

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}1"

        # Data rows
        for row_idx, item in enumerate(po_data.get("line_items", []), start=2):
            row_data = [
                item.get("line_number"),
                item.get("part_number"),
                item.get("description"),
                item.get("quantity"),
                item.get("unit_of_measure"),
                item.get("unit_price"),
                item.get("total_price"),
                item.get("supplier_part_number"),
                item.get("notes"),
            ]
            for col_idx, value in enumerate(row_data, start=1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.font = Font(name="Arial", size=10)
                cell.alignment = Alignment(wrap_text=False, vertical="center")

        self._autosize_columns(ws)
