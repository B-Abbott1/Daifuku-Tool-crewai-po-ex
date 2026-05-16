from po_processer.llm import extract_po_json
from po_processer.tools import ingest_file, write_excel


class PoProcesser:
    """Processes spare parts POs from PDF, Excel, or text files."""

    def run(self, file_path: str, output_path: str = "po_output.xlsx") -> str:
        """
        Full pipeline:
          1. Ingest the file directly in Python
          2. Extract JSON via OpenAI
          3. Write the Excel file directly in Python
        """
        print("[pipeline] Step 1: Ingesting file...")
        raw_text = ingest_file(file_path)
        if raw_text.startswith("[FileIngestorTool]"):
            raise RuntimeError(f"File ingestion failed: {raw_text}")
        print(f"[pipeline] Ingested {len(raw_text)} characters from {file_path}")

        print("[pipeline] Step 2: Extracting line items with GPT-4o-mini...")
        extracted_json = extract_po_json(raw_text)

        print("[pipeline] Step 3: Writing Excel file...")
        confirmation = write_excel(
            normalized_data=extracted_json,
            output_path=output_path,
        )
        if confirmation.startswith("[ExcelWriterTool]"):
            raise RuntimeError(confirmation)
        print(f"[pipeline] {confirmation}")
        return confirmation
