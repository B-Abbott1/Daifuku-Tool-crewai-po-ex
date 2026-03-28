#!/usr/bin/env python
import sys
import warnings
import json
import os

from pathlib import Path

from po_processer.crew import PoProcesser

warnings.filterwarnings("ignore", category=SyntaxWarning, module="pysbd")


def resolve_file_path(file_path: str) -> str:
    path = Path(file_path).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(f"PO file not found: {path}")
    if path.suffix.lower() not in (".pdf", ".xlsx", ".xls", ".txt", ".csv"):
        raise ValueError(f"Unsupported file type '{path.suffix}'. Supported: .pdf, .xlsx, .xls, .txt, .csv")
    return str(path)


def derive_output_path(file_path: str) -> str:
    """Derive the Excel output filename from the input filename.
    e.g. 'C:/Users/bradk/po_processer/acme_po.pdf' -> 'acme_po.xlsx'
    Output is written to the same directory as the input file.
    """
    p = Path(file_path)
    return str(p.parent / (p.stem + ".xlsx"))


def prompt_for_file() -> str:
    print("\n" + "=" * 52)
    print("   PO Processor -- Spare Parts Purchase Order Crew")
    print("=" * 52)
    while True:
        file_path = input("\nEnter path to PO file (PDF, Excel, or text):\n> ").strip().strip('"')
        if not file_path:
            print("  [!] No path entered. Please try again.")
            continue
        path = Path(file_path).expanduser().resolve()
        if not path.exists():
            print(f"  [!] File not found: {path}\n      Please check the path and try again.")
            continue
        if path.suffix.lower() not in (".pdf", ".xlsx", ".xls", ".txt", ".csv"):
            print(f"  [!] Unsupported file type '{path.suffix}'.")
            print("      Supported formats: .pdf  .xlsx  .xls  .txt  .csv")
            continue
        print(f"\n  [OK] File found: {path}")
        return str(path)


def run():
    if len(sys.argv) >= 3:
        file_path = sys.argv[2]
    elif os.environ.get("PO_FILE_PATH"):
        file_path = os.environ["PO_FILE_PATH"]
        print(f"[main] Using PO_FILE_PATH env var: {file_path}")
    else:
        file_path = prompt_for_file()

    try:
        file_path = resolve_file_path(file_path)
        output_path = derive_output_path(file_path)
        print(f"\n[main] Starting PO processing for: {file_path}")
        print(f"[main] Output will be saved to:    {output_path}\n")
        result = PoProcesser().run(file_path=file_path, output_path=output_path)
        print("\n[main] Done.")
        print(result)
    except FileNotFoundError as e:
        raise Exception(f"File error: {e}")
    except Exception as e:
        raise Exception(f"An error occurred: {e}")


def train():
    raise Exception("Train mode not supported in this workflow.")


def replay():
    if len(sys.argv) < 2:
        raise Exception("Usage: python main.py replay <task_id>")
    try:
        PoProcesser().crew().replay(task_id=sys.argv[1])
    except Exception as e:
        raise Exception(f"An error occurred while replaying: {e}")


def run_with_trigger():
    if len(sys.argv) < 2:
        raise Exception("Usage: python main.py trigger '<json_payload>'")
    try:
        trigger_payload = json.loads(sys.argv[1])
    except json.JSONDecodeError:
        raise Exception("Invalid JSON payload.")
    if "file_path" not in trigger_payload:
        raise Exception("Trigger payload must include 'file_path'.")
    try:
        file_path = resolve_file_path(trigger_payload["file_path"])
        output_path = derive_output_path(file_path)
        result = PoProcesser().run(file_path=file_path, output_path=output_path)
        print("\n[main] Done.")
        print(result)
        return result
    except Exception as e:
        raise Exception(f"An error occurred: {e}")


COMMANDS = {
    "run":     run,
    "train":   train,
    "replay":  replay,
    "trigger": run_with_trigger,
}

if __name__ == "__main__":
    command = sys.argv[1] if len(sys.argv) > 1 else "run"
    if command not in COMMANDS:
        print(f"Unknown command '{command}'. Available: {', '.join(COMMANDS)}")
        sys.exit(1)
    COMMANDS[command]()
