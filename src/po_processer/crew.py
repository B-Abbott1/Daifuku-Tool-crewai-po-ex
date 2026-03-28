import json
import os
from dotenv import load_dotenv
from crewai import Agent, Crew, Process, Task, LLM
from crewai.project import CrewBase, agent, crew, task
from crewai.agents.agent_builder.base_agent import BaseAgent
from po_processer.tools.file_ingestor import FileIngestorTool
from po_processer.tools.excel_writer import ExcelWriterTool

load_dotenv()


@CrewBase
class PoProcesser():
    """PoProcesser crew — processes spare parts POs from PDF, Excel, or text files."""

    agents: list[BaseAgent]
    tasks: list[Task]

    openai_llm = LLM(
        model="gpt-4o-mini",
        api_key=os.getenv("OPENAI_API_KEY"),
    )

    # ---------------------------------------------------------------------------
    # Agents
    # ---------------------------------------------------------------------------

    @agent
    def po_extraction_expert(self) -> Agent:
        # Only agent — GPT-4o-mini reads text and outputs JSON.
        # File ingestion and Excel writing are handled in Python directly.
        return Agent(
            config=self.agents_config['po_extraction_expert'],  # type: ignore[index]
            llm=self.openai_llm,
            verbose=True,
            max_iter=3,
        )

    # ---------------------------------------------------------------------------
    # Tasks
    # ---------------------------------------------------------------------------

    @task
    def extract_po_line_items_task(self) -> Task:
        return Task(
            config=self.tasks_config['extract_po_line_items_task'],  # type: ignore[index]
        )

    # ---------------------------------------------------------------------------
    # Crew
    # ---------------------------------------------------------------------------

    @crew
    def crew(self) -> Crew:
        return Crew(
            agents=self.agents,
            tasks=self.tasks,
            process=Process.sequential,
            verbose=True,
            memory=False,
        )

    # ---------------------------------------------------------------------------
    # Orchestration — called from main.py instead of crew().kickoff() directly
    # ---------------------------------------------------------------------------

    def run(self, file_path: str, output_path: str = "po_output.xlsx") -> str:
        """
        Full pipeline:
          1. Ingest the file directly in Python (no LLM needed)
          2. Pass raw text to GPT-4o-mini for JSON extraction
          3. Write the Excel file directly in Python (no LLM needed)
        Returns the excel_writer confirmation string.
        """
        # Step 1: ingest file directly — no LLM involved
        print("[crew] Step 1: Ingesting file...")
        raw_text = FileIngestorTool()._run(file_path=file_path)
        if raw_text.startswith("[FileIngestorTool]"):
            raise RuntimeError(f"File ingestion failed: {raw_text}")
        print(f"[crew] Ingested {len(raw_text)} characters from {file_path}")

        # Step 2: run only the extraction task, injecting raw_text as context
        print("[crew] Step 2: Extracting line items with GPT-4o-mini...")
        result = self.crew().kickoff(inputs={
            "raw_po_text": raw_text,
            "file_path": file_path,
        })
        extracted_json = str(result)

        # Step 3: write Excel directly in Python — no LLM involved
        print("[crew] Step 3: Writing Excel file...")
        writer = ExcelWriterTool()
        confirmation = writer._run(
            normalized_data=extracted_json,
            output_path=output_path,
        )
        print(f"[crew] {confirmation}")
        return confirmation
