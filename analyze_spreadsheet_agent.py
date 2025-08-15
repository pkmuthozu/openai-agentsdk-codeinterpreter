"""Spreadsheet analysis via OpenAI Agent SDK and Code Interpreter.

This script uploads an Excel workbook and instructs an OpenAI agent equipped
with the Code Interpreter tool to profile and analyze its contents.  The agent
follows a structured workflow:

1. **PROFILE** – inspect sheets, infer data types, and build a
   `data_dictionary.json`.
2. **PLAN** – describe how to answer the user's question using the available
   data.
3. **ANSWER** – execute the plan with pandas, optionally producing plots.

The script saves any files produced by the agent (for example,
`data_dictionary.json` or charts) into the current working directory.

Usage:

```
python analyze_spreadsheet_agent.py <workbook.xlsx> --question "What is the task?"
```

An OpenAI API key must be supplied either via the `OPENAI_API_KEY` environment
variable or the `--api-key` command-line option.
"""

from __future__ import annotations

import argparse
import os
from pathlib import Path

from openai import OpenAI


# Instructions given to the agent.  They mirror the detailed spec from the
# repository README and user request.
ANALYST_INSTRUCTIONS = """
You are a meticulous data analyst working ONLY with the Excel workbook available inside the Code Interpreter container.

STEP 1 — PROFILE:
- Use Code Interpreter to open the workbook safely (the file has been uploaded to your container).
- Enumerate sheet names. For each sheet:
  - sample up to 10 rows (head), row count, and column count
  - list columns with inferred dtype (numeric/date/text/categorical) and % missing
  - note likely primary keys or unique identifier columns (if any)
  - detect obvious date columns and normalize to ISO-8601 (YYYY-MM-DD) in memory (do not overwrite original)
- Output a compact JSON object called DATA_DICTIONARY with this structure:
  {
    "sheets": [
      {
        "name": "...",
        "rows": 12345,
        "cols": 12,
        "columns": [
          {"name": "col_a", "inferred_type": "numeric|date|text|categorical", "missing_pct": 0.12, "unique_ct": 999}
        ],
        "sample": [ ... up to 10 rows ... ]
      }
    ],
    "notes": ["any parsing warnings or assumptions"]
  }
- Save this JSON to a file named data_dictionary.json and also print it in the text output.

STEP 2 — PLAN:
- Based on the DATA_DICTIONARY, draft a short PLAN (bulleted) describing how to answer the user question.
- If required columns/sheets are missing or ambiguous, state that clearly and propose fallback options.

STEP 3 — ANSWER:
- Execute the PLAN using pandas. Prefer memory-efficient operations and sampling if the sheet is huge.
- If a plot helps, produce a simple chart (PNG). Titles and axis labels must be clear and short.
- End with a short, actionable TL;DR.

Constraints:
- No external internet access.
- Do not write back to the source file; transformations should be in-memory only.
- Be explicit about any assumptions.
"""


def run_analysis(workbook: Path, question: str, api_key: str | None = None) -> None:
    """Upload ``workbook`` and ask the agent to analyse it.

    Parameters
    ----------
    workbook:
        Path to the Excel file to be analysed.  The file is uploaded to the
        agent's Code Interpreter environment.
    question:
        The user's question that the agent should plan to answer based on the
        workbook contents.
    api_key:
        Optional OpenAI API key. If omitted, the ``OPENAI_API_KEY`` environment
        variable must be set.
    """

    client = OpenAI(api_key=api_key) if api_key else OpenAI()

    # Upload workbook so the code interpreter can access it.
    with open(workbook, "rb") as f:
        uploaded = client.files.create(file=f, purpose="assistants")

    # Create the agent with Code Interpreter capability.
    agent = client.agents.create(
        name="Spreadsheet Analyst",
        model="gpt-4.1",
        instructions=ANALYST_INSTRUCTIONS,
        tools=[{"type": "code_interpreter"}],
    )

    # Start a new thread and send the user's question with the uploaded file
    # attached.
    thread = client.threads.create()
    client.messages.create(
        thread_id=thread.id,
        role="user",
        content=[{"type": "input_text", "text": question}],
        attachments=[{"file_id": uploaded.id}],
    )

    # Run the agent and wait for completion.
    run = client.threads.runs.create_and_poll(
        thread_id=thread.id, agent_id=agent.id
    )

    # Print any text outputs produced by the agent.
    if run.output_text:
        print(run.output_text)

    # Download any files created in the Code Interpreter environment.
    for item in run.output:
        if item.type != "output_file":
            continue

        file_id = item.file_id
        filename = item.filename or f"{file_id}"
        content = client.files.content(file_id)
        with open(filename, "wb") as f:
            f.write(content)
        print(f"Saved {filename}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Analyze an Excel workbook using an OpenAI Agent with Code Interpreter"
    )
    parser.add_argument("workbook", type=Path, help="Path to the Excel workbook")
    parser.add_argument(
        "--question",
        default="What insights can you derive from this workbook?",
        help="User question guiding the analysis",
    )
    parser.add_argument(
        "--api-key",
        dest="api_key",
        help="OpenAI API key (optional if OPENAI_API_KEY env var is set)",
    )

    args = parser.parse_args()

    run_analysis(args.workbook, args.question, args.api_key)


if __name__ == "__main__":
    main()

