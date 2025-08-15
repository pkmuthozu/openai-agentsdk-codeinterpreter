# OpenAI Agent SDK Spreadsheet Analyzer

This repository contains a small example application that demonstrates how to
use the [OpenAI Agent SDK](https://github.com/openai) with the **Code
Interpreter** tool to profile and analyze Excel workbooks.  The script
`analyze_spreadsheet_agent.py` uploads a workbook, instructs an OpenAI agent to
perform a detailed exploration, and saves the results locally.

## Requirements

* Python 3.8+
* An OpenAI API key, either set in the `OPENAI_API_KEY` environment variable or
  provided via the `--api-key` flag
* The `openai` Python package

## Usage

Set your OpenAI API key and run the analyzer:

```bash
# Option 1: read the API key from the environment
export OPENAI_API_KEY="sk-..."
python analyze_spreadsheet_agent.py <workbook.xlsx> --question "What questions should be answered?"

# Option 2: pass the key on the command line
python analyze_spreadsheet_agent.py <workbook.xlsx> --api-key "sk-..." --question "What questions should be answered?"
```

The script will:

1. Upload the workbook to the agent's Code Interpreter environment.
2. Ask the agent to profile all sheets, create a `data_dictionary.json`, plan an
   analysis, and answer the provided question.
3. Download any output files produced (e.g., `data_dictionary.json` or plots).

## Notes

The agent only works with files available inside the Code Interpreter container
and has no external internet access.  All transformations occur in memory; the
original workbook is never modified.

