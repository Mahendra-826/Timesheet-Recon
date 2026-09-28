"""
schema_analyzer.py

Uses Gemini to dynamically identify the schema of an Excel worksheet.

Outputs:
- Sheet Type
- Header Row
- Employee Column
- Date Columns
"""

from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
from google import genai

from config import GEMINI_API_KEY, GEMINI_MODEL



class SchemaAnalyzer:

    def __init__(self):
        self.client = genai.Client(api_key=GEMINI_API_KEY)
        self.model = GEMINI_MODEL

    def analyze(self, excel_path: str | Path) -> dict:
        """
        Analyze an Excel sheet using Gemini.
        """

        preview_df = pd.read_excel(
            excel_path,
            header=None,
            nrows=10,
        )

        rows = (
            preview_df.fillna("")
            .astype(str)
            .values
            .tolist()
        )

        prompt = f"""
You are an expert Excel Schema Analyzer.

Below are the first 10 rows of an Excel worksheet.

{json.dumps(rows, indent=2)}

Your task is to identify the schema.

Identify:

1. Sheet Type
2. Header Row Number (0-based index)
3. Employee Name Column
4. Date Columns
5. Week Start Date Column
6. Week End Date Column
7. Billable/Total Hours Column
8. Status Column

Rules:
- The employee column contains employee names.
- If the sheet has ONE ROW PER DAY (daily calendar dates such as
  1-Jun, 2-Jun...), put those columns in "date_columns" and leave
  "week_start_column" / "week_end_column" as null.
- If the sheet has ONE ROW PER WEEK with separate start/end date
  columns (e.g. "Time Sheet Start Date" / "Time Sheet End Date"),
  leave "date_columns" as an empty list and instead identify those
  two columns as "week_start_column" and "week_end_column".
- "billable_hours_column" is the column holding the numeric hours
  worked/billed for that row (e.g. "Billable Hours (ST)", "Total
  Hours"). Use null if no such column exists.
- "status_column" is the column holding an approval/invoice status
  such as "Invoiced", "Approved", "Pending". Use null if no such
  column exists.
- Ignore all other columns.
- Return ONLY valid JSON.

Return ONLY valid JSON in this format:

{{
    "sheet_type": "",
    "header_row": 0,
    "employee_column": {{
        "name": "",
        "index": 0
    }},
    "date_columns": [
        {{
            "name": "",
            "index": 0
        }}
    ],
    "week_start_column": {{
        "name": "",
        "index": 0
    }},
    "week_end_column": {{
        "name": "",
        "index": 0
    }},
    "billable_hours_column": {{
        "name": "",
        "index": 0
    }},
    "status_column": {{
        "name": "",
        "index": 0
    }}
}}

Use null (not an object) for any of the last four fields that do
not apply to this sheet.
"""
        response = self.client.models.generate_content(
            model=self.model,
            contents=prompt,
        )

        text = response.text.strip()

        if text.startswith("```"):
            text = (
                text.replace("```json", "")
                .replace("```", "")
                .strip()
            )

        schema = json.loads(text)

        # Verify that the proposed employee field exists in the proposed header row.
        proposed_header = (
            preview_df.iloc[schema["header_row"]]
            .fillna("")
            .astype(str)
            .tolist()
        )
        employee_name = schema["employee_column"]["name"]

        if employee_name not in proposed_header:
            for row_index, row in enumerate(rows):
                if employee_name in row:
                    schema["header_row"] = row_index
                    schema["employee_column"]["index"] = row.index(employee_name)
                    break

        # Daily reports often have calendar dates above weekday headers.
        # Store the calendar dates because governance needs real weekdays.
        header_row = schema["header_row"]
        for item in schema.get("date_columns", []):
            column_index = item["index"]
            current_name = str(item.get("name", "")).split()[0]
            current_date = pd.to_datetime(
                current_name,
                format="%d-%b",
                errors="coerce",
            )

            if not pd.isna(current_date):
                continue

            for row_index in range(header_row - 1, -1, -1):
                value = preview_df.iloc[row_index, column_index]
                parsed = pd.to_datetime(
                    str(value),
                    format="%d-%b",
                    errors="coerce",
                )
                if not pd.isna(parsed):
                    item["name"] = f"{parsed.day}-{parsed.strftime('%b')}"
                    break

        header_df = pd.read_excel(
            excel_path,
            header=None,
            skiprows=schema["header_row"],
            nrows=1,
        )

        schema["columns"] = (
            header_df.iloc[0]
            .fillna("")
            .astype(str)
            .tolist()
        )

        # Re-locate each optional column against the confirmed header
        # row, so a wrong/stale index from Gemini can never point at
        # the wrong cell. Set to null when the named column isn't
        # actually present in this sheet's header.
        for field_name in (
            "week_start_column",
            "week_end_column",
            "billable_hours_column",
            "status_column",
        ):
            self._relocate_column(schema, field_name)

        usage = response.usage_metadata

        return {
            "schema": schema,
            "input_tokens": usage.prompt_token_count,
            "output_tokens": usage.candidates_token_count,
            "total_tokens": usage.total_token_count,
        }

    @staticmethod
    def _relocate_column(schema, field_name):
        """
        Verify an optional column identified by Gemini against the
        confirmed header row, fixing its index or discarding it.
        """

        info = schema.get(field_name)

        if not info or not info.get("name"):
            schema[field_name] = None
            return

        columns = schema["columns"]
        name = info["name"]

        if name in columns:
            info["index"] = columns.index(name)
            return

        lowered = [str(c).strip().lower() for c in columns]
        target = str(name).strip().lower()

        if target in lowered:
            index = lowered.index(target)
            info["name"] = columns[index]
            info["index"] = index
            return

        # Gemini named a column that isn't actually in the header;
        # let callers fall back to their own heuristics instead of
        # trusting a made-up index.
        schema[field_name] = None