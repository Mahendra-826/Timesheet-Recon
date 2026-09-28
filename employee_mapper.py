"""
employee_mapper.py

Responsibilities
----------------
1. Load Client and Parent Excel files.
2. Load schema information from schema.json.
3. Convert Excel data into Pandas DataFrames.
4. Extract unique employee names from both sheets.
5. Send ONLY employee names to Gemini for mapping.
6. Do not normalize or fuzzy-match names in Python.
7. Ask the user for confirmation when Gemini is uncertain.
8. Identify Parent employees that are not mapped to any Client employee.
9. Save confirmed mappings to employee_mapping.json.

The full DataFrames are retained because they will be used
later for timesheet hours governance.
"""

from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
from google import genai
import time

from config import (
    CLIENT_TIMESHEET,
    OUTPUT_DIR,
    PARENT_TIMESHEET,
    GEMINI_API_KEY,
    GEMINI_MODEL,
    MAPPING_FILE,
)


class EmployeeMapper:

    def __init__(
        self,
        schema_file="schema.json",
        mapping_file=MAPPING_FILE,
    ):

        self.schema_file = Path(schema_file)
        self.mapping_file = Path(mapping_file)

        self.client = genai.Client(
            api_key=GEMINI_API_KEY
        )

        self.model = GEMINI_MODEL

    # ======================================================
    # LOAD SCHEMA
    # ======================================================

    def load_schema(self):

        if not self.schema_file.exists():

            raise FileNotFoundError(
                f"Schema file not found: {self.schema_file}"
            )

        with open(
            self.schema_file,
            "r",
            encoding="utf-8",
        ) as file:

            return json.load(file)

    # ======================================================
    # LOAD EXCEL USING SCHEMA
    # ======================================================

    def load_dataframes(
        self,
        client_excel,
        parent_excel,
    ):

        schema = self.load_schema()

        client_schema = schema["client"]
        parent_schema = schema["parent"]

        # ----------------------------------------------
        # Client Excel
        # ----------------------------------------------

        client_header_row = client_schema["header_row"]

        client_df = pd.read_excel(
            client_excel,
            header=client_header_row,
        )

        # ----------------------------------------------
        # Parent Excel
        # ----------------------------------------------

        parent_header_row = parent_schema["header_row"]

        parent_df = pd.read_excel(
            parent_excel,
            header=parent_header_row,
        )

        return client_df, parent_df

    # ======================================================
    # EXTRACT EMPLOYEE NAMES
    # ======================================================

    def extract_employee_names(
        self,
        client_df,
        parent_df,
    ):

        schema = self.load_schema()

        client_employee_column = (
            schema["client"]["employee_column"]["name"]
        )

        parent_employee_column = (
            schema["parent"]["employee_column"]["name"]
        )

        # ----------------------------------------------
        # Validate columns
        # ----------------------------------------------

        if client_employee_column not in client_df.columns:

            raise ValueError(
                f"Client employee column "
                f"'{client_employee_column}' "
                f"not found in Excel."
            )

        if parent_employee_column not in parent_df.columns:

            raise ValueError(
                f"Parent employee column "
                f"'{parent_employee_column}' "
                f"not found in Excel."
            )

        # ----------------------------------------------
        # Extract names
        #
        # IMPORTANT:
        # No normalization is performed here.
        # ----------------------------------------------

        client_names = (
            client_df[client_employee_column]
            .dropna()
            .astype(str)
            .drop_duplicates()
            .tolist()
        )

        parent_names = (
            parent_df[parent_employee_column]
            .dropna()
            .astype(str)
            .drop_duplicates()
            .tolist()
        )

        return client_names, parent_names

    # ======================================================
    # BUILD GEMINI PROMPT
    # ======================================================

    def build_prompt(
        self,
        client_names,
        parent_names,
    ):

        prompt = f"""
You are an expert employee identity matching system
for a timesheet governance application.

You need to map employees from a Client Timesheet
to employees from a Parent Company Timesheet.

CLIENT EMPLOYEE NAMES
---------------------

{json.dumps(client_names, indent=2)}

PARENT EMPLOYEE NAMES
---------------------

{json.dumps(parent_names, indent=2)}

IMPORTANT RULES:

1. Perform the employee matching using your reasoning.

2. Do NOT assume that the names have the same format.

3. Names may have:
   - Different ordering
   - First name / last name reversed
   - Extra spaces
   - Commas
   - Initials
   - Abbreviated first names
   - Minor spelling variations

4. Do NOT modify the original names.

5. The parent_name in the result MUST exactly match
   one of the names from the Parent Employee list.

6. Every Client employee must be classified into
   exactly one of these categories:

   - matched
   - needs_confirmation
   - unmatched

7. If there is a clear and reliable match, classify it
   as "matched".

8. If there are multiple possible Parent employees,
   classify it as "needs_confirmation".

9. If there is no reasonable Parent employee match,
   classify it as "unmatched".

10. Do not invent employees.

11. Do not perform fuzzy matching outside your reasoning.

12. Return ONLY valid JSON.

Use this exact structure:

{{
    "mappings": [
        {{
            "client_name": "",
            "parent_name": "",
            "confidence": 0.0,
            "status": "matched",
            "reason": ""
        }}
    ],

    "needs_confirmation": [
        {{
            "client_name": "",
            "possible_parent_names": [],
            "confidence": 0.0,
            "reason": ""
        }}
    ],

    "unmatched": [
        {{
            "client_name": "",
            "reason": ""
        }}
    ]
}}

For "mappings":

- client_name must be exactly as provided.
- parent_name must be exactly as provided.
- confidence must be between 0 and 1.
- status must be "matched".

For "needs_confirmation":

- Keep the original client name.
- List all reasonable Parent employee candidates.
- Do not select one automatically.

For "unmatched":

- Keep the original client name.
- Explain why a reliable match could not be identified.

Return ONLY JSON.
"""

        return prompt

    # ======================================================
    # CALL GEMINI
    # ======================================================

    def map_employees(
        self,
        client_names,
        parent_names,
    ):

        prompt = self.build_prompt(
            client_names,
            parent_names,
        )

        # Log prompt size and measure Gemini call duration
        try:
            print(f"[{time.strftime('%H:%M:%S')}] Calling Gemini for employee mapping (prompt ~{len(prompt)} chars)...", flush=True)
            t0 = time.time()
            response = self.client.models.generate_content(
                model=self.model,
                contents=prompt,
            )
            duration = time.time() - t0
            print(f"[{time.strftime('%H:%M:%S')}] Gemini call completed in {duration:.1f}s", flush=True)
        except Exception as e:
            print(f"[{time.strftime('%H:%M:%S')}] Gemini call failed: {e}", flush=True)
            raise

        text = response.text.strip()

        # ----------------------------------------------
        # Remove Markdown JSON fences
        # ----------------------------------------------

        if text.startswith("```"):

            text = (
                text
                .replace("```json", "")
                .replace("```", "")
                .strip()
            )

        result = json.loads(text)

        usage = response.usage_metadata

        return {
            "result": result,
            "input_tokens": (
                usage.prompt_token_count or 0
            ),
            "output_tokens": (
                usage.candidates_token_count or 0
            ),
            "total_tokens": (
                usage.total_token_count or 0
            ),
        }

    # ======================================================
    # CONFIRM / FINALIZE MAPPINGS
    # ======================================================

    def confirm_mappings(
        self,
        mapping_result,
        parent_names,
    ):
        """
        Finalize Gemini mappings without blocking the governance workflow.

        Clear Gemini matches are accepted as-is.
        For uncertain matches, the first Gemini candidate is retained as a
        best-effort mapping so the existing workflow remains non-blocking.
        For unmatched Client employees, NO input() call is made. They remain
        unmapped and are later classified as Client-only employees.

        This is important for Streamlit because console input() can block the
        application and prevent the remaining governance steps from running.
        """
        confirmed_mappings = []

        # ----------------------------------------------
        # First add all clear Gemini matches
        # ----------------------------------------------
        for mapping in mapping_result.get("mappings", []):
            confirmed_mappings.append(
                {
                    "client_name": mapping["client_name"],
                    "parent_name": mapping["parent_name"],
                    "confidence": mapping.get("confidence", 0.0),
                    "status": "confirmed",
                    "source": "gemini",
                    "reason": mapping.get("reason", ""),
                }
            )

        # ----------------------------------------------
        # Handle uncertain mappings
        # ----------------------------------------------
        # Keep the existing non-blocking behavior: if Gemini provides
        # candidates, use the first candidate. If there are no candidates,
        # leave the employee unmapped so it becomes Client-only.
        for item in mapping_result.get("needs_confirmation", []):
            client_name = item["client_name"]
            candidates = item.get("possible_parent_names", [])

            print("\n" + "=" * 80)
            print("EMPLOYEE MAPPING NEEDS CONFIRMATION")
            print("=" * 80)
            print(f"\nClient Employee : {client_name}")
            print("\nPossible Parent Employees:")

            for index, candidate in enumerate(candidates, start=1):
                print(f"{index}. {candidate}")

            print("0. No match")

            if candidates:
                selected_parent = candidates[0]
                confirmed_mappings.append(
                    {
                        "client_name": client_name,
                        "parent_name": selected_parent,
                        "confidence": item.get("confidence", 0.5),
                        "status": "confirmed",
                        "source": "auto",
                        "reason": (
                            "Auto-selected first Gemini candidate in "
                            "non-interactive governance workflow."
                        ),
                    }
                )
                print(
                    f"Auto-mapped {client_name} -> {selected_parent}"
                )
            else:
                print(
                    f"No candidates for {client_name}; "
                    "classifying as Client-only."
                )

        # ----------------------------------------------
        # Handle unmatched employees
        # ----------------------------------------------
        # IMPORTANT: Never call input() here. An unmatched Client employee
        # is intentionally skipped from governance processing and is later
        # captured by find_client_only_employees().
        for item in mapping_result.get("unmatched", []):
            client_name = item["client_name"]

            print("\n" + "=" * 80)
            print("UNMATCHED EMPLOYEE - CLIENT ONLY")
            print("=" * 80)
            print(f"\nClient Employee : {client_name}")
            print(f"Reason          : {item.get('reason', '')}")
            print(
                "\nNo Parent mapping was found. "
                "Skipping governance processing for this employee."
            )
            print(
                "The employee will be added to the final output as "
                "Client-only."
            )

        return confirmed_mappings

    # ======================================================
    # FIND PARENT EMPLOYEES NOT PRESENT IN CLIENT MAPPING
    # ======================================================

    def find_parent_only_employees(
        self,
        parent_names,
        final_mappings,
    ):
        """
        Identify Parent employees who have no Client mapping.

        Employee matching is still done entirely by Gemini.
        Here we only ignore extra spaces while checking whether
        a mapped Parent employee should be excluded from the
        Parent-only report.
        """

        def clean(name):
            return " ".join(str(name).split())

        mapped_parent_names = {
            clean(mapping["parent_name"])
            for mapping in final_mappings
            if mapping.get("parent_name")
        }

        parent_only = [
            parent_name
            for parent_name in parent_names
            if clean(parent_name) not in mapped_parent_names
        ]

        return parent_only
    # ======================================================
    # FIND CLIENT EMPLOYEES NOT MAPPED TO PARENT
    # ======================================================
    def find_client_only_employees(
                self,
                client_names,
                final_mappings,
            ):
                """
                Identify Client employees who were not mapped
                to any Parent employee.
    
                No name normalization is performed.
                """
    
                mapped_client_names = {
                    mapping["client_name"]
                    for mapping in final_mappings
                    if mapping.get("client_name")
                }
    
                client_only = [
                    client_name
                    for client_name in client_names
                    if client_name not in mapped_client_names
                ]
    
                return client_only
    
    # ======================================================
    # DISPLAY PARENT-ONLY EMPLOYEES
    # ======================================================

    def display_parent_only_employees(
        self,
        parent_only,
    ):

        print("\n" + "=" * 80)
        print("PARENT EMPLOYEES NOT FOUND IN CLIENT")
        print("=" * 80)

        if not parent_only:

            print(
                "\nAll Parent employees have a Client mapping."
            )

            return

        print(
            f"\nTotal Parent-only employees : "
            f"{len(parent_only)}"
        )

        print(
            "\nThese employees are present in the "
            "Parent sheet but have no Client mapping:"
        )

        for index, name in enumerate(
            parent_only,
            start=1,
        ):

            print(
                f"{index}. {name}"
            )

        print("=" * 80)
    # ======================================================
    # DISPLAY CLIENT-ONLY EMPLOYEES
    # ======================================================

    def display_client_only_employees(
        self,
        client_only,
    ):

        print("\n" + "=" * 80)
        print("CLIENT EMPLOYEES NOT FOUND IN PARENT")
        print("=" * 80)

        if not client_only:

            print(
                "\nAll Client employees have a Parent mapping."
            )

            return

        print(
            f"\nTotal Client-only employees : "
            f"{len(client_only)}"
        )

        print(
            "\nThese employees are present in the "
            "Client sheet but have no Parent mapping:"
        )

        for index, name in enumerate(
            client_only,
            start=1,
        ):

            print(
                f"{index}. {name}"
            )

        print("=" * 80)
    # ======================================================
    # SAVE MAPPING
    # ======================================================

    def save_mapping(
        self,
        mappings,
        client_only,
        parent_only,
    ):

        output = {
            "mappings": mappings,
            "client_only": client_only,
            "parent_only": parent_only,
        }

        with open(
            self.mapping_file,
            "w",
            encoding="utf-8",
        ) as file:

            json.dump(
                output,
                file,
                indent=4,
                ensure_ascii=False,
            )

        print(
            f"\nMapping saved to: "
            f"{self.mapping_file}"
        )

    # ======================================================
    # MAIN WORKFLOW
    # ======================================================

    def run(
        self,
        client_excel,
        parent_excel,
    ):

        print("=" * 80)
        print("EMPLOYEE MAPPER")
        print("=" * 80)

        # ----------------------------------------------
        # Load Excel data
        # ----------------------------------------------

        client_df, parent_df = (
            self.load_dataframes(
                client_excel,
                parent_excel,
            )
        )  
        # Save extracted data for verification
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

        output_path = OUTPUT_DIR / "extracted_data.xlsx"

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            client_df.to_excel(
                writer,
                sheet_name="Client_Extracted",
                index=False,
            )
            parent_df.to_excel(
                writer,
                sheet_name="Parent_Extracted",
                index=False,
            )

        # Extracted data saved to file for verification
        # (Details suppressed to reduce console noise)

        # ----------------------------------------------
        # Extract employee names
        # ----------------------------------------------

        client_names, parent_names = (
            self.extract_employee_names(
                client_df,
                parent_df,
            )
        )

        print("\n" + "=" * 80)
        print("EMPLOYEE NAMES")
        print("=" * 80)

        print("\nClient Employees:")

        for name in client_names:

            print(
                f" - {name}"
            )

        print("\nParent Employees:")

        for name in parent_names:

            print(
                f" - {name}"
            )

        print(
            f"\nUnique Client Employees : "
            f"{len(client_names)}"
        )

        print(
            f"Unique Parent Employees : "
            f"{len(parent_names)}"
        )

        # ----------------------------------------------
        # Send ONLY names to Gemini
        # ----------------------------------------------

        print("\n" + "=" * 80)
        print("RUNNING GEMINI EMPLOYEE MAPPING")
        print("=" * 80)

        gemini_result = self.map_employees(
            client_names,
            parent_names,
        )

        mapping_result = gemini_result["result"]

        # ----------------------------------------------
        # Display Gemini result
        # ----------------------------------------------

        print("\nGemini mapping completed.")

        print("\nClear Matches:")

        for mapping in mapping_result.get(
            "mappings",
            [],
        ):

            print(
                f"  {mapping['client_name']}"
                f"  ->  "
                f"{mapping['parent_name']}"
            )

        print("\nNeeds Confirmation:")

        for item in mapping_result.get(
            "needs_confirmation",
            [],
        ):

            print(
                f"  {item['client_name']}"
            )

        print("\nUnmatched:")

        for item in mapping_result.get(
            "unmatched",
            [],
        ):

            print(
                f"  {item['client_name']}"
            )

        # ----------------------------------------------
        # User confirmation
        # ----------------------------------------------

        final_mappings = self.confirm_mappings(
            mapping_result,
            parent_names,
        )

        # ----------------------------------------------
        # Find Client-only employees
        # ----------------------------------------------

        client_only = (
            self.find_client_only_employees(
                client_names,
                final_mappings,
            )
        )

        # ----------------------------------------------
        # Find Parent-only employees
        # ----------------------------------------------

        parent_only = (
            self.find_parent_only_employees(
                parent_names,
                final_mappings,
            )
        )

        # ----------------------------------------------
        # Display Client-only employees
        # ----------------------------------------------

        self.display_client_only_employees(
            client_only
        )

        # ----------------------------------------------
        # Display Parent-only employees
        # ----------------------------------------------

        self.display_parent_only_employees(
            parent_only
        )

        # ----------------------------------------------
        # Save mapping
        # ----------------------------------------------

        self.save_mapping(
            final_mappings,
            client_only,
            parent_only,
        )
        # ----------------------------------------------
        # Token usage
        # ----------------------------------------------

        print("\n" + "=" * 80)
        print("EMPLOYEE MAPPER TOKEN USAGE")
        print("=" * 80)

        print(
            f"Input Tokens  : "
            f"{gemini_result['input_tokens']}"
        )

        print(
            f"Output Tokens : "
            f"{gemini_result['output_tokens']}"
        )

        print(
            f"Total Tokens  : "
            f"{gemini_result['total_tokens']}"
        )

        print("=" * 80)

        return {
            "client_df": client_df,
            "parent_df": parent_df,
            "client_names": client_names,
            "parent_names": parent_names,
            "mappings": final_mappings,
            "client_only": client_only,
            "parent_only": parent_only,
            "input_tokens": gemini_result[
                "input_tokens"
            ],
            "output_tokens": gemini_result[
                "output_tokens"
            ],
            "total_tokens": gemini_result[
                "total_tokens"
            ],
        }


# ==========================================================
# COMMAND LINE ENTRY POINT
# ==========================================================

if __name__ == "__main__":

    mapper = EmployeeMapper()

    mapper.run(
        CLIENT_TIMESHEET,
        PARENT_TIMESHEET,
    )

