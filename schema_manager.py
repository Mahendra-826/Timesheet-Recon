"""
schema_manager.py

Responsibilities
----------------
1. Load existing schema.json
2. Validate schema against current Excel
3. Re-run SchemaAnalyzer if schema changed
4. Save latest schema
"""

from __future__ import annotations

import json
from pathlib import Path

import pandas as pd

from schema_analyzer import SchemaAnalyzer


class SchemaManager:

    def __init__(self, schema_file="schema.json"):

        self.schema_file = Path(schema_file)

        self.analyzer = SchemaAnalyzer()

    # --------------------------------------------------------

    def get_schema(self, client_excel, parent_excel):

        """
        Returns a valid schema.

        If schema.json doesn't exist,
        Gemini is called.

        If schema.json exists,
        it is validated against the Excel.

        If validation fails,
        Gemini is called again.
        """

        if self.schema_file.exists():

            schema = self.load_schema()

            client_valid = self.validate_schema(
                client_excel,
                schema["client"]
            )

            parent_valid = self.validate_schema(
                parent_excel,
                schema["parent"]
            )

            if client_valid and parent_valid:

                print("Loading existing schema...")
                print("Client schema valid")
                print("Parent schema valid")

                return {
                    "client": {
                        "schema": schema["client"],
                        "input_tokens": 0,
                        "output_tokens": 0,
                        "total_tokens": 0,
                    },
                    "parent": {
                        "schema": schema["parent"],
                        "input_tokens": 0,
                        "output_tokens": 0,
                        "total_tokens": 0,
                    },
                }

            print("Schema mismatch detected.")
            print("Running Schema Analyzer...\n")

        else:

            print("No schema found.")
            print("Running Schema Analyzer...\n")

        result = self.analyzer.analyze(client_excel)
        parent_result = self.analyzer.analyze(parent_excel)

        self.save_schema({
            "client": result["schema"],
            "parent": parent_result["schema"]
        })

        return {
            "client": result,
            "parent": parent_result
        }
    # --------------------------------------------------------

    def validate_schema(self, excel_path, schema):

        df = pd.read_excel(
            excel_path,
            header=None,
        )

        header_row = schema["header_row"]

        current_columns = (
            df.iloc[header_row]
            .fillna("")
            .astype(str)
            .tolist()
        )

        saved_columns = schema["columns"]
        employee_column = schema.get("employee_column", {}).get("name")
        invalid_date_columns = any(
            pd.isna(
                pd.to_datetime(
                    item.get("name", "").split()[0],
                    format="%d-%b",
                    errors="coerce",
                )
            )
            for item in schema.get("date_columns", [])
        )

        # week_start_column / week_end_column / billable_hours_column /
        # status_column were added later. Force a one-time re-analysis
        # for any schema.json saved before they existed, and otherwise
        # make sure a column Gemini previously identified is still
        # present in the current header.
        optional_column_fields = (
            "week_start_column",
            "week_end_column",
            "billable_hours_column",
            "status_column",
        )

        missing_optional_fields = any(
            field not in schema
            for field in optional_column_fields
        )

        stale_optional_columns = any(
            schema.get(field) is not None
            and schema[field].get("name") not in current_columns
            for field in optional_column_fields
        )

        if (
            current_columns != saved_columns
            or employee_column not in current_columns
            or invalid_date_columns
            or missing_optional_fields
            or stale_optional_columns
        ):
            print("\nSaved Columns:")
            print(saved_columns)

            print("\nCurrent Columns:")
            print(current_columns)

            if employee_column not in current_columns:
                print(
                    f"\nEmployee column '{employee_column}' "
                    "is missing from the saved header row."
                )
            else:
                print("\nHeader mismatch detected.")
            return False

        return True

    # --------------------------------------------------------

    def save_schema(self, schema):

        with open(
            self.schema_file,
            "w",
            encoding="utf-8",
        ) as f:

            json.dump(
                schema,
                f,
                indent=4,
            )

    # --------------------------------------------------------

    def load_schema(self):

        with open(
            self.schema_file,
            "r",
            encoding="utf-8",
        ) as f:

            return json.load(f)