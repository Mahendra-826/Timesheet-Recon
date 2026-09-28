"""
hours_governance.py

Weekly Hours Governance

ASPIRE = Parent Mon-Fri total
FG      = Client weekly hours for Invoiced records
FG      = Client status for non-Invoiced records

Governance Rules
----------------
1. Invoiced:
       - Compare Parent Mon-Fri hours (ASPIRE)
         against Client weekly hours.
       - Match     -> GREEN
       - Mismatch  -> RED

2. Non-Invoiced:
       - Do NOT compare client hours.
       - ASPIRE -> RED
       - FG     -> RED
       - FG displays the client status.

3. Invoiced + ASPIRE 0 + Client 0:
       - GREEN / GREEN

4. Non-Invoiced + ASPIRE 0 + Client 0:
       - Still RED / RED

5. Multiple client records for the same employee/week:
       - If any record is non-Invoiced, do not silently discard
         that record in favor of an Invoiced record.
       - A non-Invoiced record must remain visible for governance.
"""

from __future__ import annotations

import json
import re
from copy import copy
from datetime import timedelta
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

from config import (
    CLIENT_TIMESHEET,
    PARENT_TIMESHEET,
    OUTPUT_FILE,
    PENDING_REVIEW_FILE,
    MATCH_COLOR,
    MISMATCH_COLOR,
    MISSING_CLIENT_COLOR,
    CLIENT_ONLY_COLOR,
    PENDING_REVIEW_COLOR,
)


# ====================================================================
# Shared governance decision logic
#
# Extracted so that the SAME rule is used both during the normal run
# (process_week) and later, when a user submits a corrected Billable
# Hours value for a flagged "needs review" week
# (apply_pending_corrections). There must only be one place that
# decides Green vs Red.
# ====================================================================

def evaluate_governance(
    is_invoiced,
    parent_total,
    client_total,
    client_status,
    available_working_days,
    green_fill,
    red_fill,
):
    """
    Returns (display_fg_value, fill, is_red) using the exact governance
    rules documented at the top of this file.
    """

    if is_invoiced:

        # ------------------------------------------------------------
        # Invoiced + 0/0
        # ------------------------------------------------------------

        if abs(parent_total) < 0.01 and abs(client_total) < 0.01:

            return 0, green_fill, False

        # ------------------------------------------------------------
        # Complete 5-day week
        # ------------------------------------------------------------

        if available_working_days == 5:

            display_fg_value = round(client_total, 2)

            if abs(parent_total - client_total) < 0.01:

                return display_fg_value, green_fill, False

            return display_fg_value, red_fill, True

        # ------------------------------------------------------------
        # Incomplete / cross-month week
        # ------------------------------------------------------------

        if parent_total <= client_total:

            return round(parent_total, 2), green_fill, False

        return round(client_total, 2), red_fill, True

    # ------------------------------------------------------------
    # NON-INVOICED
    #
    # NEVER use client_total here.
    # ------------------------------------------------------------

    display_fg_value = client_status if client_status else "Non-Invoiced"

    return display_fg_value, red_fill, True


class HoursGovernance:

    def __init__(self, client_path=None, parent_path=None):

        # ------------------------------------------------------------
        # Load schema
        # ------------------------------------------------------------

        with open("schema.json", "r", encoding="utf-8") as f:
            self.schema = json.load(f)

        # ------------------------------------------------------------
        # Load employee mapping
        # ------------------------------------------------------------

        with open("employee_mapping.json", "r", encoding="utf-8") as f:
            self.mapping_data = json.load(f)

        self.client_schema = self.schema["client"]
        self.parent_schema = self.schema["parent"]

        # ------------------------------------------------------------
        # Read Excel files
        # ------------------------------------------------------------

        client_src = (
            Path(client_path)
            if client_path is not None
            else Path(CLIENT_TIMESHEET)
        )

        parent_src = (
            Path(parent_path)
            if parent_path is not None
            else Path(PARENT_TIMESHEET)
        )

        self.client_df = pd.read_excel(
            client_src,
            header=self.client_schema["header_row"],
        )

        self.parent_df = pd.read_excel(
            parent_src,
            header=self.parent_schema["header_row"],
        )

        # ------------------------------------------------------------
        # Resolve client column names
        # ------------------------------------------------------------

        self.resolve_client_columns()

        # ------------------------------------------------------------
        # Load parent workbook
        # ------------------------------------------------------------

        self.workbook = load_workbook(parent_src)

        self.sheet = self.workbook.active

        # ------------------------------------------------------------
        # Create fills
        # ------------------------------------------------------------

        self.green_fill = PatternFill(
            fill_type="solid",
            start_color=MATCH_COLOR,
            end_color=MATCH_COLOR,
        )

        self.red_fill = PatternFill(
            fill_type="solid",
            start_color=MISMATCH_COLOR,
            end_color=MISMATCH_COLOR,
        )

        self.missing_client_fill = PatternFill(
            fill_type="solid",
            start_color=MISSING_CLIENT_COLOR,
            end_color=MISSING_CLIENT_COLOR,
        )

        self.client_only_fill = PatternFill(
            fill_type="solid",
            start_color=CLIENT_ONLY_COLOR,
            end_color=CLIENT_ONLY_COLOR,
        )

        self.pending_review_fill = PatternFill(
            fill_type="solid",
            start_color=PENDING_REVIEW_COLOR,
            end_color=PENDING_REVIEW_COLOR,
        )

        # ------------------------------------------------------------
        # Weeks with a duplicate-row correction/reversal pattern
        # (e.g. +40 then -40) that need a human to confirm the final
        # Billable Hours before a Green/Red decision is made.
        # ------------------------------------------------------------

        self.pending_review = []

        self.pending_review_cells = set()

        self.pending_review_employees = set()

        # ------------------------------------------------------------
        # Track employees with at least one RED weekly result
        # ------------------------------------------------------------

        self.red_employees = set()

        # ------------------------------------------------------------
        # Track client employees with no matching parent row
        # ------------------------------------------------------------

        self.missing_parent_employees = []

        # ------------------------------------------------------------
        # Rows appended for Client-only employees
        # ------------------------------------------------------------

        self.client_only_excel_rows = set()

        # ------------------------------------------------------------
        # Track cells that were explicitly processed as RED
        #
        # This prevents later zero-highlighting logic from accidentally
        # changing a non-Invoiced RED result to GREEN.
        # ------------------------------------------------------------

        self.red_governance_cells = set()

        # ------------------------------------------------------------
        # Determine year
        # ------------------------------------------------------------

        if self.client_start_col is None:
            raise ValueError(
                "Could not determine client start date column; "
                "update schema or column names"
            )

        valid_start_dates = pd.to_datetime(
            self.client_df[self.client_start_col],
            errors="coerce",
        ).dropna()

        if valid_start_dates.empty:
            raise ValueError(
                "Could not determine year because no valid "
                "client start dates were found."
            )

        self.year = valid_start_dates.iloc[0].year

        # ------------------------------------------------------------
        # Build parent date map
        # ------------------------------------------------------------

        self.parent_date_map = self.build_parent_date_map()

    # ================================================================
    # Build Parent Date Map
    # ================================================================

    def build_parent_date_map(self):

        mapping = {}

        for item in self.parent_schema["date_columns"]:

            dt = pd.to_datetime(
                item["name"].split()[0] + f"-{self.year}",
                format="%d-%b-%Y",
            )

            mapping[dt.date()] = item["index"]

        return mapping

    # ================================================================
    # Update Week Headers
    # ================================================================

    def ensure_governance_columns(self):
        """Create ASPIRE, FG, and the final Total Hrs columns when needed."""

        header_row = self.parent_schema["header_row"] + 1
        employee_col = self.parent_schema["employee_column"]["index"] + 1

        last_employee_row = header_row
        for row in range(header_row + 1, self.sheet.max_row + 1):
            employee_value = self.sheet.cell(row, employee_col).value
            if employee_value is not None and str(employee_value).strip() != "":
                last_employee_row = row

        date_end_col = max(
            item["index"] + 1
            for item in self.parent_schema["date_columns"]
        )

        saturday_col = None
        sunday_col = None
        total_col = None
        created_weekend_cols = []

        last_date = pd.to_datetime(
            self.parent_schema["date_columns"][-1]["name"].split()[0]
            + f"-{self.year}",
            format="%d-%b-%Y",
        )

        if last_date.weekday() == 5:
            saturday_col = date_end_col
            sunday_col = date_end_col + 1
        elif last_date.weekday() < 5:
            saturday_col = date_end_col + 1
            sunday_col = date_end_col + 2
        else:
            saturday_col = date_end_col - 1
            sunday_col = date_end_col

        for col in range(1, self.sheet.max_column + 1):
            value = self.sheet.cell(header_row, col).value

            if value == "Total Hrs":
                total_col = col

        # Missing weekend placeholders belong immediately after the last date.
        if saturday_col is not None or sunday_col is not None:
            required_last_col = max(
                col for col in (saturday_col, sunday_col) if col is not None
            )

            expected_headers = {
                column: label
                for column, label in (
                    (saturday_col, "ASPIRE"),
                    (sunday_col, "FG"),
                )
                if column is not None
            }
            needs_insert = any(
                self.sheet.cell(header_row, column).value != label
                for column, label in expected_headers.items()
            )

            can_reuse_existing_columns = (
                total_col is not None
                and total_col > required_last_col
            )

            if needs_insert and not can_reuse_existing_columns:
                columns_to_add = required_last_col - date_end_col
                self.sheet.insert_cols(
                    date_end_col + 1,
                    columns_to_add,
                )
                created_weekend_cols = list(range(
                    date_end_col + 1,
                    required_last_col + 1,
                ))

            elif needs_insert:
                created_weekend_cols = list(expected_headers)

            if saturday_col is not None:
                self.sheet.cell(header_row, saturday_col).value = "ASPIRE"
            if sunday_col is not None:
                self.sheet.cell(header_row, sunday_col).value = "FG"

            for column in created_weekend_cols:
                for row in range(1, last_employee_row + 1):
                    source_cell = self.sheet.cell(row, date_end_col)
                    target_cell = self.sheet.cell(row, column)
                    target_cell._style = copy(source_cell._style)
                    target_cell.number_format = source_cell.number_format

                if column in (saturday_col, sunday_col):
                    for row in range(header_row + 1, last_employee_row + 1):
                        self.sheet.cell(row, column).value = 0

            total_col = None
            for col in range(1, self.sheet.max_column + 1):
                if self.sheet.cell(header_row, col).value == "Total Hrs":
                    total_col = col

        # Total Hrs must exist after all governance columns.
        if total_col is None or total_col <= max(saturday_col, sunday_col):
            total_col = max(self.sheet.max_column, sunday_col) + 1
            self.sheet.cell(header_row, total_col).value = "Total Hrs"

            for row in range(1, last_employee_row + 1):
                source_cell = self.sheet.cell(row, date_end_col)
                target_cell = self.sheet.cell(row, total_col)
                target_cell._style = copy(source_cell._style)
                target_cell.number_format = source_cell.number_format

        self.governance_saturday_col = saturday_col
        self.governance_sunday_col = sunday_col
        self.total_hours_col = total_col

    def update_week_headers(self):
        """
        Replace comparison headers using schema.

        Normal week:
            Saturday -> ASPIRE
            Sunday   -> FG

        Last incomplete week:
            Remaining columns become ASPIRE and FG.
        """

        header_row = self.parent_schema["header_row"] + 1

        self.ensure_governance_columns()

        columns = self.parent_schema["columns"]

        date_columns = self.parent_schema["date_columns"]

        date_map = {}

        for item in date_columns:

            dt = pd.to_datetime(
                item["name"].split()[0] + f"-{self.year}",
                format="%d-%b-%Y",
            )

            date_map[item["index"]] = dt.weekday()

        # ------------------------------------------------------------
        # Normal weeks
        # ------------------------------------------------------------

        for index, header in enumerate(columns):

            excel_col = index + 1

            if index in date_map:

                weekday = date_map[index]

                # Saturday
                if weekday == 5:

                    self.sheet.cell(
                        header_row,
                        excel_col,
                    ).value = "ASPIRE"

                # Sunday
                elif weekday == 6:

                    self.sheet.cell(
                        header_row,
                        excel_col,
                    ).value = "FG"

        # ------------------------------------------------------------
        # Last incomplete week
        # ------------------------------------------------------------

        last_date = pd.to_datetime(
            date_columns[-1]["name"].split()[0] + f"-{self.year}",
            format="%d-%b-%Y",
        )

        last_weekday = last_date.weekday()

        if last_weekday < 5:

            next_index = date_columns[-1]["index"] + 1

            labels = ["ASPIRE", "FG"]

            label = 0

            while (
                next_index < len(columns)
                and columns[next_index] != "Total Hrs"
                and label < 2
            ):

                self.sheet.cell(
                    header_row,
                    next_index + 1,
                ).value = labels[label]

                label += 1
                next_index += 1

            self.sheet.cell(
                header_row,
                self.governance_saturday_col,
            ).value = "ASPIRE"
            self.sheet.cell(
                header_row,
                self.governance_sunday_col,
            ).value = "FG"

    # ================================================================
    # Normalize Employee Name
    # ================================================================

    def normalize_name(self, value):

        if pd.isna(value):
            return ""

        normalized = str(value).strip().lower()

        normalized = normalized.replace(",", " ")

        normalized = re.sub(
            r"\s+",
            " ",
            normalized,
        )

        return normalized

    # ================================================================
    # Resolve Client Column Names
    # ================================================================

    def find_column(self, df_columns, candidates):

        # ------------------------------------------------------------
        # Direct match
        # ------------------------------------------------------------

        for c in candidates:

            if c is None:
                continue

            if c in df_columns:
                return c

        # ------------------------------------------------------------
        # Case-insensitive / normalized match
        # ------------------------------------------------------------

        def norm(s):

            return re.sub(
                r"\s+",
                " ",
                str(s).strip().lower(),
            )

        normalized_cols = {
            norm(c): c
            for c in df_columns
        }

        for c in candidates:

            if c is None:
                continue

            nc = norm(c)

            if nc in normalized_cols:
                return normalized_cols[nc]

        return None

    # ------------------------------------------------------------

    def resolve_client_columns(self):

        cols = list(self.client_df.columns)

        # ------------------------------------------------------------
        # Employee column from schema
        # ------------------------------------------------------------

        client_emp = (
            self.client_schema
            .get("employee_column", {})
            .get("name")
        )

        # ------------------------------------------------------------
        # AI-identified columns (from schema_analyzer.py / Gemini)
        #
        # Preferred source of truth. The hardcoded candidate lists
        # below only kick in when the schema doesn't have a value
        # (older schema.json, or Gemini couldn't identify the column
        # for this particular sheet).
        # ------------------------------------------------------------

        def from_schema(field_name):

            info = self.client_schema.get(field_name)

            if not info:
                return None

            name = info.get("name")

            if name in cols:
                return name

            return None

        # ------------------------------------------------------------
        # Candidate columns (fallback only)
        # ------------------------------------------------------------

        start_candidates = [
            "Time Sheet Start Date",
            "Start Date",
            "StartDate",
            "TimeSheet Start Date",
            "Time Sheet Start",
        ]

        end_candidates = [
            "Time Sheet End Date",
            "End Date",
            "EndDate",
            "Time Sheet End",
        ]

        worker_candidates = [
            client_emp,
            "Worker",
            "Employee",
            "Emp_Name",
            "Employee Name",
            "Worker Name",
        ]

        billable_candidates = [
            "Billable Hours (ST)",
            "Billable Hours",
            "Hours",
            "Total Hours",
            "Billable",
        ]

        status_candidates = [
            "Time Sheet Status",
            "Status",
            "Timesheet Status",
        ]

        # ------------------------------------------------------------
        # Resolve columns
        # ------------------------------------------------------------

        self.client_start_col = (
            from_schema("week_start_column")
            or self.find_column(cols, start_candidates)
        )

        self.client_end_col = (
            from_schema("week_end_column")
            or self.find_column(cols, end_candidates)
        )

        self.client_worker_col = self.find_column(
            cols,
            worker_candidates,
        )

        self.client_billable_col = (
            from_schema("billable_hours_column")
            or self.find_column(cols, billable_candidates)
        )

        self.client_status_col = (
            from_schema("status_column")
            or self.find_column(cols, status_candidates)
        )

        # ------------------------------------------------------------
        # Fallback employee column using schema index
        # ------------------------------------------------------------

        if (
            self.client_worker_col is None
            and "employee_column" in self.client_schema
        ):

            idx = (
                self.client_schema["employee_column"]
                .get("index")
            )

            try:

                if (
                    idx is not None
                    and 0 <= idx < len(cols)
                ):
                    self.client_worker_col = cols[idx]

            except Exception:
                pass

        # ------------------------------------------------------------
        # Warnings
        # ------------------------------------------------------------

        missing = []

        if self.client_start_col is None:
            missing.append("start date")

        if self.client_end_col is None:
            missing.append("end date")

        if self.client_worker_col is None:
            missing.append("worker/employee")

        if missing:

            print(
                "WARNING: Could not resolve client columns: "
                + ", ".join(missing)
                + ". Update schema or column names."
            )

    # ================================================================
    # Find Parent Excel Row
    # ================================================================

    def get_parent_excel_row(self, parent_name):

        employee_col = (
            self.parent_schema["employee_column"]["name"]
        )

        target_name = self.normalize_name(
            parent_name
        )

        rows = self.parent_df[
            self.parent_df[employee_col].map(
                self.normalize_name
            ) == target_name
        ]

        if rows.empty:

            print(
                f"\nWARNING: Parent employee not found: "
                f"{parent_name}"
            )

            return None

        return (
            int(rows.index[0])
            + self.parent_schema["header_row"]
            + 2
        )

    # ================================================================
    # Get Client Rows
    # ================================================================

    @staticmethod
    def _needs_manual_review(values):
        """
        True only when a duplicate employee/week group's offsetting
        +/- Billable Hours pair (e.g. 40.00 then -40.00) completes at
        the VERY LAST row, with nothing entered afterward.

        That means the reversal was never followed by a settling
        entry, so there is no reliable final value -- a human must
        confirm it.

        If ANY row (any status, any value, including 0) was entered
        after the pair completed, that later row is trusted as the
        settled final answer and no review is needed -- this is what
        makes the representative row (the group's last row) the
        correct one to use automatically in that case.
        """

        seen = set()

        pair_complete_index = None

        for index, value in enumerate(values):

            rounded = round(float(value), 2)

            if rounded == 0:
                continue

            if -rounded in seen:
                # Consume the matched counterpart so a later,
                # unrelated repeat of the same value can't falsely
                # re-trigger this same pair.
                pair_complete_index = index
                seen.discard(-rounded)
            else:
                seen.add(rounded)

        if pair_complete_index is None:
            return False

        return pair_complete_index == len(values) - 1

    def get_client_rows(self, client_name):
        """
        Get client records for one employee.

        Returns (rows, review_map):

        - `rows` has exactly one representative row per employee/week.
        - `review_map` maps (week_start, week_end) -> the list of raw
          Billable Hours values for any week whose offsetting +/- pair
          completed at the last row with nothing after it (see
          _needs_manual_review).

        RULE:
        When an employee has multiple rows for the same
        Time Sheet Start/End Date, the row that is LAST in the
        original sheet order is the final/authoritative state for
        that week (corrections are appended after the original
        entry), regardless of its status. If that duplicate group's
        Billable Hours contain an offsetting +/- pair, the week is
        also flagged in `review_map` so governance can defer its
        Green/Red decision to a human instead of trusting the
        automatic pick.
        """

        target_name = self.normalize_name(
            client_name
        )

        # ------------------------------------------------------------
        # Find employee rows
        # ------------------------------------------------------------

        if (
            self.client_worker_col is not None
            and self.client_worker_col in self.client_df.columns
        ):

            mask = (
                self.client_df[
                    self.client_worker_col
                ].map(self.normalize_name)
                == target_name
            )

        else:

            mask = False

            for c in self.client_df.columns:

                try:

                    mask = mask | (
                        self.client_df[c].map(
                            self.normalize_name
                        )
                        == target_name
                    )

                except Exception:

                    continue

        rows = self.client_df[mask].copy()

        if rows.empty:
            return rows, {}

        # ------------------------------------------------------------
        # Validate dates
        # ------------------------------------------------------------

        if (
            self.client_start_col is None
            or self.client_end_col is None
        ):

            print(
                "WARNING: Client start/end date columns "
                "not resolved; skipping rows for this client."
            )

            return rows.iloc[0:0], {}

        # ------------------------------------------------------------
        # Normalize dates
        # ------------------------------------------------------------

        rows[self.client_start_col] = pd.to_datetime(
            rows[self.client_start_col],
            errors="coerce",
        )

        rows[self.client_end_col] = pd.to_datetime(
            rows[self.client_end_col],
            errors="coerce",
        )

        # Remove rows without valid weekly dates
        rows = rows[
            rows[self.client_start_col].notna()
            & rows[self.client_end_col].notna()
        ].copy()

        if rows.empty:
            return rows, {}

        # ------------------------------------------------------------
        # Normalize hours
        # ------------------------------------------------------------

        if (
            self.client_billable_col is not None
            and self.client_billable_col in rows.columns
        ):

            rows[self.client_billable_col] = pd.to_numeric(
                rows[self.client_billable_col],
                errors="coerce",
            ).fillna(0)

        else:

            rows["_billable_missing"] = 0

            self.client_billable_col = (
                "_billable_missing"
            )

        # ------------------------------------------------------------
        # Normalize status
        # ------------------------------------------------------------

        if (
            self.client_status_col is not None
            and self.client_status_col in rows.columns
        ):

            rows[self.client_status_col] = (
                rows[self.client_status_col]
                .fillna("")
                .astype(str)
                .str.strip()
            )

        else:

            rows["_status_missing"] = ""

            self.client_status_col = (
                "_status_missing"
            )

        # ------------------------------------------------------------
        # Normalize status for comparison
        # ------------------------------------------------------------

        rows["_normalized_status"] = (
            rows[self.client_status_col]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.lower()
        )

        # ------------------------------------------------------------
        # Identify Invoiced
        #
        # Only an EXACT normalized "invoiced" status is treated
        # as Invoiced.
        # ------------------------------------------------------------

        rows["_is_invoiced"] = (
            rows["_normalized_status"] == "invoiced"
        )

        # ------------------------------------------------------------
        # Resolve duplicate employee + week records
        #
        # For each (Start Date, End Date) group, keep the row that is
        # LAST in the original sheet order as the final state for
        # that week, and flag the week if the group's Billable Hours
        # contain an offsetting +/- pair (correction/reversal).
        # ------------------------------------------------------------

        review_map = {}

        final_rows = []

        for (week_start, week_end), group in rows.groupby(
            [
                self.client_start_col,
                self.client_end_col,
            ],
            sort=False,
        ):

            group = group.sort_index()

            final_rows.append(group.iloc[-1])

            if len(group) > 1:

                billable_values = group[
                    self.client_billable_col
                ].tolist()

                if self._needs_manual_review(billable_values):

                    review_map[
                        (week_start, week_end)
                    ] = billable_values

        rows = pd.DataFrame(
            final_rows,
            columns=rows.columns,
        )

        # ------------------------------------------------------------
        # Remove helper columns
        # ------------------------------------------------------------

        rows = rows.drop(
            columns=[
                "_normalized_status",
                "_is_invoiced",
            ],
            errors="ignore",
        )

        return rows, review_map

    # ================================================================
    # Find Week Cells
    # ================================================================

    def find_week_cells(self, start, end):

        saturday = None

        sunday = None

        current = start

        while current <= end:

            if current.date() in self.parent_date_map:

                col = (
                    self.parent_date_map[
                        current.date()
                    ]
                    + 1
                )

                if current.weekday() == 5:

                    saturday = col

                elif current.weekday() == 6:

                    sunday = col

            current += timedelta(days=1)

        return saturday, sunday

    # ================================================================
    # Process One Week
    # ================================================================

        # ================================================================
    # Process One Week
    # ================================================================

    def process_week(
        self,
        excel_row,
        client_row,
        client_name=None,
        parent_name=None,
        original_values=None,
    ):
        """
        Process one client weekly record.

        Governance rules:

        1. EXACTLY "Invoiced"
           - ASPIRE = Parent Mon-Fri total
           - FG = Client weekly hours
           - Compare hours
           - Match = GREEN
           - Mismatch = RED

        2. ANYTHING OTHER THAN "Invoiced"
           - ASPIRE = Parent Mon-Fri total
           - FG = Client status
           - ASPIRE = RED
           - FG = RED
           - Client hours are NEVER compared

        IMPORTANT:
        A non-Invoiced status can NEVER enter the hours-comparison
        logic, even for an incomplete or cross-month week.

        If `original_values` is provided, this week had multiple
        duplicate employee/week rows with an offsetting +/- Billable
        Hours pair (a correction/reversal). The Green/Red decision is
        deferred: the cell is marked for manual review instead, and
        finalized later by apply_pending_corrections() once a human
        confirms the correct Billable Hours value.
        """

        needs_review = original_values is not None

        # ------------------------------------------------------------
        # Read week dates
        # ------------------------------------------------------------

        start = pd.to_datetime(
            client_row[self.client_start_col]
        )

        end = pd.to_datetime(
            client_row[self.client_end_col]
        )

        # ------------------------------------------------------------
        # Client weekly hours
        # ------------------------------------------------------------

        try:
            raw_client_hours = client_row.get(
                self.client_billable_col,
                0,
            )

            client_total = float(
                raw_client_hours or 0
            )

        except (
            ValueError,
            TypeError,
        ):
            client_total = 0.0

        # ------------------------------------------------------------
        # Client status
        # ------------------------------------------------------------

        raw_status = client_row.get(
            self.client_status_col,
            "",
        )

        if pd.isna(raw_status):
            client_status = ""
        else:
            client_status = str(
                raw_status
            ).strip()

        normalized_status = (
            re.sub(
                r"\s+",
                " ",
                client_status,
            )
            .strip()
            .lower()
        )

        # ------------------------------------------------------------
        # CRITICAL GOVERNANCE DECISION
        #
        # ONLY exact normalized "invoiced" is Invoiced.
        #
        # Everything else is Non-Invoiced.
        # ------------------------------------------------------------

        is_invoiced = (
            normalized_status == "invoiced"
        )

        # (Debug prints removed)

        # ------------------------------------------------------------
        # Read Parent Monday-Friday hours
        # ------------------------------------------------------------

        parent_total = 0.0

        current = start

        available_days = []

        while current <= end:

            if (
                current.date()
                in self.parent_date_map
                and current.weekday() < 5
            ):

                col = (
                    self.parent_date_map[
                        current.date()
                    ]
                    + 1
                )

                value = self.sheet.cell(
                    excel_row,
                    col,
                ).value

                try:

                    parent_total += float(
                        value or 0
                    )

                except (
                    ValueError,
                    TypeError,
                ):

                    pass

                available_days.append(
                    col
                )

            current += timedelta(days=1)

        # ------------------------------------------------------------
        # No Monday-Friday columns
        # ------------------------------------------------------------

        if not available_days:

            print(
                f"\nWARNING: No working-day columns "
                f"found for employee row {excel_row}"
            )

            return

        available_working_days = len(
            available_days
        )

        # ------------------------------------------------------------
        # Find last working-day column
        # ------------------------------------------------------------

        last_working_col = max(
            available_days
        )

        # ------------------------------------------------------------
        # Find Saturday / Sunday columns
        # ------------------------------------------------------------

        saturday = None
        sunday = None

        current = start

        while current <= end:

            if (
                current.date()
                in self.parent_date_map
            ):

                col = (
                    self.parent_date_map[
                        current.date()
                    ]
                    + 1
                )

                if current.weekday() == 5:

                    saturday = col

                elif current.weekday() == 6:

                    sunday = col

            current += timedelta(days=1)

        # ------------------------------------------------------------
        # Last incomplete week
        #
        # For 27-Apr -> 03-May:
        #
        # Apr 27 = Monday
        # Apr 28 = Tuesday
        # Apr 29 = Wednesday
        # Apr 30 = Thursday
        #
        # Parent total = 32
        #
        # Friday/Saturday placeholder columns are used for:
        # ASPIRE / FG
        # ------------------------------------------------------------

        if saturday is None:

            saturday = (
                last_working_col + 1
            )

        if sunday is None:

            sunday = (
                last_working_col + 2
            )

        # ------------------------------------------------------------
        # GOVERNANCE RESULT
        #
        # A flagged "needs review" week is NOT decided here. It gets
        # the pending-review color and its inputs are recorded so
        # apply_pending_corrections() can run this SAME rule later
        # once a human confirms the correct Billable Hours.
        # ------------------------------------------------------------

        if needs_review:

            display_fg_value = round(client_total, 2)
            fill = self.pending_review_fill
            is_red = None

        else:

            display_fg_value, fill, is_red = evaluate_governance(
                is_invoiced,
                parent_total,
                client_total,
                client_status,
                available_working_days,
                self.green_fill,
                self.red_fill,
            )

        # ------------------------------------------------------------
        # Write ASPIRE
        # ------------------------------------------------------------

        aspire_cell = self.sheet.cell(
            excel_row,
            saturday,
        )

        aspire_cell.value = round(
            parent_total,
            2,
        )

        # ------------------------------------------------------------
        # Write FG
        # ------------------------------------------------------------

        FG_cell = self.sheet.cell(
            excel_row,
            sunday,
        )

        FG_cell.value = display_fg_value

        # ------------------------------------------------------------
        # Apply fill
        # ------------------------------------------------------------

        aspire_cell.fill = fill
        FG_cell.fill = fill

        # ------------------------------------------------------------
        # Track result
        # ------------------------------------------------------------

        if needs_review:

            self.pending_review_cells.add((excel_row, saturday))
            self.pending_review_cells.add((excel_row, sunday))
            self.pending_review_employees.add(excel_row)

            self.pending_review.append(
                {
                    "key": (
                        f"{parent_name}||{start.date().isoformat()}"
                        f"||{end.date().isoformat()}"
                    ),
                    "client_name": client_name,
                    "parent_name": parent_name,
                    "excel_row": excel_row,
                    "aspire_col": saturday,
                    "fg_col": sunday,
                    "week_start": start.date().isoformat(),
                    "week_end": end.date().isoformat(),
                    "parent_total": round(parent_total, 2),
                    "available_working_days": available_working_days,
                    "is_invoiced": is_invoiced,
                    "client_status": client_status,
                    "original_values": [
                        round(float(v), 2) for v in original_values
                    ],
                    "suggested_value": round(client_total, 2),
                }
            )

        elif is_red:

            self.red_employees.add(
                excel_row
            )

            self.red_governance_cells.add(
                (
                    excel_row,
                    saturday,
                )
            )

            self.red_governance_cells.add(
                (
                    excel_row,
                    sunday,
                )
            )

    # ================================================================
    # Calculate Total Hours
    # ================================================================

    def calculate_total_hours(self):
        """
        Sum only ASPIRE column values and write them into
        the 'Total Hrs' column.

        Do NOT write 0 for completely empty employee rows.
        """

        header_row = (
            self.parent_schema["header_row"]
            + 1
        )

        total_col = self.total_hours_col

        aspire_cols = []

        for col in range(
            1,
            self.sheet.max_column + 1,
        ):

            if (
                self.sheet.cell(
                    header_row,
                    col,
                ).value
                == "ASPIRE"
            ):

                aspire_cols.append(col)

        start_row = header_row + 1

        # ------------------------------------------------------------
        # Employee name column
        # ------------------------------------------------------------

        employee_col = (
            self.parent_schema[
                "employee_column"
            ]["index"]
            + 1
        )

        # ------------------------------------------------------------
        # Process employee rows
        # ------------------------------------------------------------

        for row in range(
            start_row,
            self.sheet.max_row + 1,
        ):

            employee_value = self.sheet.cell(
                row,
                employee_col,
            ).value

            # --------------------------------------------------------
            # Empty employee row
            # --------------------------------------------------------

            if (
                employee_value is None
                or str(employee_value).strip() == ""
            ):

                self.sheet.cell(
                    row,
                    total_col,
                ).value = None

                continue

            # --------------------------------------------------------
            # Calculate ASPIRE total
            # --------------------------------------------------------

            total = 0.0

            for col in aspire_cols:

                value = self.sheet.cell(
                    row,
                    col,
                ).value

                try:

                    total += float(
                        value or 0
                    )

                except (
                    ValueError,
                    TypeError,
                ):

                    pass

            self.sheet.cell(
                row,
                total_col,
            ).value = round(
                total,
                2,
            )

    # ================================================================
    # Highlight Zero Matches
    # ================================================================

    def highlight_zero_matches(self):
        """
        Highlight ASPIRE and FG green when both are numeric zeros.

        IMPORTANT:
        This function must NEVER overwrite cells that were explicitly
        marked RED by process_week().

        Therefore:
            - Invoiced 0/0 -> GREEN
            - Non-Invoiced 0/status -> remains RED
        """

        header_row = (
            self.parent_schema["header_row"]
            + 1
        )

        # ------------------------------------------------------------
        # Find all ASPIRE / FG pairs
        # ------------------------------------------------------------

        pairs = []

        for col in range(
            1,
            self.sheet.max_column,
        ):

            if (
                self.sheet.cell(
                    header_row,
                    col,
                ).value
                == "ASPIRE"
                and self.sheet.cell(
                    header_row,
                    col + 1,
                ).value
                == "FG"
            ):

                pairs.append(
                    (
                        col,
                        col + 1,
                    )
                )

        # ------------------------------------------------------------
        # Employee column
        # ------------------------------------------------------------

        employee_col = (
            self.parent_schema[
                "employee_column"
            ]["index"]
            + 1
        )

        # ------------------------------------------------------------
        # Process rows
        # ------------------------------------------------------------

        for row in range(
            header_row + 1,
            self.sheet.max_row + 1,
        ):

            emp_val = self.sheet.cell(
                row,
                employee_col,
            ).value

            if (
                emp_val is None
                or str(emp_val).strip() == ""
            ):
                continue

            for aspire_col, FG_col in pairs:

                # ----------------------------------------------------
                # Never overwrite explicit RED governance cells
                # ----------------------------------------------------

                if (
                    row,
                    aspire_col,
                ) in self.red_governance_cells:

                    continue

                if (
                    row,
                    FG_col,
                ) in self.red_governance_cells:

                    continue

                # ----------------------------------------------------
                # Never overwrite a pending-review cell either; its
                # color is not final until a human confirms the
                # Billable Hours value.
                # ----------------------------------------------------

                if (
                    row,
                    aspire_col,
                ) in self.pending_review_cells:

                    continue

                if (
                    row,
                    FG_col,
                ) in self.pending_review_cells:

                    continue

                aspire_cell = self.sheet.cell(
                    row,
                    aspire_col,
                )

                FG_cell = self.sheet.cell(
                    row,
                    FG_col,
                )

                # ----------------------------------------------------
                # Only numeric values qualify as 0/0
                # ----------------------------------------------------

                try:

                    aspire_val = float(
                        aspire_cell.value
                    )

                    FG_val = float(
                        FG_cell.value
                    )

                except (
                    ValueError,
                    TypeError,
                ):

                    continue

                # ----------------------------------------------------
                # Numeric zero / zero
                # ----------------------------------------------------

                if (
                    abs(aspire_val) < 0.01
                    and abs(FG_val) < 0.01
                ):

                    aspire_cell.fill = (
                        self.green_fill
                    )

                    FG_cell.fill = (
                        self.green_fill
                    )

    # ================================================================
    # Highlight Employee Names
    # ================================================================

    def highlight_employee_names(self):
        """
        Highlight the employee name based on weekly ASPIRE/FG results.

        Logic:
        - If ANY weekly result is RED:
            employee name = RED
        - If NO weekly result is RED:
            employee name = GREEN

        The decision is based on process_week(), not Excel colors.
        """

        header_row = (
            self.parent_schema["header_row"]
            + 1
        )

        employee_col = (
            self.parent_schema[
                "employee_column"
            ]["index"]
            + 1
        )

        for row in range(
            header_row + 1,
            self.sheet.max_row + 1,
        ):

            employee_cell = self.sheet.cell(
                row,
                employee_col,
            )

            employee_value = (
                employee_cell.value
            )

            # --------------------------------------------------------
            # Skip empty parent rows
            # --------------------------------------------------------

            if (
                employee_value is None
                or str(employee_value).strip() == ""
            ):

                continue

            # --------------------------------------------------------
            # Client-only rows
            # --------------------------------------------------------

            if (
                row
                in self.client_only_excel_rows
            ):

                continue

            # --------------------------------------------------------
            # At least one week still awaiting manual review
            #
            # The overall Red/Green verdict for this employee isn't
            # final until that week is resolved, so don't commit to
            # either color yet.
            # --------------------------------------------------------

            if row in self.pending_review_employees:

                employee_cell.fill = (
                    self.pending_review_fill
                )

            # --------------------------------------------------------
            # At least one weekly result RED
            # --------------------------------------------------------

            elif row in self.red_employees:

                employee_cell.fill = (
                    self.red_fill
                )

            # --------------------------------------------------------
            # No weekly RED result
            # --------------------------------------------------------

            else:

                employee_cell.fill = (
                    self.green_fill
                )

    # ================================================================
    # Highlight Missing Client Row
    # ================================================================

    def highlight_missing_client_row(
        self,
        excel_row,
    ):
        """
        Highlight the entire parent employee row when
        the employee has no corresponding record in the client sheet.
        """

        for col in range(
            1,
            self.sheet.max_column + 1,
        ):

            self.sheet.cell(
                excel_row,
                col,
            ).fill = (
                self.missing_client_fill
            )

    # ================================================================
    # Append Missing Parent Employees
    # ================================================================

    def append_missing_parent_employees(self):
        """
        Append Client-only employees to the end of the Parent sheet.

        Client-only employees are intentionally excluded from weekly
        governance processing because there is no Parent employee row
        to compare against.

        They are still added to the final workbook and highlighted
        with CLIENT_ONLY_COLOR.
        """

        combined = []

        seen = set()

        # ------------------------------------------------------------
        # Add item helper
        # ------------------------------------------------------------

        def add_item(item):

            if isinstance(item, str):

                client_name = item
                parent_name = None

            else:

                client_name = item.get(
                    "client_name"
                )

                parent_name = item.get(
                    "parent_name"
                )

            if (
                client_name is None
                or str(client_name).strip() == ""
            ):

                return

            key = self.normalize_name(
                client_name
            )

            if not key or key in seen:

                return

            seen.add(key)

            combined.append(
                {
                    "client_name": client_name,
                    "parent_name": parent_name,
                }
            )

        # ------------------------------------------------------------
        # Existing missing-parent employees
        # ------------------------------------------------------------

        for item in self.missing_parent_employees:

            add_item(item)

        # ------------------------------------------------------------
        # Explicit client-only employees
        # ------------------------------------------------------------

        for client_name in self.mapping_data.get(
            "client_only",
            [],
        ):

            add_item(client_name)

        if not combined:
            return

        # ------------------------------------------------------------
        # Header and employee columns
        # ------------------------------------------------------------

        header_row = (
            self.parent_schema["header_row"]
            + 1
        )

        employee_col = (
            self.parent_schema[
                "employee_column"
            ]["index"]
            + 1
        )

        # ------------------------------------------------------------
        # Find last non-empty employee row
        # ------------------------------------------------------------

        last_row = header_row

        for row in range(
            header_row + 1,
            self.sheet.max_row + 1,
        ):

            val = self.sheet.cell(
                row,
                employee_col,
            ).value

            if (
                val is not None
                and str(val).strip() != ""
            ):

                last_row = row

        start_row = last_row + 1

        # ------------------------------------------------------------
        # Append employees
        # ------------------------------------------------------------

        for item in combined:

            name = (
                item["client_name"]
                or item["parent_name"]
            )

            self.sheet.cell(
                start_row,
                employee_col,
            ).value = name

            # --------------------------------------------------------
            # Highlight complete appended row
            # --------------------------------------------------------

            for col in range(
                1,
                self.sheet.max_column + 1,
            ):

                self.sheet.cell(
                    start_row,
                    col,
                ).fill = (
                    self.client_only_fill
                )

            self.client_only_excel_rows.add(
                start_row
            )

            start_row += 1

        # ------------------------------------------------------------
        # Synchronize in-memory list
        # ------------------------------------------------------------

        self.missing_parent_employees = (
            combined
        )

    # ================================================================
    # Run Governance
    # ================================================================

    def run(self):

        # ============================================================
        # Update Parent headers
        # ============================================================

        self.update_week_headers()

        # ============================================================
        # Process employee mappings
        # ============================================================

        mappings = self.mapping_data[
            "mappings"
        ]

        for mapping in mappings:

            client_name = mapping[
                "client_name"
            ]

            parent_name = mapping[
                "parent_name"
            ]

            # --------------------------------------------------------
            # Find parent employee row
            # --------------------------------------------------------

            excel_row = (
                self.get_parent_excel_row(
                    parent_name
                )
            )

            if excel_row is None:

                self.missing_parent_employees.append(
                    {
                        "client_name": client_name,
                        "parent_name": parent_name,
                    }
                )

                continue

            # --------------------------------------------------------
            # Get client weekly rows
            # --------------------------------------------------------

            rows, review_map = self.get_client_rows(
                client_name
            )

            if rows.empty:

                print(
                    f"WARNING: No client records found "
                    f"for {client_name}"
                )

                self.highlight_missing_client_row(
                    excel_row
                )

                continue

            # --------------------------------------------------------
            # Process every weekly row
            # --------------------------------------------------------

            for _, row in rows.iterrows():

                week_key = (
                    row[self.client_start_col],
                    row[self.client_end_col],
                )

                self.process_week(
                    excel_row,
                    row,
                    client_name=client_name,
                    parent_name=parent_name,
                    original_values=review_map.get(week_key),
                )

        # ============================================================
        # Append Client-only employees
        # ============================================================

        self.append_missing_parent_employees()

        # ============================================================
        # Calculate Total Hrs
        # ============================================================

        self.calculate_total_hours()

        # ============================================================
        # Highlight zero matches
        #
        # This cannot override explicit RED governance results.
        # ============================================================

        self.highlight_zero_matches()

        # ============================================================
        # Highlight employee names
        # ============================================================

        self.highlight_employee_names()

        # ============================================================
        # Save output
        # ============================================================

        OUTPUT_FILE.parent.mkdir(
            parents=True,
            exist_ok=True,
        )

        self.workbook.save(
            OUTPUT_FILE
        )

        # ============================================================
        # Save pending reviews (weeks needing a human-confirmed
        # Billable Hours value before Green/Red can be decided)
        # ============================================================

        PENDING_REVIEW_FILE.parent.mkdir(
            parents=True,
            exist_ok=True,
        )

        with open(
            PENDING_REVIEW_FILE,
            "w",
            encoding="utf-8",
        ) as f:

            json.dump(
                self.pending_review,
                f,
                indent=4,
            )

        if self.pending_review:

            print(
                f"\n{len(self.pending_review)} week(s) need manual "
                f"review before final Green/Red colors are applied. "
                f"See {PENDING_REVIEW_FILE}"
            )

        print(
            "\nOutput saved to file"
        )


# ====================================================================
# Apply Pending Corrections
#
# Called by the UI (after a human confirms the Billable Hours for
# every flagged week) to finalize the weeks that HoursGovernance.run()
# deferred. Uses the exact same evaluate_governance() rule as the
# original run, so "the colouring logic applies the same" once real
# input is available.
# ====================================================================

def apply_pending_corrections(corrections):
    """
    corrections: dict mapping a pending entry's "key" to the
    human-confirmed Billable Hours value.

    Returns the number of weeks resolved (0 if there was nothing
    pending). Safe to call even if PENDING_REVIEW_FILE is missing or
    empty.
    """

    if not PENDING_REVIEW_FILE.exists():
        return 0

    with open(PENDING_REVIEW_FILE, "r", encoding="utf-8") as f:
        pending = json.load(f)

    if not pending:
        return 0

    with open("schema.json", "r", encoding="utf-8") as f:
        schema = json.load(f)

    parent_schema = schema["parent"]

    workbook = load_workbook(OUTPUT_FILE)
    sheet = workbook.active

    green_fill = PatternFill(
        fill_type="solid",
        start_color=MATCH_COLOR,
        end_color=MATCH_COLOR,
    )

    red_fill = PatternFill(
        fill_type="solid",
        start_color=MISMATCH_COLOR,
        end_color=MISMATCH_COLOR,
    )

    touched_rows = set()

    for entry in pending:

        try:
            corrected_value = float(
                corrections.get(entry["key"], entry["suggested_value"])
            )
        except (TypeError, ValueError):
            corrected_value = entry["suggested_value"]

        display_fg_value, fill, _ = evaluate_governance(
            entry["is_invoiced"],
            entry["parent_total"],
            corrected_value,
            entry["client_status"],
            entry["available_working_days"],
            green_fill,
            red_fill,
        )

        aspire_cell = sheet.cell(
            entry["excel_row"],
            entry["aspire_col"],
        )

        fg_cell = sheet.cell(
            entry["excel_row"],
            entry["fg_col"],
        )

        aspire_cell.value = round(entry["parent_total"], 2)
        aspire_cell.fill = fill

        fg_cell.value = display_fg_value
        fg_cell.fill = fill

        touched_rows.add(entry["excel_row"])

    # ----------------------------------------------------------------
    # Re-decide each touched employee's name color now that every
    # pending week for them has a final Green/Red result.
    # ----------------------------------------------------------------

    header_row = parent_schema["header_row"] + 1
    employee_col = parent_schema["employee_column"]["index"] + 1

    pairs = []

    for col in range(1, sheet.max_column):

        if (
            sheet.cell(header_row, col).value == "ASPIRE"
            and sheet.cell(header_row, col + 1).value == "FG"
        ):
            pairs.append((col, col + 1))

    def _is_red_fill(cell):

        try:
            rgb = cell.fill.start_color.rgb
        except Exception:
            return False

        if not rgb:
            return False

        rgb = str(rgb)[-6:]

        return rgb.upper() == MISMATCH_COLOR.upper()

    for row in touched_rows:

        employee_value = sheet.cell(row, employee_col).value

        if employee_value is None or str(employee_value).strip() == "":
            continue

        row_is_red = any(
            _is_red_fill(sheet.cell(row, aspire_col))
            or _is_red_fill(sheet.cell(row, fg_col))
            for aspire_col, fg_col in pairs
        )

        sheet.cell(row, employee_col).fill = (
            red_fill if row_is_red else green_fill
        )

    workbook.save(OUTPUT_FILE)

    # Mark every pending review as resolved.
    with open(PENDING_REVIEW_FILE, "w", encoding="utf-8") as f:
        json.dump([], f, indent=4)

    return len(pending)


# ====================================================================
# COMMAND LINE ENTRY POINT
# ====================================================================

if __name__ == "__main__":

    HoursGovernance().run()