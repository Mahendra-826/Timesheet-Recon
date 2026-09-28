import sys
import time
from pathlib import Path
from config import CLIENT_TIMESHEET, PARENT_TIMESHEET
from schema_manager import SchemaManager
from employee_mapper import EmployeeMapper
from hours_governance import HoursGovernance


def main():

    print("=" * 80, flush=True)
    print("TIMESHEET GOVERNANCE", flush=True)
    print("=" * 80, flush=True)

    # --------------------------------------------------
    # Determine input paths (allow overriding via CLI args)
    # --------------------------------------------------

    client_path = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(CLIENT_TIMESHEET)
    parent_path = Path(sys.argv[2]) if len(sys.argv) > 2 else Path(PARENT_TIMESHEET)

    print(f"\n[{time.strftime('%H:%M:%S')}] [1/3] Validating schema...", flush=True)

    manager = SchemaManager()

    manager.get_schema(client_path, parent_path)

    print(f"[{time.strftime('%H:%M:%S')}] Schema validation completed.", flush=True)

    # --------------------------------------------------
    # Step 2: Employee Mapping
    # --------------------------------------------------

    print(f"\n[{time.strftime('%H:%M:%S')}] [2/3] Running employee mapping...", flush=True)

    mapper = EmployeeMapper()

    mapper.run(client_path, parent_path)

    print(f"[{time.strftime('%H:%M:%S')}] Employee mapping completed.", flush=True)

    # --------------------------------------------------
    # Step 3: Hours Governance
    # --------------------------------------------------

    print(f"\n[{time.strftime('%H:%M:%S')}] [3/3] Running hours governance...", flush=True)

    governance = HoursGovernance(client_path, parent_path)

    governance.run()

    print(f"\n[{time.strftime('%H:%M:%S')}] " + "=" * 80, flush=True)
    print("PROCESS COMPLETED SUCCESSFULLY", flush=True)
    print("=" * 80, flush=True)


if __name__ == "__main__":
    main()