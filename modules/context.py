from dataclasses import dataclass, field
from typing import Any


@dataclass
class ModuleDataContext:
    filtered_df: Any
    planner_source: Any
    product_summary: Any
    date_summary: Any
    weekly_summary: Any
    key_findings: Any
    prev_meta: dict
    curr_meta: dict
    date_basis: str
    selected_start_date: Any
    selected_end_date: Any
    auth_user: dict = field(default_factory=dict)
    user_role: str = "Viewer"
    module_access: str = "none"
    excel_bytes: bytes | None = None
    csv_bytes: bytes | None = None
    professional_excel_bytes: bytes | None = None
    reference: dict = field(default_factory=dict)
