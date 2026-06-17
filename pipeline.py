"""
Data pipeline: parses a RingCentral Performance Report Excel export.

Methodology (locked):
- Unit of analysis = session (grouped by Session Id).
- Only Inbound call directions are included.
- Per-session outcome priority: Answered > VM/Abandoned > VM/Missed > Abandoned > Missed/Other.
- Park Off legs are ignored for outcome classification.
- Validation: every inbound session maps to exactly one outcome.
"""

from __future__ import annotations
import pandas as pd
from pathlib import Path
from dataclasses import dataclass, field
from typing import Optional


OUTCOME_ANSWERED = "Answered"
OUTCOME_VM_ABANDONED = "VM/Abandoned"
OUTCOME_VM_MISSED = "VM/Missed"
OUTCOME_ABANDONED = "Abandoned"
OUTCOME_MISSED = "Missed/Other"

IGNORED_RESULTS = {"Park Off"}


@dataclass
class QueueStats:
    name: str
    answered: int = 0
    vm_abandoned: int = 0
    vm_missed: int = 0
    abandoned: int = 0
    missed: int = 0

    @property
    def total(self) -> int:
        return self.answered + self.vm_abandoned + self.vm_missed + self.abandoned + self.missed

    @property
    def unanswered(self) -> int:
        return self.vm_abandoned + self.vm_missed + self.abandoned + self.missed

    @property
    def answer_rate(self) -> float:
        return self.answered / self.total if self.total else 0.0


@dataclass
class PipelineResult:
    total_inbound_sessions: int
    answered: int
    vm_abandoned: int
    vm_missed: int
    abandoned: int
    missed: int
    queue_stats: dict[str, QueueStats] = field(default_factory=dict)
    # Per-session detail for additional analysis
    sessions_df: Optional[pd.DataFrame] = None
    reconciliation_ok: bool = True
    reconciliation_note: str = ""

    @property
    def unanswered(self) -> int:
        return self.vm_abandoned + self.vm_missed + self.abandoned + self.missed

    @property
    def answer_rate(self) -> float:
        return self.answered / self.total_inbound_sessions if self.total_inbound_sessions else 0.0


def _classify_session(group: pd.DataFrame) -> str:
    """Return a single outcome for a session group (all legs of one session)."""
    results = set(group["Result"].dropna().str.strip())
    handle_times = group["Handle Time"].fillna(0)

    # Answered: any leg has Handle Time > 0, or Result == Answered
    if (handle_times > 0).any() or "Answered" in results:
        return OUTCOME_ANSWERED

    # Filter out ignored non-outcome results
    active_results = results - IGNORED_RESULTS

    if "VM/Abandoned" in active_results:
        return OUTCOME_VM_ABANDONED

    if "VM/Missed" in active_results:
        return OUTCOME_VM_MISSED

    if "Abandoned" in active_results:
        return OUTCOME_ABANDONED

    return OUTCOME_MISSED


def _parse_handle_time(series: pd.Series) -> pd.Series:
    """Convert Handle Time to total seconds (supports HH:MM:SS strings or numeric seconds)."""
    def to_seconds(val):
        if pd.isna(val):
            return 0.0
        if isinstance(val, (int, float)):
            return float(val)
        s = str(val).strip()
        if ":" in s:
            parts = s.split(":")
            try:
                if len(parts) == 3:
                    return int(parts[0]) * 3600 + int(parts[1]) * 60 + float(parts[2])
                if len(parts) == 2:
                    return int(parts[0]) * 60 + float(parts[1])
            except ValueError:
                return 0.0
        try:
            return float(s)
        except ValueError:
            return 0.0

    return series.apply(to_seconds)


def parse_report(path: Path) -> PipelineResult:
    # Try to find the sheet that looks like the Calls tab
    xl = pd.ExcelFile(path)
    sheet_name = xl.sheet_names[0]
    for name in xl.sheet_names:
        if "call" in name.lower():
            sheet_name = name
            break

    df = pd.read_excel(path, sheet_name=sheet_name, dtype=str)

    # Normalize column names
    df.columns = [c.strip() for c in df.columns]

    required = {"Session Id", "Call Direction", "Result", "Handle Time", "Queue"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Missing expected columns: {missing}. Found: {list(df.columns)}")

    # Filter to Inbound only
    df = df[df["Call Direction"].str.strip().str.lower() == "inbound"].copy()
    df["Handle Time"] = _parse_handle_time(df["Handle Time"])

    # Record file order for queue attribution (entry = first row, terminating = last row)
    df["_row_order"] = range(len(df))

    # Group by Session Id
    sessions = []
    for session_id, group in df.groupby("Session Id", sort=False):
        group = group.sort_values("_row_order")
        outcome = _classify_session(group)

        entry_queue = group.iloc[0]["Queue"] if pd.notna(group.iloc[0]["Queue"]) else ""
        term_queue = group.iloc[-1]["Queue"] if pd.notna(group.iloc[-1]["Queue"]) else ""

        sessions.append({
            "session_id": session_id,
            "outcome": outcome,
            "entry_queue": str(entry_queue).strip(),
            "terminating_queue": str(term_queue).strip(),
            "leg_count": len(group),
        })

    sessions_df = pd.DataFrame(sessions)

    # Aggregate totals
    outcome_counts = sessions_df["outcome"].value_counts().to_dict()
    answered = outcome_counts.get(OUTCOME_ANSWERED, 0)
    vm_abandoned = outcome_counts.get(OUTCOME_VM_ABANDONED, 0)
    vm_missed = outcome_counts.get(OUTCOME_VM_MISSED, 0)
    abandoned = outcome_counts.get(OUTCOME_ABANDONED, 0)
    missed = outcome_counts.get(OUTCOME_MISSED, 0)
    total = len(sessions_df)

    # Validation
    computed = answered + vm_abandoned + vm_missed + abandoned + missed
    recon_ok = computed == total
    recon_note = (
        f"OK: {computed} outcomes == {total} sessions"
        if recon_ok
        else f"MISMATCH: {computed} outcomes != {total} sessions"
    )

    # Per-queue stats (using entry queue as primary)
    queue_stats: dict[str, QueueStats] = {}
    for _, row in sessions_df.iterrows():
        q = row["entry_queue"] or "Unknown"
        if q not in queue_stats:
            queue_stats[q] = QueueStats(name=q)
        qs = queue_stats[q]
        o = row["outcome"]
        if o == OUTCOME_ANSWERED:
            qs.answered += 1
        elif o == OUTCOME_VM_ABANDONED:
            qs.vm_abandoned += 1
        elif o == OUTCOME_VM_MISSED:
            qs.vm_missed += 1
        elif o == OUTCOME_ABANDONED:
            qs.abandoned += 1
        else:
            qs.missed += 1

    return PipelineResult(
        total_inbound_sessions=total,
        answered=answered,
        vm_abandoned=vm_abandoned,
        vm_missed=vm_missed,
        abandoned=abandoned,
        missed=missed,
        queue_stats=queue_stats,
        sessions_df=sessions_df,
        reconciliation_ok=recon_ok,
        reconciliation_note=recon_note,
    )
