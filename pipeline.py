"""
Data pipeline: parses a RingCentral Performance Report Excel export and
produces the analysis behind the AI Receptionist business-case deck.

Methodology (locked, validated against real FBM export):
- Unit of analysis = session (grouped by Session Id), inbound only.
- Session ID deduplication removes phantom ring legs (one call routed to N
  agents counts once, not N times).
- Spam filter: any session whose maximum leg Call Length <= 5s is excluded.
- Per-session outcome priority: Answered > VM/Abandoned > VM/Missed >
  Abandoned > Missed. VM/Abandoned and VM/Missed are kept separate internally
  (never merged); the deck's "voicemail" line is a display-only grouping.
- Queue tiering (A/B/C/D) is assigned externally (Claude) and Tier D
  (back-office) queues are excluded from the headline universe.
- Business hours = Mon-Fri, 07:00-18:00 local (per Call Start Time as exported).
- Repeat callers = distinct From Number with 2+ unanswered sessions.
"""

from __future__ import annotations
import pandas as pd
from pathlib import Path
from dataclasses import dataclass, field
from typing import Optional
import datetime as dt

OUTCOME_ANSWERED = "Answered"
OUTCOME_VM_ABANDONED = "VM/Abandoned"
OUTCOME_VM_MISSED = "VM/Missed"
OUTCOME_ABANDONED = "Abandoned"
OUTCOME_MISSED = "Missed"

# Display grouping for the "how they were missed" slide
DISPLAY_ABANDONED = "Abandoned while ringing"
DISPLAY_VOICEMAIL = "Went to voicemail"
DISPLAY_RANG_OUT = "Rang out — no answer"

SPAM_MAX_SECONDS = 5
BUSINESS_DAYS = {0, 1, 2, 3, 4}  # Mon-Fri
BUSINESS_START_HOUR = 7
BUSINESS_END_HOUR = 18


@dataclass
class QueueStats:
    name: str
    inbound: int = 0
    answered: int = 0
    vm_abandoned: int = 0
    vm_missed: int = 0
    abandoned: int = 0
    missed: int = 0
    tier: str = ""             # A / B / C / D, assigned later
    classification: str = ""   # human-readable label, assigned later

    @property
    def total_missed(self) -> int:
        return self.vm_abandoned + self.vm_missed + self.abandoned + self.missed

    @property
    def miss_rate(self) -> float:
        return self.total_missed / self.inbound if self.inbound else 0.0

    @property
    def answer_rate(self) -> float:
        return self.answered / self.inbound if self.inbound else 0.0


@dataclass
class PipelineResult:
    # Raw / dedup counts
    raw_inbound_legs: int
    inbound_sessions: int          # after dedup, before spam filter
    phantom_legs_removed: int
    spam_sessions_removed: int

    # Universe used for headline figures (A+B+C queues, spam removed)
    universe_sessions: int
    answered: int
    vm_abandoned: int
    vm_missed: int
    abandoned: int
    missed: int

    # Derived analytics
    business_hours_miss_pct: float
    repeat_callers: int
    hourly_missed: dict[int, int]  # hour(7..17) -> missed count
    days_in_period: int

    queue_stats: dict[str, QueueStats] = field(default_factory=dict)
    reporting_period: str = ""
    reconciliation_ok: bool = True
    reconciliation_note: str = ""
    sessions_df: Optional[pd.DataFrame] = None

    @property
    def total_missed(self) -> int:
        return self.vm_abandoned + self.vm_missed + self.abandoned + self.missed

    @property
    def voicemail_total(self) -> int:
        return self.vm_abandoned + self.vm_missed

    @property
    def miss_rate(self) -> float:
        return self.total_missed / self.universe_sessions if self.universe_sessions else 0.0

    @property
    def answer_rate(self) -> float:
        return self.answered / self.universe_sessions if self.universe_sessions else 0.0

    @property
    def misses_per_day(self) -> float:
        return self.total_missed / self.days_in_period if self.days_in_period else 0.0


def _to_seconds(val) -> float:
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if ":" in s:
        parts = s.split(":")
        try:
            parts = [float(p) for p in parts]
        except ValueError:
            return 0.0
        if len(parts) == 3:
            return parts[0] * 3600 + parts[1] * 60 + parts[2]
        if len(parts) == 2:
            return parts[0] * 60 + parts[1]
    try:
        return float(s)
    except ValueError:
        return 0.0


def _parse_start_time(val):
    if pd.isna(val):
        return None
    for fmt in ("%m/%d/%Y %I:%M:%S %p", "%m/%d/%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S"):
        try:
            return dt.datetime.strptime(str(val).strip(), fmt)
        except ValueError:
            continue
    try:
        return pd.to_datetime(val).to_pydatetime()
    except Exception:
        return None


def _classify_session(group: pd.DataFrame) -> str:
    results = set(group["Result"].dropna().str.strip())
    handle_times = group["_handle_seconds"]

    if (handle_times > 0).any() or "Answered" in results:
        return OUTCOME_ANSWERED
    if "VM/Abandoned" in results:
        return OUTCOME_VM_ABANDONED
    if "VM/Missed" in results:
        return OUTCOME_VM_MISSED
    if "Abandoned" in results:
        return OUTCOME_ABANDONED
    return OUTCOME_MISSED


def parse_sessions(path: Path) -> pd.DataFrame:
    """Parse the export down to one row per inbound session (pre-tiering).

    Returns a DataFrame with columns:
      session_id, outcome, queue, is_spam, start_time, from_number, in_business_hours
    plus module-level counters attached as DataFrame.attrs.
    """
    xl = pd.ExcelFile(path)
    sheet_name = xl.sheet_names[0]
    for name in xl.sheet_names:
        if name.strip().lower() == "calls":
            sheet_name = name
            break
    else:
        for name in xl.sheet_names:
            if "call" in name.lower():
                sheet_name = name
                break

    df = pd.read_excel(path, sheet_name=sheet_name, dtype=str)
    df.columns = [c.strip() for c in df.columns]

    required = {"Session Id", "Call Direction", "Result", "Call Length", "Queue"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Missing expected columns: {missing}. Found: {list(df.columns)}")

    df = df[df["Call Direction"].str.strip().str.lower() == "inbound"].copy()
    raw_inbound_legs = len(df)

    df["_handle_seconds"] = df.get("Handle Time", 0)
    df["_handle_seconds"] = df["_handle_seconds"].apply(_to_seconds)
    df["_call_seconds"] = df["Call Length"].apply(_to_seconds)
    df["_start"] = df.get("Call Start Time").apply(_parse_start_time) if "Call Start Time" in df.columns else None
    df["_row_order"] = range(len(df))

    sessions = []
    for session_id, group in df.groupby("Session Id", sort=False):
        group = group.sort_values("_row_order")
        outcome = _classify_session(group)
        max_call = group["_call_seconds"].max()
        entry_queue = group.iloc[0]["Queue"]
        entry_queue = str(entry_queue).strip() if pd.notna(entry_queue) else ""
        # prefer first non-blank queue for attribution
        if not entry_queue:
            non_blank = group["Queue"].dropna()
            entry_queue = str(non_blank.iloc[0]).strip() if len(non_blank) else "Unknown"

        starts = [s for s in group["_start"].tolist() if s is not None] if "_start" in group else []
        start_time = starts[0] if starts else None
        in_bh = False
        if start_time is not None:
            in_bh = (start_time.weekday() in BUSINESS_DAYS
                     and BUSINESS_START_HOUR <= start_time.hour < BUSINESS_END_HOUR)

        from_number = ""
        if "From Number" in group.columns:
            fn = group["From Number"].dropna()
            from_number = str(fn.iloc[0]).strip() if len(fn) else ""

        sessions.append({
            "session_id": session_id,
            "outcome": outcome,
            "queue": entry_queue,
            "is_spam": max_call <= SPAM_MAX_SECONDS,
            "start_time": start_time,
            "from_number": from_number,
            "in_business_hours": in_bh,
        })

    sdf = pd.DataFrame(sessions)
    sdf.attrs["raw_inbound_legs"] = raw_inbound_legs
    return sdf


def build_result(sdf: pd.DataFrame, queue_tiers: dict[str, dict]) -> PipelineResult:
    """Apply tiering, spam filter, and compute all headline figures.

    queue_tiers: {queue_name: {"tier": "A".."D", "classification": "..."}}
    """
    raw_inbound_legs = sdf.attrs.get("raw_inbound_legs", len(sdf))
    inbound_sessions = len(sdf)
    phantom_legs_removed = raw_inbound_legs - inbound_sessions

    # Attach tier
    sdf = sdf.copy()
    sdf["tier"] = sdf["queue"].map(lambda q: queue_tiers.get(q, {}).get("tier", "C"))
    sdf["classification"] = sdf["queue"].map(lambda q: queue_tiers.get(q, {}).get("classification", ""))

    spam_mask = sdf["is_spam"]
    spam_sessions_removed = int(spam_mask.sum())
    clean = sdf[~spam_mask].copy()

    # Headline universe: A+B+C queues only (exclude Tier D back-office)
    universe = clean[clean["tier"].isin(["A", "B", "C"])].copy()

    def count(df, outcome):
        return int((df["outcome"] == outcome).sum())

    answered = count(universe, OUTCOME_ANSWERED)
    vm_abandoned = count(universe, OUTCOME_VM_ABANDONED)
    vm_missed = count(universe, OUTCOME_VM_MISSED)
    abandoned = count(universe, OUTCOME_ABANDONED)
    missed = count(universe, OUTCOME_MISSED)
    universe_sessions = len(universe)

    total_missed = vm_abandoned + vm_missed + abandoned + missed
    missed_df = universe[universe["outcome"] != OUTCOME_ANSWERED]

    # Business hours miss %
    bh_miss = int(missed_df["in_business_hours"].sum())
    business_hours_miss_pct = bh_miss / total_missed if total_missed else 0.0

    # Repeat callers: distinct From Number with 2+ unanswered sessions
    repeat_callers = 0
    if "from_number" in missed_df.columns:
        counts = missed_df[missed_df["from_number"] != ""]["from_number"].value_counts()
        repeat_callers = int((counts >= 2).sum())

    # Hourly distribution of misses (business hours window)
    hourly_missed = {h: 0 for h in range(BUSINESS_START_HOUR, BUSINESS_END_HOUR)}
    for st in missed_df["start_time"].dropna():
        if st.hour in hourly_missed:
            hourly_missed[st.hour] += 1

    # Period span
    starts = [s for s in universe["start_time"].dropna()]
    if starts:
        dmin, dmax = min(starts).date(), max(starts).date()
        days_in_period = (dmax - dmin).days + 1
        reporting_period = f"{dmin.strftime('%b %-d')}–{dmax.strftime('%-d, %Y')}"
    else:
        days_in_period = 30
        reporting_period = ""

    # Per-queue stats (A+B+C)
    queue_stats: dict[str, QueueStats] = {}
    for _, row in universe.iterrows():
        q = row["queue"]
        if q not in queue_stats:
            qs = QueueStats(name=q)
            qs.tier = row["tier"]
            qs.classification = row["classification"]
            queue_stats[q] = qs
        qs = queue_stats[q]
        qs.inbound += 1
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

    computed = answered + total_missed
    recon_ok = computed == universe_sessions
    recon_note = (f"OK: {computed} outcomes == {universe_sessions} universe sessions"
                  if recon_ok else
                  f"MISMATCH: {computed} != {universe_sessions}")

    return PipelineResult(
        raw_inbound_legs=raw_inbound_legs,
        inbound_sessions=inbound_sessions,
        phantom_legs_removed=phantom_legs_removed,
        spam_sessions_removed=spam_sessions_removed,
        universe_sessions=universe_sessions,
        answered=answered,
        vm_abandoned=vm_abandoned,
        vm_missed=vm_missed,
        abandoned=abandoned,
        missed=missed,
        business_hours_miss_pct=business_hours_miss_pct,
        repeat_callers=repeat_callers,
        hourly_missed=hourly_missed,
        days_in_period=days_in_period,
        queue_stats=queue_stats,
        reporting_period=reporting_period,
        reconciliation_ok=recon_ok,
        reconciliation_note=recon_note,
        sessions_df=sdf,
    )


def distinct_queues(sdf: pd.DataFrame) -> list[str]:
    return sorted(q for q in sdf["queue"].dropna().unique() if q and q != "Unknown")
