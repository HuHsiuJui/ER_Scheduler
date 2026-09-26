#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ER Nurse Scheduler FINAL-5.24 — AREA MATRIX + FLOW/OBSERVATION SPLIT
==============================================================

Canonical workflow
------------------
    Nurses.xlsx                           -> people / home shift / areas
    LIVE Google 預假表                   -> OFF / 特休 / mandatory D/E/N / 不D不E不N
    LIVE Google 上月正式班表             -> cross-month safety boundary
    SciPy/HiGHS MILP                     -> one canonical schedule
    validator                            -> PASS or BLOCKED
    Excel export                         -> only after hard validation passes

Important source rules
----------------------
- 預假表只讀指定 LIVE Google Sheet + 精確月份分頁；不讀 cache/snapshot/匯出 xlsx。
- 上月正式班表只讀指定 LIVE Google Sheet + 精確 ROC 月份分頁（例如 2026/08 -> 11508）。
- 不再搜尋、下載、遷移或猜測本機「上月 Excel」。
- Nurses.xlsx remains local because it is the canonical staff/qualification master.

Canonical scheduling rules
--------------------------
- D/E core: Leader, Triage, Critical, Clinic1, Clinic2, 流動, 留觀.
- N core with 5 staff: Leader+Triage, Critical, Clinic1, 流動, 留觀.
- N with 6 or more staff: Leader and Triage are split between two people; the
  sixth core position is Triage. Clinic3 is support-only after these core roles.
- E hard minimum is the seven core roles only. Clinic3 is optional support and
  never creates a staffing shortage or blocks publication.
- Working-area qualification comes ONLY from Nurses.xlsx `areas`.
- 11h minimum rest; cross-shift D/E/N transition must also have >20h rest.
- Max 5 consecutive attendance days.
- For regular/full-time staff, OFF -> WORK -> OFF singleton is forbidden as a hard rule.
- For regular/full-time staff, consecutive OFF >5 wholly inside the active month is forbidden unless covered by explicit annual leave.
- Cross-month OFF runs may be approved beyond 5 days and are excluded from the active-month max-OFF rule.
- Part-time staff are exempt from singleton/max-consecutive-OFF roster-shape rules; explicit OFF is authoritative.
- Part-time target is max(10, calendar dates remaining after explicit OFF/annual leave, explicit monthly-hours target).
- Part-time work below its computed target produces a visible WARNING but does not block review export.
- N and D/E may not coexist in one uninterrupted attendance block; a true OFF is required.
- Requested OFF is hard OFF but does NOT reduce the monthly attendance target.
- Annual leave is hard OFF and paid credit; it reduces target one 8h shift per day.
- No target-lowering or safety-solution fallback: impossible hard rules => BLOCKED / infeasible.
- Independence: service age >=2.5 months OR 18 verified formal same-shift Preceptor Observation
  sessions completed before the workday. The 18th teaching shift itself is still non-independent;
  independence starts on the next workday.
- Staff who start the month as non-independent remain on their Nurses.xlsx home shift for the
  entire month. Formal D/E teaching uses only a whitelisted, same-home-shift Preceptor who is
  qualified for both Leader and Observation.
- Formal teaching does not reduce monthly attendance target.
- Part-time: explicit Google OFF/annual leave only; no hidden weekend/holiday OFF; fixed home shift;
  regular target defaults to at least 10 shifts (or a higher explicit monthly-hours threshold);
  no H; no Clinic3 flexible count; no cross-shift support. Independent full 8h D/E/N work
  still counts as core manpower when assigned to a qualified core area.
- H is heavily de-prioritized support and never counts as D/E/N core manpower.
- Month-scoped staff exclusions remove that person from scheduling/output only for the named month;
  the Nurses.xlsx master and LIVE Google source audit remain unchanged.
- Each person may provide D/E/N cross-shift support on at most 5 days per month.

Dependencies
------------
    numpy, scipy, openpyxl, gspread, google-auth
"""
from __future__ import annotations


# ==============================================================================
# MODULE: constants.py
# ==============================================================================

from dataclasses import dataclass
from datetime import date


PROGRAM_VERSION = "ER Nurse Scheduler FINAL-5.24 AREA MATRIX + FLOW/OBSERVATION SPLIT"

SHIFT_HOURS = {"D": 8, "E": 8, "N": 8, "H": 8, "8-5": 8}
SHIFT_INTERVAL = {
    "D": (8.0, 16.0),
    "E": (16.0, 24.0),
    "N": (0.0, 8.0),
    "H": (12.0, 20.0),
    "8-5": (8.0, 17.0),
}

D_E_CORE_ROLES = (
    "Leader",
    "Triage",
    "Critical",
    "Clinic1",
    "Clinic2",
    "Observation1",
    "Observation2",
)
D_E_L_PLUS_T_CORE_ROLES = (
    "LeaderTriage",
    "Critical",
    "Clinic1",
    "Clinic2",
    "Observation1",
    "Observation2",
)
N_CORE_ROLES = (
    "LeaderTriage",
    "Critical",
    "Clinic1",
    "Observation1",
    "Observation2",
)
N_EXPANDED_CORE_ROLES = (
    "Leader",
    "Triage",
    "Critical",
    "Clinic1",
    "Observation1",
    "Observation2",
)

# Institution-specific values are intentionally empty in the public edition.
PRECEPTOR_IDS = frozenset()
PRECEPTOR_DISPLAY = {}
PART_TIME_IDS = frozenset()
PART_TIME_AREA_PREFERENCE = {}
TEMPORARY_PART_TIME_IDS_BY_MONTH = {}
MONTHLY_EXCLUDED_IDS_BY_MONTH = {}
EXPECTED_GOOGLE_SHEET_ID = ""
EXPECTED_GOOGLE_TITLE = ""
EXPECTED_PREVIOUS_ROSTER_SHEET_ID = ""
EXPECTED_PREVIOUS_ROSTER_TITLE = ""
BUILTIN_IDENTITY_ALIASES = {}
BUILTIN_TRAINING_MIGRATION = {}

# 2026 weekday holidays from the official DGPA 115-year office calendar CSV.
# Weekends are handled separately by is_holiday(); only weekday non-working dates belong here.
TAIWAN_2026_EXTRA_HOLIDAYS = frozenset({
    date(2026, 1, 1),
    date(2026, 2, 16), date(2026, 2, 17), date(2026, 2, 18), date(2026, 2, 19), date(2026, 2, 20),
    date(2026, 2, 27),
    date(2026, 4, 3), date(2026, 4, 6),
    date(2026, 5, 1),
    date(2026, 6, 19),
    date(2026, 9, 25), date(2026, 9, 28),
    date(2026, 10, 9), date(2026, 10, 26),
    date(2026, 12, 25),
})


@dataclass(frozen=True)
class ObjectiveWeights:
    cross_shift: float = 20.0
    h_shift: float = 35.0
    preferred_miss: float = 3.0
    avoid_violation: float = 12.0
    work_block_start: float = 2.0
    night_support_singleton: float = 2.0
    area_preference_miss: float = 1.0
    formal_teaching_reward: float = 4.0


# ==============================================================================
# MODULE: models.py
# ==============================================================================

from dataclasses import dataclass, field
from datetime import date
from enum import Enum
from typing import Any, Mapping


class Shift(str, Enum):
    D = "D"
    E = "E"
    N = "N"
    H = "H"
    OFF = "OFF"
    ADMIN = "8-5"
    SPECIAL = "SPECIAL"


class IndependenceStatus(str, Enum):
    INDEPENDENT = "INDEPENDENT"
    NON_INDEPENDENT = "NON_INDEPENDENT"


class RequestKind(str, Enum):
    OFF_LOCK = "OFF_LOCK"
    ANNUAL_LEAVE = "ANNUAL_LEAVE"
    MUST_SHIFT = "MUST_SHIFT"
    PREFERRED_SHIFT = "PREFERRED_SHIFT"
    AVOID_SHIFT = "AVOID_SHIFT"
    NOTE = "NOTE"


class Severity(str, Enum):
    HARD = "HARD"
    WARNING = "WARNING"
    INFO = "INFO"


class PublishStatus(str, Enum):
    PASS = "PASS"
    WARNING = "WARNING"
    BLOCKED = "BLOCKED"


@dataclass(frozen=True)
class NurseProfile:
    nurse_id: str
    name: str
    hire_date: date
    position: str
    areas: frozenset[str]
    home_shift: str
    part_time: bool = False
    # Regular monthly hours / overtime threshold. This is not a hard work cap;
    # hours above it are legal overtime and are minimized by the objective.
    max_hours: int | None = None
    max_night_shifts: int | None = None
    # Verified carry-over paid-leave/unused-rest credit from a prior period.
    # The scheduler works in 8h shifts, so only explicit multiples of 8h are accepted.
    prior_unused_hours: int = 0
    # Optional migration seed if the existing Nurses.xlsx already has a verified teaching-progress column.
    training_completed_seed: int | None = None


@dataclass(frozen=True)
class TrainingState:
    nurse_id: str
    status: IndependenceStatus
    completed_observation_sessions: int | None = None
    source: str = "verified"


@dataclass(frozen=True)
class Request:
    nurse_id: str
    day: date
    kind: RequestKind
    shift: str | None = None
    raw: str = ""
    source: str = "Google Sheet"


@dataclass(frozen=True)
class BoundaryEntry:
    nurse_id: str
    day: date
    raw: str
    shift: str
    attendance: bool
    start_hour: float | None = None
    end_hour: float | None = None


@dataclass(frozen=True)
class Assignment:
    nurse_id: str
    day: date
    shift: str
    area: str | None = None
    source: str = "SOLVER"
    locked: bool = False
    formal_teaching: bool = False
    preceptor_id: str | None = None
    independent_core: bool = True


@dataclass(frozen=True)
class Issue:
    code: str
    severity: Severity
    message: str
    nurse_id: str | None = None
    day: date | None = None
    shift: str | None = None
    details: Mapping[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class GoogleSourceAudit:
    spreadsheet_id: str
    spreadsheet_title: str
    worksheet_title: str
    year: int
    month: int
    rows: int
    cols: int
    matrix_sha256: str
    matched_people: int
    nonblank_request_cells: int
    result: str
    source_mode: str = "LIVE_GOOGLE"
    source_verified: bool = True
    checks: tuple[tuple[str, str, str], ...] = ()


@dataclass(frozen=True)
class PreviousRosterAudit:
    spreadsheet_id: str
    spreadsheet_title: str
    worksheet_title: str
    target_year: int
    target_month: int
    rows: int
    cols: int
    matrix_sha256: str
    matched_people: int
    boundary_entries: int
    result: str
    source_mode: str = "LIVE_GOOGLE_PREVIOUS_ROSTER"
    source_verified: bool = True
    checks: tuple[tuple[str, str, str], ...] = ()


@dataclass(frozen=True)
class FeasibilityReport:
    feasible: bool
    hard_issues: tuple[Issue, ...]
    warnings: tuple[Issue, ...]
    metrics: Mapping[str, Any]


@dataclass(frozen=True)
class ScheduleResult:
    year: int
    month: int
    nurses: tuple[NurseProfile, ...]
    assignments: tuple[Assignment, ...]
    target_shifts: Mapping[str, int]
    training_end: Mapping[str, TrainingState]
    google_audit: GoogleSourceAudit
    previous_audit: PreviousRosterAudit
    solver_status: str
    solver_metrics: Mapping[str, Any]
    result_sha256: str

    def assignment_map(self) -> dict[tuple[str, date], Assignment]:
        return {(a.nurse_id, a.day): a for a in self.assignments}


@dataclass(frozen=True)
class ValidationResult:
    hard_errors: tuple[Issue, ...]
    warnings: tuple[Issue, ...]
    metrics: Mapping[str, Any]
    publish_status: PublishStatus

    @property
    def publishable(self) -> bool:
        """Only PASS is directly publishable. WARNING may be exported for review, never treated as final publish."""
        return self.publish_status == PublishStatus.PASS

    @property
    def exportable(self) -> bool:
        return self.publish_status in {PublishStatus.PASS, PublishStatus.WARNING}


# ==============================================================================
# MODULE: normalization.py
# ==============================================================================

import re
import unicodedata
from datetime import date, datetime, timedelta
from typing import Any


OFF_TOKENS = {
    "OFF", "休息日", "例假", "休假班", "休假", "公假", "補休", "休D", "休E", "休N",
}


def text(value: Any) -> str:
    if value is None:
        return ""
    s=unicodedata.normalize("NFKC", str(value))
    # Remove invisible Unicode format marks (e.g. LRM/RLM) that Google Sheets may carry.
    # This is character normalization, not semantic/fuzzy matching.
    s="".join(ch for ch in s if unicodedata.category(ch)!="Cf")
    return s.strip()


def norm_id(value: Any) -> str:
    s = text(value)
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s


def canonical_area(value: str) -> str:
    s = text(value).replace(" ", "")
    mapping = {
        "Leader": "Leader", "L": "Leader",
        "Triage": "Triage", "T": "Triage",
        "Critical": "Critical", "C": "Critical",
        "Clinic1": "Clinic1", "Clinic1": "Clinic1", "診間1": "Clinic1", "診1": "Clinic1",
        "Clinic2": "Clinic2", "診間2": "Clinic2", "診2": "Clinic2",
        "Clinic3": "Clinic3", "診間3": "Clinic3", "診3": "Clinic3",
        "Observation": "Observation", "留觀": "Observation", "流動": "Observation",
    }
    return mapping.get(s, s)


def parse_areas(value: Any) -> frozenset[str]:
    raw = text(value)
    if not raw:
        return frozenset()
    parts = re.split(r"[、,;/，；]+", raw)
    return frozenset(canonical_area(p) for p in parts if text(p))


def excel_date(value: Any) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        return (datetime(1899, 12, 30) + timedelta(days=float(value))).date()
    s = text(value)
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    raise ValueError(f"無法解析日期：{value!r}")


def normalize_shift(value: Any) -> str:
    s = text(value)
    if not s:
        return ""
    u = s.upper().replace(" ", "")
    if u in {"OFF", "休息日", "例假", "休假班", "休假", "公假", "補休"} or u.startswith("休D") or u.startswith("休E") or u.startswith("休N"):
        return "OFF"
    if re.fullmatch(r"8-?1?-?5(?:\+L)?", u) or u in {"815", "815+L"}:
        return "8-5+L" if "+L" in u else "8-5"
    if u in {"D", "E", "N", "H"}:
        return u
    if re.fullmatch(r"\*?D(?:S|\d+)?", u):
        return "D"
    if re.fullmatch(r"E(?:S|\d+)?", u):
        return "E"
    if re.fullmatch(r"N(?:S|\d+)?", u):
        return "N"
    # Common legacy roster symbols that are attendance but have no DEN semantics.
    if any(k in u for k in ("外訓", "化災", "ICU", "日值", "EMT", "實證", "督考", "普渡")):
        return "SPECIAL"
    return s


def is_attendance_shift(value: Any) -> bool:
    s = normalize_shift(value)
    return s in {"D", "E", "N", "H", "8-5", "8-5+L", "SPECIAL"}


# ==============================================================================
# MODULE: calendar_tw.py
# ==============================================================================

import calendar
from datetime import date


def month_days(year: int, month: int) -> tuple[date, ...]:
    return tuple(date(year, month, d) for d in range(1, calendar.monthrange(year, month)[1] + 1))


def is_holiday(day: date, official_holidays: frozenset[date]) -> bool:
    return day.weekday() >= 5 or day in official_holidays


def official_working_days(year: int, month: int, official_holidays: frozenset[date]) -> int:
    return sum(1 for d in month_days(year, month) if not is_holiday(d, official_holidays))


# ==============================================================================
# MODULE: config.py
# ==============================================================================

from dataclasses import dataclass, field
from datetime import date
from pathlib import Path



@dataclass(frozen=True)
class SchedulerConfig:
    year: int
    month: int
    min_rest_hours: float = 11.0
    max_consecutive_attendance: int = 5
    d_min_core: int = 7
    e_min_core: int = 7
    # Clinic3 is an optional support assignment.  The E hard floor is only the
    # seven core roles; zero means Clinic3 never creates a feasibility gate.
    e_clinic3_flexible_min: int = 0
    n_min_core: int = 5
    d_max_total: int = 10
    e_max_total: int = 10
    n_max_total: int = 6
    h_person_cap: int = 6
    # Hard monthly cap on D/E/N assignments outside the Nurses.xlsx home shift.
    max_cross_shift_days_per_person: int | None = 5
    # Audited MUST_SHIFT exceptions that should not consume the discretionary
    # cross-support cap (for example, a user-requested temporary shift).  They
    # remain visible and counted in total cross-shift reporting.
    cross_shift_cap_exempt_assignments: frozenset[tuple[str, date, str]] = frozenset()
    # Monthly work above the required-hours threshold is overtime.  When set,
    # this is a hard per-person ceiling on overtime shifts for real staff.
    max_overtime_shifts_per_person: int | None = None
    allow_h: bool = True
    preceptor_ids: frozenset[str] = PRECEPTOR_IDS
    part_time_ids: frozenset[str] = PART_TIME_IDS
    part_time_min_shifts: int = 10
    official_holidays: frozenset[date] = TAIWAN_2026_EXTRA_HOLIDAYS
    objective: ObjectiveWeights = field(default_factory=ObjectiveWeights)
    require_previous_boundary_days: int = 14
    block_on_unknown_google_value: bool = True
    allow_warning_export: bool = True
    solver_time_limit_seconds: float = 120.0
    # False: required monthly hours are the regular-hours threshold. Work above
    # that threshold is permitted and reported as overtime, not made infeasible.
    require_exact_targets: bool = False
    # Intermediate solve: keep part-time staff at their individually proven
    # availability/capacity target before allowing warning-only shortfalls.
    protect_part_time_targets: bool = False
    debug_enforce_max_consecutive: bool = True
    debug_enforce_max_consecutive_off: bool = True
    debug_enforce_singleton: bool = True
    debug_enforce_rest: bool = True
    debug_enforce_cross_shift_20h: bool = True
    debug_enforce_n_recovery: bool = True
    debug_enforce_core_roles: bool = True
    # Decide OFF only after the in-house high-priority core has been protected.
    # This lexicographically prefers distinct internal staff for
    # Leader/L+T, Triage and Critical before lower-priority aesthetics or
    # cross-shift minimization may send qualified people OFF.
    prioritize_internal_high_core: bool = True
    debug_zero_objective: bool = False
    # Optional audited relief-seat optimization. Defaults preserve normal production behavior.
    relief_ids: frozenset[str] = frozenset()
    minimize_relief_only: bool = False
    fast_feasible_objective: bool = False
    relief_shift_cap: int | None = None
    real_target_equality: bool = False
    # Require every regular/full-time employee to reach at least the canonical
    # monthly hours threshold while still allowing legal overtime above it.
    # This is stricter than objective-only under-target minimization and is used
    # for publishable fairness proofs when overtime is permitted.
    real_target_minimum: bool = False
    # Optional person-specific exact monthly workday targets.  This is used for
    # audited scenario requests without forcing every regular employee to exact
    # equality.
    exact_target_ids: frozenset[str] = frozenset()
    # Ordinary requested-OFF dates that the user explicitly authorized the
    # optimizer to move.  They are removed from hard requests and audited here.
    movable_off_assignments: frozenset[tuple[str, date]] = frozenset()
    minimize_moved_off_only: bool = False
    moved_off_assignment_cap: int | None = None
    # Proof/lexicographic phases for the user-authorized zero-relief scenario.
    # The first phase minimizes only the sum of real-staff overtime.  Later
    # phases may fix that proven minimum while improving OFF movement,
    # under-target work, cross-support and roster shape.
    minimize_overtime_only: bool = False
    total_overtime_shift_cap: int | None = None
    # Internal conditional Plan-C staff are not part of the normal monthly
    # roster.  They may be scheduled only in their declared eligible role(s),
    # are excluded from ordinary overtime accounting, and are minimized in a
    # dedicated proof phase before any lower-priority optimization.
    conditional_plan_c_ids: frozenset[str] = frozenset()
    conditional_plan_c_shift_cap: int | None = None
    minimize_conditional_plan_c_only: bool = False
    # Optional externally planned shadow-training programme.  Trainees are
    # forced to the specified shift but never counted as independent core on
    # that day; the corresponding distinct Teaching roles must also be present.
    required_training_roles_by_day_shift: Mapping[tuple[date, str], tuple[str, ...]] = field(default_factory=dict)
    required_training_trainees_by_day_shift: Mapping[tuple[date, str], frozenset[str]] = field(default_factory=dict)
    required_training_assignments_by_day_shift: Mapping[tuple[date, str], Mapping[str, str]] = field(default_factory=dict)
    preferred_preceptors: Mapping[str, str] = field(default_factory=dict)
    output_dir: Path = Path(".")


# ==============================================================================
# MODULE: constraints.py
# ==============================================================================

from datetime import date, datetime, timedelta
from typing import Iterable, Mapping



def role_eligible(nurse: NurseProfile, role: str) -> bool:
    areas = nurse.areas
    if role == "Leader":
        return "Leader" in areas
    if role == "Triage":
        return "Triage" in areas
    if role == "LeaderTriage":
        return "Leader" in areas and "Triage" in areas
    if role == "Critical":
        return "Critical" in areas
    if role == "Clinic1":
        return "Clinic1" in areas
    if role == "Clinic2":
        return "Clinic2" in areas
    if role in {"Observation1", "Observation2", "Observation"}:
        return "Observation" in areas
    if role == "Clinic3":
        return "Clinic3" in areas
    return False




def service_months_on_day(nurse: NurseProfile, day: date) -> float:
    """Project-standard service age in average calendar months (365.25/12 days)."""
    if day <= nurse.hire_date:
        return 0.0
    return max(0.0, (day - nurse.hire_date).days / 30.4375)


def age_independent_on_day(nurse: NurseProfile, day: date, *, threshold_months: float = 2.5) -> bool:
    """Independence age gate only. This NEVER grants or removes working-area qualifications."""
    return service_months_on_day(nurse, day) >= threshold_months - 1e-9


def role_eligible_on_day(nurse: NurseProfile, role: str, day: date) -> bool:
    """Working-area qualification comes ONLY from NurseProfile.areas read from Nurses.xlsx.

    `day` is accepted for a stable call signature, but hire date / service age is intentionally
    irrelevant to Leader, Triage, Critical, Clinic or Observation eligibility.
    """
    del day
    return role_eligible(nurse, role)

def part_time_workday_allowed(nurse: NurseProfile, day: date, *, official_holidays: frozenset[date] = frozenset()) -> tuple[bool, str]:
    """Date-level availability for part-time staff.

    No part-time employee receives an implicit weekend / official-holiday OFF.
    Date-level nonwork is controlled ONLY by explicit hard OFF / annual leave input.
    Shift restrictions (fixed home shift, no H) remain separate in shift_allowed_static().
    `day` and `official_holidays` are accepted only for a stable call signature.
    """
    del day, official_holidays
    return True, "explicit OFF/leave only" if nurse.part_time else "not part-time"


def shift_allowed_static(nurse: NurseProfile, day: date, shift: str, *, allow_h: bool=True) -> tuple[bool,str]:
    """Single static shift-eligibility rule. Dynamic OFF/rest/consecutive rules are modeled separately."""
    s=normalize_shift(shift)
    if s not in {"D","E","N","H"}:
        return False, "unsupported shift"
    if nurse.part_time and (s=="H" or s!=nurse.home_shift):
        return False, "part-time fixed home shift / no H"
    if nurse.home_shift=="N" and s in {"D","E","H"}:
        return False, "home-N cannot D/E/H"
    if s=="N" and (day-nurse.hire_date).days < 92:
        return False, "<3 months cannot N"
    if s=="H" and (not allow_h or nurse.part_time or "Leader" in nurse.areas):
        return False, "H not eligible"
    return True, "PASS"

def shift_interval(day: date, shift: str) -> tuple[datetime, datetime] | None:
    s = normalize_shift(shift)
    if s == "8-5+L":
        s = "8-5"
    hours = SHIFT_INTERVAL.get(s)
    if hours is None:
        return None
    start_h, end_h = hours
    start = datetime.combine(day, datetime.min.time()) + timedelta(hours=start_h)
    end = datetime.combine(day, datetime.min.time()) + timedelta(hours=end_h)
    return start, end


def rest_hours(prev_day: date, prev_shift: str, next_day: date, next_shift: str) -> float | None:
    a = shift_interval(prev_day, prev_shift)
    b = shift_interval(next_day, next_shift)
    if a is None or b is None:
        return None
    return (b[0] - a[1]).total_seconds() / 3600.0


def opposite_n_family(a: str, b: str) -> bool:
    a = normalize_shift(a)
    b = normalize_shift(b)
    return (a == "N" and b in {"D", "E"}) or (b == "N" and a in {"D", "E"})


def schedule_lookup(assignments: Iterable[Assignment]) -> dict[tuple[str, date], str]:
    return {(a.nurse_id, a.day): a.shift for a in assignments}


def attendance_lookup(assignments: Iterable[Assignment]) -> dict[tuple[str, date], bool]:
    return {(a.nurse_id, a.day): is_attendance_shift(a.shift) for a in assignments}


def independent_at_start(
    nurse: NurseProfile,
    day: date,
    training_start: Mapping[str, TrainingState],
    formal_teaching_days_before: Mapping[str, set[date]],
) -> bool:
    """Independent when service age >=2.5 months OR 18 formal Preceptor sessions were completed before this workday."""
    if age_independent_on_day(nurse, day):
        return True
    state = training_start[nurse.nurse_id]
    if state.status == IndependenceStatus.INDEPENDENT:
        return True
    prior = state.completed_observation_sessions
    if prior is None:
        return False
    count = prior + sum(1 for d in formal_teaching_days_before.get(nurse.nurse_id, set()) if d < day)
    return count >= 18


def normalize_boundary_map(entries: Iterable[BoundaryEntry]) -> dict[tuple[str, date], BoundaryEntry]:
    return {(e.nurse_id, e.day): e for e in entries}


def max_attendance_capacity_under_fixed_rules(
    nurse: NurseProfile,
    days: tuple[date, ...],
    *,
    hard_off_days: set[date],
    boundary: Mapping[tuple[str, date], BoundaryEntry],
    max_consecutive: int,
    official_holidays: frozenset[date] = frozenset(),
    enforce_singleton: bool = True,
) -> tuple[int, tuple[date, ...]]:
    """Exact per-person maximum attendance under immutable day-level nonwork + max-consecutive/singleton.

    This is a tiny dynamic program used to *prove* when a nominal monthly target cannot be
    reached without violating stronger hard rules. It does not consider staffing/core needs,
    so it is an upper bound on what the full solver can legally assign.

    OFF-WORK-OFF is a strict hard prohibition when the WORK day being decided lies in
    the active scheduling month.  A singleton whose center is already in the previous
    month is historical/immutable and must not erase every DP state at the month boundary.
    This matches the MILP and final validator, which only constrain active-month center days.
    """
    if not days:
        return 0, ()
    start = days[0]
    # Boundary status is authoritative/fixed. Missing entries are treated as unknown non-attendance
    # only for DP initialization; the normal boundary loader separately requires complete lookback.
    def bstate(d: date) -> tuple[int, int]:
        e = boundary.get((nurse.nurse_id, d))
        if e is None:
            return 0, 1
        return (1 if e.attendance else 0, 0 if e.attendance else 1)

    prev2_w, prev2_fixed_off = bstate(start - timedelta(days=2))
    prev1_w, prev1_fixed_off = bstate(start - timedelta(days=1))
    run = 0
    cur = start - timedelta(days=1)
    while run < max_consecutive:
        e = boundary.get((nurse.nurse_id, cur))
        if e is None or not e.attendance:
            break
        run += 1
        cur -= timedelta(days=1)

    # state -> (count, chosen_days)
    dp: dict[tuple[int,int,int,int,int], tuple[int, tuple[date, ...]]] = {
        (run, prev2_w, prev2_fixed_off, prev1_w, prev1_fixed_off): (0, ())
    }
    for d in days:
        externally_fixed_off = (
            d in hard_off_days
            or not part_time_workday_allowed(nurse, d, official_holidays=official_holidays)[0]
        )
        choices = (0,) if externally_fixed_off else (0, 1)
        nxt: dict[tuple[int,int,int,int,int], tuple[int, tuple[date, ...]]] = {}
        for (run0,p2w,p2f,p1w,p1f),(count,chosen) in dp.items():
            for cw in choices:
                cf = 1 if (cw == 0 and externally_fixed_off) else 0
                if cw and run0 >= max_consecutive:
                    continue
                # Strict OFF -> WORK -> OFF rule applies only when the center WORK day is
                # inside the active month.  On d == first day, p1 is the previous month's
                # last day and is already immutable; rejecting that historical pattern here
                # would make the feasibility DP stricter than the MILP/final validator and
                # can incorrectly collapse capacity to zero.
                center_day = d - timedelta(days=1)
                if enforce_singleton and center_day >= start and p1w and not p2w and not cw:
                    continue
                nr = run0 + 1 if cw else 0
                state = (nr, p1w, p1f, cw, cf)
                val = (count + cw, chosen + ((d,) if cw else ()))
                if state not in nxt or val[0] > nxt[state][0]:
                    nxt[state] = val
        dp = nxt
        if not dp:
            return 0, ()
    best = max(dp.values(), key=lambda x: x[0])
    return best


# ==============================================================================
# MODULE: areas.py
# ==============================================================================

from typing import Iterable
from datetime import date


# User-authoritative core-area allocation order.  The matching algorithm fills
# these safety-critical areas with qualified in-house staff before considering
# lower-priority roles.  Anonymous RELIEF seats are always the last candidate.
CORE_AREA_PRIORITY = {
    "Leader": 0,
    "LeaderTriage": 0,
    "Triage": 1,
    "Critical": 2,
    "Clinic1": 3,
    "Clinic2": 3,
    "Observation1": 3,
    "Observation2": 3,
}


def _core_role_priority(role: str) -> int:
    return CORE_AREA_PRIORITY.get(role, 3)


def _is_relief_profile(nurse: NurseProfile) -> bool:
    return nurse.nurse_id.startswith("RELIEF-")



def _match_roles(workers: list[NurseProfile], roles: tuple[str, ...], day: date, forced: tuple[str, str] | None = None) -> dict[str, str] | None:
    by_id={n.nurse_id:n for n in workers}
    used: set[str]=set()
    result: dict[str,str]={}
    roles_left=list(roles)
    if forced:
        nid,role=forced
        n=by_id.get(nid)
        if n is None or role not in roles_left or not role_eligible_on_day(n,role,day):
            return None
        used.add(nid); result[nid]=role; roles_left.remove(role)

    # Explicit safety order first, then scarcity within the same tier.  DFS
    # backtracking still guarantees a complete feasible matching when one exists.
    roles_left.sort(key=lambda r: (
        _core_role_priority(r),
        sum(role_eligible_on_day(n,r,day) for n in workers if n.nurse_id not in used),
        r,
    ))
    assigned_role: dict[str,str]={}

    def candidates(role: str) -> list[NurseProfile]:
        out=[n for n in workers if n.nurse_id not in used and role_eligible_on_day(n,role,day)]
        def key(n: NurseProfile):
            pref=PART_TIME_AREA_PREFERENCE.get(n.nurse_id)
            display={"Observation1":"Observation","Observation2":"Observation","LeaderTriage":"Leader+Triage"}.get(role,role)
            role_priority=_core_role_priority(role)
            higher_role_qualifications=sum(
                1 for higher_role in ("Leader","LeaderTriage","Triage","Critical")
                if _core_role_priority(higher_role) < role_priority
                and role_eligible_on_day(n,higher_role,day)
            )
            return (
                1 if _is_relief_profile(n) else 0,
                0 if pref==display else 1,
                higher_role_qualifications,
                len(n.areas),
                int(n.nurse_id) if n.nurse_id.isdigit() else 10**9,
                n.name,
            )
        return sorted(out,key=key)

    def dfs(i: int) -> bool:
        if i>=len(roles_left): return True
        role=roles_left[i]
        for n in candidates(role):
            used.add(n.nurse_id); assigned_role[n.nurse_id]=role
            if dfs(i+1): return True
            used.remove(n.nurse_id); assigned_role.pop(n.nurse_id,None)
        return False
    if not dfs(0): return None
    result.update(assigned_role)
    return result


def assign_core_areas(
    workers: Iterable[NurseProfile], shift: str, *, day: date, formal_teaching: bool=False,
    preceptor_ids: frozenset[str]=frozenset(), forced_preceptor_id: str | None=None,
) -> tuple[dict[str,str], str | None]:
    workers=list(workers)
    # N uses one combined Leader+Triage position at the five-person floor.  As
    # soon as a sixth independent worker is present, Leader and Triage must be
    # assigned to separate people.  Clinic3 is never part of these core roles.
    if shift in {"D", "E"}:
        roles=D_E_CORE_ROLES
    elif len(workers) >= 6:
        roles=N_EXPANDED_CORE_ROLES
    else:
        roles=N_CORE_ROLES
    if formal_teaching and shift in {"D","E"}:
        # The MILP may select the exact teaching Preceptor.  If supplied, do not silently switch to another
        # person during deterministic area materialization.
        candidate_pids = [forced_preceptor_id] if forced_preceptor_id else sorted(preceptor_ids)
        for pid in candidate_pids:
            if not pid: continue
            if not any(n.nurse_id==pid for n in workers): continue
            for obs in ("Observation1","Observation2"):
                matched=_match_roles(workers,roles,day,(pid,obs))
                if matched is not None:
                    return ({nid:{"Observation1":"流動","Observation2":"留觀","LeaderTriage":"Leader+Triage"}.get(r,r) for nid,r in matched.items()},pid)
        raise RuntimeError(f"{shift}正式帶教找不到可被固定在Observation且其餘核心仍可匹配的Preceptor")
    matched=_match_roles(workers,roles,day)
    if matched is None:
        raise RuntimeError(f"{shift}核心角色無法完成一對一匹配")
    return ({nid:{"Observation1":"流動","Observation2":"留觀","LeaderTriage":"Leader+Triage"}.get(r,r) for nid,r in matched.items()},None)


# ==============================================================================
# MODULE: metrics.py
# ==============================================================================

from collections import Counter
from datetime import date
from typing import Iterable, Mapping



def work_assignments(assignments: Iterable[Assignment], nurse_id: str) -> tuple[Assignment, ...]:
    return tuple(a for a in assignments if a.nurse_id == nurse_id and a.shift in {"D", "E", "N", "H", "8-5", "8-5+L", "SPECIAL"})


def actual_work_days(assignments: Iterable[Assignment], nurse_id: str) -> int:
    return len(work_assignments(assignments, nurse_id))


def actual_hours(assignments: Iterable[Assignment], nurse_id: str) -> int:
    return sum(SHIFT_HOURS.get(a.shift, 8) for a in work_assignments(assignments, nurse_id))


def target_hours(target_shifts: Mapping[str, int], nurse_id: str) -> int:
    return int(target_shifts.get(nurse_id, 0)) * 8


def overtime_hours(assignments: Iterable[Assignment], target_shifts: Mapping[str, int], nurse_id: str) -> int:
    return max(0, actual_hours(assignments, nurse_id) - target_hours(target_shifts, nurse_id))


def under_target_hours(assignments: Iterable[Assignment], target_shifts: Mapping[str, int], nurse_id: str) -> int:
    return max(0, target_hours(target_shifts, nurse_id) - actual_hours(assignments, nurse_id))


def shift_counts(assignments: Iterable[Assignment], nurse_id: str) -> Counter[str]:
    return Counter(a.shift for a in assignments if a.nurse_id == nurse_id)


def cross_shift_days(assignments: Iterable[Assignment], nurses: Mapping[str, NurseProfile], nurse_id: str) -> int:
    home_shift=nurses[nurse_id].home_shift
    return sum(
        1 for a in assignments
        if a.nurse_id==nurse_id and a.shift in {"D","E","N"} and a.shift!=home_shift
    )


# ==============================================================================
# MODULE: input_identity.py
# ==============================================================================

import json
from pathlib import Path
from typing import Mapping


def load_identity_aliases(path: Path | None = None) -> dict[tuple[str, str], str]:
    """Return exact user-verified aliases.  Built-ins require no external JSON."""
    out: dict[tuple[str, str], str] = dict(BUILTIN_IDENTITY_ALIASES)
    if path is None:
        return out
    payload=json.loads(path.read_text(encoding="utf-8"))
    if payload.get("schema") != "ER_IDENTITY_ALIASES_V1":
        raise ValueError("identity aliases schema 必須是 ER_IDENTITY_ALIASES_V1")
    for row in payload.get("aliases",[]):
        nid=str(row.get("nurse_id","")).strip()
        source=str(row.get("source_name","")).strip()
        canonical=str(row.get("canonical_name","")).strip()
        verified=str(row.get("verified_by","")).strip()
        if not nid or not source or not canonical or not verified:
            raise ValueError(f"identity alias 不完整：{row}")
        key=(nid,source)
        if key in out and out[key]!=canonical:
            raise ValueError(f"identity alias 衝突：{key}")
        out[key]=canonical
    return out


# ==============================================================================
# MODULE: input_nurses.py
# ==============================================================================

from pathlib import Path
from typing import Iterable



def _month_column_candidates(month: int) -> tuple[str, ...]:
    english = ("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    return (english[month - 1], f"{month}月", str(month))


def load_nurses_xlsx(path: Path, month: int, *, part_time_ids: frozenset[str]) -> tuple[NurseProfile, ...]:
    """Read the existing Nurses.xlsx.  No new workbook or sidecar state file is required."""
    from openpyxl import load_workbook

    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if "Nurses" not in wb.sheetnames:
            raise ValueError("Nurses.xlsx 缺少工作表 'Nurses'")
        ws = wb["Nurses"]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            raise ValueError("Nurses 工作表為空")
        headers = [text(v) for v in rows[0]]
        required = {"id", "name", "hire_date", "position", "areas"}
        missing = sorted(required - set(headers))
        if missing:
            raise ValueError(f"Nurses 欄位缺失：{missing}")
        month_col = next((c for c in _month_column_candidates(month) if c in headers), None)
        if month_col is None:
            normalized = {text(h).strip(): h for h in headers}
            month_col = next((normalized[c] for c in _month_column_candidates(month) if c in normalized), None)
        if month_col is None:
            raise ValueError(f"Nurses 找不到月份欄：{_month_column_candidates(month)}")
        idx = {h: i for i, h in enumerate(headers)}
        m_idx = headers.index(month_col)

        unused_aliases = (
            "未休時數", "上期未休時數", "上月未休時數", "前期未休時數",
            "prior_unused_hours", "previous_unused_hours", "carryover_hours", "comp_hours",
        )
        unused_col = next((a for a in unused_aliases if a in headers), None)
        unused_idx = headers.index(unused_col) if unused_col is not None else None

        training_aliases = (
            "已完成帶教班數", "累積帶教完成班數", "Preceptor帶教完成班數",
            "preceptor_completed_shifts", "training_precepted_shifts",
        )
        training_col = next((a for a in training_aliases if a in headers), None)
        training_idx = headers.index(training_col) if training_col is not None else None

        max_hours_aliases = (
            "max_hours", "required_monthly_hours", "regular_hours_threshold",
            "每月應上班時數", "月應上班時數", "月工時上限", "每月最高工時", "最高工時",
        )
        max_night_aliases = ("max_night_shifts", "maximum_night_shifts", "每月最高夜班", "夜班上限", "最高夜班數")
        max_hours_col = next((a for a in max_hours_aliases if a in headers), None)
        max_night_col = next((a for a in max_night_aliases if a in headers), None)
        max_hours_idx = headers.index(max_hours_col) if max_hours_col is not None else None
        max_night_idx = headers.index(max_night_col) if max_night_col is not None else None

        out: list[NurseProfile] = []
        seen: set[str] = set()
        for row in rows[1:]:
            nid = norm_id(row[idx["id"]] if idx["id"] < len(row) else "")
            name = text(row[idx["name"]] if idx["name"] < len(row) else "")
            home = text(row[m_idx] if m_idx < len(row) else "").upper()
            if not nid and not name:
                continue
            if home not in {"D", "E", "N"}:
                continue
            if not nid:
                raise ValueError(f"人員 {name!r} 缺少 ID；正式排班禁止只靠姓名")
            if nid in seen:
                raise ValueError(f"Nurses ID 重複：{nid}")
            seen.add(nid)

            prior_unused_hours = 0
            if unused_idx is not None and unused_idx < len(row) and row[unused_idx] not in (None, ""):
                try:
                    prior_unused_hours = int(float(row[unused_idx]))
                except Exception as exc:
                    raise ValueError(f"{nid} {name} 未休時數不是數字：{row[unused_idx]!r}") from exc
                if prior_unused_hours < 0 or prior_unused_hours % 8 != 0:
                    raise ValueError(f"{nid} {name} 未休時數必須是非負8小時整數倍，目前={prior_unused_hours}h")

            training_seed = None
            if training_idx is not None and training_idx < len(row) and row[training_idx] not in (None, ""):
                try:
                    training_seed = int(float(row[training_idx]))
                except Exception as exc:
                    raise ValueError(f"{nid} {name} 帶教完成班數不是數字：{row[training_idx]!r}") from exc
                if training_seed < 0:
                    raise ValueError(f"{nid} {name} 帶教完成班數不得小於0")

            def optional_nonnegative_int(col_idx: int | None, label: str) -> int | None:
                if col_idx is None or col_idx >= len(row) or row[col_idx] in (None, ""):
                    return None
                try:
                    value = int(float(row[col_idx]))
                except Exception as exc:
                    raise ValueError(f"{nid} {name} {label}不是整數：{row[col_idx]!r}") from exc
                if value < 0:
                    raise ValueError(f"{nid} {name} {label}不得小於0")
                return value

            max_hours = optional_nonnegative_int(max_hours_idx, "每月應上班時數／加班起算線")
            max_night_shifts = optional_nonnegative_int(max_night_idx, "夜班上限")
            if max_hours is not None and max_hours % 8 != 0:
                raise ValueError(f"{nid} {name} 每月應上班時數必須是8小時整數倍，目前={max_hours}h")

            out.append(NurseProfile(
                nurse_id=nid,
                name=name,
                hire_date=excel_date(row[idx["hire_date"]]),
                position=text(row[idx["position"]]),
                areas=parse_areas(row[idx["areas"]]),
                home_shift=home,
                part_time=nid in part_time_ids,
                max_hours=max_hours,
                max_night_shifts=max_night_shifts,
                prior_unused_hours=prior_unused_hours,
                training_completed_seed=training_seed,
            ))
        if not out:
            raise ValueError("Nurses 沒有可排人員")
        return tuple(out)
    finally:
        wb.close()


# ==============================================================================
# MODULE: input_google.py
# ==============================================================================

import hashlib
import json
import time
import re
from datetime import date
from pathlib import Path
from typing import Any, Iterable, Mapping


def normalize_matrix(matrix: Iterable[Iterable[Any]]) -> list[list[str]]:
    return [[text(v) for v in row] for row in matrix]


def matrix_sha256(matrix: Iterable[Iterable[Any]]) -> str:
    normalized = normalize_matrix(matrix)
    payload = json.dumps(normalized, ensure_ascii=False, separators=(",", ":"))
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def _google_readonly_client(credentials_path: Path):
    """Authenticate with the existing service account; all Google sources stay LIVE."""
    try:
        import gspread
    except ImportError as exc:
        raise RuntimeError("缺少 gspread。請執行：pip install -U gspread google-auth") from exc

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://www.googleapis.com/auth/drive.readonly",
    ]
    try:
        return gspread.service_account(filename=str(credentials_path), scopes=scopes)
    except Exception as exc:
        raise RuntimeError(f"Google service account 認證失敗：{credentials_path}") from exc


def _open_live_google_spreadsheet(spreadsheet_id: str, credentials_path: Path, *, source_label: str):
    """Open one canonical LIVE Google Sheet strictly by spreadsheet ID."""
    client = _google_readonly_client(credentials_path)
    service_email = _service_account_email(credentials_path)
    try:
        spreadsheet = client.open_by_key(spreadsheet_id)
    except Exception as exc:
        # gspread maps API 403 to PermissionError in some versions; keep the message deterministic.
        status = getattr(getattr(exc, "response", None), "status_code", None)
        cause = getattr(exc, "__cause__", None)
        cause_status = getattr(getattr(cause, "response", None), "status_code", None)
        if isinstance(exc, PermissionError) or status == 403 or cause_status == 403:
            raise PermissionError(
                f"{source_label} 尚未授權給 Google service account。\n"
                f"請在該 Google Sheet 按『共用』→ 加入檢視者：{service_email}\n"
                "完成分享後直接重新執行本程式；不需要下載班表、不需要 OAuth、不需要改程式。"
            ) from exc
        raise RuntimeError(
            f"LIVE Google Sheet 開啟失敗：{source_label} (spreadsheet_id={spreadsheet_id})"
        ) from exc
    return spreadsheet, spreadsheet_id


def _verify_google_sources_access(credentials_path: Path) -> None:
    """Fail fast before parsing, with a precise permission instruction for either canonical source."""
    sources = (
        (EXPECTED_GOOGLE_SHEET_ID, "LIVE Google 預假表"),
        (EXPECTED_PREVIOUS_ROSTER_SHEET_ID, "LIVE Google 上月正式班表"),
    )
    for sid, label in sources:
        _open_live_google_spreadsheet(sid, credentials_path, source_label=label)


def _parse_request_cell(raw: str) -> tuple[RequestKind, str | None] | None:
    s = text(raw)
    if not s:
        return None
    u = s.upper().replace(" ", "")
    if u == "OFF":
        return RequestKind.OFF_LOCK, None
    if "特休" in s:
        return RequestKind.ANNUAL_LEAVE, None
    m = re.fullmatch(r"(?:必|MUST[:_]?)(D|E|N)", u)
    if m:
        return RequestKind.MUST_SHIFT, m.group(1)
    # A date cell containing an explicit D/E/N is an authoritative requested
    # assignment, not an aesthetic preference that the optimizer may ignore.
    if u in {"D", "E", "N"}:
        return RequestKind.MUST_SHIFT, u
    m = re.fullmatch(r"不(D|E|N)", u)
    if m:
        return RequestKind.AVOID_SHIFT, m.group(1)
    m = re.match(r"^(D|E|N)[(（]", u)
    if m:
        return RequestKind.MUST_SHIFT, m.group(1)
    if u in {"化災", "感控"}:
        return RequestKind.NOTE, None
    return None


def parse_google_matrix_strict(
    matrix: Iterable[Iterable[Any]],
    nurses: tuple[NurseProfile, ...],
    *,
    year: int,
    month: int,
    spreadsheet_id: str,
    spreadsheet_title: str,
    worksheet_title: str,
    expected_spreadsheet_id: str,
    expected_title: str,
    block_on_unknown: bool = True,
    source_mode: str = "LIVE_GOOGLE",
    source_verified: bool = True,
    name_aliases: Mapping[tuple[str, str], str] | None = None,
) -> tuple[tuple[Request, ...], GoogleSourceAudit]:
    raw_rows = [list(row) for row in matrix]
    rows = normalize_matrix(raw_rows)
    checks: list[tuple[str, str, str]] = []

    def check(name: str, ok: bool, detail: str) -> None:
        checks.append((name, "PASS" if ok else "FAIL", detail))
        if not ok:
            raise ValueError(f"Google來源檢查失敗[{name}]：{detail}")

    check("spreadsheet_id", spreadsheet_id == expected_spreadsheet_id, f"actual={spreadsheet_id}")
    check("spreadsheet_title", spreadsheet_title == expected_title, f"actual={spreadsheet_title}")
    expected_tab = f"{month}月"
    check("worksheet_title", worksheet_title == expected_tab, f"actual={worksheet_title}, expected={expected_tab}")
    check("matrix_shape", len(rows) >= 4 and max((len(r) for r in rows), default=0) >= 33, f"rows={len(rows)}")

    header_row = rows[2] if len(rows) > 2 else []
    check("identity_headers", header_row[:3] == ["班別", "姓名", "員編"], f"actual={header_row[:3]}")
    date_row = raw_rows[1]
    day_columns: dict[int, int] = {}
    for col in range(3, len(date_row)):
        raw_date = date_row[col]
        parsed_date = None
        if isinstance(raw_date, (int, float)):
            try:
                parsed_date = excel_date(raw_date)
            except Exception:
                parsed_date = None
        elif hasattr(raw_date, "year") and hasattr(raw_date, "month") and hasattr(raw_date, "day"):
            try:
                parsed_date = excel_date(raw_date)
            except Exception:
                parsed_date = None
        else:
            raw_serial = text(raw_date)
            if re.fullmatch(r"\d{5}(?:\.0+)?", raw_serial):
                try:
                    parsed_date = excel_date(float(raw_serial))
                except Exception:
                    parsed_date = None
        if parsed_date is not None:
            if parsed_date.year == year and parsed_date.month == month:
                day_columns[parsed_date.day] = col
            continue
        raw_text = text(raw_date)
        m = re.fullmatch(r"(?:\d{4}[-/])?(\d{1,2})/(\d{1,2})", raw_text)
        if m:
            mm, dd = int(m.group(1)), int(m.group(2))
            if mm == month:
                day_columns[dd] = col
        else:
            m2 = re.fullmatch(r"(\d{1,2})", raw_text)
            if m2:
                dd = int(m2.group(1))
                if 1 <= dd <= 31:
                    day_columns[dd] = col

    import calendar
    expected_days = set(range(1, calendar.monthrange(year, month)[1] + 1))
    check("date_columns", set(day_columns) == expected_days,
          f"actual={sorted(day_columns)}, expected={sorted(expected_days)}")

    nurse_by_id = {n.nurse_id: n for n in nurses}
    name_aliases = dict(name_aliases or {})
    requests: list[Request] = []
    seen_rows: set[str] = set()
    matched_ids: set[str] = set()
    unknown: list[str] = []
    alias_hits: list[str] = []
    nonblank = 0

    for row_no, row in enumerate(rows[3:], start=4):
        if len(row) < 3:
            continue
        name = text(row[1])
        nid = norm_id(row[2])
        if not name and not nid:
            continue
        if not nid:
            raise ValueError(f"Google第{row_no}列 {name!r} 缺少員編；禁止 fuzzy/姓名猜測")
        if nid in seen_rows:
            raise ValueError(f"Google員編重複：{nid}")
        seen_rows.add(nid)
        nurse = nurse_by_id.get(nid)
        if nurse is None:
            raise ValueError(f"Google員編 {nid} ({name}) 不存在於 Nurses.xlsx")
        if name != nurse.name:
            canonical = name_aliases.get((nid, name))
            if canonical != nurse.name:
                raise ValueError(
                    f"Google姓名/員編不一致：{nid} Google={name!r} Nurses={nurse.name!r}；"
                    "未提供精確verified alias"
                )
            alias_hits.append(f"{nid}:{name}->{nurse.name}")
        matched_ids.add(nid)
        for dd, col in day_columns.items():
            raw = row[col] if col < len(row) else ""
            if not text(raw):
                continue
            nonblank += 1
            parsed = _parse_request_cell(raw)
            if parsed is None:
                unknown.append(f"row={row_no},id={nid},date={year}-{month:02d}-{dd:02d},value={raw!r}")
                continue
            kind, shift = parsed
            requests.append(Request(nid, date(year, month, dd), kind, shift, text(raw), "LIVE Google 預假表"))

    if unknown and block_on_unknown:
        raise ValueError("Google存在未辨識預班內容，禁止猜測：\n" + "\n".join(unknown[:20]))
    checks.append(("identity_match", "PASS", f"matched_people={len(matched_ids)}"))
    checks.append(("explicit_aliases", "PASS", f"used={len(alias_hits)}" + (" | " + "; ".join(alias_hits) if alias_hits else "")))
    checks.append(("unknown_values", "PASS" if not unknown else "WARNING", f"count={len(unknown)}"))

    audit = GoogleSourceAudit(
        spreadsheet_id=spreadsheet_id,
        spreadsheet_title=spreadsheet_title,
        worksheet_title=worksheet_title,
        year=year,
        month=month,
        rows=len(rows),
        cols=max((len(r) for r in rows), default=0),
        matrix_sha256=matrix_sha256(rows),
        matched_people=len(matched_ids),
        nonblank_request_cells=nonblank,
        result=("PASS" if source_verified and not unknown else "WARNING" if source_verified else "DIAGNOSTIC_ONLY"),
        source_mode=source_mode,
        source_verified=source_verified,
        checks=tuple(checks + [("source_mode", "PASS" if source_verified else "WARNING", source_mode)]),
    )
    return tuple(requests), audit


def load_google_live_strict(
    spreadsheet_id: str,
    credentials_path: Path,
    nurses: tuple[NurseProfile, ...],
    *, year: int, month: int,
    expected_spreadsheet_id: str,
    expected_title: str,
    block_on_unknown: bool = True,
    name_aliases: Mapping[tuple[str, str], str] | None = None,
) -> tuple[tuple[Request, ...], GoogleSourceAudit, list[list[str]]]:
    """Read the official pre-request sheet LIVE. There is deliberately no xlsx/cache fallback."""
    spreadsheet, sid = _open_live_google_spreadsheet(spreadsheet_id, credentials_path, source_label="LIVE Google 預假表")
    if sid != expected_spreadsheet_id:
        raise ValueError(f"Google Sheet ID 不符：actual={sid}, expected={expected_spreadsheet_id}")
    if spreadsheet.title != expected_title:
        raise ValueError(f"Google Sheet 文件名稱不符：{spreadsheet.title!r}")
    tab = f"{month}月"
    try:
        ws = spreadsheet.worksheet(tab)
    except Exception as exc:
        raise ValueError(f"找不到精確分頁 {tab!r}；Strict mode 禁止 fallback") from exc
    matrix = ws.get_all_values()
    requests, audit = parse_google_matrix_strict(
        matrix, nurses, year=year, month=month,
        spreadsheet_id=sid, spreadsheet_title=spreadsheet.title,
        worksheet_title=ws.title, expected_spreadsheet_id=expected_spreadsheet_id,
        expected_title=expected_title, block_on_unknown=block_on_unknown,
        source_mode="LIVE_GOOGLE", source_verified=True, name_aliases=name_aliases,
    )
    return requests, audit, normalize_matrix(matrix)


# ==============================================================================
# MODULE: input_previous.py
# ==============================================================================

import calendar
from datetime import date, timedelta
from pathlib import Path
from typing import Any, Iterable, Mapping


def _matrix_cell(rows: list[list[str]], r: int, c: int) -> str:
    if r < 0 or r >= len(rows):
        return ""
    row = rows[r]
    if c < 0 or c >= len(row):
        return ""
    return text(row[c])


def _full_month_day_blocks_matrix(
    matrix: Iterable[Iterable[Any]], target_year: int, target_month: int
) -> list[tuple[int, int, int]]:
    """Find exact contiguous 1..month-end day-number blocks in a LIVE Google matrix.

    Returns zero-based tuples: (header_row, first_day_col, last_day_col).
    Duplicate partial month tables elsewhere on the hospital sheet cannot overwrite this block.
    """
    rows = normalize_matrix(matrix)
    last_day = calendar.monthrange(target_year, target_month)[1]
    seq = [str(i) for i in range(1, last_day + 1)]
    blocks: list[tuple[int, int, int]] = []
    for r, row in enumerate(rows[:140]):
        vals = [text(v) for v in row[:180]]
        for c in range(0, max(0, len(vals) - last_day + 1)):
            candidate=[]
            for v in vals[c:c+last_day]:
                m=re.fullmatch(r"(\d{1,2})(?:\.0+)?", v)
                candidate.append(str(int(m.group(1))) if m else "")
            if candidate == seq:
                blocks.append((r, c, c + last_day - 1))
    return blocks


def _find_identity_columns_matrix(
    rows: list[list[str]], header_row: int, first_day_col: int
) -> tuple[int | None, int | None]:
    """Find the nearest exact 姓名 + 員工編號 identity pair to the left of a calendar block."""
    for rr in range(header_row, min(len(rows), header_row + 4)):
        for c in range(first_day_col - 1, max(-1, first_day_col - 45), -1):
            v = _matrix_cell(rows, rr, c)
            if v in {"員工編號", "員編", "ID", "id"}:
                for cc in range(c - 1, max(-1, c - 10), -1):
                    if _matrix_cell(rows, rr, cc) == "姓名":
                        return cc, c
    return None, None


def parse_previous_roster_google_matrix_strict(
    matrix: Iterable[Iterable[Any]],
    nurses: tuple[NurseProfile, ...],
    *,
    target_year: int,
    target_month: int,
    lookback_days: int,
    spreadsheet_id: str,
    spreadsheet_title: str,
    worksheet_title: str,
    expected_spreadsheet_id: str,
    name_aliases: Mapping[tuple[str, str], str] | None = None,
) -> tuple[tuple[BoundaryEntry, ...], PreviousRosterAudit]:
    """Parse the previous official roster directly from LIVE Google values.

    Strictness:
    - exact previous-month ROC tab is selected before this function is called;
    - exact contiguous 1..month-end calendar block only;
    - exact employee ID, or one explicit built-in verified source-name alias when historical row has no ID;
    - every employed nurse/date in the lookback window must be present and recognized;
    - no fuzzy matching, no xlsx migration, no cache.
    """
    rows = normalize_matrix(matrix)
    checks: list[tuple[str, str, str]] = []

    def check(name: str, ok: bool, detail: str) -> None:
        checks.append((name, "PASS" if ok else "FAIL", detail))
        if not ok:
            raise ValueError(f"上月LIVE Google來源檢查失敗[{name}]：{detail}")

    check("spreadsheet_id", spreadsheet_id == expected_spreadsheet_id,
          f"actual={spreadsheet_id}, expected={expected_spreadsheet_id}")
    exact_tab = f"{target_year - 1911:03d}{target_month:02d}"
    check("worksheet_title", worksheet_title == exact_tab,
          f"actual={worksheet_title}, expected={exact_tab}")

    blocks = _full_month_day_blocks_matrix(rows, target_year, target_month)
    check("full_calendar_block", bool(blocks), f"found={len(blocks)}")

    nurse_by_id = {n.nurse_id: n for n in nurses}
    nurse_ids = set(nurse_by_id)
    aliases = dict(name_aliases or {})

    # Employee ID is the identity key. Historical Google rosters may omit the ID
    # while using one of several *explicitly verified* spellings for the same person.
    # Build an exact-name identity set per employee from BOTH sides of every verified
    # alias. This deliberately does not use fuzzy/edit-distance matching.
    verified_names_by_id: dict[str, set[str]] = {
        nid: {text(n.name)} for nid, n in nurse_by_id.items()
    }
    for (aid, source_name), canonical_name in aliases.items():
        if aid not in nurse_by_id:
            continue
        source_name = text(source_name)
        canonical_name = text(canonical_name)
        if source_name:
            verified_names_by_id[aid].add(source_name)
        if canonical_name:
            verified_names_by_id[aid].add(canonical_name)

    exact_name_to_ids: dict[str, list[str]] = {}
    for aid, names in verified_names_by_id.items():
        for verified_name in names:
            exact_name_to_ids.setdefault(verified_name, []).append(aid)

    month_end = date(target_year, target_month, calendar.monthrange(target_year, target_month)[1])
    wanted = tuple(month_end - timedelta(days=i) for i in range(lookback_days - 1, -1, -1))
    valid_candidates: list[tuple[tuple[BoundaryEntry, ...], int, int, int]] = []
    diagnostics: list[str] = []

    for header_row, first_col, last_col in blocks:
        name_col, id_col = _find_identity_columns_matrix(rows, header_row, first_col)
        if id_col is None or name_col is None:
            diagnostics.append(f"calendar@R{header_row+1}C{first_col+1}: no exact 姓名/員工編號")
            continue

        found: dict[tuple[str, date], BoundaryEntry] = {}
        seen_rows: dict[str, int] = {}
        duplicate_ids: set[str] = set()
        alias_hits: list[str] = []

        # Dates before a person's hire date are authoritative NOT_EMPLOYED non-attendance,
        # not guessed OFF. This keeps cross-month timeline complete for true new hires.
        for n in nurses:
            for d in wanted:
                if d < n.hire_date:
                    found[(n.nurse_id, d)] = BoundaryEntry(
                        n.nurse_id, d, "NOT_EMPLOYED", "OFF", False
                    )

        for r in range(header_row + 1, len(rows)):
            nid = norm_id(_matrix_cell(rows, r, id_col))
            source_name = text(_matrix_cell(rows, r, name_col))
            if nid not in nurse_ids:
                if nid:
                    continue
                # Historical row has no employee ID: resolve ONLY by an exact
                # user/project-verified name belonging to exactly one current nurse ID.
                candidates = exact_name_to_ids.get(source_name, [])
                if len(candidates) != 1:
                    continue
                nid = candidates[0]
                alias_hits.append(f"{source_name}->{nid}:{nurse_by_id[nid].name}")

            if nid in seen_rows:
                duplicate_ids.add(nid)
                continue
            seen_rows[nid] = r
            nurse = nurse_by_id[nid]
            if source_name and source_name not in verified_names_by_id.get(nid, {text(nurse.name)}):
                diagnostics.append(
                    f"{worksheet_title}: {nid} source_name={source_name!r} not in verified exact names="
                    f"{sorted(verified_names_by_id.get(nid, {text(nurse.name)}))}"
                )
                duplicate_ids.add(nid)  # invalidate this candidate deterministically
                continue

            for d in wanted:
                if d < nurse.hire_date:
                    continue
                col = first_col + d.day - 1
                raw = _matrix_cell(rows, r, col)
                if not raw:
                    diagnostics.append(f"{worksheet_title}: blank {nid} {d.isoformat()}")
                    continue
                sh = normalize_shift(raw)
                if sh not in {"D", "E", "N", "H", "OFF", "8-5", "8-5+L", "SPECIAL"}:
                    diagnostics.append(f"{worksheet_title}: unknown {nid} {d.isoformat()} {raw!r}")
                    continue
                found[(nid, d)] = BoundaryEntry(nid, d, raw, sh, is_attendance_shift(sh))

        missing = [
            (n.nurse_id, d)
            for n in nurses
            for d in wanted
            if (n.nurse_id, d) not in found
        ]
        if duplicate_ids:
            diagnostics.append(f"{worksheet_title}: duplicate/identity-invalid={','.join(sorted(duplicate_ids))}")
            continue
        if missing:
            diagnostics.append(f"{worksheet_title}: missing={len(missing)} sample={missing[:5]}")
            continue

        entries = tuple(found[k] for k in sorted(found, key=lambda x: (x[0], x[1])))
        matched_people = len({e.nurse_id for e in entries})
        valid_candidates.append((entries, matched_people, header_row, first_col))
        checks.append(("explicit_aliases", "PASS", f"used={len(alias_hits)}" + (" | " + "; ".join(alias_hits) if alias_hits else "")))

    if not valid_candidates:
        raise ValueError(
            f"LIVE上月正式班表無法建立完整{lookback_days}日邊界；"
            + " | ".join(diagnostics[:12])
        )
    if len(valid_candidates) != 1:
        raise ValueError(
            "LIVE上月正式班表存在多個同時完整的1..月末區塊，Strict mode拒絕猜測："
            + ", ".join(f"R{r+1}C{c+1}" for _,_,r,c in valid_candidates)
        )

    entries, matched_people, header_row, first_col = valid_candidates[0]
    checks.append(("boundary_complete", "PASS", f"entries={len(entries)}"))
    checks.append(("calendar_location", "PASS", f"R{header_row+1}C{first_col+1}"))
    audit = PreviousRosterAudit(
        spreadsheet_id=spreadsheet_id,
        spreadsheet_title=spreadsheet_title,
        worksheet_title=worksheet_title,
        target_year=target_year,
        target_month=target_month,
        rows=len(rows),
        cols=max((len(r) for r in rows), default=0),
        matrix_sha256=matrix_sha256(rows),
        matched_people=matched_people,
        boundary_entries=len(entries),
        result="PASS",
        source_mode="LIVE_GOOGLE_PREVIOUS_ROSTER",
        source_verified=True,
        checks=tuple(checks),
    )
    return entries, audit


def load_previous_roster_google_live_strict(
    spreadsheet_id: str,
    credentials_path: Path,
    nurses: tuple[NurseProfile, ...],
    *,
    target_year: int,
    target_month: int,
    lookback_days: int,
    expected_spreadsheet_id: str,
    name_aliases: Mapping[tuple[str, str], str] | None = None,
) -> tuple[tuple[BoundaryEntry, ...], PreviousRosterAudit, list[list[str]]]:
    """Read the previous official roster LIVE from Google. Never downloads or opens local xlsx."""
    spreadsheet, sid = _open_live_google_spreadsheet(spreadsheet_id, credentials_path, source_label="LIVE Google 預假表")
    if sid != expected_spreadsheet_id:
        raise ValueError(f"上月班表 Google Sheet ID 不符：actual={sid}, expected={expected_spreadsheet_id}")
    exact_tab = f"{target_year - 1911:03d}{target_month:02d}"
    try:
        ws = spreadsheet.worksheet(exact_tab)
    except Exception as exc:
        raise ValueError(f"上月正式班表找不到精確分頁 {exact_tab!r}；禁止 fallback 到週班/公告/其他月份") from exc
    matrix = ws.get_all_values()
    boundary, audit = parse_previous_roster_google_matrix_strict(
        matrix, nurses,
        target_year=target_year, target_month=target_month,
        lookback_days=lookback_days,
        spreadsheet_id=sid, spreadsheet_title=spreadsheet.title,
        worksheet_title=ws.title,
        expected_spreadsheet_id=expected_spreadsheet_id,
        name_aliases=name_aliases,
    )
    return boundary, audit, normalize_matrix(matrix)


# ==============================================================================
# MODULE: input_training.py
# ==============================================================================

from datetime import date


def load_training_state_canonical(
    nurses: tuple[NurseProfile, ...],
    *, year: int, month: int,
    interactive: bool = True,
) -> dict[str, TrainingState]:
    """Resolve month-start independence from canonical verified inputs only.

    Priority for staff who are still <2.5 months at month start:
      1. explicit verified training-progress column in Nurses.xlsx;
      2. one-time user-confirmed migration seed for the specified month;
      3. current-month new hire -> 0;
      4. interactive explicit integer entry (never guessed).

    The previous official Google roster is NOT used to infer teaching counts because a shift code
    alone cannot prove a qualified same-shift Preceptor Observation pairing.
    """
    start = date(year, month, 1)
    migration = BUILTIN_TRAINING_MIGRATION.get((year, month), {})
    result: dict[str, TrainingState] = {}

    for n in nurses:
        if age_independent_on_day(n, start):
            known = n.training_completed_seed
            result[n.nurse_id] = TrainingState(
                n.nurse_id,
                IndependenceStatus.INDEPENDENT,
                known,
                "age>=2.5m OR-gate" + ("; Nurses.xlsx progress preserved" if known is not None else "; teaching history not required"),
            )
            continue

        completed: int | None = None
        source = ""
        if n.training_completed_seed is not None:
            completed = int(n.training_completed_seed)
            source = "Nurses.xlsx verified training progress"
        elif n.nurse_id in migration:
            completed = int(migration[n.nurse_id])
            source = "user-confirmed one-time migration progress"
        elif n.hire_date >= start:
            completed = 0
            source = "current-month new hire; initial formal teaching=0"
        elif interactive:
            while True:
                raw = input(
                    f"{n.name}({n.nurse_id}) 月初年資<2.5月，缺少可驗證正式Preceptor累積。"
                    "請輸入已完成正式帶教班數 > "
                ).strip()
                try:
                    completed = int(raw)
                    if completed < 0:
                        raise ValueError
                    source = "interactive explicit verified migration"
                    break
                except Exception:
                    print("請輸入0以上整數，例如 10。")
        else:
            raise ValueError(
                f"{n.name}({n.nurse_id}) 年資<2.5月且沒有可驗證帶教累積；"
                "non-interactive strict mode 禁止把未知歷史當0"
            )

        status = IndependenceStatus.INDEPENDENT if int(completed or 0) >= 18 else IndependenceStatus.NON_INDEPENDENT
        result[n.nurse_id] = TrainingState(n.nurse_id, status, int(completed or 0), source)
    return result


# ==============================================================================
# MODULE: feasibility.py

# ==============================================================================

from collections import Counter
from datetime import date
from typing import Mapping



def build_target_shifts(
    nurses: tuple[NurseProfile, ...],
    requests: tuple[Request, ...],
    training: Mapping[str, TrainingState],
    *, year: int, month: int, official_holidays: frozenset[date],
    part_time_min_shifts: int = 10,
) -> dict[str, int]:
    standard = official_working_days(year, month, official_holidays)
    annual = Counter(r.nurse_id for r in requests if r.kind == RequestKind.ANNUAL_LEAVE)
    hard_off = {(r.nurse_id, r.day) for r in requests if r.kind in {RequestKind.OFF_LOCK, RequestKind.ANNUAL_LEAVE}}
    days = month_days(year, month)
    targets: dict[str, int] = {}
    for n in nurses:
        if n.part_time:
            # Every calendar date remaining after explicit Google OFF/annual
            # leave is an intended work date. Ten shifts is the floor; an
            # explicit monthly-hours threshold may set an even higher target.
            available_days=sum(
                1 for d in days
                if (n.nurse_id,d) not in hard_off
                and part_time_workday_allowed(n,d,official_holidays=official_holidays)[0]
            )
            explicit_target=n.max_hours//8 if n.max_hours is not None else 0
            targets[n.nurse_id]=max(part_time_min_shifts,available_days,explicit_target)
            continue

        # IMPORTANT: training/non-independent status NEVER reduces monthly attendance target.
        # The "18" threshold belongs only to formal Preceptor Observation progress / independence.
        # An explicit monthly-hours value replaces the calendar-derived regular-hours
        # threshold. It never prohibits necessary work above the threshold.
        base = n.max_hours // 8 if n.max_hours is not None else standard

        # Annual leave is paid leave and reduces required attendance one 8h shift per day.
        annual_credit_shifts = int(annual[n.nurse_id])

        # Prior unused-rest credit is accepted only when explicitly verified in Nurses.xlsx.
        # It is NOT inferred from requested OFF, old rosters, overtime, or cache.
        # Per current policy this credit is applicable to independent staff; trainees keep full target
        # unless they have an explicit annual-leave day in the source data.
        state = training[n.nurse_id]
        unused_credit_shifts = 0
        if state.status == IndependenceStatus.INDEPENDENT:
            unused_credit_shifts = int(getattr(n, "prior_unused_hours", 0) or 0) // 8

        targets[n.nurse_id] = max(0, base - annual_credit_shifts - unused_credit_shifts)
    return targets


def feasibility_check(
    nurses: tuple[NurseProfile, ...], requests: tuple[Request, ...], training: Mapping[str, TrainingState],
    targets: Mapping[str, int], *, year: int, month: int,
) -> FeasibilityReport:
    hard: list[Issue] = []
    warnings: list[Issue] = []
    days = month_days(year, month)
    hard_off = {(r.nurse_id, r.day) for r in requests if r.kind in {RequestKind.OFF_LOCK, RequestKind.ANNUAL_LEAVE}}
    must = {(r.nurse_id, r.day): r.shift for r in requests if r.kind == RequestKind.MUST_SHIFT}
    by_id = {n.nurse_id: n for n in nurses}

    for (nid, d), shift in must.items():
        if (nid, d) in hard_off:
            hard.append(Issue("REQUEST_CONFLICT", Severity.HARD, "同日同時存在硬OFF與MUST_SHIFT", nid, d, shift))

    # Static shift/role capacity screening before MILP. It catches obvious impossibility only;
    # dynamic trainee transitions are conservatively treated as potential future capacity.
    for (nid,d),shift in must.items():
        n=by_id[nid]
        allowed,why=shift_allowed_static(n,d,shift,allow_h=True)
        if not allowed:
            hard.append(Issue("MUST_SHIFT_INELIGIBLE",Severity.HARD,f"MUST {shift} 不合法：{why}",nid,d,shift))

    for d in days:
        independent=[n for n in nurses if (n.nurse_id,d) not in hard_off and training[n.nurse_id].status==IndependenceStatus.INDEPENDENT]
        if len(independent) < 19:
            warnings.append(Issue("STATIC_TOTAL_CORE_CAPACITY",Severity.WARNING,f"當日已知獨立且非硬OFF僅{len(independent)}人；最低核心需要19，可能依賴當月新人轉獨立",day=d))
        for shift,roles in (("D", ("Leader", "Triage", "Critical", "Clinic1", "Clinic2", "Observation1", "Observation2")),
                             ("E", ("Leader", "Triage", "Critical", "Clinic1", "Clinic2", "Observation1", "Observation2")),
                             ("N", ("LeaderTriage", "Critical", "Clinic1", "Observation1", "Observation2"))):
            available=[n for n in independent if shift_allowed_static(n,d,shift,allow_h=True)[0]]
            min_core=7 if shift in {"D","E"} else 5
            if len(available) < min_core:
                warnings.append(Issue("STATIC_SHIFT_CAPACITY",Severity.WARNING,f"{shift}靜態獨立候選{len(available)}<{min_core}；可能需等待新人轉獨立或MILP會不可行",day=d,shift=shift))
            for role in roles:
                if not any(role_eligible_on_day(n,role,d) for n in available):
                    warnings.append(Issue("STATIC_ROLE_CAPACITY",Severity.WARNING,f"靜態檢查找不到{shift}/{role}；可能需等待新人轉獨立或MILP會不可行",day=d,shift=shift))

    total_target = sum(targets.values())
    core_floor = len(days) * (7 + 7 + 5)
    unavoidable_ot_lower_bound=max(0,core_floor-total_target)
    if unavoidable_ot_lower_bound:
        warnings.append(Issue("TARGET_BELOW_CORE_FLOOR",Severity.WARNING,
            f"全員目標班數合計{total_target}<最低核心班槽{core_floor}；至少{unavoidable_ot_lower_bound}班超過個人目標才可能填滿核心"))
    metrics = {
        "nurses": len(nurses),
        "days": len(days),
        "target_total_shifts": total_target,
        "minimum_core_shift_slots": core_floor,
        "unavoidable_overtime_lower_bound_shifts": unavoidable_ot_lower_bound,
        "hard_off_cells": len(hard_off),
        "must_shift_cells": len(must),
    }
    return FeasibilityReport(not hard, tuple(hard), tuple(warnings), metrics)


def build_schedule_advice(
    *,
    feasibility: FeasibilityReport | None = None,
    result: ScheduleResult | None = None,
    validation: ValidationResult | None = None,
    config: SchedulerConfig | None = None,
    solver_error: BaseException | str | None = None,
) -> tuple[dict[str, str], ...]:
    """Translate incompatible or relief-dependent schedules into safe actions.

    Advice never edits the roster automatically.  It identifies the blocking
    evidence and presents changes that require a fresh solver run and, where
    applicable, explicit user/manager approval.
    """
    rows: list[dict[str, str]] = []
    seen: set[tuple[str, str, str]] = set()

    def add(priority: str, category: str, evidence: str, action: str, impact: str) -> None:
        key = (category, evidence, action)
        if key in seen:
            return
        seen.add(key)
        rows.append({
            "priority": priority,
            "category": category,
            "evidence": evidence,
            "recommendation": action,
            "impact": impact,
            "approval": "需使用者／主管確認；修改條件後必須重新求解與驗證",
        })

    issues: list[Issue] = []
    if feasibility is not None:
        issues.extend(feasibility.hard_issues)
        issues.extend(feasibility.warnings)
    if validation is not None:
        issues.extend(validation.hard_errors)
        issues.extend(validation.warnings)

    for issue in issues:
        where = " ".join(
            value for value in (
                issue.day.isoformat() if issue.day else "",
                issue.shift or "",
                issue.nurse_id or "",
            ) if value
        )
        evidence = f"{where}｜{issue.message}" if where else issue.message
        if issue.code == "REQUEST_CONFLICT":
            add("P1", "預假／指定班衝突", evidence,
                "回到 LIVE 預假表確認該日只能保留 OFF 或指定班其中一項；禁止由程式自行覆蓋。",
                "解除同一人同日互斥的硬條件。")
        elif issue.code == "MUST_SHIFT_INELIGIBLE":
            add("P1", "指定班別不合法", evidence,
                "取消或改列合法班別；若資格資料有誤，先修正 Nurses.xlsx areas／班組後再排。",
                "避免把沒有資格或休息不足的人強制排入。")
        elif issue.code in {"STATIC_TOTAL_CORE_CAPACITY", "STATIC_SHIFT_CAPACITY"}:
            add("P1", "可用核心人力不足", evidence,
                "依日期與班別新增外部／機動待補；其次才考慮取消非必要訓練或經核准調整預假。",
                "直接補足核心人數，不破壞休息與連續上班規則。")
        elif issue.code in {"STATIC_ROLE_CAPACITY", "CORE_ROLE_SHORTAGE", "CORE_MINIMUM_SHORTAGE"}:
            add("P1", "核心資格不足", evidence,
                "優先尋找同班且具該區域資格者；D/E 的 Leader 或 Triage 缺口可評估由同一名合格者改列 L+T；否則新增對應資格待補。",
                "補的是缺少的資格，不只是增加總人數。")
        elif issue.code == "TARGET_BELOW_CORE_FLOOR":
            add("P1", "正常工時總量低於最低需求", evidence,
                "至少補足訊息所列班數：可選外部待補，或只對具資格的少數人逐人放寬 1 班加班。",
                "這是數學下限；不調整就無法填滿每日核心。")
        elif issue.code in {"PERSON_OVERTIME_CAP", "TARGET_INFEASIBLE"}:
            add("P2", "個人工時上限衝突", evidence,
                "保持一般人上限不變，先新增待補；若外援不可得，只對能解除核心缺口的人員放寬 1 班。",
                "限制加班只落在必要且具資格的人員。")
        elif issue.code in {"PERSON_CROSS_SHIFT_CAP", "CROSS_SHIFT_20H", "REST_11H", "N_RECOVERY"}:
            add("P1", "跨班／休息衝突", evidence,
                "不要手動硬塞班；改找同班替補或調整前後 OFF。若要放寬跨班天數，仍不得放寬 11 小時與跨班 >20 小時休息規則。",
                "避免產生不安全的 D/E/N 轉換。")
        elif "TRAIN" in issue.code or "TEACH" in issue.code:
            add("P2", "訓練與核心人力衝突", evidence,
                "把該訓練移到核心已補足且有 Teaching 人員的日期；Triage 訓練必要時改由合格老師採 L+T。",
                "保留訓練總天數，同時避免學生占用獨立核心。")

    error_text = str(solver_error or "")
    if error_text:
        if "TARGET_INFEASIBLE" in error_text:
            add("P1", "可出勤日不足", error_text,
                "比較目標班數與合法可出勤日；降低目標、修正錯誤預假，或新增待補，不可用違反休息規則的班次補齊。",
                "解除個人容量與目標班數矛盾。")
        if "Teaching roles" in error_text or "訓練" in error_text or "Preceptor" in error_text:
            add("P1", "帶教配置不可行", error_text,
                "移動該日訓練、增加同班 Teaching 資格人員，或把 Triage 訓練改成合格老師 L+T；每班仍最多同時訓練 2 人。",
                "補足師生同班同區與完整核心。")
        if "roles=" in error_text or "核心角色" in error_text:
            add("P1", "核心角色一對一匹配失敗", error_text,
                "依錯誤列出的角色補入具資格人員；若為 D/E Leader＋Triage 組合，可先評估 L+T，其他角色不得互相冒充。",
                "針對真正缺少的資格補人。")
        if "MILP不可行" in error_text or "Infeasible" in error_text or "infeasible" in error_text:
            add("P2", "整體條件組合無解", error_text,
                "依序重跑三個診斷情境：①增加待補；②只對必要人員放寬加班 1 班；③移動非必要訓練日。每次只改一類條件以確認真正原因。",
                "能區分人數不足、資格不足與訓練日期衝突。")

    if result is not None:
        by_id = {n.nurse_id: n for n in result.nurses}
        relief_assignments = [
            assignment for assignment in result.assignments
            if assignment.nurse_id in by_id and _is_relief_profile(by_id[assignment.nurse_id])
        ]
        for assignment in sorted(relief_assignments, key=lambda a: (a.day, a.shift, a.area, a.nurse_id)):
            evidence = (
                f"{assignment.day.isoformat()} {assignment.shift}｜{assignment.area}｜"
                f"目前使用 {by_id[assignment.nurse_id].name}"
            )
            if assignment.area in {"Leader", "Triage"} and assignment.shift in {"D", "E"}:
                action = (
                    "先找同班具相同資格者；若當日有一人同時具 Leader＋Triage，"
                    "可經主管核准改採 L+T，否則填入具資格之外部／機動人員。"
                )
            else:
                action = f"填入具 {assignment.area} 資格的外部／機動人員；若改用院內人員，必須重新檢查工時、跨班與休息。"
            add("P1", "待補核心席位", evidence, action, "將暫用待補代號替換為可追蹤的實際人員。")

        if relief_assignments and config is not None and config.max_overtime_shifts_per_person is not None:
            add("P2", "待補替代方案", f"目前共有 {len(relief_assignments)} 個待補班次",
                f"若外援不足，可試算只對能直接補核心者放寬到最多 {config.max_overtime_shifts_per_person + 1} 班加班；不可全面放寬。",
                "可能減少待補，但會增加少數人的加班。")
        if relief_assignments and config is not None and config.max_cross_shift_days_per_person is not None:
            add("P3", "跨班替代方案", f"目前跨班上限為每人 {config.max_cross_shift_days_per_person} 天",
                "僅在外援與同班人員皆不可得時，試算具資格者增加 1 天跨班；11 小時與跨班 >20 小時休息仍是硬規則。",
                "可能減少待補，但增加跨班負擔。")

    if not rows:
        add("INFO", "未偵測到硬性衝突", "目前資料沒有可定位的硬性錯誤或待補席位。",
            "若求解器仍無解，先增加一個具全核心資格的待補席位重跑，再逐項檢查預假、訓練與跨月休息。",
            "提供保守的第一步診斷，不會自動修改條件。")

    priority_order = {"P1": 0, "P2": 1, "P3": 2, "INFO": 3}
    return tuple(sorted(rows, key=lambda row: (
        priority_order.get(row["priority"], 9), row["category"], row["evidence"]
    )))


# ==============================================================================
# MODULE: solver.py
# ==============================================================================

from collections import defaultdict
from dataclasses import replace
from datetime import date, timedelta
from typing import Any, Mapping
import hashlib
import json

import numpy as np
from scipy.optimize import Bounds, LinearConstraint, milp
from scipy.sparse import coo_matrix



class _Model:
    def __init__(self) -> None:
        self.names: list[str] = []
        self.lb: list[float] = []
        self.ub: list[float] = []
        self.integrality: list[int] = []
        self.rows: list[dict[int, float]] = []
        self.row_lb: list[float] = []
        self.row_ub: list[float] = []

    def var(self, name: str, lb: float = 0, ub: float = 1, integer: bool = True) -> int:
        i = len(self.names)
        self.names.append(name); self.lb.append(lb); self.ub.append(ub); self.integrality.append(1 if integer else 0)
        return i

    def con(self, coeff: Mapping[int, float], lb: float = -np.inf, ub: float = np.inf) -> None:
        self.rows.append(dict(coeff)); self.row_lb.append(lb); self.row_ub.append(ub)

    def compile(self) -> tuple[Bounds, LinearConstraint]:
        rr=[]; cc=[]; vv=[]
        for r, row in enumerate(self.rows):
            for c, v in row.items():
                if v:
                    rr.append(r); cc.append(c); vv.append(v)
        A = coo_matrix((vv, (rr, cc)), shape=(len(self.rows), len(self.names))).tocsr()
        return Bounds(np.array(self.lb), np.array(self.ub)), LinearConstraint(A, np.array(self.row_lb), np.array(self.row_ub))


def _result_hash(assignments: tuple[Assignment, ...], google_audit: GoogleSourceAudit, previous_audit: PreviousRosterAudit, targets: Mapping[str,int], training_end: Mapping[str,TrainingState]) -> str:
    rows = [
        (a.nurse_id, a.day.isoformat(), a.shift, a.area or "", a.source, a.formal_teaching, a.preceptor_id or "", a.independent_core)
        for a in sorted(assignments, key=lambda a: (a.nurse_id, a.day, a.shift, a.area or ""))
    ]
    payload = json.dumps({
        "assignments": rows,
        "google": google_audit.matrix_sha256,
        "previous_google": previous_audit.matrix_sha256,
        "targets": sorted((str(k),int(v)) for k,v in targets.items()),
        "training": sorted((str(k),v.status.value,None if v.completed_observation_sessions is None else int(v.completed_observation_sessions)) for k,v in training_end.items()),
    }, ensure_ascii=False, separators=(",", ":"))
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def solve_schedule(
    nurses: tuple[NurseProfile, ...], requests: tuple[Request, ...], training_start: Mapping[str, TrainingState],
    boundary: tuple[BoundaryEntry, ...], targets: Mapping[str, int], google_audit: GoogleSourceAudit,
    previous_audit: PreviousRosterAudit, config: SchedulerConfig,
) -> ScheduleResult:
    from itertools import combinations

    days = month_days(config.year, config.month)
    role_floor = {"D": len(D_E_CORE_ROLES), "E": len(D_E_CORE_ROLES), "N": len(N_CORE_ROLES)}
    configured_floor = {"D": config.d_min_core, "E": config.e_min_core, "N": config.n_min_core}
    for shift, minimum in configured_floor.items():
        if minimum < role_floor[shift]:
            raise ValueError(
                f"{shift}_min_core={minimum} 小於必要角色數{role_floor[shift]}；"
                "設定值不得與核心角色清單矛盾"
            )
    by_id = {n.nurse_id: n for n in nurses}
    relief_ids=set(config.relief_ids)
    conditional_plan_c_ids=set(config.conditional_plan_c_ids)
    unknown_relief_ids=relief_ids-set(by_id)
    if unknown_relief_ids:
        raise ValueError(f"relief_ids 含未知人員：{sorted(unknown_relief_ids)}")
    unknown_plan_c_ids=conditional_plan_c_ids-set(by_id)
    if unknown_plan_c_ids:
        raise ValueError(f"conditional_plan_c_ids 含未知人員：{sorted(unknown_plan_c_ids)}")
    if conditional_plan_c_ids & relief_ids:
        raise ValueError("Plan C院內人員不得同時列為外部待補")
    if config.relief_shift_cap is not None and config.relief_shift_cap < 0:
        raise ValueError("relief_shift_cap must be >= 0 or None")
    nurse_order = {n.nurse_id:i for i,n in enumerate(nurses)}
    hard_off = {(r.nurse_id, r.day) for r in requests if r.kind in {RequestKind.OFF_LOCK, RequestKind.ANNUAL_LEAVE}}
    must = {(r.nurse_id, r.day): r.shift for r in requests if r.kind == RequestKind.MUST_SHIFT}
    pref = {(r.nurse_id, r.day): r.shift for r in requests if r.kind == RequestKind.PREFERRED_SHIFT}
    avoid: dict[tuple[str, date], set[str]] = defaultdict(set)
    for r in requests:
        if r.kind == RequestKind.AVOID_SHIFT and r.shift:
            avoid[(r.nurse_id, r.day)].add(r.shift)
    bmap = {(e.nurse_id, e.day): e for e in boundary}

    def eligible_training_preceptors(day: date, home_shift: str) -> frozenset[str]:
        """Teaching-qualified same-home-shift staff who can occupy Observation."""
        return frozenset(
            pid for pid in config.preceptor_ids
            if pid in by_id
            and by_id[pid].home_shift == home_shift
            and "Teaching" in by_id[pid].areas
            and role_eligible_on_day(by_id[pid], "Observation", day)
        )

    # Prove each person's maximum legal attendance. Part-time uses explicit availability + max5 only; singleton is a regular/full-time roster-shape rule.
    # IMPORTANT FINAL-5.4: capacity is a FEASIBILITY GATE only. It may NOT silently lower
    # a person's monthly target. Legitimate target reductions must already be encoded by
    # explicit paid credits (annual leave / verified prior-unused hours). Requested OFF alone
    # does not reduce the target. If the remaining legal dates cannot reach the target, STOP.
    effective_targets: dict[str,int] = dict((str(k), int(v)) for k,v in targets.items())
    capacity_adjustments: list[dict[str,Any]] = []
    for n in nurses:
        n_hard = {d for nid,d in hard_off if nid == n.nurse_id}
        cap, witness = max_attendance_capacity_under_fixed_rules(
            n, days, hard_off_days=n_hard, boundary=bmap,
            max_consecutive=config.max_consecutive_attendance,
            official_holidays=config.official_holidays,
            enforce_singleton=(config.debug_enforce_singleton and not n.part_time),
        )
        nominal = int(targets[n.nurse_id])
        if cap < nominal:
            if n.part_time:
                # Keep the published/nominal availability target unchanged for
                # warning and audit, but make the exact-target solve aim at the
                # mathematically proven maximum. This prevents one unavoidable
                # part-time shortfall from forcing the whole roster into the
                # flexible stage and needlessly under-scheduling other staff.
                effective_targets[n.nurse_id] = cap
                capacity_adjustments.append({
                    "nurse_id": n.nurse_id,
                    "name": n.name,
                    "nominal_target": nominal,
                    "proven_capacity": cap,
                    "constrained_under_shifts": nominal-cap,
                    "reason": "part-time minimum unavailable under explicit OFF/safety rules; warning-only",
                })
                continue
            state = training_start[n.nurse_id]
            category = "非獨立訓練人員" if state.status == IndependenceStatus.NON_INDEPENDENT else "人員"
            raise RuntimeError(
                f"TARGET_INFEASIBLE: {category} {n.name}({n.nurse_id}) 目標{nominal}班，"
                f"在硬OFF/max{config.max_consecutive_attendance}"
                + ("/singleton" if not n.part_time else "")
                + f"規則下最多只能{cap}班。"
                "禁止以任何fallback自動降低應上班天數；請調整OFF/假別/折抵時數或規則。"
            )

    m = _Model()
    shifts = ("D", "E", "N", "H")
    x: dict[tuple[str,date,str],int] = {}
    w: dict[tuple[str,date],int] = {}
    for n in nurses:
        for d in days:
            w[(n.nurse_id,d)] = m.var(f"w|{n.nurse_id}|{d}")
            for s in shifts:
                ub = 1.0
                if (n.nurse_id,d) in hard_off: ub = 0
                allowed,_ = shift_allowed_static(n,d,s,allow_h=config.allow_h)
                date_allowed,_ = part_time_workday_allowed(n,d,official_holidays=config.official_holidays)
                if not allowed or not date_allowed: ub = 0
                # A person who is non-independent at month start remains on the
                # Nurses.xlsx home shift for the full active month.  This avoids
                # silently moving a D-group trainee to E after an in-month age/
                # session transition; any normal cross-shift support starts next month.
                if (training_start[n.nurse_id].status==IndependenceStatus.NON_INDEPENDENT
                        and s != n.home_shift):
                    ub=0
                x[(n.nurse_id,d,s)] = m.var(f"x|{n.nurse_id}|{d}|{s}", ub=ub)
            coeff = {x[(n.nurse_id,d,s)]:1 for s in shifts}; coeff[w[(n.nurse_id,d)]]=-1
            m.con(coeff,lb=0,ub=0)
            if (n.nurse_id,d) in must:
                target=must[(n.nurse_id,d)]
                if target not in shifts: raise ValueError(f"不支援 MUST_SHIFT={target}")
                m.con({x[(n.nurse_id,d,target)]:1},lb=1,ub=1)

    trainees=[n for n in nurses if training_start[n.nurse_id].status==IndependenceStatus.NON_INDEPENDENT]
    trainee_ids={n.nurse_id for n in trainees}
    z: dict[tuple[str,date],int]={}
    t: dict[tuple[str,date,str],int]={}
    q: dict[tuple[str,date,str],int]={}
    for n in trainees:
        prior_raw=training_start[n.nurse_id].completed_observation_sessions
        if prior_raw is None:
            raise ValueError(f"非獨立作業者 {n.name}({n.nurse_id}) 缺少正式Preceptor帶教進度")
        prior=int(prior_raw)
        M=max(18,len(days)+prior)
        for di,d in enumerate(days):
            z[(n.nurse_id,d)]=m.var(f"z|{n.nurse_id}|{d}")
            prev=[t[(n.nurse_id,days[k],s)] for k in range(di) for s in ("D","E") if (n.nurse_id,days[k],s) in t]
            # Externally planned training is materialized after the shift solve,
            # but it must still advance the trainee's canonical 18-session
            # independence state on later dates.  Count only sessions strictly
            # before the active day; same-day training is still a shadow seat.
            planned_prior=sum(
                1
                for (training_day, _training_shift), planned_ids
                in config.required_training_trainees_by_day_shift.items()
                if n.nurse_id in planned_ids and training_day < d
            )
            effective_prior=prior+planned_prior
            if age_independent_on_day(n,d):
                # 2.5-month OR gate: service age alone makes this and later workdays independent.
                m.con({z[(n.nurse_id,d)]:1},lb=1,ub=1)
            else:
                # Before 2.5 months, independence can still be reached by 18 formal teaching sessions.
                c1={v:1 for v in prev}; c1[z[(n.nurse_id,d)]]=-18; m.con(c1,lb=-effective_prior)
                c2={v:1 for v in prev}; c2[z[(n.nurse_id,d)]]=-M; m.con(c2,ub=17-effective_prior)
            for s in ("D","E"):
                tv=m.var(f"teach|{n.nurse_id}|{d}|{s}"); t[(n.nurse_id,d,s)]=tv
                m.con({tv:1,x[(n.nurse_id,d,s)]:-1},ub=0)
                m.con({tv:1,z[(n.nurse_id,d)]:1},ub=1)
            for s in ("D","E","N"):
                qv=m.var(f"core|{n.nurse_id}|{d}|{s}"); q[(n.nurse_id,d,s)]=qv
                # q = x AND z
                m.con({qv:1,x[(n.nurse_id,d,s)]:-1},ub=0)
                m.con({qv:1,z[(n.nurse_id,d)]:-1},ub=0)
                m.con({qv:1,x[(n.nurse_id,d,s)]:-1,z[(n.nurse_id,d)]:-1},lb=-1)
    def core_var(n: NurseProfile,d: date,s: str) -> int:
        return q[(n.nurse_id,d,s)] if n.nurse_id in trainee_ids else x[(n.nurse_id,d,s)]

    forced_shadow_keys = {
        (nurse_id, workday, shift)
        for (workday, shift), nurse_ids in config.required_training_trainees_by_day_shift.items()
        for nurse_id in nurse_ids
    }

    # Internal high-core matching layer.  The normal Hall constraints below
    # prove that the whole core can be assigned, including anonymous relief.
    # These additional role variables separately measure whether distinct
    # in-house workers can cover the highest-priority roles.  Their gaps are
    # minimized before cross-shift/aesthetic costs, so OFF is not handed to a
    # qualified senior worker while a relief seat fills Leader/Triage/Critical.
    priority_role_gap: dict[tuple[date, str, str], int] = {}
    priority_role_assignment: dict[tuple[str, date, str, str], int] = {}
    if config.prioritize_internal_high_core:
        for workday in days:
            for shift in ("D", "E", "N"):
                required_training_roles = set(
                    config.required_training_roles_by_day_shift.get((workday, shift), ())
                )
                if shift == "N" or "LeaderTriage" in required_training_roles:
                    priority_roles = ("LeaderTriage", "Critical")
                else:
                    priority_roles = ("Leader", "Triage", "Critical")

                assignments_by_nurse: dict[str, list[int]] = defaultdict(list)
                for role in priority_roles:
                    gap = m.var(f"internal_high_gap|{workday}|{shift}|{role}")
                    priority_role_gap[(workday, shift, role)] = gap
                    role_coeff = {gap: 1}
                    for nurse in nurses:
                        if (
                            nurse.nurse_id in relief_ids
                            or (nurse.nurse_id, workday, shift) in forced_shadow_keys
                            or not role_eligible_on_day(nurse, role, workday)
                        ):
                            continue
                        # When this exact core role is also the planned
                        # training role, its internal occupant must be a
                        # same-home-shift Teaching-qualified preceptor.
                        if role in required_training_roles and not (
                            nurse.home_shift == shift and "Teaching" in nurse.areas
                        ):
                            continue
                        variable = m.var(
                            f"internal_high_role|{nurse.nurse_id}|{workday}|{shift}|{role}"
                        )
                        priority_role_assignment[(nurse.nurse_id, workday, shift, role)] = variable
                        m.con({variable: 1, core_var(nurse, workday, shift): -1}, ub=0)
                        role_coeff[variable] = 1
                        assignments_by_nurse[nurse.nurse_id].append(variable)
                    # Exactly one distinct in-house occupant or one audited gap.
                    m.con(role_coeff, lb=1, ub=1)
                for variables in assignments_by_nurse.values():
                    m.con({variable: 1 for variable in variables}, ub=1)

    for (workday, shift), nurse_ids in config.required_training_trainees_by_day_shift.items():
        roles = tuple(config.required_training_roles_by_day_shift.get((workday, shift), ()))
        if shift not in {"D", "E"} or workday not in days:
            raise ValueError(f"指定訓練日期／班別不合法：{workday} {shift}")
        if len(nurse_ids) != len(roles) or len(roles) != len(set(roles)):
            raise ValueError(f"指定訓練學生與角色數量不一致或角色重複：{workday} {shift}")
        for nurse_id in nurse_ids:
            if nurse_id not in by_id:
                raise ValueError(f"指定訓練含未知人員：{nurse_id}")
            m.con({x[(nurse_id, workday, shift)]: 1}, lb=1, ub=1)
        combined_l_plus_t = "LeaderTriage" in roles
        required_core_roles = (
            D_E_L_PLUS_T_CORE_ROLES if combined_l_plus_t else D_E_CORE_ROLES
        )
        # Student shadow seats are additional to the complete independent core.
        m.con(
            {x[(n.nurse_id, workday, shift)]: 1 for n in nurses},
            lb=len(required_core_roles) + len(nurse_ids),
        )
        # Full Hall theorem over the required Teaching roles guarantees that
        # distinct same-home-shift Teaching staff can cover every student area.
        for size in range(1, len(roles) + 1):
            for subset in combinations(range(len(roles)), size):
                subset_roles = tuple(roles[index] for index in subset)
                coefficients = {
                    x[(n.nurse_id, workday, shift)]: 1
                    for n in nurses
                    if n.nurse_id not in nurse_ids
                    and n.home_shift == shift
                    and "Teaching" in n.areas
                    and any(role_eligible_on_day(n, role, workday) for role in subset_roles)
                }
                if not coefficients:
                    raise RuntimeError(
                        f"MODEL_INFEASIBLE_STATIC: {workday} {shift} Teaching roles={subset_roles}"
                    )
                m.con(coefficients, lb=size)

    # N staffing mode is driven by the actual headcount.  At five people the
    # core contains a combined Leader+Triage role.  Once a sixth person is
    # scheduled, all six must be independent core-capable and the model must
    # admit a separate Leader and Triage assignment.
    n_expanded: dict[date, int] = {}
    if config.n_max_total >= 6:
        for d in days:
            nv=m.var(f"n_expanded|{d}")
            n_expanded[d]=nv
            total_coeff={x[(n.nurse_id,d,"N")]:1 for n in nurses}
            total_coeff[nv]=-(config.n_max_total-5)
            m.con(total_coeff,ub=5)
            total_lower={x[(n.nurse_id,d,"N")]:1 for n in nurses}
            total_lower[nv]=-6
            m.con(total_lower,lb=0)
            independent_lower={core_var(n,d,"N"):1 for n in nurses}
            independent_lower[nv]=-6
            m.con(independent_lower,lb=0)

    # Formal teaching: at most one trainee per D/E shift and at least one whitelisted,
    # same-home-shift Leader-qualified Preceptor must be present. Exact "Preceptor fixed to Observation while
    # all remaining core roles still match" is enforced by iterative cut generation after
    # each candidate MILP solution, avoiding tens of thousands of conditional Hall rows.
    for d in days:
        for s in ("D","E"):
            vars_t=[t[(n.nurse_id,d,s)] for n in trainees]
            if not vars_t:
                continue
            m.con({v:1 for v in vars_t},ub=1)
            coeff={v:1 for v in vars_t}
            for pid in eligible_training_preceptors(d,s):
                coeff[x[(pid,d,s)]]=coeff.get(x[(pid,d,s)],0)-1
            m.con(coeff,ub=0)

    if config.required_training_trainees_by_day_shift:
        # The externally planned programme is materialized and validated in a
        # dedicated same-area pass.  Do not let the legacy one-trainee solver
        # variables invent extra sessions or use them to inflate core status.
        for teaching_variable in t.values():
            m.con({teaching_variable: 1}, lb=0, ub=0)

    # Hall constraints guarantee that the selected independent staff admit an exact one-person-per-role matching.
    if config.debug_enforce_core_roles:
      for d in days:
        for s,default_roles in (("D",D_E_CORE_ROLES),("E",D_E_CORE_ROLES),("N",N_CORE_ROLES)):
            required_training_roles=tuple(
                config.required_training_roles_by_day_shift.get((d,s),())
            )
            roles=tuple(
                D_E_L_PLUS_T_CORE_ROLES
                if s in {"D","E"} and "LeaderTriage" in required_training_roles
                else default_roles
            )
            required_teaching_roles=set(
                required_training_roles
            )
            # Full Hall theorem, not a partial approximation: every role subset is constrained.
            # D/E has 7 roles (127 non-empty subsets), N has 5 (31); still small enough for monthly MILP.
            for size in range(1, len(roles) + 1):
                for subset in combinations(range(len(roles)),size):
                    subset_roles=tuple(roles[i] for i in subset)
                    coeff={}
                    for n in nurses:
                        if (n.nurse_id,d,s) in forced_shadow_keys:
                            continue
                        if any(
                            role_eligible_on_day(n,r,d)
                            and (
                                r not in required_teaching_roles
                                or (
                                    n.home_shift==s
                                    and "Teaching" in n.areas
                                )
                            )
                            for r in subset_roles
                        ):
                            coeff[core_var(n,d,s)]=1
                    if not coeff:
                        raise RuntimeError(f"MODEL_INFEASIBLE_STATIC: {d} {s} roles={subset_roles}")
                    m.con(coeff,lb=size)

        # Conditional Hall theorem for the six-person N configuration.  These
        # rows activate only when actual N headcount reaches six, guaranteeing
        # that the sixth person is materialized as a separate Triage role.
        if d in n_expanded:
            roles=tuple(N_EXPANDED_CORE_ROLES)
            for size in range(1,len(roles)+1):
                for subset in combinations(range(len(roles)),size):
                    subset_roles=tuple(roles[i] for i in subset)
                    coeff={}
                    for n in nurses:
                        if any(role_eligible_on_day(n,r,d) for r in subset_roles):
                            coeff[core_var(n,d,"N")]=1
                    coeff[n_expanded[d]]=coeff.get(n_expanded[d],0)-size
                    m.con(coeff,lb=0)

    # Optional E Clinic3 support.  A positive value remains configurable for a
    # future policy change, but the current user rule sets the minimum to zero.
    if config.e_clinic3_flexible_min > 0:
        for d in days:
            eligible={x[(n.nurse_id,d,"E")]:1 for n in nurses if ("Clinic3" in n.areas and not n.part_time)}
            if len(eligible) < config.e_clinic3_flexible_min:
                raise RuntimeError(f"MODEL_INFEASIBLE_STATIC: {d} E Clinic3 eligible pool<{config.e_clinic3_flexible_min}")
            m.con(eligible,lb=config.e_clinic3_flexible_min)

    for d in days:
        for s,cap in (("D",config.d_max_total),("E",config.e_max_total),("N",config.n_max_total)):
            m.con({x[(n.nurse_id,d,s)]:1 for n in nurses},ub=cap)
        for s,minimum in configured_floor.items():
            if (
                s in {"D","E"}
                and "LeaderTriage" in config.required_training_roles_by_day_shift.get((d,s),())
            ):
                minimum=len(D_E_L_PLUS_T_CORE_ROLES)
            m.con({
                core_var(n,d,s):1 for n in nurses
                if (n.nurse_id,d,s) not in forced_shadow_keys
            },lb=minimum)

    if config.relief_shift_cap is not None:
        m.con(
            {
                x[(nurse_id,d,s)]:1
                for nurse_id in relief_ids
                for d in days
                for s in ("D","E","N")
            },
            ub=config.relief_shift_cap,
        )

    if config.moved_off_assignment_cap is not None:
        if config.moved_off_assignment_cap < 0:
            raise ValueError("moved_off_assignment_cap must be >= 0 or None")
        m.con(
            {
                w[(nurse_id, workday)]: 1
                for nurse_id, workday in config.movable_off_assignments
                if nurse_id in by_id and workday in days
            },
            ub=config.moved_off_assignment_cap,
        )

    if config.max_cross_shift_days_per_person is not None:
        if config.max_cross_shift_days_per_person < 0:
            raise ValueError("max_cross_shift_days_per_person must be >= 0 or None")
        for n in nurses:
            # Relief IDs are anonymous day-level vacancy seats, not the same
            # physical worker carried from one date to the next.
            if n.nurse_id in relief_ids:
                continue
            coeff={
                x[(n.nurse_id,d,s)]:1
                for d in days for s in ("D","E","N")
                if s!=n.home_shift
                and (n.nurse_id,d,s) not in config.cross_shift_cap_exempt_assignments
            }
            m.con(coeff,ub=config.max_cross_shift_days_per_person)

    for n in nurses:
        m.con({x[(n.nurse_id,d,"H")]:1 for d in days},ub=config.h_person_cap)
        if n.max_night_shifts is not None:
            m.con({x[(n.nurse_id,d,"N")]:1 for d in days},ub=n.max_night_shifts)

    first=days[0]
    timeline=[first-timedelta(days=i) for i in range(config.max_consecutive_attendance,0,-1)]+list(days)
    if config.debug_enforce_max_consecutive:
        for n in nurses:
            if n.nurse_id in relief_ids:
                continue
            for end_i in range(config.max_consecutive_attendance,len(timeline)):
                window=timeline[end_i-config.max_consecutive_attendance:end_i+1]; coeff={};fixed=0
                for d in window:
                    if d in days: coeff[w[(n.nurse_id,d)]]=1
                    else: fixed+=1 if (bmap.get((n.nurse_id,d)) and bmap[(n.nurse_id,d)].attendance) else 0
                m.con(coeff,ub=config.max_consecutive_attendance-fixed)

    # Regular/full-time only: strict OFF -> WORK -> OFF hard gate.
    # Part-time availability may legitimately create isolated working days, so part-time is exempt.
    if config.debug_enforce_singleton:
        for n in nurses:
            if n.part_time or n.nurse_id in relief_ids:
                continue
            for d in days:
                # w[d] <= w[d-1] + w[d+1]
                coeff={w[(n.nurse_id,d)]:1}
                fixed_attendance=0
                known=True
                for nd in (d-timedelta(days=1), d+timedelta(days=1)):
                    if nd in days:
                        coeff[w[(n.nurse_id,nd)]]=coeff.get(w[(n.nurse_id,nd)],0)-1
                    elif nd < first:
                        e=bmap.get((n.nurse_id,nd))
                        if e is None:
                            known=False
                        elif e.attendance:
                            fixed_attendance += 1
                    else:
                        # At month-end, the next-month state is unknown. Prevent a
                        # provable isolated final-day work block by requiring the
                        # previous day to be attendance when the last day is worked.
                        # A future roster may extend the block, but cannot be relied on
                        # to validate this month's hard rule.
                        known=True
                if known:
                    m.con(coeff, ub=fixed_attendance)

    # Regular/full-time only: every six-day nonwork window wholly inside the
    # active month must contain either attendance or an explicit annual-leave
    # day. Cross-month OFF runs are explicitly approvable and excluded here.
    # Part-time explicit OFF/availability patterns are authoritative and are exempt from this roster-shape rule.
    if config.debug_enforce_max_consecutive_off:
        annual_days={(r.nurse_id,r.day) for r in requests if r.kind==RequestKind.ANNUAL_LEAVE}
        off_timeline=list(days)
        for n in nurses:
            if n.part_time or n.nurse_id in relief_ids:
                continue
            for end_i in range(5, len(off_timeline)):
                window=off_timeline[end_i-5:end_i+1]
                # At least one attendance day in every 6-day window ending in the active month.
                coeff={}
                fixed_attendance=0
                fixed_annual=0
                for wd in window:
                    if (n.nurse_id,wd) in annual_days:
                        fixed_annual += 1
                    if wd in days:
                        coeff[w[(n.nurse_id,wd)]]=coeff.get(w[(n.nurse_id,wd)],0)+1
                    else:
                        e=bmap.get((n.nurse_id,wd))
                        if e and e.attendance:
                            fixed_attendance += 1
                if fixed_attendance==0 and fixed_annual==0:
                    m.con(coeff, lb=1)

    if config.debug_enforce_rest:
        for n in nurses:
            if n.nurse_id in relief_ids:
                continue
            for d in days:
                prev=d-timedelta(days=1)
                if prev in days:
                    for ps in shifts:
                        for cs in shifts:
                            rh=rest_hours(prev,ps,d,cs)
                            if rh is None:
                                continue
                            cross_den=(ps in {"D","E","N"} and cs in {"D","E","N"} and ps!=cs)
                            violates=(rh < config.min_rest_hours) or (
                                config.debug_enforce_cross_shift_20h and cross_den and rh <= 20.0
                            )
                            if violates:
                                m.con({x[(n.nurse_id,prev,ps)]:1,x[(n.nurse_id,d,cs)]:1},ub=1)
                else:
                    be=bmap.get((n.nurse_id,prev))
                    if be and be.attendance:
                        for cs in shifts:
                            rh=rest_hours(prev,be.shift,d,cs)
                            if rh is None:
                                continue
                            cross_den=(be.shift in {"D","E","N"} and cs in {"D","E","N"} and be.shift!=cs)
                            violates=(rh < config.min_rest_hours) or (
                                config.debug_enforce_cross_shift_20h and cross_den and rh <= 20.0
                            )
                            if violates:
                                m.con({x[(n.nurse_id,d,cs)]:1},ub=0)

    # N recovery: within any attendance block, N and D/E cannot coexist. H/admin/special are attendance, not recovery OFF.
    def current_endpoint(nid: str,d: date,fam: str) -> dict[int,float] | None:
        if d in days:
            return {x[(nid,d,"N")]:1} if fam=="N" else {x[(nid,d,"D")]:1,x[(nid,d,"E")]:1}
        be=bmap.get((nid,d))
        if be is None: return None
        actual="N" if be.shift=="N" else "DE" if be.shift in {"D","E"} else "OTHER"
        return {} if actual==fam else None
    seq_start=first-timedelta(days=config.max_consecutive_attendance)
    seq=[seq_start+timedelta(days=i) for i in range((days[-1]-seq_start).days+1)]
    if config.debug_enforce_n_recovery:
        # A legacy official roster may already contain an N<->D/E violation entirely before
        # the active month.  Do NOT add an impossible constant-only constraint for history we
        # cannot change.  If that invalid attendance block reaches month-end, force day 1 OFF
        # so the active roster terminates the legacy block instead of extending it.
        for n in nurses:
            if n.nurse_id in relief_ids:
                continue
            fam=set(); cur=first-timedelta(days=1)
            while True:
                be=bmap.get((n.nurse_id,cur))
                if be is None or not be.attendance:
                    break
                if be.shift=="N": fam.add("N")
                elif be.shift in {"D","E"}: fam.add("DE")
                cur-=timedelta(days=1)
            if "N" in fam and "DE" in fam:
                m.con({w[(n.nurse_id,first)]:1},ub=0)
        for n in nurses:
            if n.nurse_id in relief_ids:
                continue
            for i,a in enumerate(seq):
                for j in range(i+1,min(len(seq),i+config.max_consecutive_attendance+1)):
                    b=seq[j]
                    if b < first:
                        continue
                    distance=j-i
                    for fa,fb in (("N","DE"),("DE","N")):
                        av=current_endpoint(n.nurse_id,a,fa);bv=current_endpoint(n.nurse_id,b,fb)
                        if av is None or bv is None: continue
                        fixed_end=(1 if av=={} else 0)+(1 if bv=={} else 0);coeff={};fixed_mid=0
                        for src in (av,bv):
                            if src:
                                for k,v in src.items(): coeff[k]=coeff.get(k,0)+v
                        for md in seq[i+1:j]:
                            if md in days: coeff[w[(n.nurse_id,md)]]=coeff.get(w[(n.nurse_id,md)],0)+1
                            else:
                                e=bmap.get((n.nurse_id,md));fixed_mid+=1 if e and e.attendance else 0
                        m.con(coeff,ub=distance-fixed_end-fixed_mid)

    over={};under={};block_start={}
    for n in nurses:
        over[n.nurse_id]=m.var(f"over|{n.nurse_id}",lb=0,ub=len(days),integer=False)
        under[n.nurse_id]=m.var(f"under|{n.nurse_id}",lb=0,ub=len(days),integer=False)
        if config.require_exact_targets:
            # Fairness-first feasibility: everyone must hit the canonical monthly target exactly.
            # If this model is feasible, OT=UNDER=0 is mathematically proven without needing
            # global optimality of lower-priority aesthetic preferences.
            c={w[(n.nurse_id,d)]:1 for d in days}
            m.con(c,lb=effective_targets[n.nurse_id],ub=effective_targets[n.nurse_id])
            m.con({over[n.nurse_id]:1},lb=0,ub=0)
            m.con({under[n.nurse_id]:1},lb=0,ub=0)
        else:
            c={w[(n.nurse_id,d)]:1 for d in days};c[over[n.nurse_id]]=-1;m.con(c,ub=effective_targets[n.nurse_id])
            c={w[(n.nurse_id,d)]:-1 for d in days};c[under[n.nurse_id]]=-1;m.con(c,ub=-effective_targets[n.nurse_id])
            if (
                config.max_overtime_shifts_per_person is not None
                and n.nurse_id not in relief_ids
                and n.nurse_id not in conditional_plan_c_ids
            ):
                if config.max_overtime_shifts_per_person < 0:
                    raise ValueError("max_overtime_shifts_per_person must be >= 0 or None")
                m.con(
                    {over[n.nurse_id]: 1},
                    ub=config.max_overtime_shifts_per_person,
                )
            if config.protect_part_time_targets and n.part_time:
                m.con({w[(n.nurse_id,d)]:1 for d in days},lb=effective_targets[n.nurse_id])
            if n.nurse_id in config.exact_target_ids:
                m.con(
                    {w[(n.nurse_id,d)]:1 for d in days},
                    lb=effective_targets[n.nurse_id],
                    ub=effective_targets[n.nurse_id],
                )
                m.con({over[n.nurse_id]:1},lb=0,ub=0)
                m.con({under[n.nurse_id]:1},lb=0,ub=0)
            if (
                config.real_target_equality
                and n.nurse_id not in relief_ids
                and not n.part_time
                and n.nurse_id not in config.exact_target_ids
            ):
                m.con(
                    {w[(n.nurse_id,d)]:1 for d in days},
                    lb=effective_targets[n.nurse_id],
                    ub=effective_targets[n.nurse_id],
                )
                m.con({over[n.nurse_id]:1},lb=0,ub=0)
                m.con({under[n.nurse_id]:1},lb=0,ub=0)
            if (
                config.real_target_minimum
                and n.nurse_id not in relief_ids
                and not n.part_time
                and n.nurse_id not in config.exact_target_ids
            ):
                m.con(
                    {w[(n.nurse_id,d)]:1 for d in days},
                    lb=effective_targets[n.nurse_id],
                )
                m.con({under[n.nurse_id]:1},lb=0,ub=0)
        if not config.require_exact_targets:
            for i,d in enumerate(days):
                bs=m.var(f"blockstart|{n.nurse_id}|{d}");block_start[(n.nurse_id,d)]=bs
                c={bs:1,w[(n.nurse_id,d)]:-1}
                if i: c[w[(n.nurse_id,days[i-1])]]=1;m.con(c,lb=0)
                else:
                    prev=bmap.get((n.nurse_id,d-timedelta(days=1)));m.con(c,lb=-(1 if prev and prev.attendance else 0))

    group_max_under={}
    group_max_over={}
    for group in ("D","E","N"):
        members=[n for n in nurses if n.home_shift==group and not n.part_time]
        if members:
            gu=m.var(f"maxunder|{group}",lb=0,ub=len(days),integer=False);group_max_under[group]=gu
            go=m.var(f"maxover|{group}",lb=0,ub=len(days),integer=False);group_max_over[group]=go
            for n in members:
                m.con({under[n.nurse_id]:1,gu:-1},ub=0)
                m.con({over[n.nurse_id]:1,go:-1},ub=0)

    if config.total_overtime_shift_cap is not None:
        if config.total_overtime_shift_cap < 0:
            raise ValueError("total_overtime_shift_cap must be >= 0 or None")
        m.con(
            {
                over[n.nurse_id]: 1
                for n in nurses
                if n.nurse_id not in relief_ids
                and n.nurse_id not in conditional_plan_c_ids
            },
            ub=config.total_overtime_shift_cap,
        )
    if config.conditional_plan_c_shift_cap is not None:
        if config.conditional_plan_c_shift_cap < 0:
            raise ValueError("conditional_plan_c_shift_cap must be >= 0 or None")
        m.con(
            {
                w[(nurse_id, workday)]: 1
                for nurse_id in conditional_plan_c_ids
                for workday in days
            },
            ub=config.conditional_plan_c_shift_cap,
        )

    integ=np.array(m.integrality)

    # Single MILP with mathematically separated priorities. If HiGHS proves OPTIMAL,
    # the weight bands are lexicographic: OT > under-target > group fairness > aesthetics.
    c=np.zeros(len(m.names))
    aesthetic_coeff={}
    for n in nurses:
        for d in days:
            for s in shifts:
                v=x[(n.nurse_id,d,s)];cost=0.0
                if s=="H":cost+=config.objective.h_shift
                if s!=n.home_shift and s!="H":cost+=config.objective.cross_shift
                if pref.get((n.nurse_id,d)) and pref[(n.nurse_id,d)]!=s:cost+=config.objective.preferred_miss
                if s in avoid.get((n.nurse_id,d),set()):cost+=config.objective.avoid_violation
                if cost:aesthetic_coeff[v]=aesthetic_coeff.get(v,0.0)+cost
            if (n.nurse_id,d) in block_start:
                aesthetic_coeff[block_start[(n.nurse_id,d)]]=aesthetic_coeff.get(block_start[(n.nurse_id,d)],0.0)+config.objective.work_block_start
    for v in t.values():aesthetic_coeff[v]=aesthetic_coeff.get(v,0.0)-config.objective.formal_teaching_reward
    aesthetic_bound=sum(abs(v) for v in aesthetic_coeff.values())+1.0
    priority_gap_vars=list(priority_role_gap.values())
    priority_gap_bound=max(1.0,float(len(priority_gap_vars)))
    priority_gap_weight=aesthetic_bound+1.0
    fairness_vars=[*group_max_under.values(),*group_max_over.values()]
    fairness_bound=max(1.0,30.0*max(1,len(fairness_vars)))
    fairness_weight=priority_gap_weight*priority_gap_bound+aesthetic_bound+1.0
    under_weight=fairness_weight*fairness_bound+priority_gap_weight*priority_gap_bound+aesthetic_bound+1.0
    under_bound=float(len(days)*len(nurses))
    overtime_weight=(
        under_weight*under_bound
        + fairness_weight*fairness_bound
        + priority_gap_weight*priority_gap_bound
        + aesthetic_bound+1.0
    )
    for v in over.values():c[v]=overtime_weight
    for v in under.values():c[v]=under_weight
    for v in fairness_vars:c[v]=fairness_weight
    for v in priority_gap_vars:c[v]=priority_gap_weight
    for v,cost in aesthetic_coeff.items():c[v]+=cost
    if config.debug_zero_objective:
        c[:] = 0.0
    if config.minimize_relief_only:
        c[:] = 0.0
        for nurse_id in relief_ids:
            for d in days:
                for s in ("D","E","N"):
                    c[x[(nurse_id,d,s)]]=1.0
        # Avoid arbitrary formal-teaching selections in a relief-count proof.
        for v in t.values():
            c[v]+=1e-4
        # Tie-break equal-relief proofs toward a roster that protects the
        # internal senior core without ever making one extra relief shift
        # preferable.
        for v in priority_gap_vars:
            c[v]+=1e-5
    if config.fast_feasible_objective:
        c[:] = 0.0
        for n in nurses:
            if n.nurse_id not in relief_ids:
                # Compact production hierarchy: protect the required monthly
                # attendance first, then minimize legal overtime.  This keeps
                # the objective much easier for HiGHS to prove than the full
                # fairness formulation while preserving the user's work-hour
                # threshold policy.
                c[under[n.nurse_id]]+=100000.0
                c[over[n.nurse_id]]+=10000.0
        for n in nurses:
            for d in days:
                for s in shifts:
                    v=x[(n.nurse_id,d,s)]
                    if s=="H":
                        c[v]+=200.0
                    elif s in {"D","E","N"} and s!=n.home_shift:
                        c[v]+=100.0
                if (n.nurse_id,d) in block_start:
                    c[block_start[(n.nurse_id,d)]]+=1.0
        # Safe teaching is materialized in a dedicated verified pass after the
        # shift roster is fixed; keep solver teaching decisions at zero here.
        for v in t.values():
            c[v]+=100.0
        for v in priority_gap_vars:
            # Full core coverage is already hard-constrained.  This term only
            # prefers an in-house senior over an audited relief seat and must
            # not overwhelm monthly-hours/overtime protection.
            c[v]+=1000.0
    if config.minimize_moved_off_only:
        c[:] = 0.0
        for nurse_id, workday in config.movable_off_assignments:
            if nurse_id in by_id and workday in days:
                c[w[(nurse_id, workday)]] = 1.0
    if config.minimize_overtime_only:
        c[:] = 0.0
        for n in nurses:
            if n.nurse_id not in relief_ids and n.nurse_id not in conditional_plan_c_ids:
                c[over[n.nurse_id]] = 1.0
    if config.minimize_conditional_plan_c_only:
        c[:] = 0.0
        for nurse_id in conditional_plan_c_ids:
            for workday in days:
                c[w[(nurse_id,workday)]] = 1.0
    # Iterative cut generation for exact formal-teaching safety.
    # If a candidate solution schedules teaching on a shift where no on-shift whitelist
    # Preceptor can be forced to Observation while the other six core roles still match,
    # forbid only that selected trainee/date/shift teaching decision and re-solve.
    # Other trainees on the same date/shift may still admit a safe Preceptor/role match.
    cut_rounds=0
    while True:
        bounds,constraints=m.compile()
        integ=np.array(m.integrality)
        round_no=cut_rounds+1
        print(
            f"      MILP round {round_no}: variables={len(m.names)}, "
            f"constraints={len(m.rows)}, time_limit={config.solver_time_limit_seconds:g}s...",
            flush=True,
        )
        round_started=time.monotonic()
        res=milp(c=c,integrality=integ,bounds=bounds,constraints=[constraints],options={"time_limit":config.solver_time_limit_seconds,"presolve":True,"mip_rel_gap":0.0})
        round_elapsed=time.monotonic()-round_started
        print(
            f"      MILP round {round_no} finished in {round_elapsed:.1f}s: "
            f"status={res.status}, message={res.message}",
            flush=True,
        )
        if res.x is None:
            raise RuntimeError(f"MILP不可行/未取得可行解: {res.message}")
        sol=res.x
        unsafe=[]
        selected_teaching={(d,s):nid for (nid,d,s),v in t.items() if sol[v]>.5}
        for (d,s),nid in selected_teaching.items():
            workers=[]
            for nn in nurses:
                if (nn.nurse_id,d,s) not in forced_shadow_keys and sol[core_var(nn,d,s)]>.5:
                    workers.append(nn)
            try:
                assign_core_areas(
                    workers,s,day=d,formal_teaching=True,
                    preceptor_ids=eligible_training_preceptors(d,by_id[nid].home_shift),
                )
            except RuntimeError:
                unsafe.append((nid,d,s))
        if not unsafe:
            break
        print(
            f"      teaching safety: {len(set(unsafe))} unsafe trainee/date/shift decision(s); "
            "adding cuts and re-solving...",
            flush=True,
        )
        cut_rounds+=1
        if cut_rounds>30:
            raise RuntimeError(f"正式帶教安全cut超過30輪：{unsafe}")
        for nid,d,s in sorted(set(unsafe)):
            m.con({t[(nid,d,s)]:1},ub=0)

    opt_over=int(round(sum(
        sol[over[n.nurse_id]]
        for n in nurses
        if n.nurse_id not in conditional_plan_c_ids
    )))
    opt_under=int(round(sum(sol[v] for v in under.values())))
    fair_score=float(sum(sol[v] for v in fairness_vars))
    solver_status="OPTIMAL" if res.status==0 else "FEASIBLE_NOT_PROVEN_OPTIMAL"
    phase4_message=str(res.message)

    # Preserve exactly the solver-selected formal teaching decisions. Materialization
    # must not invent additional teaching sessions after optimization, because doing so
    # would change independence transitions outside the proven MILP solution.
    solver_teaching_by_shift={(d,s):nid for (nid,d,s),v in t.items() if sol[v]>.5}

    selected_shift={}
    for n in nurses:
        for d in days:
            chosen=[s for s in shifts if sol[x[(n.nurse_id,d,s)]]>.5]
            if chosen:selected_shift[(n.nurse_id,d)]=chosen[0]
    selected_cross_decisions={
        (nurse_id,day,shift)
        for (nurse_id,day),shift in selected_shift.items()
        if shift in {"D","E","N"} and shift != by_id[nurse_id].home_shift
    }

    start_nonind={n.nurse_id for n in trainees}
    completed_counts={nid:int(training_start[nid].completed_observation_sessions or 0) for nid in start_nonind}
    for planned_people in config.required_training_assignments_by_day_shift.values():
        for planned_nurse_id in planned_people:
            completed_counts.setdefault(
                planned_nurse_id,
                int(training_start[planned_nurse_id].completed_observation_sessions or 0),
            )
    assignments=[]
    teaching_log=[]

    def identity_independent(n: NurseProfile, d: date) -> bool:
        if n.nurse_id not in start_nonind:
            return True
        return age_independent_on_day(n,d) or completed_counts[n.nurse_id] >= 18

    def match_planned_training_core(
        workers: Sequence[NurseProfile],
        roles: Sequence[str],
        planned_roles: frozenset[str],
        shift: str,
        workday: date,
        preferred_worker_by_role: Mapping[str, str] = {},
    ) -> dict[str, str]:
        """Match the complete core while forcing a Teaching preceptor into each student area."""
        candidates_by_role: dict[str, list[NurseProfile]] = {}
        for role in roles:
            candidates_by_role[role] = [
                worker
                for worker in workers
                if role_eligible_on_day(worker, role, workday)
                and (
                    role not in planned_roles
                    or (
                        worker.home_shift == shift
                        and "Teaching" in worker.areas
                    )
                )
            ]
            if not candidates_by_role[role]:
                raise RuntimeError(
                    f"{workday} {shift} 訓練核心角色 {role} 沒有合格Preceptor"
                )
            preferred_worker_id = preferred_worker_by_role.get(role)
            candidates_by_role[role].sort(
                key=lambda worker: (
                    worker.nurse_id != preferred_worker_id,
                    worker.nurse_id,
                )
            )

        ordered_roles = sorted(
            roles,
            key=lambda role: (
                role not in preferred_worker_by_role,
                len(candidates_by_role[role]),
            ),
        )
        chosen: dict[str, str] = {}
        used: set[str] = set()

        def search(index: int) -> bool:
            if index == len(ordered_roles):
                return True
            role = ordered_roles[index]
            for worker in candidates_by_role[role]:
                if worker.nurse_id in used:
                    continue
                used.add(worker.nurse_id)
                chosen[worker.nurse_id] = role
                if search(index + 1):
                    return True
                chosen.pop(worker.nurse_id, None)
                used.remove(worker.nurse_id)
            return False

        if not search(0):
            raise RuntimeError(
                f"{workday} {shift} 訓練後完整核心角色無法完成一對一匹配"
            )
        return chosen

    for d in days:
        for s in ("D","E","N"):
            scheduled=[n for n in nurses if selected_shift.get((n.nurse_id,d))==s]
            indep=[
                n for n in scheduled
                if identity_independent(n,d)
                and (n.nurse_id,d,s) not in forced_shadow_keys
            ]
            nonind=[n for n in scheduled if not identity_independent(n,d)]
            formal_n=None; forced_preceptor=None
            planned_preceptors: dict[str, str] = {}
            planned_assignments = dict(
                config.required_training_assignments_by_day_shift.get((d, s), {})
            )
            selected_nid=solver_teaching_by_shift.get((d,s)) if s in {"D","E"} else None
            if planned_assignments:
                missing_students = set(planned_assignments) - {
                    nurse.nurse_id for nurse in scheduled
                }
                if missing_students:
                    raise RuntimeError(
                        f"TEACHING_MATERIALIZATION_MISMATCH: {d} {s} missing={sorted(missing_students)}"
                    )
                planned_roles = frozenset(planned_assignments.values())
                core_roles = (
                    D_E_L_PLUS_T_CORE_ROLES
                    if "LeaderTriage" in planned_roles
                    else D_E_CORE_ROLES
                )
                matched = match_planned_training_core(
                    indep,
                    core_roles,
                    planned_roles,
                    s,
                    d,
                    {planned_assignments[nid]: pid for nid, pid in config.preferred_preceptors.items() if nid in planned_assignments},
                )
                areas = {
                    nid: {
                        "Observation1": "流動",
                        "Observation2": "留觀",
                        "LeaderTriage": "Leader+Triage",
                    }.get(role, role)
                    for nid, role in matched.items()
                }
                preceptor_by_role = {role: nid for nid, role in matched.items()}
                planned_preceptors = {
                    nurse_id: preceptor_by_role[role]
                    for nurse_id, role in planned_assignments.items()
                }
            elif selected_nid:
                formal_n=next((nn for nn in nonind if nn.nurse_id==selected_nid),None)
                if formal_n is None:
                    raise RuntimeError(
                        f"TEACHING_MATERIALIZATION_MISMATCH: {d} {s} solver trainee={selected_nid} "
                        "在具體化時不是非獨立同班人員"
                    )
                areas,forced_preceptor=assign_core_areas(
                    indep,s,day=d,formal_teaching=True,
                    preceptor_ids=eligible_training_preceptors(d,formal_n.home_shift),
                )
            else:
                if (
                    s in {"D","E"}
                    and "LeaderTriage" in config.required_training_roles_by_day_shift.get((d,s),())
                ):
                    matched=_match_roles(indep,D_E_L_PLUS_T_CORE_ROLES,d)
                    if matched is None:
                        raise RuntimeError(f"{d} {s} L+T訓練備援核心角色無法完成一對一匹配")
                    areas={
                        nid:{
                            "Observation1":"流動",
                            "Observation2":"留觀",
                            "LeaderTriage":"Leader+Triage",
                        }.get(role,role)
                        for nid,role in matched.items()
                    }
                else:
                    areas,_=assign_core_areas(indep,s,day=d,formal_teaching=False,preceptor_ids=config.preceptor_ids)

            # Clinic3 is a support position, never a substitute for a missing
            # core area.  It may be assigned only after every applicable core
            # role has been materialized: 7 roles on D/E, 5 roles on a
            # five-person N shift, or 6 roles (separate Leader and Triage) once
            # N reaches six people.
            clinic3_support_slot = len(scheduled) > len(areas) + len(planned_preceptors)
            for n in scheduled:
                planned_preceptor = planned_preceptors.get(n.nurse_id)
                formal=bool(planned_preceptor or (formal_n and n.nurse_id==formal_n.nurse_id))
                if planned_preceptor:
                    area=areas[planned_preceptor]; indep_core=False; preceptor_id=planned_preceptor
                elif formal:
                    # A formal trainee shadows the selected Preceptor in the
                    # exact same Observation zone.  The trainee is displayed
                    # as 流動/留觀 just like the teacher, but never contributes
                    # an additional independent core count.
                    area=areas[forced_preceptor]; indep_core=False; preceptor_id=forced_preceptor
                elif n.nurse_id in areas:
                    area=areas[n.nurse_id]; indep_core=True; preceptor_id=None
                elif not identity_independent(n,d):
                    area="Clinic3" if (clinic3_support_slot and "Clinic3" in n.areas and not n.part_time) else "技術組"
                    indep_core=False; preceptor_id=None
                else:
                    area="Clinic3" if (clinic3_support_slot and "Clinic3" in n.areas and not n.part_time) else "Reserve"
                    # Clinic3/Reserve are support manpower and therefore never
                    # inflate the exported or validated core count.
                    indep_core=False; preceptor_id=None
                is_locked = must.get((n.nurse_id,d)) == s
                assignments.append(Assignment(n.nurse_id,d,s,area,"GOOGLE_MUST" if is_locked else "SOLVER",is_locked,formal,preceptor_id,indep_core))
            for trainee_id, preceptor_id in planned_preceptors.items():
                completed_counts[trainee_id]+=1
                teaching_log.append({"date":d.isoformat(),"shift":s,"trainee_id":trainee_id,"preceptor_id":preceptor_id,"completed_after":completed_counts[trainee_id]})
            if formal_n is not None:
                completed_counts[formal_n.nurse_id]+=1
                teaching_log.append({"date":d.isoformat(),"shift":s,"trainee_id":formal_n.nurse_id,"preceptor_id":forced_preceptor,"completed_after":completed_counts[formal_n.nurse_id]})

        for n in nurses:
            if selected_shift.get((n.nurse_id,d))=="H":
                assignments.append(Assignment(n.nurse_id,d,"H","H支援","SOLVER",False,False,None,False))

    training_end={}
    month_end=days[-1]
    for n in nurses:
        st=training_start[n.nurse_id]
        if n.nurse_id in start_nonind:
            total=completed_counts[n.nurse_id]
            status=IndependenceStatus.INDEPENDENT if age_independent_on_day(n,month_end) or total>=18 else IndependenceStatus.NON_INDEPENDENT
            training_end[n.nurse_id]=TrainingState(n.nurse_id,status,total,"materialized formal teaching; independence=age>=2.5m OR formal_teaching>=18")
        else:
            training_end[n.nurse_id]=st

    assignments_t=tuple(sorted(assignments,key=lambda a:(a.day,a.shift,a.nurse_id)))
    materialized_cross_assignments={
        (a.nurse_id,a.day,a.shift)
        for a in assignments_t
        if a.shift in {"D","E","N"} and a.shift != by_id[a.nurse_id].home_shift
    }
    if selected_cross_decisions != materialized_cross_assignments:
        missing=sorted(selected_cross_decisions-materialized_cross_assignments)
        extra=sorted(materialized_cross_assignments-selected_cross_decisions)
        raise RuntimeError(
            "CROSS_SHIFT_MATERIALIZATION_MISMATCH: "
            f"solver_only={missing}; export_only={extra}"
        )
    per_person_cross=Counter(nid for nid,_,_ in materialized_cross_assignments)
    discretionary_cross_assignments={
        key for key in materialized_cross_assignments
        if key not in config.cross_shift_cap_exempt_assignments
    }
    per_person_discretionary_cross=Counter(
        nid for nid,_,_ in discretionary_cross_assignments
    )
    moved_off_assignments = [
        {
            "nurse_id": nurse_id,
            "name": by_id[nurse_id].name,
            "date": workday.isoformat(),
            "shift": selected_shift.get((nurse_id, workday)),
        }
        for nurse_id, workday in sorted(config.movable_off_assignments)
        if selected_shift.get((nurse_id, workday)) in {"D", "E", "N", "H"}
    ]
    internal_high_core_gap_details = [
        {
            "date": workday.isoformat(),
            "shift": shift,
            "role": {
                "LeaderTriage": "Leader+Triage",
            }.get(role, role),
        }
        for (workday, shift, role), variable in sorted(priority_role_gap.items())
        if sol[variable] > .5
    ]
    metrics={
        "overtime_shifts": opt_over,
        "under_target_shifts": opt_under,
        "group_fairness_score": fair_score,
        "cross_shift_decision_days": len(selected_cross_decisions),
        "discretionary_cross_shift_days": len(discretionary_cross_assignments),
        "max_cross_shift_days_per_person_cap": config.max_cross_shift_days_per_person,
        "max_cross_shift_days_per_person_actual": max(per_person_cross.values(),default=0),
        "max_discretionary_cross_shift_days_per_person_actual": max(
            per_person_discretionary_cross.values(), default=0
        ),
        "max_overtime_shifts_per_person_cap": config.max_overtime_shifts_per_person,
        "total_overtime_shift_cap": config.total_overtime_shift_cap,
        "conditional_plan_c_ids": sorted(conditional_plan_c_ids),
        "conditional_plan_c_shift_cap": config.conditional_plan_c_shift_cap,
        "conditional_plan_c_shifts": sum(
            1 for assignment in assignments_t
            if assignment.nurse_id in conditional_plan_c_ids
        ),
        "cross_shift_days_by_person": dict(sorted(per_person_cross.items())),
        "discretionary_cross_shift_days_by_person": dict(
            sorted(per_person_discretionary_cross.items())
        ),
        "moved_off_assignment_count": len(moved_off_assignments),
        "moved_off_assignments": moved_off_assignments,
        "internal_high_core_gap_count": len(internal_high_core_gap_details),
        "internal_high_core_gap_details": internal_high_core_gap_details,
        "optimization_proven": solver_status == "OPTIMAL",
        "weighted_objective": float(res.fun) if res.fun is not None else None,
        "solver_message": phase4_message,
        "variables": len(m.names),
        "constraints": len(m.rows),
        "teaching_cut_rounds": cut_rounds,
        "target_equality_enforced": bool(config.require_exact_targets),
        "effective_target_shifts": dict(effective_targets),
        "capacity_adjustments": capacity_adjustments,
        "constrained_under_shifts": sum(int(x["constrained_under_shifts"]) for x in capacity_adjustments),
        "formal_teaching_materialized": len(teaching_log),
        "formal_teaching_log": teaching_log,
    }
    return ScheduleResult(
        config.year, config.month, nurses, assignments_t, dict(targets), training_end,
        google_audit, previous_audit, solver_status, metrics,
        _result_hash(assignments_t, google_audit, previous_audit, targets, training_end),
    )


# ==============================================================================
# MODULE: validator.py
# ==============================================================================

from collections import Counter, defaultdict
from datetime import date, timedelta
from typing import Mapping



def validate_result(
    result: ScheduleResult,
    requests: tuple[Request, ...],
    boundary: tuple[BoundaryEntry, ...],
    config: SchedulerConfig,
) -> ValidationResult:
    hard: list[Issue]=[]; warn: list[Issue]=[]
    days=month_days(result.year,result.month)
    amap=result.assignment_map()
    by_id={n.nurse_id:n for n in result.nurses}
    bmap={(e.nurse_id,e.day):e for e in boundary}
    relief_ids=set(config.relief_ids)

    # Both authoritative sources must be LIVE Google and strictly verified.
    if result.google_audit.result != "PASS" or not result.google_audit.source_verified or result.google_audit.source_mode != "LIVE_GOOGLE":
        hard.append(Issue("GOOGLE_SOURCE",Severity.HARD,
            f"正式發布要求預假來源 LIVE_GOOGLE/PASS；actual={result.google_audit.source_mode}/{result.google_audit.result}"))
    if (result.previous_audit.result != "PASS" or not result.previous_audit.source_verified
            or result.previous_audit.source_mode != "LIVE_GOOGLE_PREVIOUS_ROSTER"):
        hard.append(Issue("PREVIOUS_GOOGLE_SOURCE",Severity.HARD,
            f"正式發布要求上月班表 LIVE_GOOGLE_PREVIOUS_ROSTER/PASS；"
            f"actual={result.previous_audit.source_mode}/{result.previous_audit.result}"))

    # Hard requests.
    for r in requests:
        a=amap.get((r.nurse_id,r.day))
        if r.kind in {RequestKind.OFF_LOCK,RequestKind.ANNUAL_LEAVE} and a is not None:
            hard.append(Issue("HARD_OFF_VIOLATION",Severity.HARD,f"{r.raw}被排成{a.shift}",r.nurse_id,r.day,a.shift))
        if r.kind==RequestKind.MUST_SHIFT and (a is None or a.shift!=r.shift):
            hard.append(Issue("MUST_SHIFT_VIOLATION",Severity.HARD,f"要求{r.shift}，實際{a.shift if a else 'OFF'}",r.nurse_id,r.day,r.shift))

    # Daily core / exact area role coverage from immutable export state.  N is
    # dynamic: five people use L+T; at six or more people Leader and Triage must
    # occupy separate core positions.
    required_de=["Leader","Triage","Critical","Clinic1","Clinic2","流動","留觀"]
    required_de_l_plus_t=["Leader+Triage","Critical","Clinic1","Clinic2","流動","留觀"]
    required_n5=["Leader+Triage","Critical","Clinic1","流動","留觀"]
    required_n6=["Leader","Triage","Critical","Clinic1","流動","留觀"]
    configured_floor={"D":config.d_min_core,"E":config.e_min_core,"N":config.n_min_core}
    for d in days:
        for s in ("D","E","N"):
            workers=[a for a in result.assignments if a.day==d and a.shift==s]
            l_plus_t_training=(
                s in {"D","E"}
                and "LeaderTriage" in config.required_training_roles_by_day_shift.get((d,s),())
            )
            roles=(required_de_l_plus_t if l_plus_t_training else required_de) if s in {"D","E"} else (required_n6 if len(workers)>=6 else required_n5)
            core=[a for a in result.assignments if a.day==d and a.shift==s and a.independent_core]
            areas=Counter(a.area for a in core)
            need=Counter(roles)
            missing={k:v-areas.get(k,0) for k,v in need.items() if areas.get(k,0)<v}
            if missing:
                hard.append(Issue("CORE_ROLE_SHORTAGE",Severity.HARD,f"{s}核心區域不足 {missing}",day=d,shift=s))
            effective_floor=(len(required_de_l_plus_t) if l_plus_t_training else configured_floor[s])
            if len(core) < effective_floor:
                hard.append(Issue("CORE_MINIMUM_SHORTAGE",Severity.HARD,
                    f"{s}獨立核心人數{len(core)}<{effective_floor}",day=d,shift=s))
            if s=="N" and len(workers)>=6 and len(core)<6:
                hard.append(Issue("N_SIXTH_TRIAGE_SHORTAGE",Severity.HARD,
                    f"N班總人數{len(workers)}時須有6名獨立核心人員並分列Leader/Triage；實際核心={len(core)}",day=d,shift=s))

            clinic3=[a for a in workers if a.area=="Clinic3"]
            if clinic3 and len(workers)<=len(roles):
                hard.append(Issue("CLINIC3_PREMATURE_SUPPORT",Severity.HARD,
                    f"{s}班核心人力尚無餘裕（總人數={len(workers)}、核心需求={len(roles)}）卻安排Clinic3",day=d,shift=s))
            if clinic3 and missing:
                hard.append(Issue("CLINIC3_WITH_CORE_SHORTAGE",Severity.HARD,
                    f"{s}班仍缺核心區域{missing}卻安排Clinic3支援",day=d,shift=s))
        # Clinic3 is support-only under the current policy.  It is validated
        # only when a future non-zero configurable minimum is explicitly set.
        if config.e_clinic3_flexible_min > 0:
            e_workers=[a for a in result.assignments if a.day==d and a.shift=="E"]
            c3_flex=sum(1 for a in e_workers if a.area=="Clinic3")
            if c3_flex < config.e_clinic3_flexible_min:
                hard.append(Issue("E_CLINIC3_FLEX_SHORTAGE",Severity.HARD,
                    f"E Clinic3彈性可支援{c3_flex}<{config.e_clinic3_flexible_min}",day=d,shift="E"))
        for ss,cap in (("D",config.d_max_total),("E",config.e_max_total),("N",config.n_max_total)):
            total=sum(1 for a in result.assignments if a.day==d and a.shift==ss)
            if total>cap:
                hard.append(Issue(f"{ss}_CAP",Severity.HARD,f"{ss}班人數{total}>{cap}",day=d,shift=ss))

    if config.max_cross_shift_days_per_person is not None:
        for n in result.nurses:
            if n.nurse_id in relief_ids:
                continue
            count=sum(
                1
                for assignment in result.assignments
                if assignment.nurse_id==n.nurse_id
                and assignment.shift in {"D","E","N"}
                and assignment.shift!=n.home_shift
                and (assignment.nurse_id,assignment.day,assignment.shift)
                    not in config.cross_shift_cap_exempt_assignments
            )
            if count>config.max_cross_shift_days_per_person:
                hard.append(Issue(
                    "PERSON_CROSS_SHIFT_CAP",Severity.HARD,
                    f"非硬性指定跨班支援{count}天>"
                    f"{config.max_cross_shift_days_per_person}天上限",
                    n.nurse_id,
                ))

    if config.max_overtime_shifts_per_person is not None:
        worked = Counter(
            assignment.nurse_id
            for assignment in result.assignments
            if assignment.shift in {"D", "E", "N", "H"}
        )
        for nurse in result.nurses:
            if nurse.nurse_id in relief_ids:
                continue
            overtime = max(0, worked[nurse.nurse_id] - result.target_shifts[nurse.nurse_id])
            if overtime > config.max_overtime_shifts_per_person:
                hard.append(Issue(
                    "PERSON_OVERTIME_CAP", Severity.HARD,
                    f"加班{overtime}班>{config.max_overtime_shifts_per_person}班上限",
                    nurse.nurse_id,
                ))

    # Cross-month max consecutive / rest / N recovery on the FINAL assignment state.
    for n in result.nurses:
        if n.nurse_id in relief_ids:
            continue
        timeline=[days[0]-timedelta(days=i) for i in range(config.require_previous_boundary_days,0,-1)]+list(days)
        shifts={}
        attendance={}
        for d in timeline:
            if d in days:
                a=amap.get((n.nurse_id,d)); sh=a.shift if a else "OFF"; att=a is not None
            else:
                e=bmap.get((n.nurse_id,d)); sh=e.shift if e else ""; att=bool(e and e.attendance)
            shifts[d]=sh; attendance[d]=att
        run=[]
        for d in timeline:
            if attendance[d]:
                run.append(d)
                if d in days and len(run)>config.max_consecutive_attendance:
                    hard.append(Issue("MAX_CONSECUTIVE",Severity.HARD,f"連續出勤{len(run)}日：{run[0]}~{d}",n.nurse_id,d,shifts[d]))
            else:
                run=[]
        for a,b in zip(timeline,timeline[1:]):
            if not(attendance[a] and attendance[b]): continue
            rh=rest_hours(a,shifts[a],b,shifts[b])
            if rh is not None and rh<config.min_rest_hours and b in days:
                hard.append(Issue("REST_11H",Severity.HARD,f"{shifts[a]}→{shifts[b]}休息{rh:g}h",n.nurse_id,b,shifts[b]))
            if (config.debug_enforce_cross_shift_20h and rh is not None and b in days
                    and shifts[a] in {"D","E","N"} and shifts[b] in {"D","E","N"}
                    and shifts[a] != shifts[b] and rh <= 20.0):
                hard.append(Issue("CROSS_SHIFT_20H",Severity.HARD,
                    f"跨班{shifts[a]}→{shifts[b]}連續休息{rh:g}h，必須>20h",n.nurse_id,b,shifts[b]))
        # Every contiguous attendance block may contain N OR D/E, never both.
        block=[]
        for d in timeline+[days[-1]+timedelta(days=1)]:
            if d in attendance and attendance[d]:
                block.append(d); continue
            if block:
                fam={"N" if shifts[x]=="N" else "DE" if shifts[x] in {"D","E"} else "OTHER" for x in block}
                if "N" in fam and "DE" in fam and any(x in days for x in block):
                    hard.append(Issue("N_RECOVERY_OFF",Severity.HARD,f"同一出勤區塊混有N與D/E：{block[0]}~{block[-1]}",n.nurse_id,block[-1]))
                block=[]

    # Regular/full-time roster-shape validation across the previous boundary. Part-time is exempt from singleton/max-off checks.
    annual_days={(r.nurse_id,r.day) for r in requests if r.kind==RequestKind.ANNUAL_LEAVE}
    for n in result.nurses:
        if n.nurse_id in relief_ids:
            continue
        timeline=[days[0]-timedelta(days=i) for i in range(config.require_previous_boundary_days,0,-1)]+list(days)
        work={}
        for d in timeline:
            if d in days:
                work[d]=(n.nurse_id,d) in amap
            else:
                e=bmap.get((n.nurse_id,d)); work[d]=bool(e and e.attendance)

        if config.debug_enforce_singleton and not n.part_time:
            for i in range(1,len(timeline)-1):
                d=timeline[i]
                if d not in days:
                    continue
                if work[d] and not work[timeline[i-1]] and not work[timeline[i+1]]:
                    hard.append(Issue("SINGLETON_WORK",Severity.HARD,
                        "OFF→上班→OFF 單日工作區塊禁止；出勤/支援至少連續2日",n.nurse_id,d,
                        amap[(n.nurse_id,d)].shift if (n.nurse_id,d) in amap else None))
            last=days[-1]
            prev=last-timedelta(days=1)
            if work[last] and not work[prev]:
                hard.append(Issue("SINGLETON_WORK_MONTH_END",Severity.HARD,
                    "月底最後一日上班但前一日OFF；未知的下月班表不能用來證明非單日工作區塊",
                    n.nurse_id,last,amap[(n.nurse_id,last)].shift))

        if config.debug_enforce_max_consecutive_off and not n.part_time:
            active_timeline=list(days)
            for end_i in range(5,len(active_timeline)):
                window=active_timeline[end_i-5:end_i+1]
                if not any(work[d] or (n.nurse_id,d) in annual_days for d in window):
                    hard.append(Issue("MAX_CONSECUTIVE_OFF",Severity.HARD,
                        f"連續6日無出勤且無明確特休：{window[0]}~{window[-1]}",n.nurse_id,window[-1]))

    # Area qualification is sourced ONLY from Nurses.xlsx -> NurseProfile.areas.
    # Hire date / service age is never allowed to grant or remove an area qualification.
    role_for_area={"Leader":"Leader","Triage":"Triage","Leader+Triage":"LeaderTriage","Critical":"Critical",
                   "Clinic1":"Clinic1","Clinic2":"Clinic2","Clinic3":"Clinic3","Observation":"Observation",
                   "流動":"Observation","留觀":"Observation"}
    for a in result.assignments:
        role=role_for_area.get(a.area or "")
        if role and a.independent_core:
            nurse=by_id[a.nurse_id]
            if not role_eligible_on_day(nurse,role,a.day):
                hard.append(Issue("AREA_NOT_QUALIFIED",Severity.HARD,
                    f"區域{a.area}不在 Nurses.xlsx 可工作區域 {sorted(nurse.areas)}",a.nurse_id,a.day,a.shift))

    # Area-priority gate: Leader -> Triage -> Critical -> all other areas.
    # A relief seat may occupy a high-priority area only when no movable,
    # qualified in-house independent core worker is assigned to a lower tier.
    preceptor_keys={
        (a.preceptor_id,a.day,a.shift)
        for a in result.assignments
        if a.formal_teaching and a.preceptor_id
    }
    area_priority={
        "Leader":0,"Leader+Triage":0,"Triage":1,"Critical":2,
        "Clinic1":3,"Clinic2":3,"Clinic3":3,"流動":3,"留觀":3,"Reserve":3,
    }
    for relief_assignment in result.assignments:
        if relief_assignment.nurse_id not in relief_ids:
            continue
        relief_role=role_for_area.get(relief_assignment.area or "")
        relief_priority=area_priority.get(relief_assignment.area or "",3)
        if relief_role is None or relief_priority >= 3:
            continue
        avoidable=[]
        for candidate in result.assignments:
            if (
                candidate.day != relief_assignment.day
                or candidate.shift != relief_assignment.shift
                or candidate.nurse_id in relief_ids
                or not candidate.independent_core
                or (candidate.nurse_id,candidate.day,candidate.shift) in preceptor_keys
            ):
                continue
            if (
                area_priority.get(candidate.area or "",3) > relief_priority
                and role_eligible_on_day(by_id[candidate.nurse_id],relief_role,candidate.day)
            ):
                avoidable.append(f"{candidate.nurse_id}/{candidate.area}")
        if avoidable:
            hard.append(Issue(
                "AVOIDABLE_PRIORITY_RELIEF",
                Severity.HARD,
                f"待補人力不得優先占用{relief_assignment.area}；院內可互換人員={avoidable}",
                relief_assignment.nurse_id,
                relief_assignment.day,
                relief_assignment.shift,
            ))

    # Training / preceptor safety.
    formal_by_shift=Counter(
        (a.day,a.shift) for a in result.assignments if a.formal_teaching
    )
    for (training_day,training_shift),count in formal_by_shift.items():
        if count>2:
            hard.append(Issue(
                "TRAINING_SHIFT_CAP",
                Severity.HARD,
                f"同一班正式訓練{count}人>2人上限",
                day=training_day,
                shift=training_shift,
            ))
    for a in result.assignments:
        if a.formal_teaching:
            trainee=by_id[a.nurse_id]
            training_role=(
                "LeaderTriage" if a.area=="Leader+Triage"
                else ("Triage" if a.area=="Triage" else "Observation")
            )
            allowed_teacher_areas=(
                {"Leader+Triage"} if training_role=="LeaderTriage"
                else ({"Triage"} if training_role=="Triage" else {"Observation","流動","留觀"})
            )
            if not a.preceptor_id:
                hard.append(Issue("TEACHING_NO_PRECEPTOR",Severity.HARD,"正式帶教缺少Preceptor",a.nurse_id,a.day,a.shift))
            pa=amap.get((a.preceptor_id,a.day)) if a.preceptor_id else None
            if pa is None or pa.shift!=a.shift or pa.area not in allowed_teacher_areas:
                hard.append(Issue("TEACHING_PRECEPTOR_MISMATCH",Severity.HARD,
                    f"Preceptor未與學生同班同{training_role}區域",a.nurse_id,a.day,a.shift))
            elif a.area != pa.area:
                hard.append(Issue(
                    "TEACHING_AREA_MISMATCH",Severity.HARD,
                    f"帶教學生區域{a.area}與老師區域{pa.area}不同；正式帶教必須同區",
                    a.nurse_id,a.day,a.shift,
                ))
            preceptor=by_id.get(a.preceptor_id or "")
            if a.shift!=trainee.home_shift:
                hard.append(Issue("TRAINEE_HOME_SHIFT_MISMATCH",Severity.HARD,
                    f"正式帶教班別{a.shift}不等於新人原班別{trainee.home_shift}",a.nurse_id,a.day,a.shift))
            if (preceptor is None or preceptor.home_shift!=trainee.home_shift
                    or "Teaching" not in preceptor.areas
                    or not role_eligible_on_day(preceptor,training_role,a.day)):
                hard.append(Issue("PRECEPTOR_TEACHING_QUALIFICATION_MISMATCH",Severity.HARD,
                    f"Preceptor必須與學生同原班別，且areas含Teaching及{training_role}資格",
                    a.nurse_id,a.day,a.shift))
        allowed_nonindependent_areas={"Observation","Clinic3","技術組","H支援","Reserve"}
        if a.formal_teaching:
            allowed_nonindependent_areas.update({"流動","留觀","Triage","Leader+Triage"})
        if not a.independent_core and a.area not in allowed_nonindependent_areas:
            hard.append(Issue("TRAINEE_UNSAFE_AREA",Severity.HARD,f"非獨立人員區域={a.area}",a.nurse_id,a.day,a.shift))

    # Static assignment legality (including part-time shift rules; date OFF only from explicit requests).
    for a in result.assignments:
        n=by_id[a.nurse_id]
        allowed,why=shift_allowed_static(n,a.day,a.shift,allow_h=config.allow_h)
        date_allowed,dwhy=part_time_workday_allowed(n,a.day,official_holidays=config.official_holidays)
        if not allowed or not date_allowed:
            hard.append(Issue("STATIC_ASSIGNMENT_ILLEGAL",Severity.HARD,
                f"{a.shift} 不合法：{why if not allowed else dwhy}",a.nurse_id,a.day,a.shift))

    for n in result.nurses:
        worked_hours=actual_hours(result.assignments,n.nurse_id)
        nights=shift_counts(result.assignments,n.nurse_id).get("N",0)
        if n.max_night_shifts is not None and nights > n.max_night_shifts:
            hard.append(Issue("MAX_NIGHT_EXCEEDED",Severity.HARD,
                f"實際夜班{nights}>{n.max_night_shifts}",n.nurse_id))

    # Canonical fairness/overtime definition. FINAL-5.4 does not permit capacity-based
    # target reduction. The canonical target already includes only legitimate paid credits
    # (annual leave / verified prior-unused hours). Requested OFF does not lower it.
    overtime={n.nurse_id:overtime_hours(result.assignments,result.target_shifts,n.nurse_id) for n in result.nurses}
    under={n.nurse_id:under_target_hours(result.assignments,result.target_shifts,n.nurse_id) for n in result.nurses}
    total_ot=sum(overtime.values())
    total_under=sum(under.values())
    optimal=bool(result.solver_metrics.get("optimization_proven")) and result.solver_status=="OPTIMAL"
    effective_raw=result.solver_metrics.get("effective_target_shifts") or dict(result.target_shifts)
    effective={str(k):int(v) for k,v in dict(effective_raw).items()}
    constrained_by_id={}
    for row in result.solver_metrics.get("capacity_adjustments",[]) or []:
        constrained_by_id[str(row.get("nurse_id"))]=int(row.get("constrained_under_shifts",0) or 0)
    constrained_under_hours=sum(v*8 for v in constrained_by_id.values())
    regular_under_hours=sum(under[n.nurse_id] for n in result.nurses if not n.part_time)
    part_time_under_hours=sum(under[n.nurse_id] for n in result.nurses if n.part_time)
    regular_constrained_hours=sum(
        constrained_by_id.get(n.nurse_id,0)*8 for n in result.nurses if not n.part_time
    )
    unexpected_under=max(0,regular_under_hours-regular_constrained_hours)
    actual_days={n.nurse_id:actual_work_days(result.assignments,n.nurse_id) for n in result.nurses}
    exact_target_proven=(
        bool(result.solver_metrics.get("target_equality_enforced"))
        and all(actual_days[n.nurse_id]==effective.get(n.nurse_id,result.target_shifts[n.nurse_id]) for n in result.nurses)
    )

    # Overtime is legal: target hours are the regular-hours threshold. The
    # lexicographic objective minimizes it and exported metrics report it.
    if unexpected_under>0:
        hard.append(Issue("UNEXPLAINED_UNDER_TARGET",Severity.HARD,
            f"一般人員不足{regular_under_hours}h，其中{unexpected_under}h沒有硬規則容量證明"))
    for n in result.nurses:
        if not n.part_time:
            continue
        actual=actual_days[n.nurse_id]
        target=int(result.target_shifts[n.nurse_id])
        if actual < target:
            code=("PART_TIME_BELOW_MINIMUM" if actual < config.part_time_min_shifts
                  else "PART_TIME_BELOW_AVAILABILITY_TARGET")
            issue=Issue(
                code,
                Severity.HARD if config.protect_part_time_targets else Severity.WARNING,
                f"部分工時人員本月實際{actual}班，低於依預假後可上班日計算的目標{target}班"
                f"（最低門檻{config.part_time_min_shifts}班）；"
                + (
                    "已啟用部分工時硬下限，禁止輸出"
                    if config.protect_part_time_targets
                    else "班表仍可輸出，但必須告知使用者"
                ),
                n.nurse_id,
            )
            (hard if config.protect_part_time_targets else warn).append(issue)
    target_minimum_met=all(
        actual_days[n.nurse_id] >= effective.get(n.nurse_id,result.target_shifts[n.nurse_id])
        for n in result.nurses if not n.part_time
    )
    if not target_minimum_met:
        hard.append(Issue("FAIRNESS_NOT_PROVEN",Severity.HARD,
            "至少一人未達每月應上班時數／加班起算線；禁止公平Gate PASS"))
    if not optimal and (total_ot>0 or unexpected_under>0 or not target_minimum_met):
        hard.append(Issue("OPTIMALITY_NOT_PROVEN",Severity.HARD,
            "存在加班或工時不足，但solver未證明這已是全域最優結果"))

    # Verify every nominal->effective reduction is exactly covered by canonical under-hours.
    for n in result.nurses:
        nominal=int(result.target_shifts[n.nurse_id])
        eff=int(effective.get(n.nurse_id,nominal))
        if eff>nominal:
            hard.append(Issue("EFFECTIVE_TARGET_INVALID",Severity.HARD,
                f"有效目標{eff}>名義目標{nominal}",n.nurse_id))
        expected_under=max(0,nominal-eff)*8
        if not n.part_time and under[n.nurse_id] != expected_under:
            hard.append(Issue("CAPACITY_PROOF_MISMATCH",Severity.HARD,
                f"名義/有效目標預期不足{expected_under}h，實際不足{under[n.nurse_id]}h",n.nurse_id))

    fairness_gate=("FAIL" if any(i.code in {"UNEXPLAINED_UNDER_TARGET","FAIRNESS_NOT_PROVEN","CAPACITY_PROOF_MISMATCH"} for i in hard)
                   else "PASS")
    status=PublishStatus.BLOCKED if hard else PublishStatus.WARNING if warn else PublishStatus.PASS
    metrics={
        "hard_errors":len(hard),"warnings":len(warn),
        "total_overtime_hours":total_ot,"total_under_hours":total_under,
        "regular_under_hours":regular_under_hours,"part_time_under_hours":part_time_under_hours,
        "constrained_under_hours":constrained_under_hours,"unexpected_under_hours":unexpected_under,
        "fairness_gate":fairness_gate,
        "optimization_proven":optimal,
        "target_equality_proven": exact_target_proven,
        "target_minimum_met": target_minimum_met,
        "solver_overtime_shifts":result.solver_metrics.get("overtime_shifts"),
        "solver_under_target_shifts":result.solver_metrics.get("under_target_shifts"),
        "solver_group_fairness_score":result.solver_metrics.get("group_fairness_score"),
        "max_cross_shift_days_per_person_cap":config.max_cross_shift_days_per_person,
        "max_cross_shift_days_per_person_actual":result.solver_metrics.get("max_cross_shift_days_per_person_actual"),
        "capacity_adjustments": result.solver_metrics.get("capacity_adjustments",[]),
    }
    return ValidationResult(tuple(hard),tuple(warn),metrics,status)


# ==============================================================================
# MODULE: exporter.py
# ==============================================================================

import hashlib
import json
from collections import Counter
from pathlib import Path



def _schedule_digest_from_rows(rows: list[tuple[str,str,str,str]]) -> str:
    payload=json.dumps(sorted(rows),ensure_ascii=False,separators=(",",":"))
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def export_xlsx_immutable(
    path: Path,
    result: ScheduleResult,
    validation: ValidationResult,
    requests: tuple[Request,...],
    google_matrix: list[list[str]],
    *, allow_warning: bool=True,
) -> Path:
    """Export is read-only with respect to ScheduleResult. No repair/recalculation is allowed here."""
    if validation.publish_status == PublishStatus.BLOCKED:
        raise RuntimeError("BLOCKED：存在Hard Error，禁止輸出正式班表")
    if validation.publish_status == PublishStatus.WARNING and not allow_warning:
        raise RuntimeError("WARNING：目前設定禁止輸出可發布班表")
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb=Workbook(); ws=wb.active; ws.title="總班表"
    days=month_days(result.year,result.month)
    headers=[
        "ID","姓名","班組","目標班數","實際工時","加班時數","不足時數",
        "D班數","E班數","N班數",
    ]+[str(d.day) for d in days]
    ws.append(headers)
    amap=result.assignment_map()
    by_id={n.nurse_id:n for n in result.nurses}
    raw_overrides=result.solver_metrics.get("scenario_overrides") or []
    if not isinstance(raw_overrides,(list,tuple)):
        raise TypeError("scenario_overrides 必須是 list/tuple of mappings")
    scenario_name=str(result.solver_metrics.get("scenario_name") or "排班假設情境")
    scenario_rows: list[list[str]]=[]
    for item in raw_overrides:
        if not isinstance(item,Mapping):
            raise TypeError("scenario_overrides 每一筆必須是 mapping")
        raw_day=item.get("date","")
        day_text=raw_day.isoformat() if hasattr(raw_day,"isoformat") else str(raw_day)
        nurse_id=str(item.get("nurse_id","")).strip()
        action=str(item.get("scenario_action","")).strip()
        if not nurse_id or not day_text.strip() or not action:
            raise ValueError("scenario_overrides 缺少 nurse_id/date/scenario_action")
        scenario_rows.append([
            scenario_name,
            nurse_id,
            str(item.get("name") or (by_id[nurse_id].name if nurse_id in by_id else "")),
            day_text,
            str(item.get("source_kind", "")),
            str(item.get("source_raw", "")),
            action,
            str(item.get("target_effect", "")),
            str(item.get("source_note", "原始資料來自 LIVE Google 預假表；LIVE 未修改，僅本情境求解時解除硬鎖定")),
        ])
    scenario_rows.sort(key=lambda row:(row[3],row[1],row[4],row[5]))
    scenario_digest=hashlib.sha256(
        json.dumps(scenario_rows,ensure_ascii=False,separators=(",",":")).encode("utf-8")
    ).hexdigest()
    raw_monthly_excluded=result.solver_metrics.get("monthly_excluded_people") or []
    if not isinstance(raw_monthly_excluded,(list,tuple)):
        raise TypeError("monthly_excluded_people 必須是 list/tuple of mappings")
    monthly_excluded_rows: list[list[str]]=[]
    for item in raw_monthly_excluded:
        if not isinstance(item,Mapping):
            raise TypeError("monthly_excluded_people 每一筆必須是 mapping")
        nurse_id=str(item.get("nurse_id","")).strip()
        name=str(item.get("name","")).strip()
        year_text=str(item.get("year",result.year))
        month_text=str(item.get("month",result.month)).zfill(2)
        if not nurse_id or not name:
            raise ValueError("monthly_excluded_people 缺少 nurse_id/name")
        monthly_excluded_rows.append([
            f"{year_text}/{month_text}",nurse_id,name,"本月不排班且不輸出人員列",
            "使用者指定的月份限定排除；Nurses.xlsx 與 LIVE 原始資料均未刪除",
        ])
    monthly_excluded_rows.sort(key=lambda row:row[1])
    monthly_excluded_digest=hashlib.sha256(
        json.dumps(monthly_excluded_rows,ensure_ascii=False,separators=(",",":")).encode("utf-8")
    ).hexdigest()
    relief_ids={str(value) for value in (result.solver_metrics.get("relief_ids") or [])}
    unknown_relief=relief_ids-set(by_id)
    if unknown_relief:
        raise ValueError(f"relief_ids 含未知人員：{sorted(unknown_relief)}")
    relief_rows=[]
    for a in sorted(
        (item for item in result.assignments if item.nurse_id in relief_ids),
        key=lambda item:(item.day,item.shift,item.nurse_id),
    ):
        relief_rows.append([
            a.day.isoformat(),a.shift,a.area or "未配置",a.nurse_id,
            by_id[a.nurse_id].name,"外部／機動待補","待確認實際支援者",
        ])
    relief_digest=hashlib.sha256(
        json.dumps(relief_rows,ensure_ascii=False,separators=(",",":")).encode("utf-8")
    ).hexdigest()
    audit_rows=[]
    def seniority_key(n: NurseProfile) -> tuple[date,tuple[int,int|str]]:
        stable_id=(0,int(n.nurse_id)) if n.nurse_id.isdigit() else (1,n.nurse_id)
        return (date.max if n.nurse_id in relief_ids else n.hire_date),stable_id

    # Older hire date means greater seniority, so ascending hire date is
    # descending service seniority. IDs provide deterministic tie-breaking.
    seniority_order=sorted(result.nurses,key=seniority_key)
    for n in seniority_order:
        counts=shift_counts(result.assignments,n.nurse_id)
        row=[n.nurse_id,n.name,n.home_shift,result.target_shifts[n.nurse_id],actual_hours(result.assignments,n.nurse_id),
             overtime_hours(result.assignments,result.target_shifts,n.nurse_id),under_target_hours(result.assignments,result.target_shifts,n.nurse_id),
             counts.get("D",0),counts.get("E",0),counts.get("N",0)]
        for d in days:
            a=amap.get((n.nurse_id,d)); row.append(a.shift if a else "OFF")
            audit_rows.append((n.nurse_id,d.isoformat(),a.shift if a else "OFF",a.area if a and a.area else ""))
        ws.append(row)
    ws.freeze_panes="K2"

    # A personnel-by-date area roster matching the visual grammar of the total
    # schedule.  Text keeps shift and area explicit; color is a secondary cue.
    area_ws=wb.create_sheet("總區域班表")
    area_ws.append(headers)
    for n in seniority_order:
        counts=shift_counts(result.assignments,n.nurse_id)
        row=[n.nurse_id,n.name,n.home_shift,result.target_shifts[n.nurse_id],actual_hours(result.assignments,n.nurse_id),
             overtime_hours(result.assignments,result.target_shifts,n.nurse_id),under_target_hours(result.assignments,result.target_shifts,n.nurse_id),
             counts.get("D",0),counts.get("E",0),counts.get("N",0)]
        for d in days:
            a=amap.get((n.nurse_id,d))
            row.append(f"{a.shift}｜{a.area or '未配置'}" if a else "OFF")
        area_ws.append(row)
    area_ws.freeze_panes="K2"

    # Explicit D/E/N group sheets. Do not rely on visual sorting in one mixed table.
    supplement_header_rows: dict[str,int]={}
    for group in ("D","E","N"):
        gs=wb.create_sheet(f"{group}班班表")
        gs.append(headers)
        for n in seniority_order:
            if n.home_shift != group:
                continue
            counts=shift_counts(result.assignments,n.nurse_id)
            row=[n.nurse_id,n.name,n.home_shift,result.target_shifts[n.nurse_id],actual_hours(result.assignments,n.nurse_id),
                 overtime_hours(result.assignments,result.target_shifts,n.nurse_id),under_target_hours(result.assignments,result.target_shifts,n.nurse_id),
                 counts.get("D",0),counts.get("E",0),counts.get("N",0)]
            for d in days:
                a=amap.get((n.nurse_id,d)); row.append(a.shift if a else "OFF")
            gs.append(row)
        supplements=sorted(
            (a for a in result.assignments
             if a.shift==group and (
                 by_id[a.nurse_id].home_shift!=group or a.nurse_id in relief_ids
             )),
            key=lambda a:(a.day,seniority_key(by_id[a.nurse_id])),
        )
        title_row=gs.max_row+2
        gs.merge_cells(start_row=title_row,start_column=1,end_row=title_row,end_column=6)
        gs.cell(title_row,1,f"{group}班跨班／待補補充人員（含原班別或人力來源）")
        header_row=title_row+1
        supplement_header_rows[gs.title]=header_row
        for c,value in enumerate(["日期","支援班別","ID","姓名","原班別／來源","區域"],start=1):
            gs.cell(header_row,c,value)
        if supplements:
            for a in supplements:
                n=by_id[a.nurse_id]
                source="外部／機動" if a.nurse_id in relief_ids else n.home_shift
                gs.append([a.day.isoformat(),a.shift,a.nurse_id,n.name,source,a.area or ""])
        else:
            gs.append(["本月無跨班補充人員"])
        gs.freeze_panes="K2"

    # Daily staffing/gap alert.  Only core shortages are hard; Clinic3 is shown
    # as optional support information and never contributes to the gap total.
    daily=wb.create_sheet("每日人力缺口")
    daily.append(["日期","班別","核心需求","實際核心","核心缺口","總人數","缺少核心區域","Clinic3選配需求","Clinic3實際支援","Clinic3硬性缺口","替補人數","替補人員（ID/姓名/原班別）","提示"])
    required_de=["Leader","Triage","Critical","Clinic1","Clinic2","流動","留觀"]
    required_de_l_plus_t=["Leader+Triage","Critical","Clinic1","Clinic2","流動","留觀"]
    required_n5=["Leader+Triage","Critical","Clinic1","流動","留觀"]
    required_n6=["Leader","Triage","Critical","Clinic1","流動","留觀"]
    gap_rows: list[tuple[date,str,str,int]]=[]
    for d in days:
        for sh in ("D","E","N"):
            aa=[a for a in result.assignments if a.day==d and a.shift==sh]
            l_plus_t_training=(
                sh in {"D","E"}
                and any(a.formal_teaching and a.area=="Leader+Triage" for a in aa)
            )
            roles=(required_de_l_plus_t if l_plus_t_training else required_de) if sh in {"D","E"} else (required_n6 if len(aa)>=6 else required_n5)
            core=[a for a in aa if a.independent_core]
            have=Counter(a.area for a in core); need=Counter(roles)
            missing=[]
            for area,cnt in need.items():
                gap=max(0,cnt-have.get(area,0))
                if gap:
                    missing.append(f"{area}x{gap}")
                    gap_rows.append((d,sh,area,gap))
            core_gap=max(0,len(roles)-len(core))
            c3_need=0
            c3_actual=sum(1 for a in aa if a.area=="Clinic3")
            c3_gap=max(0,c3_need-c3_actual)
            replacements=sorted(
                (a for a in aa if (
                    by_id[a.nurse_id].home_shift!=sh or a.nurse_id in relief_ids
                )),
                key=lambda a:seniority_key(by_id[a.nurse_id]),
            )
            replacement_text="；".join(
                (
                    f"{a.nurse_id}/{by_id[a.nurse_id].name}/外部待補"
                    if a.nurse_id in relief_ids else
                    f"{a.nurse_id}/{by_id[a.nurse_id].name}/{by_id[a.nurse_id].home_shift}"
                )
                for a in replacements
            )
            alerts=[]
            if core_gap or missing: alerts.append(f"缺核心{core_gap}；{'/'.join(missing)}")
            if replacements:
                alerts.append(f"已安排{len(replacements)}名跨班替補")
            if core_gap or missing:
                alerts.append("仍有空缺，需另尋替補人力")
            if sh=="N":
                alerts.append("5人採L+T" if len(aa)<6 else "第6人起Leader/Triage分列")
            elif l_plus_t_training:
                alerts.append("Triage訓練人力備援：合格老師採L+T，學生同區跟訓")
            daily.append([
                d.isoformat(),sh,len(roles),len(core),core_gap,len(aa),"、".join(missing),
                c3_need,c3_actual,c3_gap,
                len(replacements),replacement_text,"；".join(alerts) if alerts else "PASS",
            ])

    ar=wb.create_sheet("區域配置"); ar.append(["日期","班別","ID","姓名","區域","核心計數","正式帶教","Preceptor","人力狀態","原班別","備註"])
    for a in sorted(result.assignments,key=lambda x:(x.day,x.shift,x.nurse_id)):
        original_shift=by_id[a.nurse_id].home_shift
        replacement=a.shift in {"D","E","N"} and original_shift!=a.shift
        relief=a.nurse_id in relief_ids
        ar.append([
            a.day.isoformat(),a.shift,a.nurse_id,by_id[a.nurse_id].name,a.area,
            "1" if a.independent_core else "0","Y" if a.formal_teaching else "N",a.preceptor_id or "",
            "待補人力" if relief else ("替補" if replacement else "正常"),
            "外部／機動" if relief else (original_shift if replacement else ""),
            "需確認實際支援者" if relief else (f"由{original_shift}班支援{a.shift}班" if replacement else ""),
        ])
    for d,sh,area,count in gap_rows:
        for idx in range(1,count+1):
            ar.append([d.isoformat(),sh,"人力空缺","待補",area,"0","N","","待補","",f"空缺{idx}/{count}：請另尋替補人力"])

    fair=wb.create_sheet("工時公平"); fair.append(["ID","姓名","目標班數","實際班數","目標工時","實際工時","加班時數","不足時數","特休天數","未休折抵時數","Solver OT班數","Solver不足班數","OptimalProven","跨班支援天數"])
    annual_by_id=Counter(r.nurse_id for r in requests if r.kind==RequestKind.ANNUAL_LEAVE)
    for n in result.nurses:
        fair.append([n.nurse_id,n.name,result.target_shifts[n.nurse_id],sum(shift_counts(result.assignments,n.nurse_id).values()),result.target_shifts[n.nurse_id]*8,
                     actual_hours(result.assignments,n.nurse_id),overtime_hours(result.assignments,result.target_shifts,n.nurse_id),under_target_hours(result.assignments,result.target_shifts,n.nurse_id),
                     annual_by_id[n.nurse_id],getattr(n,"prior_unused_hours",0),
                     result.solver_metrics.get("overtime_shifts"),result.solver_metrics.get("under_target_shifts"),result.solver_metrics.get("optimization_proven"),
                     cross_shift_days(result.assignments,by_id,n.nurse_id)])

    tr=wb.create_sheet("新人帶教"); tr.append(["日期","班別","訓練類型","訓練區域","學員ID","學員姓名","PreceptorID","Preceptor姓名","正式同區帶教"])
    for a in result.assignments:
        if a.formal_teaching:
            training_type="Triage" if a.area in {"Triage","Leader+Triage"} else "Observation"
            tr.append([a.day.isoformat(),a.shift,training_type,a.area,a.nurse_id,by_id[a.nurse_id].name,a.preceptor_id,by_id[a.preceptor_id].name if a.preceptor_id in by_id else "","Y"])

    # Rolling cumulative training state is embedded in the SAME output workbook, not a sidecar JSON.
    summary_headers=["ID","姓名","月初累積","本月正式帶教","月底累積","月底狀態","來源"]
    for c,v in enumerate(summary_headers,start=10): tr.cell(1,c,v)
    teaching_added=Counter(
        a.nurse_id for a in result.assignments
        if a.formal_teaching and a.area in {"Observation","流動","留觀"}
    )
    sr=2
    for n in result.nurses:
        end_state=result.training_end[n.nurse_id]
        end_count=end_state.completed_observation_sessions
        added=int(teaching_added.get(n.nurse_id,0))
        if end_count is None and added==0:
            continue
        start_count="" if end_count is None else max(0,int(end_count)-added)
        values=[n.nurse_id,n.name,start_count,added,"" if end_count is None else int(end_count),end_state.status.value,end_state.source]
        for c,v in enumerate(values,start=10): tr.cell(sr,c,v)
        sr+=1

    scenario_sheet=None
    if scenario_rows:
        scenario_sheet=wb.create_sheet("情境調整")
        scenario_sheet.append([
            "情境","ID","姓名","日期","原始類型","原始內容",
            "本情境調整","目標影響","來源／備註",
        ])
        for row in scenario_rows:
            scenario_sheet.append(row)

    exclusion_sheet=None
    if monthly_excluded_rows:
        exclusion_sheet=wb.create_sheet("月份排除")
        exclusion_sheet.append(["年月","ID","姓名","處理","來源／備註"])
        for row in monthly_excluded_rows:
            exclusion_sheet.append(row)

    relief_sheet=None
    if relief_rows:
        relief_sheet=wb.create_sheet("待補人力")
        relief_sheet.append(["日期","班別","核心區域","暫用ID","暫用名稱","人力來源","處理狀態"])
        for row in relief_rows:
            relief_sheet.append(row)

    advice_sheet=wb.create_sheet("排班衝突建議")
    advice_sheet.append(["優先級","衝突類型","判定依據","建議處理方式","預期影響","核准／重跑要求"])
    advice_rows=build_schedule_advice(
        result=result,
        validation=validation,
        config=SchedulerConfig(
            result.year,
            result.month,
            max_cross_shift_days_per_person=result.solver_metrics.get(
                "max_cross_shift_days_per_person_cap", 5
            ),
            max_overtime_shifts_per_person=result.solver_metrics.get(
                "max_overtime_shifts_per_person_cap"
            ),
        ),
    )
    for item in advice_rows:
        advice_sheet.append([
            item["priority"], item["category"], item["evidence"],
            item["recommendation"], item["impact"], item["approval"],
        ])
    advice_sheet.freeze_panes="A2"
    advice_sheet.auto_filter.ref=f"A1:F{advice_sheet.max_row}"
    advice_sheet.sheet_view.showGridLines=False
    for cell in advice_sheet[1]:
        cell.fill=PatternFill("solid",fgColor="1F4E78")
        cell.font=Font(bold=True,color="FFFFFF")
        cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    for row_index in range(2,advice_sheet.max_row+1):
        priority=str(advice_sheet.cell(row_index,1).value or "")
        fill_color={"P1":"FCE4D6","P2":"FFF2CC","P3":"D9EAF7","INFO":"E2F0D9"}.get(priority,"FFFFFF")
        for cell in advice_sheet[row_index]:
            cell.fill=PatternFill("solid",fgColor=fill_color)
            cell.alignment=Alignment(vertical="top",wrap_text=True)
            cell.border=Border(
                left=Side(style="thin",color="D9E2F3"),
                right=Side(style="thin",color="D9E2F3"),
                top=Side(style="thin",color="D9E2F3"),
                bottom=Side(style="thin",color="D9E2F3"),
            )
        advice_sheet.row_dimensions[row_index].height=58
    for column,width in {"A":10,"B":22,"C":42,"D":64,"E":38,"F":38}.items():
        advice_sheet.column_dimensions[column].width=width

    gv=wb.create_sheet("Google來源檢查"); gv.append(["項目","結果","內容"])
    ga=result.google_audit
    meta=[
        ("Source Mode","PASS" if ga.source_verified else "WARNING",ga.source_mode),
        ("Source Verified","PASS" if ga.source_verified else "FAIL",str(ga.source_verified)),
        ("Spreadsheet ID","PASS",ga.spreadsheet_id),("文件名稱","PASS",ga.spreadsheet_title),
        ("分頁","PASS",ga.worksheet_title),("Matrix SHA256","PASS",ga.matrix_sha256),
        ("匹配人數",ga.result,str(ga.matched_people)),("非空白預班格",ga.result,str(ga.nonblank_request_cells))]
    for r in meta: gv.append(r)
    for name,status,detail in ga.checks: gv.append([name,status,detail])
    pa=result.previous_audit
    gv.append([])
    gv.append(["上月正式班表 Source Mode","PASS" if pa.source_verified else "FAIL",pa.source_mode])
    gv.append(["上月正式班表 Spreadsheet ID","PASS",pa.spreadsheet_id])
    gv.append(["上月正式班表 文件名稱","PASS",pa.spreadsheet_title])
    gv.append(["上月正式班表 分頁","PASS",pa.worksheet_title])
    gv.append(["上月正式班表 Matrix SHA256","PASS",pa.matrix_sha256])
    gv.append(["上月正式班表 Boundary Entries","PASS",str(pa.boundary_entries)])
    for name,status,detail in pa.checks: gv.append([f"previous::{name}",status,detail])
    if scenario_rows:
        gv.append([])
        gv.append(["情境模式","INFO",scenario_name])
        gv.append(["情境調整筆數","INFO",str(len(scenario_rows))])
        gv.append(["原始 LIVE 資料","PASS","Google_LIVE_RAW_AUDIT 完整保留；情境未回寫 Google Sheet"])
    if monthly_excluded_rows:
        gv.append([])
        gv.append(["月份排除","INFO","；".join(f"{row[1]} {row[2]}" for row in monthly_excluded_rows)])
        gv.append(["排除範圍","PASS","僅本月排班與輸出；Nurses.xlsx／LIVE Google 原始資料未刪除"])
    if relief_rows:
        gv.append([])
        gv.append(["待補人力","WARNING",f"最少需要 {len(relief_rows)} 班外部／機動支援；須在正式發布前填入實際人員"])
        gv.append(["安全規則","PASS","所有既有人員仍遵守連續上班、休息間隔與每人跨班上限"])

    raw=wb.create_sheet("Google_LIVE_RAW_AUDIT")
    for row in google_matrix: raw.append(list(row))

    val=wb.create_sheet("驗證總覽"); val.append(["Publish Status",validation.publish_status.value]); val.append(["Result SHA256",result.result_sha256]); val.append(["Fairness Gate",validation.metrics.get("fairness_gate")]);
    val.append(["Hard Errors",len(validation.hard_errors)]); val.append(["Warnings",len(validation.warnings)])
    val.append([]); val.append(["Severity","Code","ID","Date","Shift","Message"])
    for i in (*validation.hard_errors,*validation.warnings):
        val.append([i.severity.value,i.code,i.nurse_id or "",i.day.isoformat() if i.day else "",i.shift or "",i.message])

    meta_ws=wb.create_sheet("_IMMUTABLE_META"); meta_ws.sheet_state="hidden"
    schedule_digest=_schedule_digest_from_rows(audit_rows)
    meta_ws.append(["result_sha256",result.result_sha256])
    meta_ws.append(["schedule_digest",schedule_digest])
    meta_ws.append(["google_matrix_sha256",ga.matrix_sha256])
    meta_ws.append(["previous_google_matrix_sha256",result.previous_audit.matrix_sha256])
    meta_ws.append(["previous_google_sheet_id",result.previous_audit.spreadsheet_id])
    meta_ws.append(["previous_google_tab",result.previous_audit.worksheet_title])
    meta_ws.append(["publish_status",validation.publish_status.value])
    meta_ws.append(["scenario_name",scenario_name if scenario_rows else ""])
    meta_ws.append(["scenario_override_count",len(scenario_rows)])
    meta_ws.append(["scenario_override_digest",scenario_digest])
    meta_ws.append(["monthly_excluded_count",len(monthly_excluded_rows)])
    meta_ws.append(["monthly_excluded_digest",monthly_excluded_digest])
    meta_ws.append(["relief_shift_count",len(relief_rows)])
    meta_ws.append(["relief_assignment_digest",relief_digest])
    meta_ws.append(["minimum_relief_proven",result.solver_metrics.get("minimum_relief_proven",False)])
    meta_ws.append(["minimum_cross_shift_days",result.solver_metrics.get("minimum_cross_shift_days","")])
    meta_ws.append(["max_cross_shift_days_per_person_cap",result.solver_metrics.get("max_cross_shift_days_per_person_cap","")])
    meta_ws.append(["area_matrix_format","SHIFT｜AREA"])
    meta_ws.append(["observation_core_split","Observation1=流動;Observation2=留觀"])

    # Styling only. No schedule mutation after this point.
    for sheet in wb.worksheets:
        if sheet.max_row>=1:
            for cell in sheet[1]:
                cell.font=Font(bold=True,color="FFFFFF"); cell.fill=PatternFill("solid",fgColor="1F4E78"); cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
            sheet.row_dimensions[1].height=32
            sheet.freeze_panes="A2"
            sheet.sheet_view.showGridLines=False
            sheet.sheet_view.zoomScale=85

            # Content-aware widths with a conservative cap. This keeps IDs,
            # dates, audit labels and validation messages visible without
            # creating unusably wide sheets.
            for col_idx in range(1,sheet.max_column+1):
                max_len=0
                for row_idx in range(1,min(sheet.max_row,700)+1):
                    value=sheet.cell(row_idx,col_idx).value
                    if value is not None:
                        max_len=max(max_len,len(str(value)))
                sheet.column_dimensions[get_column_letter(col_idx)].width=min(max(max_len+2,6),45)

    # Schedule tabs stay compact enough to scan an entire month.
    for name in ("總班表","D班班表","E班班表","N班班表"):
        sheet=wb[name]
        for col,width in {"A":10,"B":12,"C":8,"D":10,"E":10,"F":10,"G":10,"H":8,"I":8,"J":8}.items():
            sheet.column_dimensions[col].width=width
        for col_idx in range(11,11+len(days)):
            sheet.column_dimensions[get_column_letter(col_idx)].width=5
        sheet.freeze_panes="K2"
    for col,width in {"A":10,"B":12,"C":8,"D":10,"E":10,"F":10,"G":10,"H":8,"I":8,"J":8}.items():
        area_ws.column_dimensions[col].width=width
    for col_idx in range(11,11+len(days)):
        area_ws.column_dimensions[get_column_letter(col_idx)].width=15
    for row_idx in range(2,area_ws.max_row+1):
        area_ws.row_dimensions[row_idx].height=30
    area_ws.freeze_panes="K2"
    area_ws.sheet_properties.tabColor="5B9BD5"

    # Shift colors are identical in the total roster, group rosters and total
    # area roster.  Status columns remain textual so color is never the only cue.
    shift_fills={
        "D":PatternFill("solid",fgColor="D9EAF7"),
        "E":PatternFill("solid",fgColor="FCE4D6"),
        "N":PatternFill("solid",fgColor="E4DFEC"),
        "H":PatternFill("solid",fgColor="E2F0D9"),
        "OFF":PatternFill("solid",fgColor="E7E6E6"),
    }
    for name in ("總班表","D班班表","E班班表","N班班表"):
        sheet=wb[name]
        for row in sheet.iter_rows(min_row=2,min_col=11,max_col=10+len(days)):
            for cell in row:
                value=str(cell.value or "")
                if value in shift_fills:
                    cell.fill=shift_fills[value]
                    cell.alignment=Alignment(horizontal="center",vertical="center")
    for row in area_ws.iter_rows(min_row=2,min_col=11,max_col=10+len(days)):
        for cell in row:
            value=str(cell.value or "")
            shift="OFF" if value=="OFF" else value.split("｜",1)[0]
            if shift in shift_fills:
                cell.fill=shift_fills[shift]
            cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    for row_idx in range(2,ar.max_row+1):
        shift=str(ar.cell(row_idx,2).value or "")
        if shift in shift_fills:
            for cell in ar[row_idx]:
                cell.fill=shift_fills[shift]
        status=str(ar.cell(row_idx,9).value or "")
        if status=="替補":
            ar.cell(row_idx,9).fill=PatternFill("solid",fgColor="FFF2CC")
            ar.cell(row_idx,10).fill=PatternFill("solid",fgColor="FFF2CC")
        elif status=="待補":
            for cell in ar[row_idx]:
                cell.fill=PatternFill("solid",fgColor="F4CCCC")
        elif status=="待補人力":
            for cell in ar[row_idx]:
                cell.fill=PatternFill("solid",fgColor="F4B183")

    # Make the new hard cap auditable at a glance.  Any cross-shift support is
    # yellow; a person exactly at the configured maximum is orange.
    cross_cap=result.solver_metrics.get("max_cross_shift_days_per_person_cap")
    for row_idx in range(2,fair.max_row+1):
        cell=fair.cell(row_idx,14)
        count=int(cell.value or 0)
        if count>0:
            cell.fill=PatternFill("solid",fgColor="FFF2CC")
        if cross_cap is not None and count==int(cross_cap):
            cell.fill=PatternFill("solid",fgColor="F4B183")

    # Daily shortage rows are red; rows already covered by cross-shift support
    # are yellow. This explicitly marks the affected date.
    for row_idx in range(2,daily.max_row+1):
        has_gap=any(int(daily.cell(row_idx,col).value or 0)>0 for col in (5,10))
        has_replacement=int(daily.cell(row_idx,11).value or 0)>0
        if has_gap:
            for cell in daily[row_idx]:
                cell.fill=PatternFill("solid",fgColor="F4CCCC")
        elif has_replacement:
            for cell in daily[row_idx]:
                cell.fill=PatternFill("solid",fgColor="FFF2CC")

    for sheet_name,header_row in supplement_header_rows.items():
        sheet=wb[sheet_name]
        title=sheet.cell(header_row-1,1)
        title.font=Font(bold=True,color="FFFFFF")
        title.fill=PatternFill("solid",fgColor="548235")
        for cell in sheet[header_row][:6]:
            cell.font=Font(bold=True,color="FFFFFF")
            cell.fill=PatternFill("solid",fgColor="70AD47")
            cell.alignment=Alignment(horizontal="center",vertical="center")

    # High-value audit/report sheets need room for full labels and messages.
    for col,width in {"A":14,"B":9,"C":13,"D":13,"E":13,"F":13,"G":18,"H":16,"I":13,"J":16,"K":12,"L":42,"M":38}.items():
        wb["每日人力缺口"].column_dimensions[col].width=width
    for col,width in {"A":14,"B":8,"C":12,"D":12,"E":16,"F":10,"G":10,"H":12,"I":10,"J":10,"K":36}.items():
        wb["區域配置"].column_dimensions[col].width=width
    wb["工時公平"].column_dimensions["N"].width=16
    for col,width in {"A":32,"B":12,"C":80}.items():
        wb["Google來源檢查"].column_dimensions[col].width=width
    for col,width in {"A":20,"B":24,"C":12,"D":14,"E":12,"F":70}.items():
        wb["驗證總覽"].column_dimensions[col].width=width
    if scenario_sheet is not None:
        scenario_sheet.sheet_properties.tabColor="F4B183"
        for col,width in {"A":24,"B":10,"C":12,"D":14,"E":18,"F":20,"G":42,"H":34,"I":70}.items():
            scenario_sheet.column_dimensions[col].width=width
        for row_idx in range(2,scenario_sheet.max_row+1):
            for cell in scenario_sheet[row_idx]:
                cell.fill=PatternFill("solid",fgColor="FFF2CC")
                cell.alignment=Alignment(vertical="top",wrap_text=True)
            scenario_sheet.row_dimensions[row_idx].height=44
    if exclusion_sheet is not None:
        exclusion_sheet.sheet_properties.tabColor="C00000"
        for col,width in {"A":12,"B":10,"C":12,"D":26,"E":70}.items():
            exclusion_sheet.column_dimensions[col].width=width
        for row_idx in range(2,exclusion_sheet.max_row+1):
            for cell in exclusion_sheet[row_idx]:
                cell.fill=PatternFill("solid",fgColor="F4CCCC")
                cell.alignment=Alignment(vertical="top",wrap_text=True)
            exclusion_sheet.row_dimensions[row_idx].height=36
    if relief_sheet is not None:
        relief_sheet.sheet_properties.tabColor="ED7D31"
        for col,width in {"A":14,"B":8,"C":18,"D":14,"E":16,"F":20,"G":28}.items():
            relief_sheet.column_dimensions[col].width=width
        for row_idx in range(2,relief_sheet.max_row+1):
            for cell in relief_sheet[row_idx]:
                cell.fill=PatternFill("solid",fgColor="FCE4D6")
                cell.alignment=Alignment(vertical="top",wrap_text=True)
    wrap_sheets=["Google來源檢查","驗證總覽"]
    if scenario_sheet is not None:
        wrap_sheets.append("情境調整")
    if exclusion_sheet is not None:
        wrap_sheets.append("月份排除")
    if relief_sheet is not None:
        wrap_sheets.append("待補人力")
    for name in wrap_sheets:
        for row in wb[name].iter_rows():
            for cell in row:
                cell.alignment=Alignment(vertical="top",wrap_text=True)
    path.parent.mkdir(parents=True,exist_ok=True)
    wb.save(path); wb.close()

    # Independent post-export readback gate.
    check=load_workbook(path,read_only=True,data_only=True)
    try:
        if "_IMMUTABLE_META" not in check.sheetnames:
            raise RuntimeError("POST_EXPORT_FAIL: missing immutable meta")
        ms=check["_IMMUTABLE_META"]
        stored={str(r[0]):str(r[1]) for r in ms.iter_rows(values_only=True) if r and r[0]}
        if stored.get("result_sha256")!=result.result_sha256:
            raise RuntimeError("POST_EXPORT_FAIL: result hash mismatch")
        if stored.get("google_matrix_sha256")!=ga.matrix_sha256:
            raise RuntimeError("POST_EXPORT_FAIL: Google pre-request hash mismatch")
        if stored.get("previous_google_matrix_sha256")!=result.previous_audit.matrix_sha256:
            raise RuntimeError("POST_EXPORT_FAIL: Google previous-roster hash mismatch")
        if stored.get("previous_google_sheet_id")!=result.previous_audit.spreadsheet_id:
            raise RuntimeError("POST_EXPORT_FAIL: Google previous-roster sheet-id mismatch")
        if stored.get("previous_google_tab")!=result.previous_audit.worksheet_title:
            raise RuntimeError("POST_EXPORT_FAIL: Google previous-roster tab mismatch")
        if stored.get("scenario_override_count")!=str(len(scenario_rows)):
            raise RuntimeError("POST_EXPORT_FAIL: scenario override count mismatch")
        if stored.get("scenario_override_digest")!=scenario_digest:
            raise RuntimeError("POST_EXPORT_FAIL: scenario override digest mismatch")
        if stored.get("monthly_excluded_count")!=str(len(monthly_excluded_rows)):
            raise RuntimeError("POST_EXPORT_FAIL: monthly exclusion count mismatch")
        if stored.get("monthly_excluded_digest")!=monthly_excluded_digest:
            raise RuntimeError("POST_EXPORT_FAIL: monthly exclusion digest mismatch")
        if stored.get("relief_shift_count")!=str(len(relief_rows)):
            raise RuntimeError("POST_EXPORT_FAIL: relief shift count mismatch")
        if stored.get("relief_assignment_digest")!=relief_digest:
            raise RuntimeError("POST_EXPORT_FAIL: relief assignment digest mismatch")
        if stored.get("area_matrix_format")!="SHIFT｜AREA":
            raise RuntimeError("POST_EXPORT_FAIL: area matrix format metadata mismatch")
        if stored.get("observation_core_split")!="Observation1=流動;Observation2=留觀":
            raise RuntimeError("POST_EXPORT_FAIL: observation split metadata mismatch")
        if scenario_rows:
            if "情境調整" not in check.sheetnames:
                raise RuntimeError("POST_EXPORT_FAIL: missing scenario audit sheet")
            scenario_check=check["情境調整"]
            exported_scenario_rows=[
                [str(value or "") for value in row[:9]]
                for row in scenario_check.iter_rows(min_row=2,values_only=True)
                if any(value not in (None,"") for value in row[:9])
            ]
            exported_scenario_digest=hashlib.sha256(
                json.dumps(exported_scenario_rows,ensure_ascii=False,separators=(",",":")).encode("utf-8")
            ).hexdigest()
            if exported_scenario_digest!=scenario_digest:
                raise RuntimeError("POST_EXPORT_FAIL: scenario audit sheet digest mismatch")
        elif "情境調整" in check.sheetnames:
            raise RuntimeError("POST_EXPORT_FAIL: unexpected scenario audit sheet")
        if monthly_excluded_rows:
            if "月份排除" not in check.sheetnames:
                raise RuntimeError("POST_EXPORT_FAIL: missing monthly exclusion sheet")
            exclusion_check=check["月份排除"]
            exported_exclusion_rows=[
                [str(value or "") for value in row[:5]]
                for row in exclusion_check.iter_rows(min_row=2,values_only=True)
                if any(value not in (None,"") for value in row[:5])
            ]
            exported_exclusion_digest=hashlib.sha256(
                json.dumps(exported_exclusion_rows,ensure_ascii=False,separators=(",",":")).encode("utf-8")
            ).hexdigest()
            if exported_exclusion_digest!=monthly_excluded_digest:
                raise RuntimeError("POST_EXPORT_FAIL: monthly exclusion sheet digest mismatch")
        elif "月份排除" in check.sheetnames:
            raise RuntimeError("POST_EXPORT_FAIL: unexpected monthly exclusion sheet")
        if relief_rows:
            if "待補人力" not in check.sheetnames:
                raise RuntimeError("POST_EXPORT_FAIL: missing relief sheet")
            relief_check=check["待補人力"]
            exported_relief_rows=[
                [str(value or "") for value in row[:7]]
                for row in relief_check.iter_rows(min_row=2,values_only=True)
                if any(value not in (None,"") for value in row[:7])
            ]
            exported_relief_digest=hashlib.sha256(
                json.dumps(exported_relief_rows,ensure_ascii=False,separators=(",",":")).encode("utf-8")
            ).hexdigest()
            if exported_relief_digest!=relief_digest:
                raise RuntimeError("POST_EXPORT_FAIL: relief sheet digest mismatch")
        elif "待補人力" in check.sheetnames:
            raise RuntimeError("POST_EXPORT_FAIL: unexpected relief sheet")
        # Re-read schedule AND area cells independently; never borrow area values from in-memory result.
        s=check["總班表"]
        head=list(next(s.iter_rows(values_only=True)))
        day_cols={int(v):i for i,v in enumerate(head) if str(v).isdigit() and 1<=int(v)<=31}
        shift_map={}
        for rr in s.iter_rows(min_row=2,values_only=True):
            nid=str(rr[0])
            for d in days:
                shift_map[(nid,d.isoformat())]=str(rr[day_cols[d.day]] or "OFF")
        area_map={}
        ar_check=check["區域配置"]
        for rr in ar_check.iter_rows(min_row=2,values_only=True):
            d,sh,nid,_,area,*_=rr
            area_map[(str(nid),str(d),str(sh))]=str(area or "")
        if "總區域班表" not in check.sheetnames:
            raise RuntimeError("POST_EXPORT_FAIL: missing matrix area roster")
        area_matrix_check=check["總區域班表"]
        area_head=list(next(area_matrix_check.iter_rows(values_only=True)))
        area_day_cols={int(v):i for i,v in enumerate(area_head) if str(v).isdigit() and 1<=int(v)<=31}
        area_matrix_rows={str(rr[0]):rr for rr in area_matrix_check.iter_rows(min_row=2,values_only=True)}
        rows=[]
        for nid in by_id:
            for d in days:
                ds=d.isoformat(); sh=shift_map[(nid,ds)]
                area="" if sh=="OFF" else area_map.get((nid,ds,sh),"")
                matrix_value=str(area_matrix_rows[nid][area_day_cols[d.day]] or "OFF")
                expected_matrix="OFF" if sh=="OFF" else f"{sh}｜{area}"
                if matrix_value!=expected_matrix:
                    raise RuntimeError(
                        f"POST_EXPORT_FAIL: area matrix mismatch {nid} {ds}: "
                        f"actual={matrix_value!r}, expected={expected_matrix!r}"
                    )
                rows.append((nid,ds,sh,area))
        if _schedule_digest_from_rows(rows)!=stored.get("schedule_digest"):
            raise RuntimeError("POST_EXPORT_FAIL: schedule/area digest mismatch")

        # Independently re-hash the exported Google RAW sheet.
        raw_check=check["Google_LIVE_RAW_AUDIT"]
        raw_matrix=[list(r) for r in raw_check.iter_rows(values_only=True)]
        # Trim only trailing empty cells/rows introduced by xlsx storage, preserving meaningful matrix content.
        while raw_matrix and not any(v not in (None,"") for v in raw_matrix[-1]): raw_matrix.pop()
        maxcol=max((max((i for i,v in enumerate(r) if v not in (None,"")),default=-1)+1 for r in raw_matrix),default=0)
        raw_matrix=[r[:maxcol] for r in raw_matrix]
        if matrix_sha256(raw_matrix)!=ga.matrix_sha256:
            raise RuntimeError("POST_EXPORT_FAIL: Google RAW matrix hash mismatch")
    except Exception:
        check.close()
        try: path.unlink()
        except Exception: pass
        raise
    check.close()
    return path


# ==============================================================================
# MODULE: reporting.py
# ==============================================================================

import json
from pathlib import Path



def write_validation_report(
    path: Path,
    feasibility: FeasibilityReport,
    result: ScheduleResult | None,
    validation: ValidationResult | None,
    *,
    config: SchedulerConfig | None = None,
    solver_error: BaseException | str | None = None,
) -> Path:
    lines=[f"{PROGRAM_VERSION} — Validation Report","="*72]
    lines.append(f"Pre-feasibility: {'PASS' if feasibility.feasible else 'BLOCKED'}")
    for k,v in feasibility.metrics.items(): lines.append(f"  {k}: {v}")
    for i in feasibility.hard_issues: lines.append(f"[HARD] {i.code} {i.message}")
    for i in feasibility.warnings: lines.append(f"[WARN] {i.code} {i.message}")
    if result:
        lines += ["",f"Solver: {result.solver_status}",f"Result SHA256: {result.result_sha256}"]
        for k,v in result.solver_metrics.items(): lines.append(f"  {k}: {v}")
        lines += [
            "", "Google Pre-request Source:",
            f"  result: {result.google_audit.result}",
            f"  id: {result.google_audit.spreadsheet_id}",
            f"  title: {result.google_audit.spreadsheet_title}",
            f"  tab: {result.google_audit.worksheet_title}",
            f"  sha256: {result.google_audit.matrix_sha256}",
            "", "Google Previous Official Roster Source:",
            f"  result: {result.previous_audit.result}",
            f"  id: {result.previous_audit.spreadsheet_id}",
            f"  title: {result.previous_audit.spreadsheet_title}",
            f"  tab: {result.previous_audit.worksheet_title}",
            f"  sha256: {result.previous_audit.matrix_sha256}",
            f"  boundary_entries: {result.previous_audit.boundary_entries}",
        ]
    if validation:
        lines += ["",f"Publish Status: {validation.publish_status.value}",f"Fairness Gate: {validation.metrics.get('fairness_gate')}",f"Hard errors: {len(validation.hard_errors)}",f"Warnings: {len(validation.warnings)}"]
        for i in validation.hard_errors: lines.append(f"[HARD] {i.code} {i.nurse_id or ''} {i.day or ''} {i.message}")
        for i in validation.warnings: lines.append(f"[WARN] {i.code} {i.nurse_id or ''} {i.day or ''} {i.message}")
    advice = build_schedule_advice(
        feasibility=feasibility,
        result=result,
        validation=validation,
        config=config,
        solver_error=solver_error,
    )
    lines += ["", "Actionable schedule advice", "-"*72]
    for index, item in enumerate(advice, start=1):
        lines.append(
            f"{index}. [{item['priority']}] {item['category']}\n"
            f"   Evidence: {item['evidence']}\n"
            f"   Recommendation: {item['recommendation']}\n"
            f"   Impact: {item['impact']}\n"
            f"   Approval: {item['approval']}"
        )
    path.write_text("\n".join(lines)+"\n",encoding="utf-8")
    return path


# ==============================================================================
# MODULE: main.py
# ==============================================================================

import argparse
import os
from pathlib import Path



if __name__ == "__main__":
    raise SystemExit("Run python main.py --demo from the project folder.")

