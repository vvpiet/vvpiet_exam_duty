import datetime
from typing import List, Dict
import pandas as pd

def generate_exam_dates(start_date: datetime.date, end_date: datetime.date, exclude_weekends: bool, holidays: List[datetime.date]) -> List[datetime.date]:
    dates = []
    cur = start_date
    while cur <= end_date:
        # When exclude_weekends is True we skip Sundays only (weekday()==6).
        # Historically we skipped Sat/Sun; updated to skip only Sundays (weekday()==6)
        if exclude_weekends and cur.weekday() == 6:
            cur += datetime.timedelta(days=1)
            continue
        if cur in holidays:
            cur += datetime.timedelta(days=1)
            continue
        dates.append(cur)
        cur += datetime.timedelta(days=1)
    return dates

def generate_schedule(dates: List[datetime.date], default_blocks: int, special_blocks: Dict[datetime.date,int], staff_df: pd.DataFrame) -> pd.DataFrame:
    """Generates a schedule DataFrame with columns: date, session (Morning/Evening), assigned (list of supervisors)
    Rules: For each session we assign 1 main supervisor and extras: if blocks == 1 then add 1 extra, else add 2 extras.
    Supervisors are assigned in round-robin order.
    """
    names = staff_df.iloc[:,1].fillna("Unnamed").tolist()
    if len(names) == 0:
        raise ValueError("No supervisors available")
    schedule_rows = []
    idx = 0
    for d in dates:
        blocks = special_blocks.get(d, default_blocks)
        # Number of supervisors per session is based on number of blocks selected by the user
        # Rule: if blocks == 1 then add 1 extra supervisor, else add 2 extra supervisors
        extras = 1 if blocks == 1 else 2
        for session in ["Morning", "Evening"]:
            # Total supervisors = blocks + extras
            count = blocks + extras
            assigned = []
            for i in range(count):
                assigned.append(names[idx % len(names)])
                idx += 1
            schedule_rows.append({"date": d, "session": session, "assigned": assigned})
    return pd.DataFrame(schedule_rows)

def build_supervisor_table(supervisor_name: str, schedule_df: pd.DataFrame) -> pd.DataFrame:
    """Build per-supervisor table with Sr. No., Date, Morning, Evening (ticks)"""
    rows = []
    # Get unique dates
    dates = sorted(schedule_df["date"].unique())
    sr = 1
    for d in dates:
        morning = schedule_df[(schedule_df["date"]==d) & (schedule_df["session"]=="Morning")]
        evening = schedule_df[(schedule_df["date"]==d) & (schedule_df["session"]=="Evening")]
        m_tick = "✓" if supervisor_name in (morning.iloc[0]["assigned"] if not morning.empty else []) else ""
        e_tick = "✓" if supervisor_name in (evening.iloc[0]["assigned"] if not evening.empty else []) else ""
        if m_tick or e_tick:
            rows.append({"Sr. No.": sr, "Date": d.strftime('%Y-%m-%d'), "Morning": m_tick, "Evening": e_tick})
            sr += 1
    return pd.DataFrame(rows)
