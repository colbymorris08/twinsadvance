#!/usr/bin/env python3
"""Minnesota Twins hitter tracking sheet for every position player
who was on the active roster at any point in September 2026.

Downloads are cached under cache/ so a later run can recompute without
re-fetching. Stats are calculated from Baseball Savant Statcast (pybaseball)
and the public MLB Stats API. Nothing in the sheet is filled by assumption:
a rate with a zero denominator is left blank.

Rerun from this folder:
    python3 build_twins_hitter_sheet.py
"""

from __future__ import annotations

import json
import math
import time
import traceback
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
import requests
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from pybaseball import statcast

ROOT = Path(__file__).resolve().parent
CACHE = ROOT / "cache"
SC_DIR = CACHE / "statcast"
PBP_DIR = CACHE / "pbp"
OUT = ROOT / "output"
DATA = ROOT / "data"
CANVAS = Path(
    "/Users/colbymorris/.cursor/projects/Users-colbymorris-twins-hitter-sheet"
    "/canvases/twins-hitter-tracking.canvas.tsx"
)

ROSTER_DATE = "2026-09-30"
SEASON_END = "2026-09-30"
SEPTEMBER_START = "2026-09-01"
TEAM_ID = 142
EXTRACT_VERSION = 1

# Named examples. Applied after the exit-velocity rule.
LEE_ID = 686797
KEASCHALL_ID = 807712
BUXTON_ID = 621439
EXAMPLE_IDS = {LEE_ID, KEASCHALL_ID, BUXTON_ID}
OVERRIDE_ARCHETYPE = {
    LEE_ID: "Damage",  # Brooks Lee, power hitter, damage example
    KEASCHALL_ID: "Barrel",  # Luke Keaschall, average-EV hitter, barrel example
    BUXTON_ID: "Damage",  # Byron Buxton, power / high-EV hitter
}
# Earlier sheet names that were not on any September 2026 active roster.
PRIOR_SHEET_NAMES = {
    "Brooks Lee",
    "Kody Clemens",
    "Ryan Kreidler",
    "Trevor Larnach",
    "Ryan Fitzgerald",
    "Austin Martin",
    "Luke Keaschall",
    "Matt Wallner",
    "Royce Lewis",
    "James Outman",
    "Byron Buxton",
    "Josh Bell",
    "Mickey Gasper",
    "Alex Jackson",
}
MONTH_WINDOWS = [
    ("april", "April", "2026-04-01", "2026-04-30", "april.json"),
    ("may", "May", "2026-05-01", "2026-05-31", "may.json"),
    ("june", "June", "2026-06-01", "2026-06-30", "june.json"),
    ("july", "July", "2026-07-01", "2026-07-31", "july.json"),
    ("august", "August", "2026-08-01", "2026-08-31", "august.json"),
    ("september", "September", "2026-09-01", "2026-09-30", "september.json"),
]

PA_EVENTS = {
    "single",
    "double",
    "triple",
    "home_run",
    "walk",
    "intent_walk",
    "hit_by_pitch",
    "strikeout",
    "strikeout_double_play",
    "strikeout_triple_play",
    "field_out",
    "force_out",
    "grounded_into_double_play",
    "grounded_into_triple_play",
    "fielders_choice",
    "fielders_choice_out",
    "double_play",
    "triple_play",
    "sac_fly",
    "sac_bunt",
    "sac_fly_double_play",
    "sac_bunt_double_play",
    "catcher_interf",
    "field_error",
    "batter_interference",
    "fan_interference",
}

BUNT_DESC = {"foul_bunt", "missed_bunt", "bunt_foul_tip"}
SWING_DESC = {
    "swinging_strike",
    "swinging_strike_blocked",
    "foul",
    "foul_tip",
    "hit_into_play",
}
WHIFF_DESC = {"swinging_strike", "swinging_strike_blocked"}
SAFE_ADVANCE = {"3B", "score", "4B", "home"}

# White band: within 10% of the comparison league rate, inclusive.
BAND = 0.10
SMALL_BB = 8
SMALL_N = 10
ARCHETYPE_MIN_BB = 30
MODEL_MIN_BB = 100
EV_CLEAR_GAP = 1.0  # mph above pooled league average EV = high exit velocity
RESIDUAL_GAP = 0.02  # barrel rate points above the EV prediction

NAVY = "122C6F"
GOLD = "E6C35C"
FILL = {
    "green": PatternFill("solid", fgColor="D8F3E4"),
    "red": PatternFill("solid", fgColor="FDE2E4"),
    "white": PatternFill("solid", fgColor="FFFFFF"),
    "small": PatternFill("solid", fgColor="F4F4F5"),
    "blank": PatternFill("solid", fgColor="FFFFFF"),
}
FONT_CELL = {
    "green": Font(name="Calibri", color="0B6B3A", bold=True, size=11),
    "red": Font(name="Calibri", color="9F1239", bold=True, size=11),
    "white": Font(name="Calibri", color="111111", size=11),
    "small": Font(name="Calibri", color="6B7280", italic=True, size=11),
    "blank": Font(name="Calibri", color="111111", size=11),
}
THIN = Border(
    left=Side(style="thin", color="E5E7EB"),
    right=Side(style="thin", color="E5E7EB"),
    top=Side(style="thin", color="E5E7EB"),
    bottom=Side(style="thin", color="E5E7EB"),
)
HEADER_FILL = PatternFill("solid", fgColor=NAVY)
HEADER_FONT = Font(name="Calibri", color="FFFFFF", bold=True, size=11)
TITLE_FONT = Font(name="Calibri", color=NAVY, bold=True, size=16)
SUB_FONT = Font(name="Calibri", color="1A2744", size=11)
WRAP = Alignment(wrap_text=True, vertical="center", horizontal="center")
LEFT = Alignment(wrap_text=True, vertical="center", horizontal="left")


def log(msg: str) -> None:
    print(msg, flush=True)


def get_json(url: str, timeout: int = 60) -> dict:
    r = requests.get(url, timeout=timeout)
    r.raise_for_status()
    return r.json()


def roster_on(date: str) -> list[dict]:
    url = (
        "https://statsapi.mlb.com/api/v1/teams/"
        f"{TEAM_ID}/roster?rosterType=active&date={date}&season=2026"
    )
    data = get_json(url)
    rows = []
    for p in data.get("roster") or []:
        rows.append(
            {
                "id": int(p["person"]["id"]),
                "name": p["person"]["fullName"],
                "pos": p["position"]["abbreviation"],
                "pos_type": p["position"]["type"],
                "status": (p.get("status") or {}).get("code"),
            }
        )
    return rows


def load_september_roster() -> tuple[list[dict], list[dict], dict]:
    """Every player on the Twins active roster on at least one day in September 2026."""
    CACHE.mkdir(parents=True, exist_ok=True)
    days = []
    cursor = date(2026, 9, 1)
    while cursor <= date(2026, 9, 30):
        days.append(cursor.isoformat())
        cursor += timedelta(days=1)
    daily: dict[str, list[dict]] = {}
    for day in days:
        path = CACHE / f"roster_{day}.json"
        rows = None
        if path.exists():
            try:
                cached = json.loads(path.read_text())
            except json.JSONDecodeError:
                cached = None
            if (
                isinstance(cached, list)
                and cached
                and isinstance(cached[0], dict)
                and "pos_type" in cached[0]
            ):
                rows = cached
        if rows is None:
            rows = roster_on(day)
            path.write_text(json.dumps(rows, indent=2))
            log(f"roster {day}: {len(rows)}")
        daily[day] = rows
    by_id: dict[int, dict] = {}
    for day, rows in daily.items():
        for row in rows:
            rec = by_id.setdefault(
                row["id"],
                {
                    "id": row["id"],
                    "name": row["name"],
                    "pos": row["pos"],
                    "pos_type": row["pos_type"],
                    "dates": [],
                    "hitter_days": 0,
                },
            )
            rec["dates"].append(day)
            if row["pos_type"] != "Pitcher":
                rec["hitter_days"] += 1
                rec["name"] = row["name"]
                rec["pos"] = row["pos"]
                rec["pos_type"] = row["pos_type"]
    hitters = []
    pitchers = []
    for rec in by_id.values():
        rec["dates"].sort()
        rec["roster_days"] = len(rec["dates"])
        rec["roster_first"] = rec["dates"][0]
        rec["roster_last"] = rec["dates"][-1]
        rec.pop("dates")
        if rec["hitter_days"] > 0:
            hitters.append(rec)
        else:
            pitchers.append(rec)
    hitters.sort(key=lambda p: (last_name(p["name"]).lower(), p["name"].lower()))
    pitchers.sort(key=lambda p: (last_name(p["name"]).lower(), p["name"].lower()))
    sept30 = daily["2026-09-30"]
    sept30_hitters = {p["id"] for p in sept30 if p["pos_type"] != "Pitcher"}
    union_ids = {p["id"] for p in hitters}
    names = {p["name"] for p in hitters}
    meta = {
        "days": len(days),
        "active_on_sept30": len(sept30),
        "hitters_on_sept30": len(sept30_hitters),
        "added_vs_sept30": [p["name"] for p in hitters if p["id"] not in sept30_hitters],
        "missing_from_union_but_on_sept30": [
            p["name"] for p in sept30 if p["pos_type"] != "Pitcher" and p["id"] not in union_ids
        ],
        "added_vs_prior_sheet": [p["name"] for p in hitters if p["name"] not in PRIOR_SHEET_NAMES],
        "removed_vs_prior_sheet": sorted(PRIOR_SHEET_NAMES - names),
    }
    (CACHE / "roster_september_2026_union.json").write_text(
        json.dumps({"hitters": hitters, "pitchers": pitchers, "meta": meta}, indent=2)
    )
    return hitters, pitchers, meta


def _played_games(games: list[dict]) -> list[dict]:
    kept = []
    for g in games:
        status = (g.get("status") or "").lower()
        if "cancel" in status or "postpon" in status:
            continue
        kept.append(g)
    return kept


def load_schedule() -> list[dict]:
    path = CACHE / "schedule_2026_rs.json"
    if path.exists():
        return _played_games(json.loads(path.read_text()))
    games = []
    # Month chunks keep the schedule payload small.
    spans = [
        ("2026-03-01", "2026-04-30"),
        ("2026-05-01", "2026-06-30"),
        ("2026-07-01", "2026-08-31"),
        ("2026-09-01", "2026-09-30"),
    ]
    for start, end in spans:
        url = (
            "https://statsapi.mlb.com/api/v1/schedule?sportId=1&gameType=R"
            f"&startDate={start}&endDate={end}"
        )
        data = get_json(url, timeout=90)
        for d in data.get("dates") or []:
            for g in d.get("games") or []:
                state = (g.get("status") or {}).get("abstractGameState") or ""
                detailed = (g.get("status") or {}).get("detailedState") or ""
                if state not in {"Final"} and not detailed.startswith("Final") and detailed not in {
                    "Completed Early",
                    "Game Over",
                }:
                    continue
                games.append(
                    {
                        "game_pk": int(g["gamePk"]),
                        "date": d["date"],
                        "status": detailed or state,
                    }
                )
        log(f"schedule {start}..{end}: running total {len(games)} finals")
    # One row per game pk.
    dedup = {g["game_pk"]: g for g in games}
    games = _played_games(sorted(dedup.values(), key=lambda g: (g["date"], g["game_pk"])))
    CACHE.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(games))
    return games


def pull_statcast_day(day: str) -> pd.DataFrame:
    try:
        df = statcast(day, day, verbose=False, parallel=False)
    except TypeError:
        df = statcast(day, day)
    return df


def download_statcast(dates: list[str]) -> None:
    SC_DIR.mkdir(parents=True, exist_ok=True)
    todo = [d for d in dates if not (SC_DIR / f"{d}.parquet").exists()]
    log(f"statcast days cached {len(dates) - len(todo)} / {len(dates)}; downloading {len(todo)}")
    if not todo:
        return

    def one(day: str) -> tuple[str, int, str | None]:
        dest = SC_DIR / f"{day}.parquet"
        err = None
        for attempt in range(4):
            try:
                df = pull_statcast_day(day)
                if df is None or len(df) == 0:
                    raise RuntimeError("empty statcast frame")
                keep = [
                    c
                    for c in [
                        "game_date",
                        "game_pk",
                        "game_type",
                        "batter",
                        "events",
                        "description",
                        "zone",
                        "type",
                        "bb_type",
                        "launch_speed",
                        "launch_angle",
                        "launch_speed_angle",
                        "at_bat_number",
                        "pitch_number",
                        "outs_when_up",
                        "on_1b",
                        "on_2b",
                        "on_3b",
                        "des",
                        "inning",
                        "inning_topbot",
                    ]
                    if c in df.columns
                ]
                missing = {"launch_speed_angle", "zone", "description", "launch_speed"} - set(keep)
                if missing:
                    raise RuntimeError(f"statcast missing columns {sorted(missing)}")
                slim = df[keep].copy()
                slim["game_date"] = pd.to_datetime(slim["game_date"]).dt.strftime("%Y-%m-%d")
                slim.to_parquet(dest, index=False)
                return day, len(slim), None
            except Exception as exc:  # noqa: BLE001
                err = f"{type(exc).__name__}: {exc}"
                time.sleep(1.5 * (attempt + 1))
        return day, 0, err

    failures = []
    done = 0
    with ThreadPoolExecutor(max_workers=4) as pool:
        futs = {pool.submit(one, d): d for d in todo}
        for fut in as_completed(futs):
            day, n, err = fut.result()
            done += 1
            if err:
                failures.append((day, err))
                log(f"statcast FAIL {day} {err} ({done}/{len(todo)})")
            elif done % 10 == 0 or done == len(todo):
                log(f"statcast {done}/{len(todo)} last {day} rows {n}")
    if failures:
        # One sequential retry for the failures.
        still = []
        for day, err in failures:
            day2, n, err2 = one(day)
            if err2:
                still.append((day2, err2))
                log(f"statcast STILL FAIL {day2} {err2}")
            else:
                log(f"statcast recovered {day2} rows {n}")
        if still:
            (CACHE / "statcast_failures.json").write_text(json.dumps(still, indent=2))
            raise SystemExit(f"{len(still)} statcast days failed. See cache/statcast_failures.json")


def extract_plays(game_pk: int, data: dict) -> list[dict]:
    rows = []
    for p in data.get("allPlays") or []:
        about = p.get("about") or {}
        result = p.get("result") or {}
        batter = (p.get("matchup") or {}).get("batter") or {}
        runners = []
        for rn in p.get("runners") or []:
            det = rn.get("details") or {}
            mv = rn.get("movement") or {}
            rid = (det.get("runner") or {}).get("id")
            runners.append(
                {
                    "id": None if rid is None else int(rid),
                    "origin": mv.get("originBase"),
                    "start": mv.get("start"),
                    "end": mv.get("end"),
                    "is_out": bool(mv.get("isOut")),
                }
            )
        rows.append(
            {
                "game_pk": int(game_pk),
                "at_bat_index": int(about["atBatIndex"]),
                "batter_id": None if batter.get("id") is None else int(batter["id"]),
                "event_type": result.get("eventType"),
                "rbi": result.get("rbi"),
                "runners": runners,
            }
        )
    return rows


def download_pbp(games: list[dict]) -> None:
    PBP_DIR.mkdir(parents=True, exist_ok=True)
    todo = []
    for g in games:
        dest = PBP_DIR / f"{g['game_pk']}.json"
        if dest.exists():
            try:
                blob = json.loads(dest.read_text())
                if blob.get("v") == EXTRACT_VERSION and "plays" in blob:
                    continue
            except Exception:
                pass
        todo.append(g)
    log(f"pbp cached {len(games) - len(todo)} / {len(games)}; downloading {len(todo)}")
    if not todo:
        return

    def one(g: dict) -> tuple[int, str | None]:
        game_pk = g["game_pk"]
        dest = PBP_DIR / f"{game_pk}.json"
        err = None
        for attempt in range(4):
            try:
                data = get_json(
                    f"https://statsapi.mlb.com/api/v1/game/{game_pk}/playByPlay",
                    timeout=45,
                )
                plays = extract_plays(game_pk, data)
                dest.write_text(json.dumps({"v": EXTRACT_VERSION, "plays": plays}))
                return game_pk, None
            except Exception as exc:  # noqa: BLE001
                err = f"{type(exc).__name__}: {exc}"
                time.sleep(0.8 * (attempt + 1))
        return game_pk, err

    failures = []
    done = 0
    with ThreadPoolExecutor(max_workers=10) as pool:
        futs = [pool.submit(one, g) for g in todo]
        for fut in as_completed(futs):
            game_pk, err = fut.result()
            done += 1
            if err:
                failures.append((game_pk, err))
            if done % 200 == 0 or done == len(todo):
                log(f"pbp {done}/{len(todo)} failures {len(failures)}")
    if failures:
        still = []
        for game_pk, _err in failures:
            g = {"game_pk": game_pk}
            game_pk2, err2 = one(g)
            if err2:
                still.append((game_pk2, err2))
        if still:
            (CACHE / "pbp_failures.json").write_text(json.dumps(still))
            log(f"pbp still failing: {len(still)}")
        else:
            log("pbp failures recovered on retry")


def load_pitches(dates: list[str]) -> pd.DataFrame:
    frames = []
    for d in dates:
        path = SC_DIR / f"{d}.parquet"
        if not path.exists():
            continue
        frames.append(pd.read_parquet(path))
    if not frames:
        raise SystemExit("no statcast cache")
    df = pd.concat(frames, ignore_index=True)
    df = df[df["game_type"].fillna("R") == "R"].copy()
    df["game_pk"] = pd.to_numeric(df["game_pk"], errors="coerce").astype("int64")
    df["batter"] = pd.to_numeric(df["batter"], errors="coerce").astype("int64")
    df["at_bat_number"] = pd.to_numeric(df["at_bat_number"], errors="coerce")
    df["pitch_number"] = pd.to_numeric(df["pitch_number"], errors="coerce")
    df = df.dropna(subset=["at_bat_number", "pitch_number"])
    df["at_bat_number"] = df["at_bat_number"].astype("int64")
    df["pitch_number"] = df["pitch_number"].astype("int64")
    df["game_date"] = df["game_date"].astype(str).str.slice(0, 10)
    df = df.drop_duplicates(["game_pk", "at_bat_number", "pitch_number", "batter"])
    df["zone_n"] = pd.to_numeric(df["zone"], errors="coerce")
    df["ev"] = pd.to_numeric(df["launch_speed"], errors="coerce")
    df["la"] = pd.to_numeric(df["launch_angle"], errors="coerce")
    df["lsa"] = pd.to_numeric(df["launch_speed_angle"], errors="coerce")
    df["outs0"] = pd.to_numeric(df["outs_when_up"], errors="coerce")
    for col in ("on_1b", "on_2b", "on_3b"):
        df[col] = pd.to_numeric(df[col], errors="coerce")
    df["description"] = df["description"].fillna("").astype(str)
    df["bb_type"] = df["bb_type"].where(df["bb_type"].notna(), None)
    df["events"] = df["events"].where(df["events"].notna(), None)
    df["type"] = df["type"].fillna("").astype(str)
    return df


def load_api_index(game_pks: list[int]) -> dict[tuple[int, int], dict]:
    grouped: dict[tuple[int, int], list[dict]] = {}
    missing = 0
    for pk in game_pks:
        path = PBP_DIR / f"{pk}.json"
        if not path.exists():
            missing += 1
            continue
        blob = json.loads(path.read_text())
        for play in blob.get("plays") or []:
            key = (int(play["game_pk"]), int(play["at_bat_index"]) + 1)
            grouped.setdefault(key, []).append(play)
    merged = {}
    for key, plays in grouped.items():
        pa_plays = [p for p in plays if p.get("event_type") in PA_EVENTS]
        primary = pa_plays[-1] if pa_plays else plays[-1]
        runners = []
        for p in plays:
            runners.extend(p.get("runners") or [])
        merged[key] = {
            "batter_id": primary.get("batter_id"),
            "event_type": primary.get("event_type"),
            "rbi": primary.get("rbi"),
            "runners": runners,
        }
    log(f"api index plays {len(merged):,} missing pbp files {missing}")
    return merged


def runner_advanced(on_2b, runners: list[dict]) -> bool:
    """End state of the runner who started on second: safe at 3B or home."""
    if on_2b is None or (isinstance(on_2b, float) and math.isnan(on_2b)):
        return False
    rid = int(on_2b)
    base = "2B"
    seen = False
    for mv in runners:
        if mv.get("id") != rid:
            continue
        seen = True
        if mv.get("is_out"):
            return False
        end = mv.get("end")
        if end:
            base = end
    if not seen:
        return False
    return base in SAFE_ADVANCE


def build_pa_table(pitches: pd.DataFrame, api: dict) -> tuple[pd.DataFrame, dict]:
    ordered = pitches.sort_values(["game_pk", "at_bat_number", "pitch_number"])
    counts = (
        ordered.groupby(["game_pk", "at_bat_number"], sort=False)
        .size()
        .rename("n_pitches")
        .reset_index()
    )
    first = ordered.drop_duplicates(["game_pk", "at_bat_number"], keep="first")[
        ["game_pk", "at_bat_number", "outs0", "on_1b", "on_2b", "on_3b"]
    ]
    last = ordered.drop_duplicates(["game_pk", "at_bat_number"], keep="last")[
        [
            "game_pk",
            "at_bat_number",
            "game_date",
            "batter",
            "events",
        ]
    ]
    pa = last.merge(first, on=["game_pk", "at_bat_number"], how="left").merge(
        counts, on=["game_pk", "at_bat_number"], how="left"
    )
    event_counts = pa["events"].fillna("(none)").value_counts().to_dict()
    unknown = sorted(e for e in event_counts if e not in PA_EVENTS and e != "(none)")
    pa = pa[pa["events"].isin(PA_EVENTS)].copy()

    rbi = []
    matched = []
    batter_mismatch = 0
    move_success = []
    move_opp = []
    risp = []
    for rec in pa.itertuples(index=False):
        key = (int(rec.game_pk), int(rec.at_bat_number))
        play = api.get(key)
        ok = play is not None and play.get("batter_id") == int(rec.batter)
        if play is not None and play.get("batter_id") != int(rec.batter):
            batter_mismatch += 1
            ok = False
        matched.append(ok)
        rbi_val = play.get("rbi") if ok else None
        if rbi_val is not None:
            rbi_val = int(rbi_val)
        rbi.append(rbi_val)
        on2 = rec.on_2b
        on3 = rec.on_3b
        outs = rec.outs0
        has2 = pd.notna(on2)
        has3 = pd.notna(on3)
        is_risp = bool(has2 or has3)
        risp.append(is_risp)
        is_move = bool(has2 and not has3 and pd.notna(outs) and int(outs) == 0)
        move_opp.append(is_move)
        if ok and is_move:
            move_success.append(bool(runner_advanced(on2, play.get("runners") or [])))
        else:
            move_success.append(pd.NA)
    pa["api_ok"] = matched
    pa["rbi"] = pd.array(rbi, dtype="Int64")
    pa["is_risp"] = risp
    pa["is_move"] = move_opp
    pa["move_success"] = pd.array(move_success, dtype="boolean")
    pa["qab"] = pa["n_pitches"] >= 6
    info = {
        "pa_rows": int(len(pa)),
        "api_matched": int(sum(matched)),
        "api_match_rate": (sum(matched) / len(pa)) if len(pa) else 0,
        "batter_mismatch": batter_mismatch,
        "unknown_events": {e: int(event_counts[e]) for e in unknown},
        "event_counts_head": {k: int(v) for k, v in list(event_counts.items())[:30]},
    }
    log(
        f"PAs {info['pa_rows']:,} api match {info['api_match_rate']:.3%} "
        f"mismatches {batter_mismatch}"
    )
    if unknown:
        log("non-PA terminal events (excluded): " + json.dumps(info["unknown_events"]))
    return pa, info


def _rate(num: int, den: int) -> float | None:
    if den <= 0:
        return None
    return num / den


def pitch_masks(p: pd.DataFrame) -> dict:
    not_bunt = ~p["description"].isin(BUNT_DESC)
    in_zone = p["zone_n"].between(1, 9, inclusive="both")
    out_zone = p["zone_n"].isin([11, 12, 13, 14])
    swing = p["description"].isin(SWING_DESC) & not_bunt
    whiff = p["description"].isin(WHIFF_DESC) & not_bunt
    tracked = (p["type"] == "X") & p["ev"].notna() & p["la"].notna()
    bunt_bb = tracked & (p["bb_type"] == "bunt")
    competitive = tracked & ~bunt_bb
    return {
        "in_zone": in_zone,
        "out_zone": out_zone,
        "swing": swing,
        "whiff": whiff,
        "not_bunt": not_bunt,
        "tracked": tracked,
        "competitive": competitive,
        "barrel": tracked & (p["lsa"] == 6),
        "damage": tracked & (p["ev"] >= 95) & (p["la"] >= 10),
    }


def summarize(pitches: pd.DataFrame, pas: pd.DataFrame) -> dict:
    """Opportunity-weighted rates. None when the denominator is zero."""
    m = pitch_masks(pitches) if len(pitches) else None
    if m is None:
        bb = dmg = brl = zsw = zwh = oz = ozs = 0
        ev_sum = 0.0
        comp_n = 0
    else:
        bb = int(m["tracked"].sum())
        brl = int(m["barrel"].sum())
        dmg = int(m["damage"].sum())
        zsw = int((m["swing"] & m["in_zone"]).sum())
        zwh = int((m["whiff"] & m["in_zone"]).sum())
        oz_pitches = m["out_zone"] & m["not_bunt"]
        oz = int(oz_pitches.sum())
        ozs = int((m["swing"] & oz_pitches).sum())
        comp = pitches.loc[m["competitive"], "ev"]
        comp_n = int(comp.shape[0])
        ev_sum = float(comp.sum()) if comp_n else 0.0

    risp_pas = pas[pas["is_risp"] & pas["api_ok"] & pas["rbi"].notna()] if len(pas) else pas
    move_pas = pas[pas["is_move"] & pas["api_ok"] & pas["move_success"].notna()] if len(pas) else pas
    risp_n = int(len(risp_pas))
    risp_k = int((risp_pas["rbi"] >= 1).sum()) if risp_n else 0
    move_n = int(len(move_pas))
    move_k = int(move_pas["move_success"].sum()) if move_n else 0
    pa_n = int(len(pas))
    qab_k = int(pas["qab"].sum()) if pa_n else 0
    games = int(pas["game_pk"].nunique()) if pa_n else 0
    if pa_n:
        d0 = str(pas["game_date"].min())
        d1 = str(pas["game_date"].max())
    else:
        d0 = d1 = None
    return {
        "pa": pa_n,
        "games": games,
        "date_start": d0,
        "date_end": d1,
        "bb": bb,
        "barrel": _rate(brl, bb),
        "barrel_n": brl,
        "damage": _rate(dmg, bb),
        "damage_n": dmg,
        "whiff": _rate(zwh, zsw),
        "whiff_n": zsw,
        "whiff_k": zwh,
        "chase": _rate(ozs, oz),
        "chase_n": oz,
        "chase_k": ozs,
        "risp": _rate(risp_k, risp_n),
        "risp_n": risp_n,
        "risp_k": risp_k,
        "move": _rate(move_k, move_n),
        "move_n": move_n,
        "move_k": move_k,
        "qab": _rate(qab_k, pa_n),
        "qab_n": pa_n,
        "qab_k": qab_k,
        "ev": (ev_sum / comp_n) if comp_n else None,
        "ev_n": comp_n,
    }


def composite(summary: dict, league: dict) -> float | None:
    parts = []
    for key in ("risp", "move", "qab"):
        r = summary[key]
        lg = league[key]
        if r is None or lg is None or lg == 0:
            return None
        parts.append(r / lg)
    return 100.0 * sum(parts) / 3.0


def tone_for(rate: float | None, league: float | None, higher_better: bool, small: bool) -> str:
    if rate is None or league is None or league == 0:
        return "blank"
    if small:
        return "small"
    ratio = rate / league
    if (1 - BAND) <= ratio <= (1 + BAND):
        return "white"
    better = ratio > 1 if higher_better else ratio < 1
    return "green" if better else "red"


def tone_composite(value: float | None, small: bool) -> str:
    if value is None:
        return "blank"
    if small:
        return "small"
    # 100 is league average; within 10% is 90–110.
    if (1 - BAND) * 100 <= value <= (1 + BAND) * 100:
        return "white"
    return "green" if value > 100 else "red"


def last_name(name: str) -> str:
    parts = name.split()
    if len(parts) >= 2 and parts[-1].rstrip(".").upper() in {"JR", "SR", "II", "III", "IV", "V"}:
        return parts[-2]
    return parts[-1]


def pct_text(rate: float | None) -> str | None:
    if rate is None:
        return None
    return f"{rate * 100:.1f}%"


def num_text(value: float | None, digits: int = 1) -> str | None:
    if value is None:
        return None
    return f"{value:.{digits}f}"


def build_rows(hitters: list[dict], pitches: pd.DataFrame, pas: pd.DataFrame, league_season: dict, month_leagues: dict, archetypes: dict) -> list[dict]:
    rows = []
    for h in hitters:
        pid = h["id"]
        arch = archetypes[pid]
        metric = arch["archetype"]  # Barrel, Damage, or None
        p_player = pitches[pitches["batter"] == pid]
        a_player = pas[pas["batter"] == pid]
        games = (
            a_player.groupby("game_pk", as_index=False)["game_date"]
            .min()
            .sort_values(["game_date", "game_pk"])
        )
        last_games = games.tail(10)
        last_pks = set(int(x) for x in last_games["game_pk"])
        windows = []
        for key, label, start, end, _fname in MONTH_WINDOWS:
            windows.append(
                (
                    key,
                    label,
                    p_player[p_player["game_date"].between(start, end)],
                    a_player[a_player["game_date"].between(start, end)],
                    month_leagues[key],
                    key,
                )
            )
        windows.append(
            (
                "last10",
                "Last 10 games" if len(last_games) == 0 else f"Last {len(last_games)} game" + ("s" if len(last_games) != 1 else ""),
                p_player[p_player["game_pk"].isin(last_pks)],
                a_player[a_player["game_pk"].isin(last_pks)],
                league_season,
                "season",
            )
        )
        for key, label, pp, aa, league, league_key in windows:
            summary = summarize(pp, aa)
            comp = composite(summary, league)
            contact_rate = None
            contact_n = summary["bb"]
            if metric == "Barrel":
                contact_rate = summary["barrel"]
                contact_lg = league["barrel"]
            elif metric == "Damage":
                contact_rate = summary["damage"]
                contact_lg = league["damage"]
            else:
                contact_lg = None
            contact_small = contact_n < SMALL_BB
            whiff_small = summary["whiff_n"] < SMALL_N
            chase_small = summary["chase_n"] < SMALL_N
            risp_small = summary["risp_n"] < SMALL_N
            move_small = summary["move_n"] < SMALL_N
            qab_small = summary["pa"] < SMALL_N
            comp_small = (
                summary["pa"] < SMALL_N
                or summary["risp_n"] < SMALL_N
                or summary["move_n"] < SMALL_N
            )
            # Undefined inputs are not "small"; the composite stays blank.
            if comp is None:
                comp_small = False
            small_bits = []
            if metric and contact_rate is not None and contact_small:
                small_bits.append("contact")
            if summary["whiff"] is not None and whiff_small:
                small_bits.append("in-zone whiff")
            if summary["chase"] is not None and chase_small:
                small_bits.append("chase")
            if summary["risp"] is not None and risp_small:
                small_bits.append("RISP knock-in")
            if summary["move"] is not None and move_small:
                small_bits.append("move-over")
            if summary["qab"] is not None and qab_small:
                small_bits.append("QAB")
            if comp is not None and comp_small:
                small_bits.append("execution")
            undefined = []
            if metric is None:
                undefined.append("contact (archetype not classified)")
            elif contact_rate is None:
                undefined.append("contact (0 tracked batted balls)")
            if summary["whiff"] is None:
                undefined.append("in-zone whiff (0 in-zone swings)")
            if summary["chase"] is None:
                undefined.append("chase (0 out-of-zone pitches)")
            if summary["risp"] is None:
                undefined.append("RISP knock-in (0 RISP PA with an RBI tag)")
            if summary["move"] is None:
                undefined.append("move-over (0 eligible PA)")
            if summary["qab"] is None:
                undefined.append("QAB (0 PA)")
            if comp is None:
                undefined.append("execution (a component rate is undefined)")
            no_pa = summary["pa"] == 0
            if no_pa:
                contact_rate = None
                comp = None
                small_bits = []
                undefined = ["no plate appearances in this window"]
            rows.append(
                {
                    "id": pid,
                    "name": h["name"],
                    "pos": h["pos"],
                    "roster_days": h.get("roster_days"),
                    "roster_first": h.get("roster_first"),
                    "roster_last": h.get("roster_last"),
                    "archetype": metric,
                    "archetype_mark": "1" if pid in EXAMPLE_IDS else None,
                    "archetype_reason": arch.get("rule_reason") or "",
                    "window_key": key,
                    "window": label,
                    "no_pa": no_pa,
                    "games": None if no_pa else summary["games"],
                    "games_available": None if no_pa else (int(len(games)) if key == "last10" else summary["games"]),
                    "date_start": None if no_pa else summary["date_start"],
                    "date_end": None if no_pa else summary["date_end"],
                    "league_key": league_key,
                    "contact_metric": metric,
                    "contact": None if no_pa else contact_rate,
                    "contact_n": None if no_pa else contact_n,
                    "contact_text": None if no_pa else pct_text(contact_rate),
                    "contact_tone": "blank" if no_pa else tone_for(contact_rate, contact_lg, True, contact_small and contact_rate is not None),
                    "whiff": None if no_pa else summary["whiff"],
                    "whiff_n": None if no_pa else summary["whiff_n"],
                    "whiff_text": None if no_pa else pct_text(summary["whiff"]),
                    "whiff_tone": "blank" if no_pa else tone_for(summary["whiff"], league["whiff"], False, whiff_small and summary["whiff"] is not None),
                    "chase": None if no_pa else summary["chase"],
                    "chase_n": None if no_pa else summary["chase_n"],
                    "chase_text": None if no_pa else pct_text(summary["chase"]),
                    "chase_tone": "blank" if no_pa else tone_for(summary["chase"], league["chase"], False, chase_small and summary["chase"] is not None),
                    "execution": None if no_pa or comp is None else round(comp, 1),
                    "execution_text": None if no_pa or comp is None else f"{comp:.1f}",
                    "execution_tone": "blank" if no_pa else tone_composite(comp, comp_small),
                    "risp": None if no_pa else summary["risp"],
                    "risp_n": None if no_pa else summary["risp_n"],
                    "risp_text": None if no_pa else pct_text(summary["risp"]),
                    "risp_tone": "blank" if no_pa else tone_for(summary["risp"], league["risp"], True, risp_small and summary["risp"] is not None),
                    "move": None if no_pa else summary["move"],
                    "move_n": None if no_pa else summary["move_n"],
                    "move_text": None if no_pa else pct_text(summary["move"]),
                    "move_tone": "blank" if no_pa else tone_for(summary["move"], league["move"], True, move_small and summary["move"] is not None),
                    "qab": None if no_pa else summary["qab"],
                    "qab_n": None if no_pa else summary["pa"],
                    "qab_text": None if no_pa else pct_text(summary["qab"]),
                    "qab_tone": "blank" if no_pa else tone_for(summary["qab"], league["qab"], True, qab_small and summary["qab"] is not None),
                    "pa": None if no_pa else summary["pa"],
                    "small_notes": ", ".join(small_bits),
                    "undefined_notes": "; ".join(undefined),
                }
            )
    return rows


def fit_barrel_model(pitches: pd.DataFrame) -> dict:
    """OLS of Statcast barrel% on average EV, fit to qualified MLB hitters."""
    masks = pitch_masks(pitches)
    comp = pitches.loc[masks["competitive"], ["batter", "ev"]]
    tracked = pitches.loc[masks["tracked"], ["batter", "lsa"]]
    ev = comp.groupby("batter")["ev"].agg(ev="mean", ev_n="size")
    br = tracked.groupby("batter").agg(bb=("lsa", "size"), barrels=("lsa", lambda s: int((s == 6).sum())))
    g = ev.join(br, how="inner")
    g = g[g["ev_n"] >= MODEL_MIN_BB]
    x = g["ev"].to_numpy(dtype=float)
    y = (g["barrels"] / g["bb"]).to_numpy(dtype=float)
    slope, intercept = np.polyfit(x, y, 1)
    pred = intercept + slope * x
    ss_res = float(((y - pred) ** 2).sum())
    ss_tot = float(((y - y.mean()) ** 2).sum())
    r2 = 1 - ss_res / ss_tot if ss_tot else None
    return {
        "intercept": float(intercept),
        "slope": float(slope),
        "r2": None if r2 is None else float(r2),
        "n_hitters": int(len(g)),
        "min_bb": MODEL_MIN_BB,
        "residual_gap": RESIDUAL_GAP,
    }


def classify(hitters, pitches, league_season) -> dict:
    """High exit velocity is Damage. Average exit velocity is Barrel."""
    out = {}
    league_ev = league_season.get("ev")
    line = None if league_ev is None else league_ev + EV_CLEAR_GAP
    season_p = pitches
    for h in hitters:
        p = season_p[season_p["batter"] == h["id"]]
        m = pitch_masks(p) if len(p) else None
        if m is None:
            ev = None
            ev_n = 0
            barrel = None
            bb = 0
        else:
            comp = p.loc[m["competitive"], "ev"]
            ev_n = int(comp.shape[0])
            ev = float(comp.mean()) if ev_n else None
            bb = int(m["tracked"].sum())
            barrel = _rate(int(m["barrel"].sum()), bb)
        rule = None
        reason = ""
        if ev is None or ev_n < ARCHETYPE_MIN_BB or barrel is None or line is None:
            reason = (
                f"Fewer than {ARCHETYPE_MIN_BB} competitive batted balls (n={ev_n}), "
                "so the contact cell stays blank."
            )
        elif ev >= line:
            rule = "Damage"
            reason = (
                f"Avg EV {ev:.1f} is at least {EV_CLEAR_GAP:.1f} mph above the 2026 MLB average "
                f"{league_ev:.1f}, so he is a high-EV hitter graded on damage rate."
            )
        else:
            rule = "Barrel"
            reason = (
                f"Avg EV {ev:.1f} is below the high-EV line of {line:.1f} "
                f"(league {league_ev:.1f}), so he is an average-EV hitter graded on barrel rate."
            )
        final = rule
        override = h["id"] in OVERRIDE_ARCHETYPE
        if override:
            final = OVERRIDE_ARCHETYPE[h["id"]]
        if h["id"] == LEE_ID and ev is not None:
            reason = (
                f"Avg EV {ev:.1f} mph is near the league average {league_ev:.1f}. "
                "Damage because he is a power hitter, the damage example."
            )
        elif h["id"] == BUXTON_ID and ev is not None and league_ev is not None:
            reason = (
                f"Avg EV {ev:.1f} mph is {ev - league_ev:+.1f} versus the league average {league_ev:.1f}. "
                "Damage because he is a power / high-EV hitter."
            )
        elif h["id"] == KEASCHALL_ID and ev is not None and league_ev is not None:
            reason = (
                f"Avg EV {ev:.1f} mph is below the league average {league_ev:.1f}. "
                "Barrel because he is an average-EV hitter, the barrel example."
            )
        if override and rule == final:
            note = "Named example. The exit-velocity rule assigns the same label."
        elif override:
            note = (
                f"Named example set to {final}. "
                f"The exit-velocity rule assigned {rule or 'unclassified'}."
            )
        else:
            note = ""
        out[h["id"]] = {
            "id": h["id"],
            "name": h["name"],
            "pos": h["pos"],
            "ev": ev,
            "ev_n": ev_n,
            "bb": bb,
            "barrel": barrel,
            "pred_barrel": None,
            "residual": None,
            "rule": rule,
            "rule_reason": reason,
            "override": override,
            "archetype": final,
            "override_note": note,
        }
    return out


def json_safe(obj):
    if isinstance(obj, dict):
        return {k: json_safe(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [json_safe(v) for v in obj]
    if obj is None or isinstance(obj, (str, int, bool)):
        return obj
    if isinstance(obj, float):
        if math.isnan(obj) or math.isinf(obj):
            return None
        return obj
    return obj


def write_qa(payload: dict) -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    (OUT / "qa.json").write_text(json.dumps(json_safe(payload), indent=2))


def apply_rate_cell(cell, value, tone: str, kind: str) -> None:
    cell.fill = FILL[tone]
    cell.font = FONT_CELL[tone]
    cell.border = THIN
    cell.alignment = Alignment(vertical="center", horizontal="center")
    if value is None:
        cell.value = None
        return
    if kind == "pct":
        cell.value = value
        cell.number_format = "0.0%"
    else:
        cell.value = value
        cell.number_format = "0.0"


def write_workbook(payload: dict) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Tracking"

    subtitle = payload["subtitle"]
    ws["A1"] = "Minnesota Twins hitters"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:W1")
    ws["A2"] = subtitle
    ws["A2"].font = SUB_FONT
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.merge_cells("A2:W2")
    ws.row_dimensions[2].height = 48
    ws["A3"] = (
        "Color: green is better than the comparison league rate, white is within 10% of that rate "
        "(inclusive), red is worse. Higher is better for contact quality, execution, RISP knock-in, "
        "move-over, and QAB. Lower is better for in-zone whiff and chase. Gray italic is a small sample "
        "and is not green or red (under 8 tracked batted balls, or under 10 swings / out-of-zone pitches / "
        "plate appearances, whichever denominator the column uses). A blank cell could not be computed. "
        "Each month uses that month’s 2026 MLB rates. Last-10 rows use full-season 2026 MLB rates. "
        "See Notes for footnotes 1 and 2."
    )
    ws["A3"].font = Font(name="Calibri", size=10, color="374151")
    ws["A3"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.merge_cells("A3:W3")
    ws.row_dimensions[3].height = 48

    headers = [
        "Hitter",
        "Pos",
        "Archetype",
        "Window",
        "Games",
        "First game",
        "Last game",
        "Contact metric",
        "Contact rate",
        "Tracked batted balls",
        "In-zone whiff",
        "In-zone swings",
        "Chase rate",
        "Out-of-zone pitches",
        "Overall Execution",
        "RISP knock-in",
        "RISP PA",
        "Move-over",
        "Move-over PA",
        "QAB rate",
        "PA",
        "Small-sample columns",
        "Undefined columns",
        "League set",
        "MLBAM id",
    ]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(5, col, header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = WRAP
        cell.border = Border(bottom=Side(style="medium", color=GOLD))
    ws.row_dimensions[5].height = 32
    ws.auto_filter.ref = f"A5:Y{5 + len(payload['rows'])}"
    ws.freeze_panes = "D6"
    ws.auto_filter.ref = f"A5:Y{5 + len(payload['rows'])}"

    rate_cols = {
        9: ("contact", "pct"),
        11: ("whiff", "pct"),
        13: ("chase", "pct"),
        15: ("execution", "idx"),
        16: ("risp", "pct"),
        18: ("move", "pct"),
        20: ("qab", "pct"),
    }
    for i, row in enumerate(payload["rows"]):
        r = 6 + i
        arch = row["archetype"] or ""
        if row["archetype_mark"]:
            arch = f"{arch}¹"
        values = [
            row["name"],
            row["pos"],
            arch,
            row["window"],
            row["games"],
            row["date_start"],
            row["date_end"],
            row["contact_metric"],
            row["contact"],
            row["contact_n"],
            row["whiff"],
            row["whiff_n"],
            row["chase"],
            row["chase_n"],
            row["execution"],
            row["risp"],
            row["risp_n"],
            row["move"],
            row["move_n"],
            row["qab"],
            row["pa"],
            row["small_notes"],
            row["undefined_notes"],
            row["league_key"],
            row["id"],
        ]
        for c, val in enumerate(values, 1):
            cell = ws.cell(r, c, None if val is None else val)
            cell.border = THIN
            cell.font = Font(name="Calibri", size=11, color="111111")
            if c in (1, 4, 22, 23):
                cell.alignment = LEFT
            else:
                cell.alignment = Alignment(vertical="center", horizontal="center")
        for col, (key, kind) in rate_cols.items():
            apply_rate_cell(ws.cell(r, col), row[key], row[f"{key}_tone"], kind)
        ws.cell(r, 10).number_format = "#,##0"
        ws.cell(r, 12).number_format = "#,##0"
        ws.cell(r, 14).number_format = "#,##0"
        ws.cell(r, 17).number_format = "#,##0"
        ws.cell(r, 19).number_format = "#,##0"
        ws.cell(r, 21).number_format = "#,##0"
        ws.cell(r, 5).number_format = "0"
        if row["archetype_mark"]:
            ws.cell(r, 3).comment = Comment(
                "Footnote 1: Luke Keaschall is Barrel by user override. The rule is on the Notes sheet.",
                "Twins sheet",
                width=240,
                height=70,
            )
        if i % 2 == 0:
            ws.row_dimensions[r].height = 18
    ws.cell(5, 3).comment = Comment("Footnote 1 on the Notes sheet. Superscript 1 marks Luke Keaschall.", "Twins sheet", width=220, height=50)
    ws.cell(5, 15).comment = Comment("Footnote 2 on the Notes sheet. 100 = league average.", "Twins sheet", width=220, height=50)

    widths = {
        1: 22, 2: 8, 3: 14, 4: 16, 5: 10, 6: 14, 7: 14, 8: 16,
        9: 14, 10: 16, 11: 14, 12: 14, 13: 12, 14: 16, 15: 18,
        16: 14, 17: 12, 18: 12, 19: 14, 20: 12, 21: 10, 22: 28, 23: 36, 24: 14, 25: 12,
    }
    for c, w in widths.items():
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.paperSize = ws.PAPERSIZE_TABLOID
    ws.sheet_properties.tabColor = NAVY
    ws.oddHeader.left.text = "Twins hitters · 2026-09-30 active roster"
    ws.print_title_rows = "1:5"
    ws.page_setup.horizontalCentered = True
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 110
    ws.auto_filter.ref = f"A5:Y{5 + len(payload['rows'])}"

    # League averages
    lg = wb.create_sheet("League averages")
    lg["A1"] = "2026 MLB pooled rates used to index this sheet"
    lg["A1"].font = TITLE_FONT
    lg.merge_cells("A1:F1")
    lg["A2"] = (
        "Pooled means every opportunity is weighted equally (total successes / total chances), "
        "not the average of player rates. Regular-season games only, through September 30, 2026. "
        "There were no regular-season games on September 28–30."
    )
    lg["A2"].alignment = Alignment(wrap_text=True)
    lg.merge_cells("A2:F2")
    lg_headers = ["Metric", "Season rate", "Season opportunities", "September rate", "September opportunities", "Better"]
    for c, h in enumerate(lg_headers, 1):
        cell = lg.cell(4, c, h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = WRAP
    season = payload["league"]["season"]
    sept = payload["league"]["september"]
    lg_rows = [
        ("Barrel% (Statcast launch_speed_angle = 6)", "barrel", "bb", "higher"),
        ("Damage rate (EV ≥ 95 mph and LA ≥ 10°)", "damage", "bb", "higher"),
        ("In-zone whiff (whiffs / in-zone swings)", "whiff", "whiff_n", "lower"),
        ("Chase rate (out-of-zone swings / out-of-zone pitches)", "chase", "chase_n", "lower"),
        ("RISP knock-in (RBI PA / RISP PA)", "risp", "risp_n", "higher"),
        ("Move-over (safe 2nd-to-3rd or score / eligible PA)", "move", "move_n", "higher"),
        ("QAB (PA with 6+ pitches / PA)", "qab", "pa", "higher"),
        ("Avg EV on competitive contact (mph)", "ev", "ev_n", "context"),
    ]
    for i, (label, key, nkey, direction) in enumerate(lg_rows):
        r = 5 + i
        lg.cell(r, 1, label).alignment = LEFT
        s_rate, m_rate = season[key], sept[key]
        if key == "ev":
            lg.cell(r, 2, None if s_rate is None else round(s_rate, 2)).number_format = "0.00"
            lg.cell(r, 4, None if m_rate is None else round(m_rate, 2)).number_format = "0.00"
        else:
            lg.cell(r, 2, s_rate).number_format = "0.00%"
            lg.cell(r, 4, m_rate).number_format = "0.00%"
        lg.cell(r, 3, season[nkey]).number_format = "#,##0"
        lg.cell(r, 5, sept[nkey]).number_format = "#,##0"
        lg.cell(r, 6, direction)
        for c in range(1, 7):
            lg.cell(r, c).font = Font(name="Calibri", size=11)
            lg.cell(r, c).border = THIN
    lg.column_dimensions["A"].width = 62
    for col in "BCDEF":
        lg.column_dimensions[col].width = 24
    lg.sheet_properties.tabColor = NAVY
    lg.freeze_panes = "A5"
    lg.sheet_view.showGridLines = False

    # Archetypes
    ar = wb.create_sheet("Archetypes")
    ar["A1"] = "2026 season archetype assignment"
    ar["A1"].font = TITLE_FONT
    ar.merge_cells("A1:N1")
    ar["A2"] = payload["footnote1"]
    ar["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ar.merge_cells("A2:N2")
    ar.row_dimensions[2].height = 120
    ar_headers = [
        "Hitter", "Pos", "MLBAM id", "Competitive batted balls", "Avg EV",
        "League EV", "EV minus league", "Barrel%", "League barrel%",
        "EV-predicted barrel%", "Barrel residual",
        "Rule result", "Final archetype", "Why",
    ]
    for c, h in enumerate(ar_headers, 1):
        cell = ar.cell(4, c, h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = WRAP
    for i, a in enumerate(payload["archetypes"]):
        r = 5 + i
        ev_gap = None if a["ev"] is None or payload["league"]["season"]["ev"] is None else a["ev"] - payload["league"]["season"]["ev"]
        vals = [
            a["name"], a["pos"], a["id"], a["ev_n"], a["ev"],
            payload["league"]["season"]["ev"], ev_gap, a["barrel"],
            payload["league"]["season"]["barrel"], a.get("pred_barrel"), a.get("residual"),
            a["rule"], a["archetype"],
            (a["override_note"] + " " + a["rule_reason"]).strip(),
        ]
        for c, v in enumerate(vals, 1):
            cell = ar.cell(r, c, v)
            cell.font = Font(name="Calibri", size=11)
            cell.border = THIN
            cell.alignment = LEFT if c in (1, 14) else Alignment(vertical="center", horizontal="center")
        ar.cell(r, 4).number_format = "#,##0"
        ar.cell(r, 5).number_format = "0.00"
        ar.cell(r, 6).number_format = "0.00"
        ar.cell(r, 7).number_format = "0.00"
        ar.cell(r, 8).number_format = "0.0%"
        ar.cell(r, 9).number_format = "0.0%"
        ar.cell(r, 10).number_format = "0.0%"
        ar.cell(r, 11).number_format = "+0.0%;-0.0%;0.0%"
        if a["id"] in EXAMPLE_IDS and a["archetype"]:
            ar.cell(r, 13, f"{a['archetype']}¹")
    for c, w in enumerate([22, 8, 12, 18, 12, 12, 16, 12, 16, 18, 14, 14, 16, 78], 1):
        ar.column_dimensions[get_column_letter(c)].width = w
    ar.freeze_panes = "B5"
    ar.row_dimensions[4].height = 30
    ar.sheet_properties.tabColor = NAVY
    ar.sheet_view.showGridLines = False

    # Data dictionary
    dd = wb.create_sheet("Data dictionary")
    dd["A1"] = "Data dictionary"
    dd["A1"].font = TITLE_FONT
    dd_headers = ["Field", "Definition", "Numerator", "Denominator", "Source", "Direction"]
    for c, h in enumerate(dd_headers, 1):
        cell = dd.cell(3, c, h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = WRAP
    dictionary = [
        ("Active roster", "Players returned by the MLB Stats API active roster for Minnesota on 2026-09-30.", "—", "—", "MLB Stats API teams/142/roster rosterType=active", "—"),
        ("Hitter", "Position players on that roster. Pitchers (position type Pitcher) are excluded.", "—", "—", "MLB Stats API roster", "—"),
        ("September row", "Regular-season games with game_date from 2026-09-01 through 2026-09-30. The last games were 2026-09-27.", "—", "—", "Stats API schedule gameType=R", "—"),
        ("Last-10 row", "The hitter's most recent regular-season games on or before 2026-09-30 in which he had at least one plate appearance, up to 10. Fewer games are used when that is all he played, and the row label shows the count.", "—", "—", "Statcast plate appearances", "—"),
        ("Barrel rate", "Statcast barrel classification. A batted ball is a barrel when launch_speed_angle equals 6.", "Barrels", "Tracked fair batted balls: Statcast type X with non-null exit velocity and launch angle, including bunts", "Baseball Savant Statcast via pybaseball", "Higher"),
        ("Damage rate", "Share of tracked fair batted balls with exit velocity of at least 95 mph and launch angle of at least 10 degrees.", "Batted balls with EV ≥ 95 and LA ≥ 10", "Same tracked fair batted balls as barrel rate", "Statcast launch_speed and launch_angle", "Higher"),
        ("Archetype", "Chooses Barrel rate or Damage rate for both rows. See footnote 1.", "—", "—", "2026 regular-season Statcast, plus two user overrides", "—"),
        ("Competitive contact", "Tracked fair batted balls excluding bb_type = bunt. Used only for average exit velocity in the archetype rule.", "—", "Competitive batted balls", "Statcast", "—"),
        ("In-zone whiff", "Swinging strikes divided by swings, on pitches in Statcast zones 1–9. Swings are swinging_strike, swinging_strike_blocked, foul, foul_tip, and hit_into_play. Bunt attempts (foul_bunt, missed_bunt, bunt_foul_tip) are removed from both the swing count and the zone counts.", "In-zone whiffs", "In-zone swings", "Statcast description and zone", "Lower"),
        ("Chase rate", "Swings at pitches in Statcast zones 11–14, divided by pitches in those zones. The same swing and bunt exclusions as in-zone whiff.", "Out-of-zone swings", "Out-of-zone pitches, bunt attempts excluded", "Statcast description and zone", "Lower"),
        ("RISP knock-in", "Plate appearances that begin with a runner on second and/or third (state on the first tracked pitch) and on which the batter is credited with at least one RBI.", "RISP PA with result.rbi ≥ 1", "RISP PA whose play-by-play RBI field matched this batter", "Statcast base state; MLB Stats API result.rbi", "Higher"),
        ("Move-over", "See footnote 2 for the exact success rule.", "Eligible PA where the runner who started on second ends safe at third or home", "PA that begin with a runner on second, third empty, and 0 outs", "Statcast base-out state; MLB Stats API runner movement", "Higher"),
        ("QAB rate", "Plate appearances that last more than 5 pitches, meaning 6 or more Statcast pitch rows in the plate appearance.", "PA with pitch count ≥ 6", "Plate appearances", "Statcast pitch rows", "Higher"),
        ("Plate appearance", "A Statcast at-bat whose final pitch event is a completed batter result (hit, out, walk, HBP, sacrifice, error, catcher's interference, and the double-play and triple-play forms of those). Baserunning-only events such as stolen bases, wild pitches, and pickoffs are not plate appearances.", "—", "—", "Statcast events", "—"),
        ("Overall Execution", "100 × ((RISP / lg RISP) + (move-over / lg move-over) + (QAB / lg QAB)) / 3. Equal weights. 100 = league average. Blank if any input rate is undefined.", "—", "—", "The three rates and the league rates on the League averages sheet", "Higher"),
        ("League set", "september = September 1–30, 2026 MLB pooled rates. season = all 2026 regular-season games on the schedule through September 30. Last-10 rows use the season set because each hitter's 10 games cover different dates.", "—", "—", "Computed in this workbook", "—"),
        ("Small sample", "Contact quality: under 8 tracked batted balls. Whiff: under 10 in-zone swings. Chase: under 10 out-of-zone pitches. RISP, move-over, QAB, and execution: under 10 opportunities (for execution, any of the three denominators under 10, or under 10 PA). The number is shown and the cell is not colored.", "—", "—", "—", "—"),
        ("White band", "Within 10% of the comparison league rate, inclusive: 0.90 ≤ rate/league ≤ 1.10. For the composite, 90 through 110.", "—", "—", "—", "—"),
    ]
    for i, rec in enumerate(dictionary):
        for c, v in enumerate(rec, 1):
            cell = dd.cell(4 + i, c, v)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.font = Font(name="Calibri", size=10)
            cell.border = THIN
        dd.row_dimensions[4 + i].height = 48
    for c, w in enumerate([24, 72, 42, 42, 42, 14], 1):
        dd.column_dimensions[get_column_letter(c)].width = w
    dd.freeze_panes = "A4"
    dd.sheet_properties.tabColor = NAVY
    dd.sheet_view.showGridLines = False
    dd.row_dimensions[3].height = 22

    notes = wb.create_sheet("Notes")
    notes["A1"] = "Notes"
    notes["A1"].font = TITLE_FONT
    notes.column_dimensions["A"].width = 28
    notes.column_dimensions["B"].width = 140
    blocks = [
        ("Roster", payload["roster_note"]),
        ("Pitchers excluded", payload["pitchers_excluded"]),
        ("Hitters included", payload["hitters_included"]),
        ("Transactions after the last game", payload["transactions_note"]),
        ("Footnote 1 — archetype", payload["footnote1"]),
        ("Footnote 2 — execution composite", payload["footnote2"]),
        ("Color", payload["color_note"]),
        ("Sources", payload["sources_note"]),
        ("Coverage", payload["coverage_note"]),
        ("What was not guessed", payload["blank_note"]),
    ]
    for i, (title, body) in enumerate(blocks):
        r = 3 + i
        notes.cell(r, 1, title).font = Font(name="Calibri", bold=True, color=NAVY, size=12)
        notes.cell(r, 1).alignment = Alignment(vertical="top")
        notes.cell(r, 2, body).alignment = Alignment(wrap_text=True, vertical="top")
        notes.cell(r, 2).font = Font(name="Calibri", size=11)
        notes.row_dimensions[r].height = 72
    notes.sheet_properties.tabColor = GOLD
    notes.sheet_view.showGridLines = False
    notes.page_setup.fitToPage = True
    notes.page_setup.fitToWidth = 1
    notes.page_setup.orientation = "landscape"

    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "Twins_Hitter_Tracking_2026-09-30.xlsx"
    wb.save(path)
    return path


def write_canvas(payload: dict) -> None:
    data = {
        "subtitle": payload["subtitle"],
        "footnote1": payload["footnote1"],
        "footnote2": payload["footnote2"],
        "colorNote": payload["color_note"],
        "rosterNote": payload["roster_note"],
        "coverageNote": payload["coverage_note"],
        "league": payload["league"],
        "rows": payload["rows"],
        "archetypes": payload["archetypes"],
        "chart": payload["chart"],
        "counts": payload["counts"],
    }
    data_json = json.dumps(json_safe(data), ensure_ascii=False)
    tsx = f'''import {{
  BarChart,
  Callout,
  Divider,
  Grid,
  H1,
  H2,
  H3,
  Pill,
  Row,
  Stack,
  Stat,
  Table,
  Text,
  useCanvasState,
  useHostTheme,
}} from "cursor/canvas";

type Tone = "green" | "red" | "white" | "small" | "blank";

type SheetRow = {{
  id: number;
  name: string;
  pos: string;
  archetype: string | null;
  archetype_mark: string | null;
  window_key: "september" | "last10";
  window: string;
  games: number;
  date_start: string | null;
  date_end: string | null;
  contact_metric: string | null;
  contact_text: string | null;
  contact_n: number;
  contact_tone: Tone;
  whiff_text: string | null;
  whiff_n: number;
  whiff_tone: Tone;
  chase_text: string | null;
  chase_n: number;
  chase_tone: Tone;
  execution: number | null;
  execution_text: string | null;
  execution_tone: Tone;
  risp_text: string | null;
  risp_n: number;
  risp_tone: Tone;
  move_text: string | null;
  move_n: number;
  move_tone: Tone;
  qab_text: string | null;
  qab_n: number;
  qab_tone: Tone;
  pa: number;
  small_notes: string;
  undefined_notes: string;
  league_key: string;
}};

type ArchetypeRow = {{
  name: string;
  pos: string;
  ev: number | null;
  ev_n: number;
  barrel: number | null;
  pred_barrel: number | null;
  residual: number | null;
  rule: string | null;
  archetype: string | null;
  override: boolean;
  rule_reason: string;
  override_note: string;
}};

type LeagueBlock = {{
  barrel: number | null;
  bb: number;
  damage: number | null;
  whiff: number | null;
  whiff_n: number;
  chase: number | null;
  chase_n: number;
  risp: number | null;
  risp_n: number;
  move: number | null;
  move_n: number;
  qab: number | null;
  pa: number;
  ev: number | null;
  ev_n: number;
}};

const DATA = {data_json} as {{
  subtitle: string;
  footnote1: string;
  footnote2: string;
  colorNote: string;
  rosterNote: string;
  coverageNote: string;
  league: {{ season: LeagueBlock; september: LeagueBlock }};
  rows: SheetRow[];
  archetypes: ArchetypeRow[];
  chart: {{ name: string; value: number }}[];
  counts: {{ roster: number; hitters: number; pitchers: number }};
}};

function pct(rate: number | null, digits = 1): string {{
  if (rate === null || Number.isNaN(rate)) return "—";
  return `${{(rate * 100).toFixed(digits)}}%`;
}}

function Rate({{ text, nLabel, tone }}: {{ text: string | null; nLabel: string; tone: Tone }}) {{
  const theme = useHostTheme();
  if (tone === "blank" || text === null) {{
    return <Text tone="tertiary" size="small">—</Text>;
  }}
  const background =
    tone === "green" ? theme.diff.insertedLine :
    tone === "red" ? theme.diff.removedLine :
    tone === "small" ? theme.fill.tertiary :
    theme.bg.editor;
  const color = tone === "small" ? theme.text.secondary : theme.text.primary;
  return (
    <span style={{{{ background, color, padding: "2px 6px", display: "inline-block", minWidth: 72, textAlign: "right" }}}}>
      <span style={{{{ fontWeight: 650 }}}}{{text}}</span>
      <span style={{{{ display: "block", fontSize: 10, color: theme.text.tertiary }}}}>
        {{tone === "small" ? `small sample · ${{nLabel}}` : nLabel}}
      </span>
    </span>
  );
}}

export default function TwinsHitterTracking() {{
  const theme = useHostTheme();
  const [which, setWhich] = useCanvasState<string>("window", "both");
  const rows = DATA.rows.filter((row) => which === "both" || row.window_key === which);
  const season = DATA.league.season;
  const sept = DATA.league.september;

  return (
    <Stack gap={{20}}>
      <Stack gap={{6}}>
        <H1>Twins hitter tracking</H1>
        <Text tone="secondary">{{DATA.subtitle}}</Text>
      </Stack>

      <Row gap={{16}} wrap>
        <Stat value={{String(DATA.counts.hitters)}} label="Position players included" />
        <Stat value={{String(DATA.counts.roster)}} label="Active roster on Sept 30, 2026" />
        <Stat value={{String(DATA.counts.pitchers)}} label="Pitchers excluded" />
      </Row>

      <Callout tone="info" title="Roster date">
        {{DATA.rosterNote}}
      </Callout>

      <Stack gap={{8}}>
        <H2>2026 MLB benchmarks</H2>
        <Text size="small" tone="secondary">
          Pooled regular-season rates: successes divided by chances, not the average of player rates.
          September rows on the sheet are indexed to the September column. Last-10 rows are indexed to the season column.
          Source: Baseball Savant Statcast and MLB Stats API play-by-play, game type R, through Sept 30, 2026.
        </Text>
        <Table
          headers={{["Rate", "2026 season", "Season n", "September", "September n", "Better"]}}
          columnAlign={{["left", "right", "right", "right", "right", "left"]}}
          rows={{[
            ["Barrel%", pct(season.barrel), season.bb.toLocaleString(), pct(sept.barrel), sept.bb.toLocaleString(), "Higher"],
            ["Damage rate (EV ≥ 95 and LA ≥ 10°)", pct(season.damage), season.bb.toLocaleString(), pct(sept.damage), sept.bb.toLocaleString(), "Higher"],
            ["In-zone whiff", pct(season.whiff), season.whiff_n.toLocaleString(), pct(sept.whiff), sept.whiff_n.toLocaleString(), "Lower"],
            ["Chase", pct(season.chase), season.chase_n.toLocaleString(), pct(sept.chase), sept.chase_n.toLocaleString(), "Lower"],
            ["RISP knock-in", pct(season.risp), season.risp_n.toLocaleString(), pct(sept.risp), sept.risp_n.toLocaleString(), "Higher"],
            ["Move-over, 0 out, 2nd to 3rd", pct(season.move), season.move_n.toLocaleString(), pct(sept.move), sept.move_n.toLocaleString(), "Higher"],
            ["QAB (6+ pitches)", pct(season.qab), season.pa.toLocaleString(), pct(sept.qab), sept.pa.toLocaleString(), "Higher"],
            ["Avg EV, competitive contact (mph)", season.ev === null ? "—" : season.ev.toFixed(1), season.ev_n.toLocaleString(), sept.ev === null ? "—" : sept.ev.toFixed(1), sept.ev_n.toLocaleString(), "Archetype input"],
          ]}}
          striped
          stickyHeader
        />
      </Stack>

      <Stack gap={{8}}>
        <H2>September execution composite</H2>
        <Text size="small" tone="secondary">
          Category axis: Twins hitters with a computable September composite that is not a small sample.
          Value axis: execution composite, where 100 equals the September 2026 MLB average.
          Hitters under 10 PA, or under 10 RISP or move-over chances, are left off this chart and stay uncolored on the sheet.
        </Text>
        <BarChart
          horizontal
          height={{Math.max(280, DATA.chart.length * 28)}}
          categories={{DATA.chart.map((d) => d.name)}}
          series={{[{{ name: "September execution composite (100 = league)", data: DATA.chart.map((d) => d.value), tone: "info" }}]}}
          referenceLines={{[{{ value: 100, label: "League 100", tone: "neutral" }}]}}
          beginAtZero
          showValues
        />
      </Stack>

      <Divider />

      <Stack gap={{8}}>
        <Row gap={{8}} align="center">
          <H2>Tracking sheet</H2>
          <Pill active={{which === "both"}} onClick={{() => setWhich("both")}}>Both windows</Pill>
          <Pill active={{which === "september"}} onClick={{() => setWhich("september")}}>September</Pill>
          <Pill active={{which === "last10"}} onClick={{() => setWhich("last10")}}>Last 10 games</Pill>
        </Row>
        <Text size="small" tone="secondary">
          Contact quality is either Statcast barrel% or damage rate (EV ≥ 95 mph and LA ≥ 10°),
          and the same choice is used on both rows. Overall Execution is the headline; the three rates beside it are the equal-weight inputs.
          Neutral cells are the white band (within 10% of the comparison league rate). Green is better. Red is worse.
          {{DATA.colorNote}}
        </Text>
        <Table
          stickyHeader
          headers={{[
            "Hitter",
            "Window",
            "Contact quality",
            "In-zone whiff",
            "Chase",
            <>Overall Execution<sup>2</sup></>,
            "RISP knock-in",
            "Move-over",
            "QAB",
          ]}}
          columnAlign={{["left", "left", "right", "right", "right", "right", "right", "right", "right"]}}
          rows={{rows.map((row) => [
            <Stack gap={{2}} key={{`${{row.id}}-${{row.window_key}}`}}>
              <Text weight="semibold">{{row.name}} · {{row.pos}}</Text>
              <Text size="small" tone="secondary">
                {{row.archetype ? row.archetype : "Unclassified"}}
                {{row.archetype_mark ? <sup>1</sup> : null}}
                {{row.contact_metric ? ` · ${{row.contact_metric === "Damage" ? "damage rate" : "barrel rate"}}` : ""}}
              </Text>
            </Stack>,
            <Stack gap={{2}} key={{`w-${{row.id}}-${{row.window_key}}`}}>
              <Text weight="semibold">{{row.window}}</Text>
              <Text size="small" tone="tertiary">
                {{row.date_start && row.date_end ? `${{row.date_start}} – ${{row.date_end}}` : "No PA in window"}}
                {{` · ${{row.games}} G · ${{row.pa}} PA`}}
              </Text>
            </Stack>,
            <Rate text={{row.contact_text}} nLabel={{`${{row.contact_n}} BB`}} tone={{row.contact_tone}} />,
            <Rate text={{row.whiff_text}} nLabel={{`${{row.whiff_n}} swings`}} tone={{row.whiff_tone}} />,
            <Rate text={{row.chase_text}} nLabel={{`${{row.chase_n}} pitches`}} tone={{row.chase_tone}} />,
            <Rate text={{row.execution_text}} nLabel={{row.league_key === "season" ? "vs 2026 MLB" : "vs month MLB"}} tone={{row.execution_tone}} />,
            <Rate text={{row.risp_text}} nLabel={{`${{row.risp_n}} PA`}} tone={{row.risp_tone}} />,
            <Rate text={{row.move_text}} nLabel={{`${{row.move_n}} PA`}} tone={{row.move_tone}} />,
            <Rate text={{row.qab_text}} nLabel={{`${{row.qab_n}} PA`}} tone={{row.qab_tone}} />,
          ])}}
        />
        <Text size="small" tone="tertiary">
          Source: MLB Stats API roster and play-by-play; Baseball Savant Statcast pitch data. Regular season only.
          {{DATA.coverageNote}}
        </Text>
      </Stack>

      <Grid columns={{2}} gap={{16}}>
        <Stack gap={{6}}>
          <H3>Footnote 1 · Archetype</H3>
          <Text size="small">{{DATA.footnote1}}</Text>
        </Stack>
        <Stack gap={{6}}>
          <H3>Footnote 2 · Overall Execution</H3>
          <Text size="small">{{DATA.footnote2}}</Text>
        </Stack>
      </Grid>

      <Stack gap={{8}}>
        <H2>Archetype audit</H2>
        <Text size="small" tone="secondary">
          Season competitive-contact average exit velocity. League average EV was {{season.ev === null ? "—" : season.ev.toFixed(1)}} mph.
          Damage is a power or high-exit-velocity hitter: average EV at least 1.0 mph above that league average.
          Barrel is an average-exit-velocity hitter: average EV below that line.
          Brooks Lee is the damage example, Luke Keaschall is the barrel example, and Byron Buxton is Damage because he is a power / high-EV hitter.
        </Text>
        <Table
          headers={{["Hitter", "Competitive BB", "Avg EV", "vs league EV", "Barrel%", "EV prediction", "Residual", "Rule", "Used"]}}
          columnAlign={{["left", "right", "right", "right", "right", "right", "right", "left", "left"]}}
          striped
          rows={{DATA.archetypes.map((a) => {{
            const gap = a.ev === null || season.ev === null ? null : a.ev - season.ev;
            const resid = a.residual === null ? "—" : `${{a.residual > 0 ? "+" : ""}}${{(a.residual * 100).toFixed(1)}} pts`;
            return [
              `${{a.name}} · ${{a.pos}}`,
              a.ev_n.toLocaleString(),
              a.ev === null ? "—" : a.ev.toFixed(1),
              gap === null ? "—" : `${{gap > 0 ? "+" : ""}}${{gap.toFixed(1)}} mph`,
              a.barrel === null ? "—" : pct(a.barrel),
              a.pred_barrel === null ? "—" : pct(a.pred_barrel),
              resid,
              a.rule ?? "Unclassified",
              a.override ? `${{a.archetype ?? "—"}} (override)` : (a.archetype ?? "—"),
            ];
          }})}}
        />
      </Stack>

      <Text size="small" tone="tertiary" style={{{{ color: theme.text.tertiary }}}}>
        Blank cells are rates that could not be computed from the cached Statcast and play-by-play files.
        They are not zeros. Small-sample cells show the rate and are not colored.
      </Text>
    </Stack>
  );
}}
'''
    # The f-string doubled braces; DATA json is inserted raw. Fix the accidental
    # double-escaping of the template by having written doubled braces.
    CANVAS.parent.mkdir(parents=True, exist_ok=True)
    CANVAS.write_text(tsx)
    log(f"canvas {CANVAS}")


def transactions_note() -> str:
    url = (
        "https://statsapi.mlb.com/api/v1/transactions?teamId=142"
        "&startDate=2026-09-01&endDate=2026-09-30"
    )
    try:
        data = get_json(url)
    except Exception as exc:  # noqa: BLE001
        return f"Transactions could not be loaded ({exc})."
    lines = []
    for t in data.get("transactions") or []:
        person = (t.get("person") or {}).get("fullName") or "Unknown"
        lines.append(
            f"{t.get('date')} — {t.get('typeDesc')}: {person}"
            + (f" ({(t.get('description') or '').strip()})" if t.get("description") else "")
        )
    if not lines:
        return "No transactions were returned for September 2026."
    return "MLB Stats API transactions for the Twins from Sept 1 through Sept 30, 2026:\n" + "\n".join(lines)


def main() -> None:
    CACHE.mkdir(parents=True, exist_ok=True)
    OUT.mkdir(parents=True, exist_ok=True)
    log("loading roster and schedule")
    hitters, pitchers, roster_meta = load_september_roster()
    log(
        f"september union hitters {len(hitters)} pitchers {len(pitchers)} "
        f"added vs Sept 30 {roster_meta['added_vs_sept30']} "
        f"removed vs prior sheet {roster_meta['removed_vs_prior_sheet']}"
    )
    games = load_schedule()
    dates = sorted({g["date"] for g in games})
    log(f"games {len(games)} dates {len(dates)} {dates[0]} .. {dates[-1]}")
    download_statcast(dates)
    download_pbp(games)
    log("loading pitches")
    pitches = load_pitches(dates)
    sched_pks = {g["game_pk"] for g in games}
    have_pks = set(int(x) for x in pitches["game_pk"].unique())
    coverage = len(sched_pks & have_pks) / len(sched_pks) if sched_pks else 0
    log(f"pitches {len(pitches):,} games in statcast {len(have_pks):,} coverage {coverage:.3%}")
    api = load_api_index(sorted(sched_pks))
    pas, pa_info = build_pa_table(pitches, api)
    missing_games = sorted(sched_pks - have_pks)
    if coverage < 0.98 or pa_info["api_match_rate"] < 0.98:
        write_qa({"coverage": coverage, "missing_games": missing_games[:50], "pa": pa_info})
        raise SystemExit(
            f"coverage {coverage:.3%} or api match {pa_info['api_match_rate']:.3%} "
            "below 98%. qa.json written; sheet not published."
        )

    log("league rates")
    league_season = summarize(pitches, pas)
    month_leagues = {}
    for key, _label, start, end, _fname in MONTH_WINDOWS:
        month_leagues[key] = summarize(
            pitches[pitches["game_date"].between(start, end)],
            pas[pas["game_date"].between(start, end)],
        )
        lg = month_leagues[key]
        log(f"{key} pa {lg['pa']} bb {lg['bb']} ev {lg['ev']}")
    league_sept = month_leagues["september"]
    log("season " + json.dumps({k: league_season[k] for k in ("barrel", "damage", "whiff", "chase", "risp", "move", "qab", "ev", "bb", "pa", "risp_n", "move_n")}))
    log("sept " + json.dumps({k: league_sept[k] for k in ("barrel", "damage", "whiff", "chase", "risp", "move", "qab", "ev", "bb", "pa", "risp_n", "move_n")}))

    log("archetypes and rows")
    archetypes = classify(hitters, pitches, league_season)
    for h in hitters:
        a = archetypes[h["id"]]
        log(f"ARCH {a['name']}: {a['archetype'] or 'unclassified'} | {a['rule_reason']}")
    rows = build_rows(hitters, pitches, pas, league_season, month_leagues, archetypes)

    def pct(x):
        return "n/a" if x is None else f"{x * 100:.1f}%"

    s, m = league_season, league_sept
    footnote1 = (
        "Footnote 1. Archetype picks the contact-quality rate shown for that hitter. "
        "Damage rate is the share of tracked fair batted balls with exit velocity of at least 95 mph "
        "and launch angle of at least 10 degrees. Tracked fair batted balls are Statcast pitches with "
        "type X and non-null exit velocity and launch angle. Barrel rate is Statcast barrel% "
        "(launch_speed_angle = 6) on that same denominator. "
        "The groups use 2026 regular-season competitive contact, which is those batted balls excluding bunts. "
        f"A hitter is Damage when average exit velocity is at least {EV_CLEAR_GAP:.1f} mph above the "
        f"pooled MLB average of {s['ev']:.1f} mph. That is the power and high-exit-velocity group, "
        "graded on damage rate. A hitter is Barrel when average exit velocity is below that line. "
        "That is the average-exit-velocity group, graded on barrel rate. "
        "The two groups do not have to be the same size. "
        f"Hitters with fewer than {ARCHETYPE_MIN_BB} competitive batted balls stay unclassified "
        "and the contact cells are blank. "
        "The superscript 1 marks Brooks Lee, Luke Keaschall, and Byron Buxton. "
        "Brooks Lee is the damage example. Luke Keaschall is the barrel example. "
        "Byron Buxton is Damage because he is a power and high-exit-velocity hitter. "
        f"2026 MLB average EV on competitive contact was {s['ev']:.1f} mph "
        f"(n={s['ev_n']:,}). 2026 MLB barrel rate was {pct(s['barrel'])} "
        f"(n={s['bb']:,} tracked batted balls, bunts included in that denominator). "
        + _player_blurb(archetypes, LEE_ID, "Brooks Lee")
        + " "
        + _player_blurb(archetypes, KEASCHALL_ID, "Luke Keaschall")
        + " "
        + _player_blurb(archetypes, BUXTON_ID, "Byron Buxton")
    )
    footnote2 = (
        "Overall Execution uses three equally weighted inputs. Each rate is divided by the MLB pooled league "
        "rate for the same definition, the three ratios are averaged, and the result is multiplied by 100. "
        "Composite = 100 × ((RISP knock-in / league RISP knock-in) + (move-over / league move-over) + "
        "(QAB / league QAB)) / 3. A composite of 100 is league average. The raw rates are not averaged "
        "together, because they sit on different scales. "
        "RISP knock-in is the share of plate appearances that begin with a runner on second and/or third "
        "in which the batter is credited with at least one RBI (MLB Stats API result.rbi on that plate appearance). "
        "Move-over denominator: plate appearances that begin with a runner on second, nobody on third, and 0 outs. "
        "A runner on first is allowed. Base-out state is the first tracked pitch of the plate appearance. "
        "Success: the runner who began the plate appearance on second has an end-of-PA state of safe at third "
        "(API movement end 3B) or safe at home (end score). He is not successful if he is out, if he never "
        "leaves second, or if he only touches third and is then retired. Advancement on the batted ball, a walk, "
        "hit-by-pitch, wild pitch, passed ball, balk, or stolen base during the plate appearance all count. "
        "A strikeout on which that runner does not end safe at third or home is a failure. "
        "QAB is the share of plate appearances with more than 5 pitches (6 or more Statcast pitches). "
        "Each month is indexed to that month’s MLB rates. Last-10 rows use full-season 2026 MLB rates, "
        "because those games are not one shared month. "
        f"2026 MLB RISP knock-in rate was {pct(s['risp'])} (n={s['risp_n']:,} PA). "
        f"2026 MLB 0-out, second-to-third move-over rate was {pct(s['move'])} (n={s['move_n']:,} PA). "
        f"2026 MLB QAB rate (more than 5 pitches) was {pct(s['qab'])} (n={s['pa']:,} PA). "
        f"September 2026 MLB RISP knock-in rate was {pct(m['risp'])} (n={m['risp_n']:,}). "
        f"September move-over rate was {pct(m['move'])} (n={m['move_n']:,}). "
        f"September QAB rate was {pct(m['qab'])} (n={m['pa']:,}). "
        f"Also indexed the same way: 2026 barrel% {pct(s['barrel'])}, damage rate {pct(s['damage'])}, "
        f"in-zone whiff {pct(s['whiff'])}, chase {pct(s['chase'])}. "
        f"September barrel% {pct(m['barrel'])}, damage rate {pct(m['damage'])}, "
        f"in-zone whiff {pct(m['whiff'])}, chase {pct(m['chase'])}."
    )
    color_note = (
        f"White band is within {BAND:.0%} of the comparison league rate, inclusive "
        f"(ratio 0.90 to 1.10). For the composite, that is 90 to 110. "
        f"Small sample: under {SMALL_BB} tracked batted balls for contact quality; "
        f"under {SMALL_N} in-zone swings, out-of-zone pitches, or plate-appearance opportunities "
        "for the other columns. Overall Execution is a small sample when PA, RISP PA, or move-over PA is under "
        f"{SMALL_N}. Santurce hitter cards color a continuous blue-to-red magnitude scale "
        "(rateHeatStyle in pregame.html). This sheet uses the requested direction instead: "
        "white at the league band, green when the rate is better, red when it is worse. "
        "Header navy #122C6F matches the Santurce licey navy. Cell green and red are light fills "
        "so the numbers stay readable; they are not the Santurce magnitude gradient."
    )
    roster_note = (
        f"MLB Stats API active roster, every day from 2026-09-01 through 2026-09-30: "
        f"{len(hitters)} position players were active on at least one of those days, and {len(pitchers)} pitchers are excluded. "
        f"September 30 alone listed {roster_meta['hitters_on_sept30']} position players. "
        f"Added versus that Sept 30 snapshot: {', '.join(roster_meta['added_vs_sept30']) or 'none'}. "
        f"Removed versus the earlier 14-name sheet, because they were not on the active roster on any September day: "
        f"{', '.join(roster_meta['removed_vs_prior_sheet']) or 'none'}. "
        f"Added versus that earlier sheet: {', '.join(roster_meta['added_vs_prior_sheet']) or 'none'}. "
        "Last-10 games can predate September when that is when the hitter last played. "
        "A month with no plate appearances is a blank row."
    )
    coverage_note = (
        f"Schedule finals: {len(games):,} regular-season games, {dates[0]} through {dates[-1]}. "
        f"Statcast covered {len(sched_pks & have_pks):,} of those game pks ({coverage:.1%}). "
        f"Plate appearances: {pa_info['pa_rows']:,}. "
        f"API join on game and at-bat (Statcast at_bat_number = API atBatIndex + 1, same batter): "
        f"{pa_info['api_match_rate']:.1%}. Batter-id mismatches: {pa_info['batter_mismatch']:,}."
    )
    blank_note = (
        "Cells are blank when the denominator is zero or, for the composite, when any of the three "
        "input rates is undefined. A month with no plate appearances is a blank row, with no zero rates. "
        f"Archetype contact cells are blank when the hitter has fewer than {ARCHETYPE_MIN_BB} competitive batted balls. "
        "No rate was imputed from a nearby player, from a career norm, or from memory."
    )
    chart = []
    for row in rows:
        if row["window_key"] != "september":
            continue
        if row["execution"] is None or row["execution_tone"] == "small":
            continue
        chart.append({"name": row["name"], "value": row["execution"]})
    chart.sort(key=lambda d: d["value"], reverse=True)

    payload = {
        "subtitle": (
            f"Every Twins position player on the active roster at any point in September 2026: {len(hitters)} hitters. "
            f"{len(pitchers)} pitchers excluded. "
            "April through September are regular-season games in that month. "
            "Last 10 is each hitter’s last 10 games with a plate appearance on or before September 30. "
            "A month with no plate appearances is blank."
        ),
        "footnote1": footnote1,
        "footnote2": footnote2,
        "color_note": color_note,
        "roster_note": roster_note,
        "coverage_note": coverage_note,
        "sources_note": (
            "Roster, schedule, transactions, RBI, and runner movement: MLB Stats API "
            "(statsapi.mlb.com), public endpoints, no login. "
            "Pitches, zones, swing descriptions, exit velocity, launch angle, and the barrel flag "
            "(launch_speed_angle = 6): Baseball Savant Statcast, downloaded with pybaseball and cached "
            f"under {SC_DIR}. Play-by-play extracts cached under {PBP_DIR}. "
            "Regular season only (schedule gameType=R and Statcast game_type=R)."
        ),
        "blank_note": blank_note,
        "transactions_note": transactions_note(),
        "pitchers_excluded": ", ".join(f"{p['name']} ({p['pos']})" for p in sorted(pitchers, key=lambda p: last_name(p["name"]))),
        "hitters_included": ", ".join(f"{p['name']} ({p['pos']})" for p in hitters),
        "league": {"season": league_season, **month_leagues},
        "rows": rows,
        "archetypes": [archetypes[h["id"]] for h in hitters],
        "chart": chart,
        "counts": {
            "roster": len(hitters) + len(pitchers),
            "hitters": len(hitters),
            "pitchers": len(pitchers),
        },
        "roster_audit": roster_meta,
        "pa_info": pa_info,
        "coverage": coverage,
        "missing_games": missing_games,
        "barrel_model": None,
    }
    write_qa(
        {
            "coverage": coverage,
            "missing_games": missing_games,
            "pa_info": pa_info,
            "league_season": league_season,
            "league_september": league_sept,
            "archetypes": payload["archetypes"],
            "hitters": [
                {
                    "name": r["name"],
                    "window": r["window"],
                    "games": r["games"],
                    "pa": r["pa"],
                    "contact": r["contact_text"],
                    "whiff": r["whiff_text"],
                    "chase": r["chase_text"],
                    "execution": r["execution_text"],
                    "risp": r["risp_text"],
                    "move": r["move_text"],
                    "qab": r["qab_text"],
                    "undefined": r["undefined_notes"],
                }
                for r in rows
            ],
        }
    )
    path = write_workbook(payload)
    (OUT / "tracking.json").write_text(json.dumps(json_safe(payload), indent=2))
    DATA.mkdir(parents=True, exist_ok=True)
    (DATA / "tracking.json").write_text(json.dumps(json_safe(payload), indent=2) + "\n")
    write_site(rows, month_leagues, league_season, payload)
    write_canvas(payload)
    log(f"wrote {path}")


def write_site(rows: list[dict], month_leagues: dict, league_season: dict, payload: dict) -> None:
    DATA.mkdir(parents=True, exist_ok=True)
    views = []
    for key, label, start, end, fname in MONTH_WINDOWS:
        subset = [r for r in rows if r["window_key"] == key]
        blob = {
            "id": key,
            "label": f"{label} 2026",
            "start": start,
            "end": end,
            "comparison": f"{label} 2026 MLB",
            "league": month_leagues[key],
            "hitters": subset,
        }
        (DATA / fname).write_text(json.dumps(json_safe(blob), indent=2) + "\n")
        views.append({"id": key, "label": f"{label} 2026", "file": fname})
        log(f"site {fname} hitters {len(subset)} no-pa {sum(1 for r in subset if r['no_pa'])}")
    last = [r for r in rows if r["window_key"] == "last10"]
    (DATA / "last10.json").write_text(
        json.dumps(
            json_safe(
                {
                    "id": "last10",
                    "label": "Last 10 games",
                    "start": None,
                    "end": SEASON_END,
                    "comparison": "2026 MLB season",
                    "league": league_season,
                    "hitters": last,
                }
            ),
            indent=2,
        )
        + "\n"
    )
    views.append({"id": "last10", "label": "Last 10 games", "file": "last10.json"})
    manifest = {
        "default": "september",
        "views": views,
        "footnote1": payload["footnote1"],
        "footnote2": payload["footnote2"],
        "roster_note": payload["roster_note"],
        "color_note": payload["color_note"],
        "blank_note": payload["blank_note"],
        "subtitle": payload["subtitle"],
        "hitters_included": payload["hitters_included"],
        "archetypes": payload["archetypes"],
        "roster_audit": payload["roster_audit"],
    }
    (DATA / "manifest.json").write_text(json.dumps(json_safe(manifest), indent=2) + "\n")
    log(f"site manifest {DATA / 'manifest.json'}")


def _player_blurb(archetypes: dict, pid: int, name: str) -> str:
    a = archetypes[pid]
    ev = "n/a" if a["ev"] is None else f"{a['ev']:.1f} mph"
    brl = "n/a" if a["barrel"] is None else f"{a['barrel'] * 100:.1f}%"
    return (
        f"{name}’s 2026 competitive-contact average EV was {ev} "
        f"(n={a['ev_n']:,}) and his barrel rate was {brl}. {a['rule_reason']} "
        f"Final label: {a['archetype'] or 'unclassified'}."
    )


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception:
        traceback.print_exc()
        raise
