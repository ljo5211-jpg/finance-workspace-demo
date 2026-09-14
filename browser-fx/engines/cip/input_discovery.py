from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from storage_paths import input_workbooks


ASSET_REGISTER_RE = re.compile(r"^(?P<year>\d{2})\.(?P<month>\d{1,2})\s+자산대장$", re.IGNORECASE)
FINAL_RE = re.compile(r"^(?P<year>\d{2})\.(?P<month>\d{1,2})\.?\s+건설중인자산.*$", re.IGNORECASE)
LEDGER_RE = re.compile(r"^(?P<year>\d{2})\s+건설중인자산원장$", re.IGNORECASE)


@dataclass(frozen=True, order=True)
class Period:
    year: int
    month: int

    @property
    def label(self) -> str:
        return f"{self.year:02d}.{self.month}"

    @property
    def korean_label(self) -> str:
        return f"{self.year:02d}년 {self.month}월"


def workbook_paths(root: Path):
    yield from input_workbooks(root)


def previous_final(root: Path, period: Period):
    """Compare the immediately preceding month; never silently guess a version."""
    previous = Period(period.year, period.month - 1) if period.month > 1 else Period(period.year - 1, 12)
    candidates = []
    for path in workbook_paths(root):
        match = FINAL_RE.match(path.stem.strip())
        if match and '검증' not in path.stem and '자산대장' not in path.stem:
            if Period(int(match['year']), int(match['month'])) == previous:
                candidates.append(path)
    note = '' if len(candidates) == 1 else (
        f'{previous.label} 전월 최종본이 없어 대체 비교를 하지 못했습니다.' if not candidates else
        f'{previous.label} 전월 최종본이 여러 개입니다. 비교할 파일 하나만 적재/최종본에 남겨 주세요.')
    return previous, candidates[0] if len(candidates) == 1 else None, note


def discover_inputs(root: Path):
    registers: dict[Period, Path] = {}
    finals: dict[Period, Path] = {}
    ledgers: dict[int, Path] = {}
    for path in workbook_paths(root):
        stem = path.stem.strip()
        if match := ASSET_REGISTER_RE.match(stem):
            period = Period(int(match.group("year")), int(match.group("month")))
            if period in registers:
                raise RuntimeError(f"동일 기간 자산대장이 중복되었습니다: {registers[period].name}, {path.name}")
            registers[period] = path
            continue
        if match := LEDGER_RE.match(stem):
            year = int(match.group("year"))
            if year in ledgers:
                raise RuntimeError(f"동일 연도 건설중인자산원장이 중복되었습니다: {ledgers[year].name}, {path.name}")
            ledgers[year] = path
            continue
        if "자산대장" not in stem and "검증" not in stem and (match := FINAL_RE.match(stem)):
            period = Period(int(match.group("year")), int(match.group("month")))
            current = finals.get(period)
            if current is None or path.stat().st_mtime > current.stat().st_mtime:
                finals[period] = path
    if not registers:
        raise RuntimeError("자산대장을 찾을 수 없습니다. 예: 26.8 자산대장.XLSX")
    latest_period = max(registers)
    ledger = ledgers.get(latest_period.year)
    if ledger is None:
        raise RuntimeError(f"{latest_period.year:02d}년 건설중인자산원장을 찾을 수 없습니다.")
    eligible_finals = {period: path for period, path in finals.items() if period <= latest_period}
    if not eligible_finals:
        raise RuntimeError("승계할 건설중인자산 최종본을 찾을 수 없습니다.")
    reference_period = max(eligible_finals)
    return {
        "registers": dict(sorted(registers.items())),
        "ledger": ledger,
        "reference_final": eligible_finals[reference_period],
        "reference_period": reference_period,
        "report_period": latest_period,
    }
