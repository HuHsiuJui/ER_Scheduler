"""Consent-first leave accounting. All amounts are hours; no payroll is issued.

Four-week totals must come from a verified complete cycle, not monthly OFF counts.
The rules here represent the supplied institutional workflow, not legal advice.
"""
from dataclasses import dataclass, asdict
from datetime import date, timedelta
from decimal import Decimal, ROUND_HALF_UP


def hours(value):
    value = Decimal(str(value))
    if not value.is_finite() or value < 0:
        raise ValueError('時數必須是有限且非負的數字')
    return value.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)


def cycle_for(day: date, anchor: date = date(2026, 9, 10)):
    start = anchor + timedelta(days=((day - anchor).days // 28) * 28)
    return start, start + timedelta(days=27)


@dataclass(frozen=True)
class LeaveInput:
    nurse_id: str
    annual_opening: float | None
    comp_opening: float | None
    overtime_hours: float = 0
    overtime_choice: str = '未選擇'  # 未選擇 / 領加班費 / 轉補休
    conversion_rate: float = 1
    approved_annual_hours: float = 0
    approved_comp_hours: float = 0
    cycle_verified: bool = False
    regular_rest_remaining: float | None = None
    additional_leave_hours: float = 0
    comp_consent: str = '未確認'
    annual_asked: bool = False
    annual_consent: str = '未確認'
    opening_verified: bool = False
    entries_verified: bool = False
    # Earned and confirmed credits only. Future scheduled overtime is not spendable.
    overtime_earned: bool = False


def settle(item: LeaveInput) -> dict:
    if item.overtime_choice not in {'未選擇', '領加班費', '轉補休'}:
        raise ValueError('不支援的加班處理選項')
    for consent in (item.comp_consent, item.annual_consent):
        if consent not in {'未確認', '同意', '不同意'}:
            raise ValueError('意願必須為未確認、同意或不同意')
    ot, rate, annual_used, comp_used, need = map(hours, (
        item.overtime_hours, item.conversion_rate, item.approved_annual_hours,
        item.approved_comp_hours, item.additional_leave_hours))
    if rate == 0:
        raise ValueError('補休換算率必須大於零')
    warnings = []
    out = dict(nurse_id=item.nurse_id, annual_opening=item.annual_opening,
               comp_opening=item.comp_opening, overtime_hours=float(ot),
               overtime_choice=item.overtime_choice, additional_leave_hours=float(need),
               suggested_comp_hours=None, annual_shortfall_hours=None,
               annual_projected=None, comp_projected=None, annual_official=None,
               comp_official=None, comp_consent=item.comp_consent,
               annual_asked=item.annual_asked, annual_consent=item.annual_consent)
    if item.annual_opening is None or item.comp_opening is None:
        return {**out, 'warnings': ['缺少期初餘額，不以零代替']}
    annual, comp = hours(item.annual_opening), hours(item.comp_opening)
    credit = (ot * rate).quantize(Decimal('0.01')) if item.overtime_choice == '轉補休' else Decimal(0)
    available = max(Decimal(0), comp + (credit if item.overtime_earned else 0) - comp_used)
    if annual_used > annual or comp_used > comp + (credit if item.overtime_earned else 0):
        warnings.append('已核准扣假超過可用餘額，須修正')
    if ot and item.overtime_choice == '未選擇':
        warnings.append('請本人選擇領加班費或轉補休')
    if not item.cycle_verified or item.regular_rest_remaining is None:
        warnings.append('四週完整週期未核定，不自動認定超休')
    elif hours(item.regular_rest_remaining) > 0 and need > 0:
        warnings.append('仍有一般休假，先核對一般休假班，不扣補休或特休')
    else:
        suggestion = min(available, need)
        shortage = need if item.comp_consent == '不同意' else need - suggestion
        out.update(suggested_comp_hours=float(suggestion), annual_shortfall_hours=float(shortage))
        if suggestion > 0 and item.comp_consent == '未確認':
            warnings.append('請詢問本人是否使用加班補休')
        if shortage > 0:
            if not item.annual_asked:
                warnings.append('補休不足或未採用，必須詢問本人是否使用特休')
            elif item.annual_consent != '同意':
                warnings.append('特休尚未同意或已拒絕，不自動扣假；需另行調整')
        if item.comp_consent == '同意':
            comp_used += suggestion
        comp_resolved = suggestion == 0 or item.comp_consent in {'同意', '不同意'}
        if shortage and comp_resolved and item.annual_asked and item.annual_consent == '同意':
            if shortage <= annual - annual_used:
                annual_used += shortage
            else:
                warnings.append('特休餘額不足，不得自動扣成負數')
    out.update(annual_projected=float(annual - annual_used),
               comp_projected=float(comp + credit - comp_used))
    if not item.opening_verified:
        warnings.append('期初餘額基準日與是否已扣本月假待確認')
    if not item.entries_verified:
        warnings.append('本月出勤與扣假紀錄待核定')
    if credit and not item.overtime_earned:
        warnings.append('加班轉補休為預估，尚未實際取得')
    if not warnings:
        out.update(annual_official=out['annual_projected'], comp_official=out['comp_projected'])
    out['warnings'] = warnings
    return out


def settle_many(rows):
    items = [LeaveInput(**row) for row in rows]
    if len({x.nurse_id for x in items}) != len(items):
        raise ValueError('休假資料員編重複')
    return [settle(x) for x in items]

