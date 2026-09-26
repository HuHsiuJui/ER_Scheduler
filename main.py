"""Portable, offline research/review entrypoint; no private data required.

Run: python main.py --demo
Custom reviewed input: python main.py --input data.json --leave-input balances.json
"""
from __future__ import annotations
import argparse
import hashlib
import json
import sys
from collections import Counter
from dataclasses import asdict
from datetime import date, timedelta
from pathlib import Path
import scheduler_core as core
from leave_accounting import settle_many

BASE = Path(__file__).resolve().parent


def synthetic_input():
    """Construct 38 fictional staff on a known feasible 3-on/3-off pattern.

    OFF patterns intentionally simplify the problem: this is an executable smoke
    demonstration, not a realistic staffing benchmark or proof for a hospital.
    """
    start = date(2026, 9, 1)
    staff, requests, boundary = [], [], []
    for shift, count in [('D', 14), ('E', 14), ('N', 10)]:
        for i in range(count):
            nid = f'{shift}{i+1:03}'
            group = int(i >= count // 2)
            staff.append(dict(nurse_id=nid, name=f'示範人員{nid}', home_shift=shift,
                hire_date='2020-01-01', position='Nurse', max_hours=120,
                areas=['Leader', 'Triage', 'Critical', 'Clinic1', 'Clinic2', 'Observation', 'Teaching']))
            for offset in range(-14, 30):
                day = start + timedelta(days=offset)
                on = (offset % 6 < 3) == (group == 0)
                if offset < 0:
                    boundary.append(dict(nurse_id=nid, day=str(day), shift=shift if on else 'OFF'))
                elif not on:
                    requests.append(dict(nurse_id=nid, day=str(day), kind='OFF_LOCK', raw='示範預假'))
    return dict(year=2026, month=9, nurses=staff, requests=requests, boundary=boundary,
        official_holidays=[], max_cross_shift_days=0, max_overtime_shifts=1,
        preferred_preceptors={}, training_plan=[], demo=True)


def parse_input(data):
    year, month = int(data['year']), int(data['month'])
    days = core.month_days(year, month)
    staff = []
    for item in data['nurses']:
        n = dict(item)
        n['hire_date'] = date.fromisoformat(n['hire_date'])
        n['areas'] = core.parse_areas('、'.join(n['areas']))
        if n['home_shift'] not in {'D', 'E', 'N'}:
            raise ValueError('home_shift 必須是 D/E/N')
        if n.get('max_hours') is not None and (n['max_hours'] < 0 or n['max_hours'] % 8):
            raise ValueError('排班核心以完整 8 小時班為單位，max_hours 必須是 8 的非負倍數')
        if n.get('prior_unused_hours', 0):
            raise ValueError('累積補休餘額不能直接當成已同意扣休或降低目標，請另用休假帳本')
        staff.append(core.NurseProfile(**n))
    nurses = tuple(sorted(staff, key=lambda n: (n.hire_date, n.nurse_id)))
    ids = {n.nurse_id for n in nurses}
    if len(ids) != len(nurses) or not nurses:
        raise ValueError('人員不得為空或有重複員編')
    reqs, boundaries = [], []
    seen = set()
    for item in data.get('requests', []):
        r = dict(item); r['day'] = date.fromisoformat(r['day']); r['kind'] = core.RequestKind(r['kind'])
        if r['nurse_id'] not in ids or r['day'] not in days:
            raise ValueError('預假員編或日期不在本月名單內')
        if r['kind'] in {core.RequestKind.MUST_SHIFT, core.RequestKind.PREFERRED_SHIFT, core.RequestKind.AVOID_SHIFT} and r.get('shift') not in {'D', 'E', 'N', 'H'}:
            raise ValueError('指定班別缺少有效 shift')
        key = (r['nurse_id'], r['day'], r['kind'], r.get('shift'))
        if key in seen:
            raise ValueError('重複預假紀錄，不得重複扣抵特休')
        seen.add(key); r['source'] = 'LOCAL_REVIEW_JSON'
        reqs.append(core.Request(**r))
    for item in data.get('boundary', []):
        b = dict(item); b['day'] = date.fromisoformat(b['day'])
        shift = b['shift']
        if shift not in {'D','E','N','H','8-5','OFF'}:
            raise ValueError('跨月班別不明；行政班請明列起迄與 attendance')
        attendance = b.pop('attendance', shift != 'OFF')
        if not isinstance(attendance, bool) or attendance != (shift != 'OFF'):
            raise ValueError('跨月出勤旗標與班別不一致')
        boundaries.append(core.BoundaryEntry(raw=b.pop('raw', shift), attendance=attendance, **b))
    bkeys = {(b.nurse_id,b.day) for b in boundaries}
    expected = {(nid, days[0]-timedelta(days=k)) for nid in ids for k in range(1,15)}
    if bkeys != expected or len(boundaries) != len(expected):
        raise ValueError('須逐人提供上月最後 14 天，未知不可當 OFF；行政班也算出勤')
    training = core.load_training_state_canonical(nurses, year=year, month=month, interactive=False)
    roles, trainees, assignments = {}, {}, {}
    for row in data.get('training_plan', []):
        day, shift, nid, role = date.fromisoformat(row['day']), row['shift'], row['nurse_id'], row['role']
        if day not in days or nid not in ids or shift not in {'D','E','N'}:
            raise ValueError('訓練人員、日期或班別錯誤')
        key = (day,shift)
        if nid in assignments.get(key,{}) or any(nid in a for k,a in assignments.items() if k[0]==day):
            raise ValueError('同日重複訓練')
        if role not in {'Observation1','Observation2','Triage','LeaderTriage','Critical','Clinic1','Clinic2'}:
            raise ValueError('未支援的訓練區域')
        assignments.setdefault(key,{})[nid] = role
        if len(assignments[key]) > 2:
            raise ValueError('每班同時訓練不得超過 2 人')
    for key, value in assignments.items():
        roles[key] = tuple(value.values()); trainees[key] = frozenset(value)
    config = core.SchedulerConfig(year, month,
        official_holidays=frozenset(date.fromisoformat(d) for d in data['official_holidays']),
        max_cross_shift_days_per_person=int(data.get('max_cross_shift_days',5)),
        max_overtime_shifts_per_person=int(data.get('max_overtime_shifts',1)),
        protect_part_time_targets=True, real_target_minimum=True, allow_h=False,
        preceptor_ids=frozenset(n.nurse_id for n in nurses if 'Teaching' in n.areas),
        part_time_ids=frozenset(n.nurse_id for n in nurses if n.part_time),
        preferred_preceptors=data.get('preferred_preceptors',{}),
        required_training_roles_by_day_shift=roles,
        required_training_trainees_by_day_shift=trainees,
        required_training_assignments_by_day_shift=assignments,
        solver_time_limit_seconds=float(data.get('time_limit_seconds',60)))
    if config.max_cross_shift_days_per_person < 0 or config.max_overtime_shifts_per_person < 0 or config.solver_time_limit_seconds <= 0:
        raise ValueError('上限不得為負，求解秒數必須大於零')
    for nid,pid in config.preferred_preceptors.items():
        if nid not in ids or pid not in config.preceptor_ids:
            raise ValueError('指定老師必須存在且 areas 含 Teaching')
    return nurses, tuple(reqs), tuple(boundaries), training, config


def run(data):
    nurses, requests, boundary, training, config = parse_input(data)
    fingerprint = hashlib.sha256(json.dumps(data,sort_keys=True,ensure_ascii=False).encode()).hexdigest()
    audit = core.GoogleSourceAudit('LOCAL', '研究示範/待核對資料', 'LOCAL', config.year, config.month,
        len(nurses), 0, fingerprint, len(nurses), len(requests), 'REVIEW', 'LOCAL_REVIEW_JSON', False)
    previous = core.PreviousRosterAudit('LOCAL','本機提供跨月資料','LOCAL',
        (date(config.year,config.month,1)-timedelta(days=1)).year,
        (date(config.year,config.month,1)-timedelta(days=1)).month,
        len(boundary),0,fingerprint,len(nurses),len(boundary),'REVIEW','LOCAL_REVIEW_JSON',False)
    targets = core.build_target_shifts(nurses, requests, training, year=config.year, month=config.month,
        official_holidays=config.official_holidays, part_time_min_shifts=config.part_time_min_shifts)
    feasibility = core.feasibility_check(nurses,requests,training,targets,year=config.year,month=config.month)
    if not feasibility.feasible:
        raise ValueError('前置檢查失敗：'+'; '.join(i.message for i in feasibility.hard_issues))
    result = core.solve_schedule(nurses,requests,training,boundary,targets,audit,previous,config)
    validation = core.validate_result(result,requests,boundary,config)
    # Preserve source failures in the report. Local/demo never gets a clinical PASS.
    errors = [asdict(i) for i in validation.hard_errors]
    clinical = [e for e in errors if e['code'] not in {'GOOGLE_SOURCE','PREVIOUS_GOOGLE_SOURCE'}]
    report = dict(mode='DEMO' if data.get('demo') else 'LOCAL_REVIEW', clinical_publishable=False,
        solver_status=result.solver_status, input_sha256=fingerprint, result_sha256=result.result_sha256,
        hard_errors=errors, schedule_constraint_errors=clinical,
        warnings=[asdict(i) for i in validation.warnings], metrics=dict(result.solver_metrics),
        notice='僅研究/人工審查，不是已核定正式班表；未執行四週休假合法性認證。')
    return result, requests, report


def json_default(value):
    if isinstance(value,date): return value.isoformat()
    if isinstance(value,(set,frozenset)): return sorted(value)
    if hasattr(value,'value'): return value.value
    if hasattr(value,'item'): return value.item()
    raise TypeError(type(value).__name__)


def main(argv=None):
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, 'reconfigure'):
            stream.reconfigure(encoding='utf-8')
    parser = argparse.ArgumentParser(description='急診護理排班研究專案（匿名展示／本機審查）')
    parser.add_argument('--demo',action='store_true',help='使用 38 人完全虛構資料')
    parser.add_argument('--input',type=Path,help='本機已整理之 JSON 輸入')
    parser.add_argument('--leave-input',type=Path,help='逐人餘額與本人意願 JSON')
    parser.add_argument('--output-dir',type=Path,default=BASE/'results')
    parser.add_argument('--write-demo-input',type=Path,help='只建立可修改的匿名輸入範例')
    args = parser.parse_args(argv)
    if args.write_demo_input:
        if args.write_demo_input.exists(): parser.error('目的檔已存在，請使用新檔名')
        args.write_demo_input.parent.mkdir(parents=True,exist_ok=True)
        args.write_demo_input.write_text(json.dumps(synthetic_input(),ensure_ascii=False,indent=2),encoding='utf-8')
        return 0
    if args.demo == bool(args.input): parser.error('請擇一指定 --demo 或 --input')
    args.output_dir.mkdir(parents=True,exist_ok=True)
    # Each run gets its own folder, so failed reruns cannot leave a stale success workbook.
    from datetime import datetime
    from uuid import uuid4
    out = args.output_dir / (datetime.now().strftime('%Y%m%d_%H%M%S')+'_'+uuid4().hex[:6])
    out.mkdir()
    try:
        data = synthetic_input() if args.demo else json.loads(args.input.read_text(encoding='utf-8-sig'))
        print('正在求解並檢查排班約束…',flush=True)
        result, requests, report = run(data)
        leave_rows = json.loads(args.leave_input.read_text(encoding='utf-8-sig')) if args.leave_input else []
        # Derive this month's overtime and original annual leave from the actual solved result.
        ids = {n.nurse_id for n in result.nurses}
        for row in leave_rows:
            if row['nurse_id'] not in ids: raise ValueError('休假資料有未知員編')
            row['overtime_hours'] = core.overtime_hours(result.assignments,result.target_shifts,row['nurse_id'])
            planned_al = 8 * sum(r.nurse_id==row['nurse_id'] and r.kind==core.RequestKind.ANNUAL_LEAVE for r in requests)
            row['approved_annual_hours'] = planned_al + row.get('approved_annual_hours',0)
        ledger = settle_many(leave_rows)
        report['leave_balances'] = ledger
        report['assignments'] = [asdict(a) for a in result.assignments]
        (out/'report.json').write_text(json.dumps(report,default=json_default,ensure_ascii=False,indent=2),encoding='utf-8')
        if report['schedule_constraint_errors']:
            print('存在排班硬條件錯誤，僅輸出診斷，禁止輸出班表。',flush=True)
            return 2
        from workbook_export import export_review
        export_review(out/'研究展示班表.xlsx',result,requests,report,ledger)
        print(f"完成：{out}\n求解狀態：{result.solver_status}；研究展示，不作正式發布。",flush=True)
        return 0
    except (ValueError,RuntimeError,KeyError,TypeError,OSError) as exc:
        (out/'failure.json').write_text(json.dumps({'status':'NO_ROSTER','reason':str(exc),
            'advice':['確認預假與指定班是否互相衝突','核對高級核心及帶教資格人力',
                      '檢查跨月連班與部分工時可上班天數','先檢查求解時限，逾時不等於已證明無解',
                      '調整條件須另得同意，不會自動刪除特休或訓練']},ensure_ascii=False,indent=2),encoding='utf-8')
        print(f'未產生班表：{exc}\n診斷：{out}',file=sys.stderr)
        return 2


if __name__=='__main__':
    raise SystemExit(main())

