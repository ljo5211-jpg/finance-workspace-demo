"""File boundary for the trial. Business calculations stay in module engines."""
import json
import sys
import sqlite3
import importlib.util
from datetime import date, datetime
from pathlib import Path


def load(name, path):
    sys.path.insert(0, str(path.parent))
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


def calculate_files(module, root_string, options_json='{}'):
    root = Path(root_string)
    options = json.loads(options_json)
    engines = Path(__file__).parent / 'engines'
    sys.modules.pop('input_discovery', None)
    engine = load('build_dashboard', engines / module / 'build_dashboard.py')
    if module == 'receivables':
        db = root / 'receivables.db'
        if db.exists():
            db.unlink()
        connection = sqlite3.connect(db)
        connection.row_factory = sqlite3.Row
        engine.create_schema(connection)
        source = root / '적재'
        inputs = sorted(source.glob('aging_*.xlsx'))
        if not inputs:
            raise ValueError('Aging 적재파일이 없습니다.')
        for path in inputs:
            engine.import_aging(connection, path)
        credits = sorted(source.glob('credit_*.xlsx'))
        if not credits:
            raise ValueError('여신 적재파일이 없습니다.')
        asof = date.fromisoformat(inputs[-1].stem.removeprefix('aging_'))
        credit_date = date.fromisoformat(credits[-1].stem.removeprefix('credit_'))
        engine.import_credit(connection, credits[-1], credit_date)
        connection.commit()
        settings = json.loads((source / 'settings.json').read_text(encoding='utf8'))
        payload = engine.advanced_dashboard_data(connection, settings['rate'], '가상 적재 환율', settings['rateDate'], str(asof), settings['allowanceModel'])
        controls = [dict(row) for row in connection.execute('SELECT snapshot_date,currency,SUM(amount) amount,COUNT(*) count FROM ar_snapshot GROUP BY snapshot_date,currency ORDER BY snapshot_date,currency')]
        connection.close()
        live = root / 'live.html'
        live.write_text('const dashboardData = '+json.dumps(payload, ensure_ascii=False), encoding='utf8')
        dev = load('trial_ar_dev', engines / module / 'build_dev.py')
        dev.ROOT = root
        development = dev.build(db, live)
        dev.validate(development)
        return json.dumps({'payload':payload, 'dev':development, 'sourceControls':controls}, ensure_ascii=False, allow_nan=False, default=str)
    if module == 'fx':
        (root / 'dashboard-settings.json').write_text('{}', encoding='utf8')
        # Optional scenario changes apply to an in-memory copy after real XLSX parsing.
        rate, repeat = options.get('rate'), options.get('repeat', 1)
        if rate is not None or repeat != 1:
            from math import isfinite
            if rate is not None and (not isinstance(rate,(int,float)) or not isfinite(rate) or not 100<=rate<=10000):
                raise ValueError('가상 평가환율은 100~10,000 사이여야 합니다.')
            if repeat not in (1,100,500):
                raise ValueError('지원하지 않는 시험 크기입니다.')
            original = engine.load_single_sheet
            files = engine.discover_inputs(root/'적재')
            monthly = [key for key in files if '월' in key and '말' not in key]
            latest = max(monthly, key=lambda key:int(key.split('년')[1].strip().removesuffix('월')))
            latest_path = files[latest]
            class Rows:
                def __init__(self, rows):self.rows=rows
                def iter_rows(self, **kwargs):return iter(self.rows)
                def close(self):pass
            def transformed(path, **kwargs):
                book,sheet=original(path, **kwargs)
                rows=[list(row) for row in sheet.iter_rows(values_only=True)];book.close()
                if path==latest_path and rate is not None:
                    for row in rows[1:]:
                        if row[11]!='USD':continue
                        delta=row[12]*rate-row[16]
                        row[10]=rate;row[16]+=delta;row[17]+=delta;row[18]+=delta
                if repeat>1 and any(name in path.name for name in ('외화계정원장','환차손익원장')):
                    expanded=[rows[0]]
                    for i in range(repeat):
                        for source in rows[1:]:
                            row=source.copy();row[1]=str(row[1])+'-COPY-'+str(i);row[2]=str(row[2])+'-COPY-'+str(i);expanded.append(row)
                    rows=expanded
                return Rows([]),Rows(rows)
            engine.load_single_sheet=transformed
        payload=engine.build_payload(root)
        for c in payload['checks']:
            if abs(c['valuationDifference'])>.01 or abs(c['realizedDifference'])>.01:raise ValueError('외화 원장 Net 불일치')
        payload['source']['sourceWorkbook']='서버 가상 XLSX 적재파일'
        return json.dumps(payload, ensure_ascii=False, allow_nan=False)
    if module == 'cip':
        payload=engine.build_payload(root,edited_final=True)
        if options.get('review'):
            apply_review(engine,payload,options['review'])
        if any(c.get('blocking') and c['status']!='일치' for c in payload['checks']):
            raise ValueError('건설중인자산 명세 대사 실패')
        return json.dumps(payload, ensure_ascii=False, allow_nan=False, default=str)
    raise ValueError('등록되지 않은 계산 모듈')


def apply_review(engine,payload,review):
    """Apply approved metadata to parsed amounts, reusing original depreciation rules."""
    from math import isfinite
    from collections import defaultdict
    from management_rules import SHEETS
    if not review.get('approved') or not str(review.get('confirmedBy','')).strip():
        raise ValueError('분류 확인자와 승인이 필요합니다.')
    choices={c['assetId']:c['sheet'] for c in review['choices']}
    components={c['rowKey']:c for c in review['componentChoices']}
    if set(choices)!={p['assetId'] for p in payload['projects']}:
        raise ValueError('자산 분류 입력이 원장과 일치하지 않습니다.')
    groups={g['assetId']:g for groups in payload['validationSheets'].values() for g in groups}
    summary=defaultdict(lambda:dict(count=0,ending=0,monthAdditions=0,monthTransfers=0,expectedMonthlyDepreciation=0))
    validation=defaultdict(list)
    for p in payload['projects']:
        sheet=choices[p['assetId']]
        if sheet not in SHEETS:raise ValueError('잘못된 관리분류입니다.')
        changes={};details=[]
        for i,c in enumerate(p['depreciationComponents']):
            ch=components.get(p['assetId']+'-'+str(i))
            if not ch or ch['assetType'] not in engine.USEFUL_LIFE_YEARS:raise ValueError('본자산 분류를 확인하세요.')
            life=ch.get('usefulLife');kind=ch['assetType']
            if kind not in ('토지','비용처리') and (not isinstance(life,(int,float)) or not isfinite(life) or not .1<=life<=100):raise ValueError('내용연수를 확인하세요.')
            life=None if kind in ('토지','비용처리') else life
            changes[(c['assetType'],c['usefulLife'])]={**ch,'usefulLife':life}
            details.append({'assetType':kind,'usefulLife':life,'paidAmount':c['amount']})
        p['sheet']=sheet;p['business']=engine.infer_business(p,sheet)
        p['depreciationComponents']=engine.depreciation_components(p,{'details':details},use_defaults=False)
        p['expectedMonthlyDepreciation']=sum(c['monthlyDepreciation'] for c in p['depreciationComponents'])
        totals=defaultdict(float)
        for c in p['depreciationComponents']:totals[c['assetType']]+=c['amount']
        p['typeTotals']=dict(totals);p['plannedAssetType']=engine.infer_asset_type(p,totals)
        group=groups.get(p['assetId'])
        if group:
            for d in group['details']:
                change=changes.get((d['assetType'],d.get('usefulLife')))
                if change:d.update(assetType=change['assetType'],usefulLife=change['usefulLife'],vendorName=change.get('vendorName',''))
            group['plannedAssetType']=p['plannedAssetType'];validation[sheet].append(group)
        if p.get('transferBasis')=='prior-final-disappearance':continue
        b=summary[p['business']];b['count']+=1
        for key in ('ending','monthAdditions','monthTransfers','expectedMonthlyDepreciation'):b[key]+=p[key]
    payload['businessSummary']=dict(summary);payload['validationSheets']=dict(validation)
