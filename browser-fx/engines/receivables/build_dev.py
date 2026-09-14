"""Read-only cohort analysis. Writes DEV artifacts only; never rebuilds LIVE or SQLite."""
from __future__ import annotations
import calendar
import json
import sqlite3
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent

def month_end(year, month):
    return date(year, month, calendar.monthrange(year, month)[1]).isoformat()

def shifted_month_end(day, months):
    d = date.fromisoformat(day)
    y, m = divmod(d.year * 12 + d.month - 1 + months, 12)
    return month_end(y, m + 1)

def is_month_end(day):
    d = date.fromisoformat(day)
    return day == month_end(d.year, d.month)

def related(code):
    return len(code) == 4 or code in {'880010', '880011'}

def net_customer(rows, asof, fx, currency):
    docs = {}
    name = rows[0]['customer_name']
    accounts = set()
    for r in rows:
        amount = float(r['amount']) * (fx if currency == 'KRW_EQ' and r['currency'] == 'USD' else 1)
        accounts.add(r['scope'])
        key = '|'.join(str(r[k] or '') for k in ('currency', 'document_no', 'invoice_date', 'recon_account'))
        age = (date.fromisoformat(asof) - date.fromisoformat(r['invoice_date'])).days if r['invoice_date'] else None
        if key not in docs:
            docs[key] = {'key': key, 'id': r['document_no'], 'invoice': r['invoice_date'], 'age': age, 'amount': 0., 'currency': r['currency']}
        docs[key]['amount'] += amount
    signed = sum(d['amount'] for d in docs.values())
    pool = -sum(min(0., d['amount']) for d in docs.values())
    net = {}
    for d in sorted(docs.values(), key=lambda x: (-(x['age'] if x['age'] is not None else -1), x['key'])):
        if d['amount'] <= 0:
            continue
        offset = min(pool, d['amount'])
        pool -= offset
        amount = d['amount'] - offset
        if amount > 1e-7:
            net[d['key']] = {**d, 'amount': amount}
    positive = sum(d['amount'] for d in net.values())
    buckets = [0.] * 5
    unknown = 0.
    for d in net.values():
        if d['age'] is None:
            unknown += d['amount']
        else:
            age = max(0, d['age'])
            bucket = 0 if age <= 90 else 1 if age <= 180 else 2 if age <= 270 else 3 if age <= 360 else 4
            buckets[bucket] += d['amount']
    known = positive - unknown
    return {'name': name, 'total': signed, 'positive': positive, 'credit': max(0., -signed), 'buckets': buckets,
            'old90': sum(buckets[1:]), 'old180': sum(buckets[2:]), 'old360': buckets[4],
            'avg': sum(max(0, d['age']) * d['amount'] for d in net.values() if d['age'] is not None) / known if known > 0 else None,
            'max': max((max(0, d['age']) for d in net.values() if d['age'] is not None), default=None),
            'unknown': unknown, 'count': len(net), 'docs': net, 'accounts': sorted(accounts)}

EMPTY = {'docs': {}, 'positive': 0., 'total': 0., 'credit': 0., 'old90': 0., 'old180': 0., 'old360': 0., 'buckets': [0.] * 5, 'avg': None, 'unknown': 0., 'name': ''}

def cohort(start, end):
    retained = sum(min(d['amount'], end['docs'].get(k, {}).get('amount', 0.)) for k, d in start['docs'].items())
    opening = start['positive']
    recovered = max(0., opening - retained)
    additions = max(0., end['positive'] - retained)
    old_start = {k: d for k, d in start['docs'].items() if d['age'] is not None and d['age'] > 90}
    old_retained = sum(min(d['amount'], end['docs'].get(k, {}).get('amount', 0.)) for k, d in old_start.items())
    old_recovered = sum(d['amount'] for d in old_start.values()) - old_retained
    entered = end['old90'] - old_retained
    return {'opening': opening, 'retained': retained, 'recovered': recovered, 'additions': additions,
            'credit': end['credit'], 'closing': end['total'], 'oldStart': start['old90'],
            'oldRecovered': max(0., old_recovered), 'oldEntered': max(0., entered), 'oldEnd': end['old90']}

def serial_customer(code, c):
    return {'id': code, 'related': related(code), **{k: v for k, v in c.items() if k != 'docs'}}

def allowance_details(rows, asof, fx, currency, model, code, include_documents=False):
    """Reproduce the operating signed invoice-bucket basis, NOT risk FIFO netting."""
    buckets = [0.] * 5
    other = excluded = unknown = 0.
    documents = []
    for r in rows:
        if currency != 'KRW_EQ' and r['currency'] != currency:
            continue
        amount = float(r['amount']) * (fx if currency == 'KRW_EQ' and r['currency'] == 'USD' else 1)
        age = (date.fromisoformat(asof)-date.fromisoformat(r['invoice_date'])).days if r['invoice_date'] else None
        i = None if age is None else 0 if age <= 90 else 1 if age <= 180 else 2 if age <= 270 else 3 if age <= 360 else 4
        category = 'included'
        if r['scope'] != 'receivable':
            other += amount
            category = 'other'
        elif code in model.get('excludedCustomerCodes', []):
            excluded += amount
            category = 'excluded'
        elif i is None:
            unknown += amount
            category = 'unknown'
        else:
            buckets[i] += amount
        if include_documents:
            rate = model['buckets'][i]['appliedRate'] if category == 'included' else None
            documents.append({'id': r['document_no'], 'invoice': r['invoice_date'], 'age': age,
                              'currency': r['currency'], 'account': r['recon_account'], 'amount': amount,
                              'bucket': i, 'category': category, 'rate': rate,
                              'loss': amount*rate/100 if rate is not None else None})
    losses = [a*b['appliedRate']/100 for a,b in zip(buckets, model['buckets'])]
    out = {'buckets': buckets, 'losses': losses, 'exposure': sum(buckets), 'loss': sum(losses),
           'old180': sum(buckets[2:]), 'old360': buckets[4], 'other': other, 'excluded': excluded, 'unknown': unknown}
    if include_documents:
        out['documents'] = sorted(documents, key=lambda d: -(d['loss'] or 0))
    return out

def build(db_path=None, live_path=None):
    db_path = db_path or ROOT / 'private/receivables.db'
    live_path = live_path or ROOT / 'publish/채권리스크대시보드_최신.html'
    text = live_path.read_text(encoding='utf-8')
    live = json.JSONDecoder().raw_decode(text.split('const dashboardData =', 1)[1].lstrip())[0]
    meta = live['meta']
    model = meta['allowanceModel']
    fx = float(meta['usdKrwRate'])
    if fx <= 0:
        raise ValueError('운영본 환율이 올바르지 않습니다.')
    connection = sqlite3.connect(db_path.as_uri() + '?mode=ro', uri=True)
    connection.row_factory = sqlite3.Row
    raw = defaultdict(lambda: defaultdict(list))
    excluded = defaultdict(lambda: {'rows': 0, 'first': None, 'last': None})
    for r in connection.execute('SELECT snapshot_date,customer_code,customer_name,document_no,invoice_date,currency,recon_account,scope,amount FROM ar_snapshot ORDER BY snapshot_date,source_row'):
        if r['currency'] not in ('KRW', 'USD'):
            entry = excluded[r['currency']]
            entry['rows'] += 1
            entry['first'] = entry['first'] or r['snapshot_date']
            entry['last'] = r['snapshot_date']
            continue
        raw[r['snapshot_date']][r['customer_code']].append(dict(r))
    connection.close()
    dates = sorted(raw)
    latest = dates[-1]
    if latest != meta['asOf']:
        raise ValueError('운영본과 원시 적재자료 기준일이 다릅니다. 운영 데이터 갱신을 먼저 완료하세요.')
    months = [d for d in dates if is_month_end(d)]
    previous = shifted_month_end(latest, -1)
    if previous not in raw:
        previous = None
    baseline_dates = [d for d in months if d < latest][-12:]
    result = {'meta': {'asOf': latest, 'creditAsOf': meta.get('creditAsOf'), 'sourceGeneratedAt': meta['generatedAt'],
                      'generatedAt': datetime.now().strftime('%Y-%m-%d %H:%M'), 'fx': fx, 'rateDate': meta.get('rateDate'),
                      'snapshotCount': len(dates), 'startDate': dates[0], 'defaultBaseline': previous,
                      'modelDate': meta.get('allowanceModel', {}).get('modelDate'), 'method': '월말 고정 전표집단 추적',
                      'pendingCredit': None, 'excludedCurrencies': dict(excluded), 'allowanceModel': model}, 'views': {}}
    if any(e['last'] == latest for e in excluded.values()):
        raise ValueError('최신 자료에 KRW·USD 외 통화가 있습니다. 환산기준 추가 후 생성하세요.')
    # Filename inspection only. New source files are imported by the existing LIVE workflow.
    import sys
    sys.path.insert(0, str(ROOT / 'scripts'))
    from build_dashboard import parse_source_filename
    credit_dates = [parsed[1].isoformat() for p in (ROOT/'적재').glob('*.xlsx') if (parsed := parse_source_filename(p.name)) and parsed[0] == 'credit']
    newest_credit = max(credit_dates, default='')
    if newest_credit > (meta.get('creditAsOf') or ''):
        result['meta']['pendingCredit'] = newest_credit
    for currency in ('KRW_EQ', 'KRW', 'USD'):
        snapshots = {}
        for day, codes in raw.items():
            snapshots[day] = {}
            for code, rows in codes.items():
                selected = rows if currency == 'KRW_EQ' else [r for r in rows if r['currency'] == currency]
                if selected:
                    snapshots[day][code] = net_customer(selected, day, fx, currency)
        now = snapshots[latest]
        # Resolved customers from the comparison year remain visible for complete reconciliation.
        codes = sorted({code for day in snapshots.values() for code in day})
        live_period = live['currencyData'][currency]['periods'][-1]
        live_customers = {c['id']: c for c in live_period['customers']}
        current = []
        for code in codes:
            c = now.get(code)
            if c is None:
                old = next(snapshots[d][code] for d in reversed(dates) if code in snapshots[d])
                c = {**EMPTY, 'name': old['name'], 'max': None, 'count': 0, 'accounts': old['accounts']}
            row = serial_customer(code, c)
            old = live_customers.get(code, {})
            # All native financial inputs below are in their source currency, not displayed units.
            # LIVE stores USD amounts in million USD; DEV money() receives native USD.
            unit = 1e6 if currency == 'USD' else 1e8
            available = bool(old.get('behavior', {}).get('creditAvailable'))
            row['creditData'] = {'available': available, 'limit': old.get('limit', 0)*unit if available else None, 'used': old.get('used', 0)*unit if available else None}
            row['legacy'] = {'grade': old.get('grade'), 'score': old.get('riskScore'), 'behavior': old.get('behavior', {})}
            row['allowance'] = allowance_details(raw[latest].get(code, []), latest, fx, currency, model, code, True)
            row['allowance']['operatingLoss'] = (old.get('allowanceExpectedLoss') or 0)*unit
            row['allowance']['baseline'] = allowance_details(raw[model['modelDate']].get(code, []), model['modelDate'], model['modelFxRate'], currency, model, code) if model['modelDate'] in raw else None
            row['allowance']['constantFxLoss'] = allowance_details(raw[latest].get(code, []), latest, model['modelFxRate'], currency, model, code)['loss']
            row['documents'] = sorted(c['docs'].values(), key=lambda d: -d['amount'])
            current.append(row)
        baselines = []
        for day in reversed(baseline_dates):
            rows = {code: cohort(snapshots[day].get(code, EMPTY), now.get(code, EMPTY)) for code in codes}
            baselines.append({'date': day, 'days': (date.fromisoformat(latest)-date.fromisoformat(day)).days, 'rows': rows})
        trends = []
        for day in (months + ([] if latest in months else [latest])):
            # Lightweight totals per customer keep scope filters and long history correct.
            trends.append({'date': day, 'rows': {code: [c['total'], c['old90'], c['old180'], c['old360']] for code, c in snapshots[day].items()}})
        closed_cohorts = []
        for day in months[-25:]:
            if day >= latest:
                continue
            item = {'date': day, 'horizons': {}}
            for horizon in (1, 3):
                endpoint = shifted_month_end(day, horizon)
                if endpoint in snapshots and endpoint <= latest:
                    item['horizons'][str(horizon)] = {'end': endpoint, 'days': (date.fromisoformat(endpoint)-date.fromisoformat(day)).days,
                        'rows': {code: [c['positive'], cohort(c, snapshots[endpoint].get(code, EMPTY))['retained']] for code, c in snapshots[day].items()}}
            closed_cohorts.append(item)
        loss_history = []
        for day in (months[-12:] + ([] if latest in months else [latest])):
            loss_history.append({'date': day, 'rows': {code: allowance_details(rows, day, fx, currency, model, code)['loss'] for code, rows in raw[day].items()}})
        result['views'][currency] = {'customers': current, 'baselines': baselines, 'trends': trends, 'cohorts': closed_cohorts,
                                    'lossHistory': loss_history, 'operatingLoss': live_period['allowance']['expectedLoss']*unit}
    return result

def validate(data):
    checked = 0
    for currency, view in data['views'].items():
        # LIVE rounded native KRW / million-USD displays before consolidation.
        rounding = 10000 if currency == 'KRW' else 100 if currency == 'USD' else 10000 + 100*data['meta']['fx']
        for c in view['customers']:
            a = c['allowance']
            assert abs(sum(d['amount'] for d in c['documents'])-c['positive']) < .05
            assert abs(sum(a['buckets'])-a['exposure']) < .05
            assert abs(sum(d['loss'] or 0 for d in a['documents'])-a['loss']) < .05
            assert abs(a['exposure']+a['other']+a['excluded']+a['unknown']-c['total']) < .05
            assert abs(a['loss']-a['operatingLoss']) <= rounding, f'대손 운영모형 대사 오류: {currency} {c["id"]}'
        for baseline in view['baselines']:
            for c in baseline['rows'].values():
                assert abs(c['opening']-c['recovered']+c['additions']-c['credit']-c['closing']) < .05
                assert abs(c['oldStart']-c['oldRecovered']+c['oldEntered']-c['oldEnd']) < .05
                assert -.01 <= c['retained'] <= c['opening']+.01
                checked += 1
    return checked

def main():
    import sys
    if '--layout-only' in sys.argv:
        existing = (HERE/'publish/receivables-dev.html').read_text(encoding='utf-8')
        data = json.JSONDecoder().raw_decode(existing.split('const DATA=', 1)[1].lstrip())[0]
    else:
        data = build()
    checks = validate(data)
    payload = json.dumps(data, ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/')
    template = (HERE/'dashboard.html').read_text(encoding='utf-8')
    assert template.count('__RECEIVABLES_DEV_DATA__') == 1
    html = template.replace('__RECEIVABLES_DEV_DATA__', payload)
    if (HERE/'workspace.css').exists():
        html = html.replace('</style>', (HERE/'workspace.css').read_text(encoding='utf-8')+'\n</style>', 1)
    if (HERE/'workspace.js').exists():
        html = html.replace('refreshBaseline();render();\n</script>', (HERE/'workspace.js').read_text(encoding='utf-8')+'\nrefreshBaseline();render();\n</script>')
    output = HERE/'publish'
    output.mkdir(exist_ok=True)
    temp = output/'receivables-dev.html.tmp'
    temp.write_text(html, encoding='utf-8')
    temp.replace(output/'receivables-dev.html')
    (output/'validation.json').write_text(json.dumps({'asOf': data['meta']['asOf'], 'checks': checks, 'bytes': len(html.encode('utf-8')), 'method': data['meta']['method']}, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'채권 DEV 생성 완료 · {checks:,}개 고객·기준일 대사 · {len(html.encode("utf-8"))/1024/1024:.1f}MB')

if __name__ == '__main__':
    main()
