"""Management classification with explicit provenance, never a current answer lookup."""
from collections import defaultdict
import json
import re

AREA = {'1100': 'HQ', '1120': '화학', '1130': '디스플레이'}
SHEETS = {'HQ', '화학_QL', '화학_기타', 'DP_8.6G', 'DP_기타'}


def confirmed_rules(path, period):
    """Optional user-confirmed classifications, effective-dated to prevent leakage."""
    if not path.exists():
        return []
    payload = json.loads(path.read_text(encoding='utf-8-sig'))
    if payload.get('version') != 1 or not isinstance(payload.get('rules'), list):
        raise ValueError('관리분류 기준 파일 형식 오류')
    selected, keys = {}, set()
    for rule in payload['rules']:
        if rule.get('keyType') not in ('assetId', 'wbs') or not rule.get('key') or rule.get('sheet') not in SHEETS:
            raise ValueError('관리분류 기준의 키 또는 시트 오류')
        if not re.fullmatch(r'\d{4}-(0[1-9]|1[0-2])', rule.get('effectiveFrom', '')):
            raise ValueError('관리분류 적용월 오류')
        if not rule.get('reason') or not rule.get('confirmedBy'):
            raise ValueError('관리분류 확정 근거 및 확인자 필요')
        if rule['effectiveFrom'] > f'{period[0]}-{period[1]:02}':
            continue
        key = (rule['keyType'], rule['key'].upper())
        dated_key = (*key, rule['effectiveFrom'])
        if dated_key in keys:
            raise ValueError(f'관리분류 확정 기준 중복: {dated_key}')
        keys.add(dated_key)
        if key not in selected or selected[key]['effectiveFrom'] < rule['effectiveFrom']:
            selected[key] = rule
    return list(selected.values())


def business_of_sheet(sheet):
    return ('HQ' if sheet.startswith('HQ') else '화학' if sheet.startswith('화학')
            else '디스플레이' if sheet.startswith(('DP', '6G')) else None)


def project_index(groups, prior_assets):
    index = defaultdict(set)
    for asset in prior_assets:
        group = groups.get(asset['assetId'])
        if group and asset.get('wbs'):
            index[asset['wbs'].upper()].add(group['sheet'])
    return dict(index)


def classify(asset, group=None, projects=None, confirmed=None):
    asset = asset or {}
    name = asset.get('name', '').upper()
    area = AREA.get(asset.get('businessArea'))
    named = {business for token, business in [('CT_IN_CH', '화학'), ('CT_IN_DP', '디스플레이')]
             if token in name}
    conflicts = []
    # Explicit confirmation is scoped to an exact asset or entire WBS, never
    # a name fragment. It enables classification without a current final file.
    for key_type in ('assetId', 'wbs'):
        matches = [r for r in confirmed or [] if r['keyType'] == key_type
                   and r['key'].upper() == asset.get(key_type, '').upper()]
        if matches:
            r = matches[0]
            return dict(sheet=r['sheet'], business=business_of_sheet(r['sheet']),
                        basis=f'확정 분류기준 {key_type}: {r["reason"]}',
                        status='기준표 확정', conflicts=[], confirmation=r)
    if area and named and named != {area}:
        conflicts.append('사업영역 코드와 자산명 조직표시 불일치')
    sheet = group.get('sheet') if group else None
    basis = '전월 확정 자산번호 분류' if sheet else None
    if not sheet:
        matches = (projects or {}).get(asset.get('wbs', '').upper(), set())
        if len(matches) > 1:
            return dict(sheet=None, business=None, basis='동일 WBS 분류 충돌',
                        status='확인 필요', conflicts=sorted(matches))
        if len(matches) == 1:
            sheet = next(iter(matches))
            basis = '전월 확정 동일 WBS 분류'
    if sheet:
        business = business_of_sheet(sheet)
        if (area and area != business) or (named and named != {business}):
            conflicts.append('확정 관리분류와 원천 조직정보 불일치: 확정 분류 유지')
        return dict(sheet=sheet, business=business, basis=basis,
                    status='승계·예외 확인' if conflicts else '승계', conflicts=conflicts)
    if conflicts or len(named) > 1:
        return dict(sheet=None, business=None, basis='원천 분류 충돌',
                    status='확인 필요', conflicts=conflicts)
    business = area or next(iter(named), None)
    if business == 'HQ':
        sheet = 'HQ'
    elif business == '화학':
        sheet = '화학_QL' if re.search(r'(?<![A-Z])QL(?![A-Z])', name) else '화학_기타'
    elif business == '디스플레이':
        sheet = 'DP_8.6G' if '8.6G' in name else 'DP_기타'
    return dict(sheet=sheet, business=business,
                basis='자산대장 BusA / 자산명 조직·프로젝트 표시',
                status='신규 분류안·확인 필요' if sheet else '확인 필요', conflicts=[])
