"""Only replaces file access with supplied synthetic rows; accounting code is unchanged."""
import json
import math
from pathlib import Path
import build_dashboard as engine


class MemorySheet:
    def __init__(self, rows):
        self.rows = rows

    def iter_rows(self, **kwargs):
        return iter(self.rows)

    def close(self):
        pass


def calculate(input_json, settings_json='{}'):
    rows = json.loads(input_json)
    if not isinstance(rows, dict) or not rows:
        raise ValueError('가상 원장 입력이 없습니다.')
    for name, sheet in rows.items():
        if Path(name).name != name or not name.endswith('.xlsx'):
            raise ValueError('잘못된 논리 시트명입니다.')
        if not isinstance(sheet, list) or not sheet:
            raise ValueError('원장 행이 없습니다: ' + name)
    settings = json.loads(settings_json)
    if settings:
        raise ValueError('시제품에서는 운영 설정을 받지 않습니다.')
    root = Path('/tmp/finance-browser-fx')
    root.mkdir(parents=True, exist_ok=True)
    (root / 'dashboard-settings.json').write_text('{}', encoding='utf8')
    engine.discover_inputs = lambda _: {Path(name).stem: Path(name) for name in rows}
    engine.load_single_sheet = lambda path, **kwargs: (MemorySheet([]), MemorySheet(rows[Path(path).name]))
    payload = engine.build_payload(root)
    for check in payload['checks']:
        if not all(math.isfinite(check[k]) and abs(check[k]) < .01 for k in ('valuationDifference', 'realizedDifference')):
            raise ValueError('원장 Net 대사 불일치: ' + check['month'])
    payload['source']['sourceWorkbook'] = '가상 원장 · 체험용'
    return json.dumps(payload, ensure_ascii=False, allow_nan=False, separators=(',', ':'))
