"""분류 비교용 요약 자료(data/comparison.json) 생성 스크립트.

사용법 (저장소 최상위에서):
    python tools/build_comparison.py "상품분류.xlsx 경로"

필요: python3, numpy, openpyxl. data/ 폴더에 io-2023-{large,medium,small,basic}.xlsx가 있어야 합니다.
웹앱과 같은 model.py의 계산 함수를 그대로 사용하므로, 앱에서 각 분류를 불러왔을 때의 값과 동일합니다.
"""
import json, os, sys, importlib.util
import openpyxl

root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
spec = importlib.util.spec_from_file_location('model', os.path.join(root, 'model.py'))
model = importlib.util.module_from_spec(spec)
spec.loader.exec_module(model)

LEVELS = [('large', '대분류'), ('medium', '중분류'), ('small', '소분류'), ('basic', '기본부문')]
clean = lambda x: str(x).strip()


def build_hierarchy(path):
    """상품분류표(부문분류 시트)에서 기본부문 → 소·중·대분류 대응표를 만든다."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    rows = list(wb['부문분류'].values)[3:]
    wb.close()
    small = medium = large = None
    hierarchy = []
    for r in rows:
        if r[2] is not None: small = clean(r[2])
        if r[4] is not None: medium = clean(r[4])
        if r[6] is not None: large = clean(r[6])
        b = clean(r[0]) if r[0] is not None else ''
        if b.isdigit() and int(b) < 9000:
            hierarchy.append({'basic': b, 'small': small, 'medium': medium, 'large': large})
    return hierarchy


def main():
    class_path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(root, 'data', 'class.xlsx')
    out = {'source': '한국은행 2023년 산업연관표(연장표), 생산자가격', 'levels': {}}
    for key, label in LEVELS:
        raw = open(os.path.join(root, 'data', 'io-2023-%s.xlsx' % key), 'rb').read()
        model.initialize(raw, label)
        sectors = json.loads(model.analysis())['rows']
        for s in sectors:
            s.pop('kind', None)
        items = json.loads(model.demand_items())['items']
        out['levels'][key] = {'label': label, 'sectors': sectors, 'items': items}
        print(label, len(sectors), '개 부문')
    hierarchy = build_hierarchy(class_path)
    codes = {k: {s['code'] for s in v['sectors']} for k, v in out['levels'].items()}
    assert len(hierarchy) == len(codes['basic']), '기본부문 수 불일치'
    for h in hierarchy:
        for k in ('basic', 'small', 'medium', 'large'):
            assert h[k] in codes[k], '분류 대응 오류: %s %s' % (k, h)
    for k in ('small', 'medium', 'large'):
        assert {h[k] for h in hierarchy} == codes[k], k + ' 분류에 연결되지 않은 부문이 있습니다.'
    out['hierarchy'] = hierarchy
    path = os.path.join(root, 'data', 'comparison.json')
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, separators=(',', ':'))
    print('저장:', path, os.path.getsize(path), 'bytes')


if __name__ == '__main__':
    main()
