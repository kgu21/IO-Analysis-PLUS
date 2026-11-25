import json, io
import numpy as np
import openpyxl

book = None
codes, names, sheets = [], [], []
L = M = V = None
valid = False
level_label = ''


def emit(x):
    return json.dumps(x, ensure_ascii=False, allow_nan=False)


def clean(x):
    return str(x).strip()


def is_total_code(c):
    # 9090(중간수요계), 9590(소계), 9690(부가가치계) 등 합계·집계 코드
    s = clean(c)
    return s.isdigit() and int(s) >= 9000


def initialize(raw, label=''):
    """엑셀(대/중/소분류·기본부문 공통 구조)을 읽고 계수를 검증한다."""
    global book, codes, names, L, M, V, sheets, valid, level_label
    level_label = label
    wb = openpyxl.load_workbook(io.BytesIO(bytes(raw)), read_only=True, data_only=True)
    book = {s: list(wb[s].values) for s in wb.sheetnames}
    wb.close()
    sheets = list(book)
    needed = ['국산투입계수(Ad)', '수입투입계수(Am)', '총투입계수(A)',
              '생산유발계수', '수입유발계수', '부가가치유발계수']
    missing = [s for s in needed if s not in book]
    if missing:
        raise ValueError('필수 시트가 없습니다: ' + ', '.join(missing))

    base = book['국산투입계수(Ad)']
    # 산업 코드는 5번째 행(인덱스 4)의 3열부터 합계 코드(9000 이상) 직전까지
    codes = []
    for c in base[4][2:]:
        if c is None:
            continue
        if is_total_code(c):
            break
        codes.append(clean(c))
    if len(codes) < 2 or len(set(codes)) != len(codes):
        raise ValueError('산업 코드를 읽지 못했습니다. 투입산출표 구조를 확인하세요.')
    nm = {clean(r[0]): clean(r[1]) for r in base[6:]
          if r and r[0] is not None and len(r) > 1 and r[1] is not None}
    names = [nm[c] for c in codes]

    def matrix(s):
        rows = book[s]
        ri = {clean(r[0]): i for i, r in enumerate(rows) if r and r[0] is not None}
        ci = {clean(c): i for i, c in enumerate(rows[4]) if c is not None}
        a = np.array([[rows[ri[c]][ci[k]] for k in codes] for c in codes], dtype=float)
        if not np.isfinite(a).all():
            raise ValueError(s + ': 비수치 또는 결측 데이터')
        return a

    L = matrix('생산유발계수'); M = matrix('수입유발계수'); V = matrix('부가가치유발계수')
    ad = matrix('국산투입계수(Ad)'); am = matrix('수입투입계수(Am)')
    rows = book['총투입계수(A)']
    ci = {clean(c): i for i, c in enumerate(rows[4]) if c is not None}
    vr = next(r for r in rows if r and r[0] is not None and clean(r[0]) == '9690')
    v = np.array([vr[ci[c]] for c in codes], dtype=float)
    n = len(codes)
    residuals = {'생산': float(np.max(np.abs((np.eye(n) - ad) @ L - np.eye(n)))),
                 '수입': float(np.max(np.abs(am @ L - M))),
                 '부가가치': float(np.max(np.abs(v[:, None] * L - V)))}
    valid = max(residuals.values()) < 1e-8
    if not valid:
        raise ValueError('계수 검증 실패: ' + str(residuals))
    return emit({'sectors': [{'code': c, 'name': nme} for c, nme in zip(codes, names)],
                 'sheets': sheets, 'residuals': residuals, 'label': label})


def explore(sheet, rowcode):
    rows = book[sheet]
    choices = [{'code': clean(r[0]), 'name': clean(r[1])} for r in rows[6:]
               if r and len(r) > 1 and r[0] is not None and r[1] is not None]
    chosen = next((r for r in rows[6:] if r and r[0] is not None and clean(r[0]) == rowcode), None)
    if chosen is None:
        chosen = next(r for r in rows[6:] if r and r[0] is not None and r[1] is not None)
    cells = []
    for i in range(2, len(rows[4])):
        if rows[4][i] is None:
            continue
        val = chosen[i] if i < len(chosen) else None
        if isinstance(val, (float, int)):
            val = float(val)
        elif val is not None:
            val = str(val)
        cells.append({'code': clean(rows[4][i]), 'name': str(rows[5][i] or '').strip(), 'value': val})
    return emit({'rows': choices, 'selected': clean(chosen[0]), 'name': clean(chosen[1]),
                 'unit': str(rows[3][0] or '係数').replace('係数', '계수 / 비율: 행·열 명칭 확인'),
                 'cells': cells})


def simulate(payload):
    if not valid:
        raise ValueError('검증된 데이터가 없습니다.')
    shocks = json.loads(payload); f = np.zeros(len(codes))
    for s in shocks:
        a = float(s['amount'])
        if not np.isfinite(a) or a < 0:
            raise ValueError('수요 증가액은 유한한 0 이상의 수여야 합니다.')
        f[codes.index(s['code'])] += a
    if not f.sum() > 0:
        raise ValueError('한 산업 이상의 양의 수요 증가액을 입력하세요.')
    x = L @ f; m = M @ f; v = V @ f
    result = [{'code': c, 'name': nme, 'production': float(a), 'imports': float(b), 'valueAdded': float(d)}
              for c, nme, a, b, d in zip(codes, names, x, m, v)]
    result.sort(key=lambda r: r['production'], reverse=True)
    return emit({'demand': float(f.sum()), 'production': float(x.sum()), 'imports': float(m.sum()),
                 'valueAdded': float(v.sum()), 'multiplier': float(x.sum() / f.sum()), 'rows': result})


def analysis():
    """산업별 유발계수와 전·후방 연쇄효과(영향력·감응도계수)."""
    if not valid:
        raise ValueError('검증된 데이터가 없습니다.')
    colL = L.sum(axis=0); rowL = L.sum(axis=1)
    influence = colL / colL.mean()      # 영향력계수: 후방연쇄효과 (열합 / 열합 평균)
    sensitivity = rowL / rowL.mean()    # 감응도계수: 전방연쇄효과 (행합 / 행합 평균)
    colM = M.sum(axis=0); colV = V.sum(axis=0)
    rows = []
    for i, (c, nme) in enumerate(zip(codes, names)):
        hi_b = influence[i] >= 1; hi_f = sensitivity[i] >= 1
        kind = ('전·후방 연쇄 모두 강함' if hi_b and hi_f else
                '후방 연쇄 강함(영향력 우위)' if hi_b else
                '전방 연쇄 강함(감응도 우위)' if hi_f else '전·후방 연쇄 모두 약함')
        rows.append({'code': c, 'name': nme, 'production': float(colL[i]), 'imports': float(colM[i]),
                     'valueAdded': float(colV[i]), 'influence': float(influence[i]),
                     'sensitivity': float(sensitivity[i]), 'kind': kind})
    return emit({'rows': rows})


def demand_items():
    """최종수요 항목(소비·투자·수출 등)별 유발액·유발계수와 산업별 상세."""
    if not valid:
        raise ValueError('검증된 데이터가 없습니다.')
    need = ['최종수요 항목별 생산유발액', '최종수요 항목별 수입유발액', '최종수요 항목별 부가가치유발액',
            'A표_국산거래표(생산자)', 'A표_수입거래표(생산자)']
    missing = [s for s in need if s not in book]
    if missing:
        raise ValueError('필수 시트가 없습니다: ' + ', '.join(missing))

    def index(rows):
        ci = {clean(c): i for i, c in enumerate(rows[4]) if c is not None}
        ri = {clean(r[0]): r for r in rows if r and r[0] is not None}
        return ci, ri

    prod, imp, vad = (book[s] for s in need[:3])
    pci, pri = index(prod); mci, mri = index(imp); vci, vri = index(vad)
    dci, dri = index(book[need[3]]); ici, iri = index(book[need[4]])
    item_codes = [clean(c) for c in prod[4][2:] if c is not None]
    code_set = set(codes)

    def num(v):
        return float(v) if isinstance(v, (int, float)) else 0.0

    items, detail = [], {}
    for ic in item_codes:
        name = ' '.join(str(prod[5][pci[ic]] or ic).split())
        dom = num(dri['9590'][dci[ic]]); imf = num(iri['9590'][ici[ic]])
        P = num(pri['9819'][pci[ic]]); Mv = num(mri['9819'][mci[ic]]); Vv = num(vri['9819'][vci[ic]])
        ok = dom > 0
        items.append({
            'code': ic, 'name': name, 'domestic': dom, 'imported': imf,
            'production': P, 'imports': Mv, 'valueAdded': Vv,
            'coefProduction': P / dom if ok else None,
            'coefValueAdded': Vv / dom if ok else None,
            'coefImportIntermediate': (Mv - imf) / dom if ok else None,
            'coefImportPublished': Mv / (dom + imf) if (dom + imf) > 0 else None})
        rows = []
        for c, nme in zip(codes, names):
            rows.append({'code': c, 'name': nme,
                         'production': num(pri[c][pci[ic]]), 'imports': num(mri[c][mci[ic]]),
                         'valueAdded': num(vri[c][vci[ic]])})
        detail[ic] = rows
    return emit({'items': items, 'detail': detail})
