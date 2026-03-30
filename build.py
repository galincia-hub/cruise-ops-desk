#!/usr/bin/env python3
"""
build.py — 모두의 크루즈 운영 데스크 HTML 빌더 v3.0
Usage: python build.py
  data/*.json → index.html
"""

import json, os, sys, re, glob as globmod

BASE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(BASE, 'data')
OUT  = os.path.join(BASE, 'index.html')
ARCHIVE_DIR = os.path.join(BASE, 'archive')
GITHUB_PAGES_BASE = 'https://galincia-hub.github.io/cruise-ops-desk'

# ── 어두운 배경색 목록 (헤더/제목 행 판별용) ──────────────────────────
DARK_BGS = {c.upper() for c in [
    '#0D1B2A','#1B2A4A','#2C3E50','#34495E',
    '#E67E22','#E74C3C','#27AE60','#8E44AD',
    '#2980B9','#1ABC9C','#6C5CE7','#D35400',
    '#00695C','#1565C0','#C62828','#F39C12',
    '#00BCD4','#455A64','#37474F','#263238',
    '#AD1457','#00838F','#558B2F','#E65100',
    '#4527A0','#283593','#00695C','#2E7D32',
    '#BF360C','#4E342E','#37474F','#546E7A',
]}

# ── 섹션 타이틀 채도 낮춘 모듈 색상 ─────────────────────────────────
MODULE_SECTION_COLORS = {
    'master':     '#4A5568',
    'org':        '#2B6CB0',
    'prog':       '#C05621',
    'sup':        '#2F855A',
    'emb':        '#4C51BF',
    'port':       '#276749',
    'qa':         '#4A5568',
    'orgv2':      '#1565C0',
    'orgv3':      '#283593',
    'supplement': '#00695C',
    'teamdocs':   '#4527A0',
    'matrix':     '#BF360C',
}

# ── 모듈별 메타정보 ───────────────────────────────────────────────────
MODULE_META = {
    'master':     {'icon':'📋','color':'#0D1B2A','desc':'D-90부터 D-Day까지 전체 준비사항'},
    'org':        {'icon':'🏢','color':'#2980B9','desc':'HQ 중심 선내운영 조직 체계'},
    'prog':       {'icon':'🎭','color':'#E67E22','desc':'공연·이벤트·환경재단행사 통합운영'},
    'sup':        {'icon':'💻','color':'#1ABC9C','desc':'IT·총무·홍보 통합 후방지원'},
    'emb':        {'icon':'🚢','color':'#6C5CE7','desc':'2,400명 승선·하선 총괄 운영'},
    'port':       {'icon':'⚓','color':'#27AE60','desc':'하코다테·오타루 기항 운영'},
    'qa':         {'icon':'❓','color':'#34495E','desc':'선행 항차 경험자 자문용'},
    'orgv2':      {'icon':'📊','color':'#1565C0','desc':'확정 조직도 + 포지션별 역량 정의'},
    'orgv3':      {'icon':'📝','color':'#283593','desc':'기안·공문·온보드미팅 준비'},
    'supplement': {'icon':'📎','color':'#00695C','desc':'무전기·Wi-Fi·비상연락 등 보강자료'},
    'teamdocs':   {'icon':'📑','color':'#4527A0','desc':'팀별 업무분장서 6종'},
    'matrix':     {'icon':'📐','color':'#BF360C','desc':'전체 업무 일람 + 진행상태 추적'},
}

# ── 사이드바 메뉴 구조 정의 ───────────────────────────────────────────
# (section_title, [(module_key, label_override)], ...)
# label_override가 None이면 manifest label 사용
SIDEBAR_SECTIONS = [
    ('운영 매뉴얼 (초안)', [
        ('master', None),
        ('org', None),
        ('prog', None),
        ('sup', 'IT홍보팀'),
        ('emb', None),
        ('port', None),
    ]),
    ('조직·기획', [
        ('orgv2', None),
        ('orgv3', '기안·온보드미팅'),
        ('teamdocs', None),
        ('matrix', None),
    ]),
    ('보강자료', [
        ('supplement', None),
    ]),
    ('참고자료', [
        ('qa', None),
    ]),
]

# master에서 제외할 시트
MASTER_EXCLUDE_SHEETS = ['인력운영안(원본)']

# ── 유틸리티 ─────────────────────────────────────────────────────────
def esc(v):
    """HTML 이스케이프 + \\n → <br>"""
    if v is None:
        return ''
    s = str(v).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
    return s.replace('\n','<br>')

def is_dark(bg):
    return bool(bg) and bg.upper() in DARK_BGS

def cell_style(c, include_bg=True, include_color=True):
    parts = []
    if include_bg and c.get('bg'):
        parts.append(f"background:{c['bg']}")
    if include_color and c.get('fc'):
        parts.append(f"color:{c['fc']}")
    if c.get('bold'):
        parts.append('font-weight:700')
    return ';'.join(parts)

# ── 아카이브 스캔 ────────────────────────────────────────────────────
def scan_archives():
    """archive/ 폴더를 스캔하여 (version, date, filename) 리스트를 최신순 반환."""
    archives = []
    if not os.path.isdir(ARCHIVE_DIR):
        return archives
    for fname in os.listdir(ARCHIVE_DIR):
        m = re.match(r'v([\d.]+)_(\d{4}-\d{2}-\d{2})\.html$', fname)
        if m:
            archives.append((m.group(1), m.group(2), fname))
    archives.sort(key=lambda x: x[1], reverse=True)
    return archives

# ── 역량 정의서 전치 렌더링 ──────────────────────────────────────────
def render_competency_transpose(sheet, module_key):
    data = sheet.get('data', [])
    if not data:
        return '', ''

    page_title = data[0][0].get('value', '')

    positions = []
    cur = None
    for row in data[1:]:
        if not row:
            continue
        single = (len(row) == 1)
        all_dark = all(is_dark(c.get('bg', '')) for c in row)
        if single and (all_dark or is_dark(row[0].get('bg', ''))):
            cur = {'name': row[0].get('value', '').replace('\n', ' '), 'rows': {}}
            positions.append(cur)
        elif cur is not None and len(row) >= 2:
            label = row[0].get('value', '').replace('\n', ' ').strip()
            value = row[1].get('value', '')
            cur['rows'][label] = value

    if not positions:
        return page_title, ''

    row_labels = list(positions[0]['rows'].keys())

    sec_color = MODULE_SECTION_COLORS.get(module_key, '#4A5568')
    out = []
    out.append('<div class="table-wrap">')
    out.append(f'<div class="section-title" style="background:{sec_color}">포지션별 역량 비교표</div>')
    out.append('<table class="ops-table">')

    out.append('<thead><tr>')
    out.append('<th style="min-width:140px">구분</th>')
    for label in row_labels:
        out.append(f'<th>{esc(label)}</th>')
    out.append('</tr></thead>')

    out.append('<tbody>')
    for pos in positions:
        name = esc(pos['name'].lstrip('▣').strip())
        out.append('<tr>')
        out.append(f'<td style="font-weight:700;white-space:nowrap;text-align:center" class="cell-wrap">{name}</td>')
        for label in row_labels:
            val = pos['rows'].get(label, '')
            out.append(f'<td style="text-align:left" class="cell-wrap">{esc(val)}</td>')
        out.append('</tr>')
    out.append('</tbody>')

    out.append('</table>')
    out.append('</div>')

    return page_title, '\n'.join(out)

# ── 시트 렌더링 핵심 로직 ─────────────────────────────────────────────
def render_sheet(sheet, module_key):
    data = sheet.get('data', [])
    out = []
    page_title = ''
    sections = []
    cur = None

    for i, row in enumerate(data):
        if not row:
            continue

        if i == 0:
            page_title = row[0].get('value', '')
            continue

        all_dark = all(is_dark(c.get('bg', '')) for c in row)
        single   = (len(row) == 1)

        if single and all_dark:
            cur = {'title': row[0], 'rows': []}
            sections.append(cur)
        else:
            if cur is None:
                cur = {'title': None, 'rows': []}
                sections.append(cur)
            rtype = 'h' if all_dark else 'd'
            cur['rows'].append((rtype, row))

    for sec in sections:
        out.append('<div class="table-wrap">')

        if sec['title']:
            sec_color = MODULE_SECTION_COLORS.get(module_key, '#4A5568')
            val = esc(sec['title'].get('value', ''))
            out.append(f'<div class="section-title" style="background:{sec_color}">{val}</div>')

        rows = sec['rows']
        if rows:
            col_count = next((len(cells) for rtype, cells in rows if rtype == 'h'), 0)

            LEFT_ALIGN_KEYWORDS = ['세부', '내용', '체크리스트', '항목', '업무']
            left_cols = set()
            for rtype, cells in rows:
                if rtype == 'h':
                    for ci, c in enumerate(cells):
                        hval = str(c.get('value', ''))
                        if any(kw in hval for kw in LEFT_ALIGN_KEYWORDS):
                            left_cols.add(ci)

            out.append('<table class="ops-table">')
            in_head = False
            in_body = False

            for rtype, cells in rows:
                if rtype == 'h':
                    if in_body:
                        out.append('</tbody>')
                        in_body = False
                    if not in_head:
                        out.append('<thead>')
                        in_head = True
                    out.append('<tr>')
                    for c in cells:
                        out.append(f'<th style="{cell_style(c, include_bg=False)}">{esc(c.get("value",""))}</th>')
                    out.append('</tr>')

                else:
                    if in_head:
                        out.append('</thead>')
                        in_head = False
                    if not in_body:
                        out.append('<tbody>')
                        in_body = True

                    n = len(cells)

                    if col_count > 0 and n < col_count / 2:
                        val = '<br>'.join(
                            esc(c.get('value', '')) for c in cells if c.get('value')
                        )
                        out.append('<tr>')
                        out.append(f'<td colspan="{col_count}" class="cell-wrap merged-desc">{val}</td>')
                        out.append('</tr>')

                    else:
                        rb = cells[0].get('bg', '')
                        row_style = f' style="--row-bg:{rb}"' if rb else ''
                        out.append(f'<tr{row_style}>')
                        for idx, c in enumerate(cells):
                            cls = "cell-wrap text-left" if idx in left_cols else "cell-wrap"
                            out.append(f'<td style="{cell_style(c, include_bg=False, include_color=False)}" class="{cls}">{esc(c.get("value",""))}</td>')
                        out.append('</tr>')

            if in_head: out.append('</thead>')
            if in_body: out.append('</tbody>')
            out.append('</table>')

        out.append('</div>')

    return page_title, '\n'.join(out)

# ── 데이터 로드 ───────────────────────────────────────────────────────
def load_all():
    with open(os.path.join(DATA, 'manifest.json'), encoding='utf-8') as f:
        manifest = json.load(f)

    modules = {}
    for key in manifest:
        path = os.path.join(DATA, f'{key}.json')
        if os.path.exists(path):
            with open(path, encoding='utf-8') as f:
                modules[key] = json.load(f)
            print(f'  loaded {key}.json ({len(modules[key]["sheets"])} sheets)')
        else:
            print(f'  [WARN] {key}.json not found', file=sys.stderr)

    return manifest, modules

# ── 사이드바 네비게이션 생성 ──────────────────────────────────────────
def _nav_group(lines, key, manifest, label_override=None):
    """단일 nav-group을 lines에 추가."""
    if key not in manifest:
        return
    info  = manifest[key]
    meta  = MODULE_META.get(key, {'icon':'📄','color':'#2C3E50'})
    icon  = meta['icon']
    label = label_override or info['label']

    # master에서 제외할 시트 필터링
    sheets = info['sheets']
    if key == 'master':
        sheets = [s for s in sheets if s not in MASTER_EXCLUDE_SHEETS]

    lines.append(f'    <div class="nav-group">')
    lines.append(f'      <div class="nav-group-header" onclick="toggleGroup(this)"><span class="icon">{icon}</span>{esc(label)}<span class="arrow">▶</span></div>')
    lines.append(f'      <div class="nav-sub">')

    # 시트 인덱스는 원본 manifest 기준 (page-id 매칭)
    for orig_idx, sheet_name in enumerate(info['sheets'], 1):
        if sheet_name in MASTER_EXCLUDE_SHEETS and key == 'master':
            continue
        pid = f'{key}-{orig_idx}'
        lines.append(f'        <div class="nav-item" onclick="showPage(\'{pid}\')">{esc(sheet_name)}</div>')
    lines.append(f'      </div>')
    lines.append(f'    </div>')

def build_nav(manifest):
    lines = [
        '<nav class="sidebar" id="sidebar">',
        '  <div class="nav-section">',
        '    <div class="nav-section-title">대시보드</div>',
        '    <div class="nav-item active" onclick="showPage(\'home\')"><span class="icon">🏠</span>홈</div>',
        '  </div>',
    ]

    for section_title, modules_list in SIDEBAR_SECTIONS:
        lines.append('  <div class="nav-section">')
        lines.append(f'    <div class="nav-section-title">{esc(section_title)}</div>')
        for key, label_override in modules_list:
            _nav_group(lines, key, manifest, label_override)
        lines.append('  </div>')

    # ── 아카이브 섹션 ──
    archives = scan_archives()
    lines.append('  <div class="nav-section">')
    lines.append('    <div class="nav-section-title">📁 아카이브 (버전 기록)</div>')
    for ver, date, fname in archives:
        url = f'{GITHUB_PAGES_BASE}/archive/{fname}'
        lines.append(f'    <div class="nav-item" onclick="window.open(\'{url}\',\'_blank\')">v{ver} ({date})</div>')
    if not archives:
        lines.append('    <div class="nav-item" style="color:var(--text-secondary);cursor:default;font-style:italic">아직 아카이브 없음</div>')
    lines.append('  </div>')

    lines.append('</nav>')
    return '\n'.join(lines)

# ── 조직도 시각화 HTML ────────────────────────────────────────────────
def build_org_chart():
    return """
    <div class="org-chart-section">
      <h2 style="font-size:18px;font-weight:800;margin-bottom:16px;color:var(--text-primary)">📊 조직도 v4 — 선내운영 조직 구조</h2>
      <div class="org-chart">
        <!-- HQ -->
        <div class="org-node org-hq">
          <div class="org-node-title">HQ 운영사령실</div>
          <div class="org-node-sub">공동부서장: CI + 모두투어</div>
        </div>

        <div class="org-branches">
          <!-- 직속 그룹 -->
          <div class="org-branch">
            <div class="org-branch-label">직속</div>

            <div class="org-node org-support">
              <div class="org-node-title">운영지원팀</div>
              <div class="org-node-members">황지애 대리<span class="org-tag tag-ci">CI원</span> · 양은희 매니저<span class="org-tag tag-modo">모두</span><br>조아라 매니저<span class="org-tag tag-派遣">웅진파견</span> · 김지은 선임<span class="org-tag tag-派遣">재단파견</span></div>
              <div class="org-node-role">고객응대 + 방송안내 + 정산관리 (팀 과업, 유동배치)</div>
              <div class="org-node-note">VIP: 부서장 직속 관리</div>
            </div>

            <div class="org-node org-fb">
              <div class="org-node-title">식음료파트</div>
              <div class="org-node-members">문형식 매니저 <span class="org-tag tag-pre">★사전탑승</span></div>
            </div>

            <div class="org-node org-it">
              <div class="org-node-title">IT홍보팀</div>
              <div class="org-node-members">최구철 매니저 <span class="org-tag tag-派遣">IT팀+대외협력팀 파견</span></div>
            </div>

            <div class="org-node org-alba">
              <div class="org-node-title">알바 12명</div>
              <div class="org-node-role">탄력 배치</div>
            </div>
          </div>

          <!-- 팀 그룹 -->
          <div class="org-branch">
            <div class="org-branch-label">팀</div>

            <div class="org-node org-event">
              <div class="org-node-title">공연/행사팀</div>
              <div class="org-node-members">이상민 매니저 <span class="org-tag tag-pre">★사전탑승</span> + 팀원4</div>
              <div class="org-node-role">환경재단 15명 협업 (재단 운영팀장 포함)</div>
            </div>

            <div class="org-node org-port">
              <div class="org-node-title">기항지운영팀</div>
              <div class="org-node-members">이수일 팀장 <span class="org-tag tag-dual">★듀얼</span> + 팀원1 + 타이요(6+가이드70)</div>
              <div class="org-sub-node">
                <div class="org-node org-embark">
                  <div class="org-node-title">승하선팀 (기항지팀 산하)</div>
                  <div class="org-node-members">김남민 매니저 <span class="org-tag tag-off">미탑승/터미널전담</span><br>이수일 <span class="org-tag tag-dual">★듀얼</span> + 알바 3~4</div>
                </div>
              </div>
            </div>
          </div>

          <!-- 임시 조직 -->
          <div class="org-branch">
            <div class="org-branch-label">임시</div>
            <div class="org-node org-temp">
              <div class="org-node-title">부산지점 사전준비팀</div>
              <div class="org-node-sub">(탑승 전 임시조직)</div>
              <div class="org-node-members">이수일 / 김남민 / 정은혜</div>
            </div>
          </div>
        </div>

        <!-- 범례 -->
        <div class="org-legend">
          <span class="org-tag tag-pre">★사전탑승자</span>
          <span class="org-tag tag-dual">★듀얼임무자</span>
          <span class="org-tag tag-派遣">파견</span>
          <span class="org-tag tag-off">미탑승</span>
        </div>
      </div>
    </div>
"""

# ── 홈 대시보드 생성 ─────────────────────────────────────────────────
def build_dashboard(manifest):
    # 섹션별 카드 그룹
    card_sections = [
        ('운영 매뉴얼 (초안)', ['master','org','prog','sup','emb','port']),
        ('조직·기획', ['orgv2','orgv3','teamdocs','matrix']),
        ('보강자료', ['supplement']),
        ('참고자료', ['qa']),
    ]

    all_cards = []
    for section_name, keys in card_sections:
        cards = []
        for key in keys:
            if key not in manifest:
                continue
            info  = manifest[key]
            meta  = MODULE_META.get(key, {'icon':'📄','color':'#555','desc':''})
            icon  = meta['icon']
            color = meta['color']
            # sup 메뉴명 오버라이드
            label = 'IT홍보팀' if key == 'sup' else info['label']
            if key == 'orgv3':
                label = '기안·온보드미팅'
            desc  = meta['desc']
            first_pid = f'{key}-1'

            sheet_lis = '\n'.join(
                f'            <li onclick="showPage(\'{key}-{i}\');event.stopPropagation()">{esc(s)}</li>'
                for i, s in enumerate(info['sheets'], 1)
                if not (key == 'master' and s in MASTER_EXCLUDE_SHEETS)
            )

            cards.append(f'''      <div class="dash-card" onclick="openGroup('{first_pid}')">
        <div class="dash-card-header" style="background:{color}">
          <h3>{icon} {esc(label)}</h3>
          <p>{esc(desc)}</p>
        </div>
        <div class="dash-card-body">
          <ul class="sheet-list">
{sheet_lis}
          </ul>
        </div>
      </div>''')
        if cards:
            all_cards.append(f'    <h3 class="dash-section-title">{esc(section_name)}</h3>')
            all_cards.append('    <div class="dashboard">')
            all_cards.extend(cards)
            all_cards.append('    </div>')

    return '\n'.join(all_cards)

# ── 콘텐츠 페이지 전체 생성 ──────────────────────────────────────────
def build_pages(manifest, modules):
    pages = []
    order = list(manifest.keys())

    for key in order:
        if key not in modules:
            continue
        info   = manifest[key]
        meta   = MODULE_META.get(key, {'icon':'📄','color':'#2C3E50','desc':''})
        icon   = meta['icon']
        color  = meta['color']
        sheets = modules[key].get('sheets', [])

        for idx, sheet in enumerate(sheets, 1):
            pid        = f'{key}-{idx}'
            sheet_name = info['sheets'][idx-1] if idx-1 < len(info['sheets']) else sheet.get('name','')
            label      = info['label']

            # master에서 제외할 시트는 페이지 생성 스킵
            if key == 'master' and sheet_name in MASTER_EXCLUDE_SHEETS:
                continue

            # 역량 정의서 전치 렌더링 판별
            is_competency = (pid == 'orgv2-2') or ('역량 정의서' in sheet_name)
            if is_competency:
                page_title, content_html = render_competency_transpose(sheet, key)
            else:
                page_title, content_html = render_sheet(sheet, key)

            # 페이지 제목: JSON 첫 행 값 우선, 없으면 sheet_name
            display_title = page_title.strip() or sheet_name

            # matrix-2 페이지에는 필터 UI 삽입
            extra_html = ''
            if pid == 'matrix-2':
                extra_html = MATRIX_FILTER_HTML

            pages.append(f'''  <div class="page" id="page-{pid}">
    <div class="breadcrumb"><a href="#" onclick="showPage(\'home\')">홈</a> / {esc(label)} / {esc(sheet_name)}</div>
    <div class="page-header">
      <h1>{icon} {esc(display_title)}</h1>
    </div>
{extra_html}
{content_html}
  </div>''')

    return '\n'.join(pages)

# ── 매트릭스 필터 HTML ───────────────────────────────────────────────
MATRIX_FILTER_HTML = """
    <div class="matrix-toolbar" id="matrix-toolbar">
      <div class="matrix-search-row">
        <input type="text" id="matrix-search" class="matrix-search" placeholder="🔍 업무명·내용·담당자 통합 검색..." oninput="matrixFilter()">
        <button class="matrix-btn" onclick="matrixResetAll()">필터 초기화</button>
        <button class="matrix-btn matrix-btn-dl" onclick="matrixDownloadCSV()">📥 CSV</button>
        <button class="matrix-btn matrix-btn-dl" onclick="matrixDownloadXlsx()">📥 Excel</button>
        <span class="matrix-counter" id="matrix-counter"></span>
      </div>
    </div>
"""

# ── CSS ──────────────────────────────────────────────────────────────
CSS = r"""/* ── 레이아웃 상수 ── */
:root {
  --sidebar-w: 272px;
  --header-h: 56px;
  --radius: 12px;
  --radius-sm: 8px;
  --accent: #10B981;
  --accent-blue: #3B82F6;
  --transition: 0.2s ease;
}

/* ── 라이트 모드 토큰 ── */
:root {
  --bg-main:       #FAFAFA;
  --bg-header:     #FFFFFF;
  --bg-sidebar:    #FFFFFF;
  --bg-card:       #FFFFFF;
  --bg-table:      #FFFFFF;
  --bg-thead:      #F5F5F5;
  --bg-hover:      #EFF6FF;
  --border-main:   #E5E5E5;
  --border-table:  #EBEBEB;
  --text-primary:  #1A1A1A;
  --text-secondary:#666666;
  --text-nav:      #555555;
  --text-nav-head: #333333;
  --nav-active-bg: #F0F0F0;
  --nav-hover-bg:  #F7F7F7;
  --nav-section-c: #AAAAAA;
  --nav-dot:       #CCCCCC;
  --nav-main-bg:   #F0F1F4;
  --shadow-card:   0 1px 3px rgba(0,0,0,.06), 0 4px 16px rgba(0,0,0,.06);
  --shadow-lg:     0 4px 24px rgba(0,0,0,.10);
  --thead-color:   #333333;
}

/* ── 다크 모드 토큰 ── */
body.dark {
  --bg-main:       #0F0F1A;
  --bg-header:     #12121F;
  --bg-sidebar:    #1A1A2E;
  --bg-card:       #1E1E30;
  --bg-table:      #1E1E30;
  --bg-thead:      #252540;
  --bg-hover:      #252545;
  --border-main:   #2A2A40;
  --border-table:  #2A2A40;
  --text-primary:  #E0E0E0;
  --text-secondary:#8888AA;
  --text-nav:      #A0A0C0;
  --text-nav-head: #C0C0D8;
  --nav-active-bg: #252545;
  --nav-hover-bg:  #1E1E38;
  --nav-section-c: #555570;
  --nav-dot:       #404060;
  --nav-main-bg:   #222240;
  --shadow-card:   0 1px 3px rgba(0,0,0,.3), 0 4px 16px rgba(0,0,0,.3);
  --shadow-lg:     0 4px 24px rgba(0,0,0,.5);
  --thead-color:   #E0E0E0;
}

/* ── 전역 전환 ── */
*, *::before, *::after { margin:0; padding:0; box-sizing:border-box; }
body {
  font-family:'Noto Sans KR', sans-serif;
  background: var(--bg-main);
  color: var(--text-primary);
  line-height: 1.65;
  transition: background var(--transition), color var(--transition);
}

/* ── 헤더 ── */
.header {
  position: fixed; top:0; left:0; right:0; height: var(--header-h); z-index:100;
  background: var(--bg-header);
  border-bottom: 1px solid var(--border-main);
  display: flex; align-items: center; padding: 0 20px; gap: 12px;
  transition: background var(--transition), border-color var(--transition);
}
.header-burger {
  display:none; cursor:pointer; padding:6px 8px; border:none;
  background:none; color: var(--text-primary); font-size:18px; border-radius:6px;
}
.header-burger:hover { background: var(--nav-hover-bg); }
.header-logo { font-weight:900; font-size:17px; letter-spacing:-.5px; color: var(--text-primary); white-space:nowrap; }
.header-logo span { color: var(--accent); }
.header-sub { font-size:12px; color: var(--text-secondary); font-weight:300; }
.header-right { margin-left:auto; display:flex; align-items:center; gap:10px; }
.badge-ver {
  padding:3px 10px; border-radius:20px; font-size:11px; font-weight:600;
  background: var(--nav-active-bg); color: var(--text-secondary);
}
.dark-toggle {
  width:36px; height:36px; border-radius:50%; border:1px solid var(--border-main);
  background: var(--nav-active-bg); cursor:pointer;
  display:flex; align-items:center; justify-content:center;
  font-size:16px; transition: all var(--transition);
}
.dark-toggle:hover { border-color: var(--accent); transform: scale(1.08); }

/* ── 사이드바 ── */
.sidebar {
  position: fixed; top: var(--header-h); left:0; bottom:0; width: var(--sidebar-w);
  background: var(--bg-sidebar);
  border-right: 1px solid var(--border-main);
  overflow-y: auto; z-index:90;
  transition: transform .3s ease, background var(--transition), border-color var(--transition);
}
.sidebar::-webkit-scrollbar { width:4px; }
.sidebar::-webkit-scrollbar-thumb { background: var(--border-main); border-radius:2px; }

.nav-section { padding:14px 0 4px; }
.nav-section-title {
  padding:0 16px; font-size:10px; text-transform:uppercase;
  letter-spacing:1.5px; color: var(--nav-section-c); font-weight:600; margin-bottom:6px;
}
.nav-item {
  display:flex; align-items:center; padding:9px 16px; cursor:pointer;
  color: var(--text-nav); font-size:13px; font-weight:400;
  border-left:3px solid transparent;
  transition: background .12s, color .12s, border-color .12s;
}
.nav-item:hover { background: var(--nav-hover-bg); color: var(--text-nav-head); }
.nav-item.active {
  background: var(--nav-active-bg); color: var(--accent);
  border-left-color: var(--accent); font-weight:600;
}
.nav-sub { padding-left:24px; }
.nav-sub .nav-item { font-size:12px; padding:7px 16px; }
.nav-sub .nav-item::before {
  content:''; display:inline-block; width:5px; height:5px;
  border-radius:50%; background: var(--nav-dot); margin-right:10px; flex-shrink:0;
}
.nav-sub .nav-item.active::before { background: var(--accent); }

.nav-group { border-bottom:1px solid var(--border-main); }
.nav-group-header {
  display:flex; align-items:center; padding:10px 16px; cursor:pointer;
  color: var(--text-nav-head); font-size:13px; font-weight:600;
  background: var(--nav-main-bg);
  margin: 2px 8px; border-radius: 6px;
  transition: background .12s;
}
.nav-group-header:hover { background: var(--nav-hover-bg); }
.nav-group-header .icon { width:22px; margin-right:10px; text-align:center; font-size:14px; }
.nav-group-header .arrow { margin-left:auto; font-size:10px; color: var(--nav-section-c); transition:transform .2s; }
.nav-group.open .arrow { transform:rotate(90deg); }
.nav-group .nav-sub { display:none; }
.nav-group.open .nav-sub { display:block; }

/* ── 메인 콘텐츠 ── */
.main {
  margin-left: var(--sidebar-w); margin-top: var(--header-h);
  padding:28px; min-height:calc(100vh - var(--header-h));
  transition: background var(--transition);
}

/* ── 대시보드 카드 ── */
.dash-section-title {
  font-size:15px; font-weight:700; color:var(--text-primary);
  margin:28px 0 12px; padding-bottom:8px;
  border-bottom:2px solid var(--border-main);
}
.dash-section-title:first-child { margin-top:0; }
.dashboard { display:grid; grid-template-columns:repeat(auto-fill, minmax(300px,1fr)); gap:20px; margin-bottom:8px; }
.dash-card {
  background: var(--bg-card);
  border: 1px solid var(--border-main);
  border-radius: var(--radius);
  overflow:hidden;
  box-shadow: var(--shadow-card);
  cursor:pointer;
  transition: box-shadow var(--transition), transform var(--transition), background var(--transition), border-color var(--transition);
}
.dash-card:hover { box-shadow: var(--shadow-lg); transform:translateY(-3px); }
.dash-card-header { padding:18px 20px; color:white; }
.dash-card-header h3 { font-size:14px; font-weight:700; }
.dash-card-header p { font-size:11px; opacity:.85; margin-top:3px; }
.dash-card-body { padding:14px 20px; }
.sheet-list { list-style:none; }
.sheet-list li {
  padding:8px 0; border-bottom:1px solid var(--border-main);
  font-size:13px; display:flex; align-items:center;
  cursor:pointer; color: var(--text-secondary);
  transition: color .12s;
}
.sheet-list li:last-child { border:none; }
.sheet-list li:hover { color: var(--accent-blue); }
.sheet-list li::before { content:'→'; margin-right:8px; color: var(--nav-dot); font-size:11px; }

/* ── 콘텐츠 페이지 ── */
.page { display:none; }
.page.active { display:block; }
.page-header { margin-bottom:24px; }
.page-header h1 { font-size:21px; font-weight:900; color: var(--text-primary); }
.breadcrumb { font-size:12px; color: var(--text-secondary); margin-bottom:8px; }
.breadcrumb a { color: var(--accent-blue); text-decoration:none; }
.breadcrumb a:hover { text-decoration:underline; }

/* ── 테이블 ── */
.table-wrap {
  overflow-x:auto; margin-bottom:28px;
  border-radius: var(--radius-sm);
  border: 1px solid var(--border-main);
  box-shadow: var(--shadow-card);
  transition: border-color var(--transition);
}

.section-title {
  padding: 11px 18px;
  font-size: 13px; font-weight: 700;
  color: #FFFFFF;
  letter-spacing: .2px;
}

.ops-table {
  width:100%; border-collapse:collapse;
  background: var(--bg-table); font-size:13px;
  transition: background var(--transition);
}

.ops-table thead th {
  padding: 10px 12px;
  font-size: 12px; font-weight: 700;
  text-align: center; letter-spacing: .3px;
  white-space: nowrap;
  background: #F5F6F8 !important;
  color: #333333 !important;
  border-bottom: 2px solid #E0E4EA;
}

.ops-table tbody tr {
  background: color-mix(in srgb, var(--row-bg, transparent) 35%, #FFFFFF);
  min-height: 40px;
}
.ops-table tbody td {
  padding: 10px 12px;
  font-size: 13px; line-height: 1.6;
  vertical-align: top;
  border-bottom: 1px solid #EEF2F7;
  color: var(--text-primary);
  text-align: center;
}
.ops-table tbody td.text-left { text-align: left; }
.ops-table tbody td:first-child { font-weight: 600; white-space: nowrap; }
.ops-table tbody tr:last-child td { border-bottom: none; }
.ops-table tbody tr:hover { background: #F0F7FF !important; }
.ops-table tbody tr:hover td { border-bottom-color: #D8E8F8; }
.cell-wrap { white-space: pre-line; }

.ops-table tbody td.merged-desc {
  padding: 14px 20px;
  font-size: 13px;
  line-height: 1.7;
  color: var(--text-secondary);
  background: #F9FAFB;
  border-bottom: 1px solid #EEF2F7;
}
body.dark .ops-table tbody td.merged-desc {
  background: #1a1a2e !important;
  color: #8888AA !important;
  border-bottom-color: #2A2A40;
}

/* ── 다크모드 테이블 ── */
body.dark .ops-table thead th {
  background: #252540 !important;
  color: #D8D8F0 !important;
  border-bottom-color: #2A2A40;
}
body.dark .ops-table tbody tr {
  background: #1E1E30 !important;
}
body.dark .ops-table tbody td {
  color: #D8D8E8 !important;
  border-bottom-color: #2A2A40;
}
body.dark .ops-table tbody td:first-child { color: #88C8A8 !important; }
body.dark .ops-table tbody tr:hover { background: #282848 !important; }
body.dark .ops-table tbody tr:hover td { border-bottom-color: #303058; }

/* ── 조직도 차트 ── */
.org-chart-section {
  margin-bottom: 36px;
  padding: 24px;
  background: var(--bg-card);
  border: 1px solid var(--border-main);
  border-radius: var(--radius);
  box-shadow: var(--shadow-card);
}
.org-chart { display:flex; flex-direction:column; align-items:center; gap:16px; }
.org-node {
  border: 2px solid var(--border-main);
  border-radius: 10px;
  padding: 12px 16px;
  background: var(--bg-card);
  min-width: 220px;
  transition: all var(--transition);
}
.org-node-title { font-size:14px; font-weight:800; color:var(--text-primary); }
.org-node-sub { font-size:11px; color:var(--text-secondary); margin-top:2px; }
.org-node-members { font-size:12px; color:var(--text-primary); margin-top:6px; line-height:1.6; }
.org-node-role { font-size:11px; color:var(--text-secondary); margin-top:4px; font-style:italic; }
.org-node-note { font-size:11px; color:var(--accent); margin-top:4px; font-weight:600; }

.org-hq { border-color:#0D1B2A; background:linear-gradient(135deg,#0D1B2A,#1B2A4A); }
.org-hq .org-node-title, .org-hq .org-node-sub { color:white; }

.org-branches { display:flex; gap:16px; flex-wrap:wrap; justify-content:center; width:100%; }
.org-branch { display:flex; flex-direction:column; gap:10px; flex:1; min-width:260px; max-width:380px; }
.org-branch-label {
  font-size:10px; font-weight:700; text-transform:uppercase; letter-spacing:1.5px;
  color:var(--nav-section-c); padding:4px 8px; text-align:center;
  border-bottom:2px solid var(--border-main);
}

.org-support { border-left:4px solid #2980B9; }
.org-fb { border-left:4px solid #27AE60; }
.org-it { border-left:4px solid #1ABC9C; }
.org-alba { border-left:4px solid #95A5A6; }
.org-event { border-left:4px solid #E67E22; }
.org-port { border-left:4px solid #27AE60; }
.org-embark { border-left:4px solid #6C5CE7; margin-left:20px; }
.org-temp { border-left:4px solid #F39C12; border-style:dashed; }

.org-sub-node { margin-top:8px; }

.org-tag {
  display:inline-block; font-size:10px; padding:1px 6px; border-radius:10px;
  font-weight:600; margin-left:4px; vertical-align:middle;
}
.tag-pre { background:#FEF3C7; color:#92400E; }
.tag-dual { background:#DBEAFE; color:#1E40AF; }
.tag-ci { background:#E0E7FF; color:#3730A3; }
.tag-modo { background:#D1FAE5; color:#065F46; }
.tag-\6D3E\9063 { background:#FEE2E2; color:#991B1B; }
.tag-off { background:#F3F4F6; color:#374151; }

body.dark .tag-pre { background:#78350F; color:#FDE68A; }
body.dark .tag-dual { background:#1E3A5F; color:#93C5FD; }
body.dark .tag-ci { background:#312E81; color:#C7D2FE; }
body.dark .tag-modo { background:#064E3B; color:#6EE7B7; }
body.dark .tag-\6D3E\9063 { background:#7F1D1D; color:#FCA5A5; }
body.dark .tag-off { background:#374151; color:#D1D5DB; }

.org-legend {
  display:flex; gap:12px; flex-wrap:wrap; justify-content:center;
  padding:12px; margin-top:8px;
  border-top:1px solid var(--border-main);
}

/* ── 매트릭스 필터 ── */
.matrix-toolbar {
  margin-bottom:16px; padding:14px 16px;
  background:var(--bg-card); border:1px solid var(--border-main);
  border-radius:var(--radius-sm);
}
.matrix-search-row { display:flex; gap:8px; align-items:center; flex-wrap:wrap; }
.matrix-search {
  flex:1; min-width:200px; padding:8px 12px; border:1px solid var(--border-main);
  border-radius:6px; font-size:13px; background:var(--bg-main); color:var(--text-primary);
  outline:none; transition:border-color .2s;
}
.matrix-search:focus { border-color:var(--accent-blue); }
.matrix-btn {
  padding:7px 14px; border:1px solid var(--border-main); border-radius:6px;
  background:var(--bg-card); color:var(--text-primary); font-size:12px; font-weight:600;
  cursor:pointer; transition:all .15s; white-space:nowrap;
}
.matrix-btn:hover { background:var(--nav-hover-bg); border-color:var(--accent-blue); }
.matrix-btn-dl { background:var(--accent-blue); color:white; border-color:var(--accent-blue); }
.matrix-btn-dl:hover { opacity:.85; }
.matrix-counter { font-size:12px; color:var(--text-secondary); font-weight:600; white-space:nowrap; }

/* 필터 드롭다운 */
.col-filter-wrap { position:relative; display:inline-block; }
.col-filter-btn {
  background:none; border:none; cursor:pointer; font-size:11px;
  color:var(--text-secondary); padding:0 2px; vertical-align:middle;
}
.col-filter-btn.active { color:var(--accent-blue); }
.col-filter-drop {
  display:none; position:absolute; top:100%; left:0; z-index:200;
  background:var(--bg-card); border:1px solid var(--border-main);
  border-radius:8px; box-shadow:var(--shadow-lg); padding:8px;
  min-width:180px; max-height:260px; overflow-y:auto;
}
.col-filter-drop.open { display:block; }
.col-filter-drop label {
  display:block; font-size:12px; padding:4px 6px; cursor:pointer;
  color:var(--text-primary); border-radius:4px;
}
.col-filter-drop label:hover { background:var(--nav-hover-bg); }
.col-filter-actions { display:flex; gap:6px; margin-bottom:6px; border-bottom:1px solid var(--border-main); padding-bottom:6px; }
.col-filter-actions button {
  font-size:11px; padding:2px 8px; border:1px solid var(--border-main);
  border-radius:4px; background:var(--bg-main); color:var(--text-primary);
  cursor:pointer;
}

/* 카테고리/상태 배지 */
.badge-cat {
  display:inline-block; font-size:10px; padding:2px 8px; border-radius:10px;
  font-weight:600; white-space:nowrap;
}
.badge-status {
  display:inline-block; font-size:10px; padding:2px 8px; border-radius:10px;
  font-weight:700; white-space:nowrap;
}
.badge-status-done { background:#D1FAE5; color:#065F46; }
.badge-status-progress { background:#DBEAFE; color:#1E40AF; }
.badge-status-todo { background:#FEF3C7; color:#92400E; }
body.dark .badge-status-done { background:#064E3B; color:#6EE7B7; }
body.dark .badge-status-progress { background:#1E3A5F; color:#93C5FD; }
body.dark .badge-status-todo { background:#78350F; color:#FDE68A; }

/* ── 오버레이(모바일) ── */
.overlay { display:none; position:fixed; inset:0; background:rgba(0,0,0,.4); z-index:85; }

/* ── 반응형 ── */
@media (max-width:768px) {
  .sidebar { transform:translateX(-100%); box-shadow: none; border-right:none; }
  .sidebar.open { transform:translateX(0); box-shadow: 4px 0 20px rgba(0,0,0,.15); }
  .overlay.open { display:block; }
  .header-burger { display:flex; }
  .main { margin-left:0; padding:16px; }
  .header-sub { display:none; }
  .dashboard { grid-template-columns:1fr; }
  .ops-table { font-size:11px; }
  .ops-table thead th, .ops-table tbody td { padding:8px 6px; }
  .org-branches { flex-direction:column; }
  .org-branch { max-width:100%; }
}
@media (min-width:769px) and (max-width:1024px) {
  :root { --sidebar-w:240px; }
  .dashboard { grid-template-columns:repeat(2,1fr); }
}"""

JS = r"""/* ── 다크모드 ── */
(function() {
  if (localStorage.getItem('dark') === '1') {
    document.body.classList.add('dark');
  }
})();

function toggleDark() {
  const isDark = document.body.classList.toggle('dark');
  localStorage.setItem('dark', isDark ? '1' : '0');
  document.getElementById('dark-btn').textContent = isDark ? '☀️' : '🌙';
}

window.addEventListener('DOMContentLoaded', function() {
  const isDark = document.body.classList.contains('dark');
  document.getElementById('dark-btn').textContent = isDark ? '☀️' : '🌙';
  // 매트릭스 필터 초기화
  if (document.getElementById('matrix-toolbar')) {
    matrixInit();
  }
});

/* ── 페이지 전환 ── */
function showPage(id) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));

  const page = document.getElementById('page-' + id);
  if (page) { page.classList.add('active'); }
  else { document.getElementById('page-home').classList.add('active'); }

  document.querySelectorAll('.nav-item').forEach(n => {
    if (n.getAttribute('onclick') && n.getAttribute('onclick').includes("'" + id + "'")) {
      n.classList.add('active');
    }
  });

  document.getElementById('sidebar').classList.remove('open');
  document.getElementById('overlay').classList.remove('open');
  window.scrollTo(0, 0);
}

function toggleGroup(el) {
  el.parentElement.classList.toggle('open');
}

function toggleSidebar() {
  document.getElementById('sidebar').classList.toggle('open');
  document.getElementById('overlay').classList.toggle('open');
}

function openGroup(pageId) {
  showPage(pageId);
  document.querySelectorAll('.nav-group').forEach(g => {
    g.querySelectorAll('.nav-item').forEach(item => {
      if (item.getAttribute('onclick') && item.getAttribute('onclick').includes(pageId)) {
        g.classList.add('open');
      }
    });
  });
}

/* ══════════════════════════════════════════════════════════════════
   매트릭스 필터 + 다운로드 (matrix-2 페이지 전용)
   ══════════════════════════════════════════════════════════════════ */
var _mxTable = null;
var _mxFilters = {};        // colIdx → Set of checked values
var _mxFilterCols = {};     // colIdx → header text
var _mxFilterable = [1, 2, 5, 7]; // 시기, 카테고리, 담당팀, 상태

// 카테고리별 색상 (16종)
var CAT_COLORS = {
  '계약·행정':'#EBF5FB','인력·조직':'#E8F8F5','VIP·의전':'#FDEDEC',
  '물류·총무':'#FEF9E7','IT·통신':'#EBF5FB','프로그램·행사':'#FDF2E9',
  '식음료':'#E8F8F5','기항지':'#EAFAF1','승하선':'#F4ECF7',
  '홍보·마케팅':'#FDEBD0','코스타소통':'#D6EAF8','안전·비상':'#FADBD8',
  '데이터·보고':'#F2F3F4','사후업무':'#F5EEF8',
  '★공연매니지먼트':'#FAE5D3','★코스타행정':'#D4EFDF'
};
var CAT_COLORS_DARK = {
  '계약·행정':'#1B4F72','인력·조직':'#0E6655','VIP·의전':'#78281F',
  '물류·총무':'#7D6608','IT·통신':'#1A5276','프로그램·행사':'#784212',
  '식음료':'#0B5345','기항지':'#145A32','승하선':'#4A235A',
  '홍보·마케팅':'#784212','코스타소통':'#1B4F72','안전·비상':'#78281F',
  '데이터·보고':'#2C3E50','사후업무':'#4A235A',
  '★공연매니지먼트':'#6E2C00','★코스타행정':'#1E8449'
};

function matrixInit() {
  var page = document.getElementById('page-matrix-2');
  if (!page) return;
  _mxTable = page.querySelector('.ops-table');
  if (!_mxTable) return;

  var ths = _mxTable.querySelectorAll('thead th');
  _mxFilterable.forEach(function(ci) {
    if (ci >= ths.length) return;
    var th = ths[ci];
    _mxFilterCols[ci] = th.textContent.trim();

    // 고유값 수집
    var vals = new Set();
    _mxTable.querySelectorAll('tbody tr').forEach(function(tr) {
      var td = tr.querySelectorAll('td')[ci];
      if (td) {
        var t = td.textContent.trim();
        if (t) vals.add(t);
      }
    });

    _mxFilters[ci] = new Set(vals);

    // 필터 버튼+드롭다운 생성
    var wrap = document.createElement('span');
    wrap.className = 'col-filter-wrap';
    var btn = document.createElement('button');
    btn.className = 'col-filter-btn';
    btn.textContent = ' ▼';
    btn.setAttribute('data-col', ci);
    btn.onclick = function(e) {
      e.stopPropagation();
      var drop = wrap.querySelector('.col-filter-drop');
      document.querySelectorAll('.col-filter-drop.open').forEach(function(d) {
        if (d !== drop) d.classList.remove('open');
      });
      drop.classList.toggle('open');
    };

    var drop = document.createElement('div');
    drop.className = 'col-filter-drop';

    var actions = document.createElement('div');
    actions.className = 'col-filter-actions';
    var selAll = document.createElement('button');
    selAll.textContent = '전체';
    selAll.onclick = function(e) { e.stopPropagation(); toggleAllChecks(drop, true, ci); };
    var selNone = document.createElement('button');
    selNone.textContent = '해제';
    selNone.onclick = function(e) { e.stopPropagation(); toggleAllChecks(drop, false, ci); };
    actions.appendChild(selAll);
    actions.appendChild(selNone);
    drop.appendChild(actions);

    Array.from(vals).sort().forEach(function(v) {
      var lbl = document.createElement('label');
      var cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.checked = true;
      cb.value = v;
      cb.setAttribute('data-col', ci);
      cb.onchange = function(e) { e.stopPropagation(); onFilterChange(ci); };
      lbl.appendChild(cb);
      lbl.appendChild(document.createTextNode(' ' + v));
      drop.appendChild(lbl);
    });

    wrap.appendChild(btn);
    wrap.appendChild(drop);
    th.appendChild(wrap);
  });

  // 배지 적용
  applyBadges();

  // 카운터 초기화
  updateCounter();

  // 외부 클릭 시 드롭다운 닫기
  document.addEventListener('click', function() {
    document.querySelectorAll('.col-filter-drop.open').forEach(function(d) {
      d.classList.remove('open');
    });
  });
}

function toggleAllChecks(drop, state, colIdx) {
  drop.querySelectorAll('input[type=checkbox]').forEach(function(cb) {
    cb.checked = state;
  });
  onFilterChange(colIdx);
}

function onFilterChange(colIdx) {
  var checked = new Set();
  document.querySelectorAll('input[data-col="'+colIdx+'"]:checked').forEach(function(cb) {
    checked.add(cb.value);
  });
  _mxFilters[colIdx] = checked;

  // 아이콘 색상
  var allVals = new Set();
  document.querySelectorAll('input[data-col="'+colIdx+'"]').forEach(function(cb) {
    allVals.add(cb.value);
  });
  var btn = document.querySelector('.col-filter-btn[data-col="'+colIdx+'"]');
  if (btn) {
    btn.classList.toggle('active', checked.size < allVals.size);
  }

  matrixFilter();
}

function matrixFilter() {
  if (!_mxTable) return;
  var query = (document.getElementById('matrix-search').value || '').toLowerCase();
  var rows = _mxTable.querySelectorAll('tbody tr');

  rows.forEach(function(tr) {
    var tds = tr.querySelectorAll('td');
    // 섹션 구분 행 (merged-desc)은 항상 표시
    if (tr.querySelector('.merged-desc')) {
      tr.style.display = '';
      return;
    }

    var show = true;

    // 열 필터
    for (var ci in _mxFilters) {
      ci = parseInt(ci);
      if (ci >= tds.length) continue;
      var cellText = tds[ci].textContent.trim();
      if (!_mxFilters[ci].has(cellText)) { show = false; break; }
    }

    // 텍스트 검색 (업무명[3] + 세부내용[4] + 담당자[6])
    if (show && query) {
      var searchText = '';
      [3,4,6].forEach(function(i) {
        if (i < tds.length) searchText += ' ' + tds[i].textContent;
      });
      if (searchText.toLowerCase().indexOf(query) === -1) show = false;
    }

    tr.style.display = show ? '' : 'none';
  });

  updateCounter();
}

function updateCounter() {
  if (!_mxTable) return;
  var rows = _mxTable.querySelectorAll('tbody tr');
  var total = 0, visible = 0;
  rows.forEach(function(tr) {
    if (tr.querySelector('.merged-desc')) return;
    total++;
    if (tr.style.display !== 'none') visible++;
  });
  var el = document.getElementById('matrix-counter');
  if (el) el.textContent = visible + ' / ' + total + '건';
}

function matrixResetAll() {
  document.getElementById('matrix-search').value = '';
  document.querySelectorAll('#page-matrix-2 input[type=checkbox]').forEach(function(cb) {
    cb.checked = true;
  });
  document.querySelectorAll('.col-filter-btn').forEach(function(b) {
    b.classList.remove('active');
  });
  for (var ci in _mxFilters) {
    var all = new Set();
    document.querySelectorAll('input[data-col="'+ci+'"]').forEach(function(cb) { all.add(cb.value); });
    _mxFilters[ci] = all;
  }
  matrixFilter();
}

function applyBadges() {
  if (!_mxTable) return;
  var isDark = document.body.classList.contains('dark');
  _mxTable.querySelectorAll('tbody tr').forEach(function(tr) {
    var tds = tr.querySelectorAll('td');
    // 카테고리 배지 (col 2)
    if (tds.length > 2) {
      var cat = tds[2].textContent.trim();
      var bgColor = isDark ? (CAT_COLORS_DARK[cat]||'#2C3E50') : (CAT_COLORS[cat]||'#F2F3F4');
      var textColor = isDark ? '#E0E0E0' : '#1A1A1A';
      if (cat) tds[2].innerHTML = '<span class="badge-cat" style="background:'+bgColor+';color:'+textColor+'">'+cat+'</span>';
    }
    // 상태 배지 (col 7)
    if (tds.length > 7) {
      var st = tds[7].textContent.trim();
      var cls = '';
      if (st === '완료') cls = 'badge-status-done';
      else if (st === '진행중') cls = 'badge-status-progress';
      else if (st === '미착수') cls = 'badge-status-todo';
      if (cls) tds[7].innerHTML = '<span class="badge-status '+cls+'">'+st+'</span>';
    }
  });
}

/* ── 다운로드 ── */
function getVisibleRows() {
  if (!_mxTable) return [];
  var headers = [];
  _mxTable.querySelectorAll('thead th').forEach(function(th) {
    headers.push(th.childNodes[0].textContent.trim());
  });
  var data = [headers];
  _mxTable.querySelectorAll('tbody tr').forEach(function(tr) {
    if (tr.style.display === 'none') return;
    if (tr.querySelector('.merged-desc')) return;
    var row = [];
    tr.querySelectorAll('td').forEach(function(td) { row.push(td.textContent.trim()); });
    data.push(row);
  });
  return data;
}

function matrixDownloadCSV() {
  var data = getVisibleRows();
  var csv = data.map(function(r) {
    return r.map(function(c) { return '"' + c.replace(/"/g,'""') + '"'; }).join(',');
  }).join('\n');
  var bom = '\uFEFF';
  var blob = new Blob([bom + csv], {type:'text/csv;charset=utf-8'});
  var a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = '업무매트릭스_' + new Date().toISOString().slice(0,10) + '.csv';
  a.click();
}

function matrixDownloadXlsx() {
  if (typeof XLSX === 'undefined') {
    var s = document.createElement('script');
    s.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
    s.onload = function() { doXlsxDownload(); };
    document.head.appendChild(s);
  } else {
    doXlsxDownload();
  }
}
function doXlsxDownload() {
  var data = getVisibleRows();
  var ws = XLSX.utils.aoa_to_sheet(data);
  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '업무매트릭스');
  XLSX.writeFile(wb, '업무매트릭스_' + new Date().toISOString().slice(0,10) + '.xlsx');
}
"""

# ── 최종 HTML 조립 ────────────────────────────────────────────────────
def build_html(manifest, modules):
    nav       = build_nav(manifest)
    org_chart = build_org_chart()
    dashboard = build_dashboard(manifest)
    pages     = build_pages(manifest, modules)

    return f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>모두의 크루즈 운영 데스크</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;700;900&family=JetBrains+Mono:wght@400;600&display=swap" rel="stylesheet">
<style>
{CSS}
</style>
</head>
<body>

<!-- Header -->
<header class="header">
  <button class="header-burger" onclick="toggleSidebar()">☰</button>
  <div class="header-logo">모두의 <span>크루즈</span> 운영 데스크</div>
  <div class="header-sub">2026 코스타 세레나 한일전세선 (6.19~6.25)</div>
  <div class="header-right">
    <span class="badge-ver">v3.0</span>
    <button class="dark-toggle" id="dark-btn" onclick="toggleDark()" title="다크모드 전환">🌙</button>
  </div>
</header>

<!-- Sidebar -->
{nav}

<!-- Overlay (mobile) -->
<div class="overlay" id="overlay" onclick="toggleSidebar()"></div>

<!-- Main -->
<main class="main">

  <!-- HOME -->
  <div class="page active" id="page-home">
    <div class="page-header">
      <h1>🚢 모두의 크루즈 운영 데스크</h1>
      <p>2026 코스타 세레나 한일전세선 (6.19~6.25) — 총괄운영 매뉴얼 대시보드</p>
    </div>
{org_chart}
{dashboard}
  </div>

  <!-- Content Pages -->
{pages}

</main>

<script>
{JS}
</script>
</body>
</html>"""

# ── 메인 ─────────────────────────────────────────────────────────────
def main():
    test_key = sys.argv[1] if len(sys.argv) > 1 else None

    print('=== 크루즈 운영 데스크 빌더 v3.0 ===')
    manifest, modules = load_all()

    if test_key:
        print(f'  [TEST MODE] {test_key}만 렌더링')
        for k in list(modules.keys()):
            if k != test_key:
                modules.pop(k)

    html = build_html(manifest, modules)

    with open(OUT, 'w', encoding='utf-8') as f:
        f.write(html)

    size_kb = os.path.getsize(OUT) // 1024
    print(f'  → {OUT} ({size_kb} KB) 생성 완료')

if __name__ == '__main__':
    main()
