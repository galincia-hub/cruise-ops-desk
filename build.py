#!/usr/bin/env python3
"""
build.py — 모두의 크루즈 운영 데스크 HTML 빌더
Usage: python build.py
  data/*.json → index.html
"""

import json, os, sys

BASE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(BASE, 'data')
OUT  = os.path.join(BASE, 'index.html')

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
    'master':     '#4A5568',   # 차분한 슬레이트
    'org':        '#2B6CB0',   # 차분한 블루
    'prog':       '#C05621',   # 차분한 오렌지
    'sup':        '#2F855A',   # 차분한 그린
    'emb':        '#4C51BF',   # 차분한 인디고
    'port':       '#276749',   # 차분한 딥그린
    'qa':         '#4A5568',   # 슬레이트
    'orgv2':      '#1565C0',   # 조직도 블루
    'supplement': '#00695C',   # 보강자료 틸
}

# ── 모듈별 메타정보 ───────────────────────────────────────────────────
MODULE_META = {
    'master': {'icon':'📋','color':'#0D1B2A','desc':'D-90부터 D-Day까지 전체 준비사항'},
    'org':    {'icon':'🏢','color':'#2980B9','desc':'HQ 중심 선내운영 조직 체계'},
    'prog':   {'icon':'🎭','color':'#E67E22','desc':'공연·이벤트·환경재단행사 통합운영'},
    'sup':    {'icon':'🔧','color':'#1ABC9C','desc':'IT·총무·홍보 통합 후방지원'},
    'emb':    {'icon':'🚢','color':'#6C5CE7','desc':'2,400명 승선·하선 총괄 운영'},
    'port':   {'icon':'⚓','color':'#27AE60','desc':'하코다테·오타루 기항 운영'},
    'qa':         {'icon':'❓','color':'#34495E','desc':'선행 항차 경험자 자문용'},
    'orgv2':      {'icon':'📊','color':'#1565C0','desc':'확정 조직도 + 포지션별 역량 정의'},
    'supplement': {'icon':'📎','color':'#00695C','desc':'무전기·Wi-Fi·비상연락 등 보강자료'},
}

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

# ── 시트 렌더링 핵심 로직 ─────────────────────────────────────────────
def render_sheet(sheet, module_key):
    """
    sheet['data']의 row_group 배열을 HTML로 변환.
    row_group = 한 엑셀 행에 속한 셀 객체들의 배열.

    판별 규칙:
      1) 첫 번째 row_group → 페이지 제목 (page-header h1)
      2) 셀 1개 + 어두운 bg → 섹션 제목 (section-title div)
      3) 전체 셀 어두운 bg → thead 헤더 행
      4) 나머지 → tbody 데이터 행
    """
    data = sheet.get('data', [])
    out = []
    page_title = ''
    sections = []  # [{title_cell, rows: [(type, cells)]}]
    cur = None

    for i, row in enumerate(data):
        if not row:
            continue

        # ① 첫 행: 페이지 제목
        if i == 0:
            page_title = row[0].get('value', '')
            continue

        all_dark = all(is_dark(c.get('bg', '')) for c in row)
        single   = (len(row) == 1)

        if single and all_dark:
            # ② 섹션 제목 → 새 섹션 시작
            cur = {'title': row[0], 'rows': []}
            sections.append(cur)
        else:
            if cur is None:
                cur = {'title': None, 'rows': []}
                sections.append(cur)
            rtype = 'h' if all_dark else 'd'
            cur['rows'].append((rtype, row))

    # 섹션별 HTML 생성
    for sec in sections:
        out.append('<div class="table-wrap">')

        # 섹션 제목 — JSON 원색 대신 모듈별 차분한 색상 사용
        if sec['title']:
            sec_color = MODULE_SECTION_COLORS.get(module_key, '#4A5568')
            val = esc(sec['title'].get('value', ''))
            out.append(f'<div class="section-title" style="background:{sec_color}">{val}</div>')

        rows = sec['rows']
        if rows:
            # ── 헤더 먼저 스캔해서 col_count 선결정 ──────────────────────
            # 설명 행이 헤더보다 앞에 오는 경우에도 올바른 colspan 적용 가능
            col_count = next((len(cells) for rtype, cells in rows if rtype == 'h'), 0)

            # ── 헤더 텍스트 기반 좌측정렬 컬럼 판별 ──────────────────
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

                else:  # 'd'
                    if in_head:
                        out.append('</thead>')
                        in_head = False
                    if not in_body:
                        out.append('<tbody>')
                        in_body = True

                    n = len(cells)

                    # ── 병합 행 판별: 실제 셀 수 < 헤더 컬럼 수의 절반 ──────
                    # (헤더가 없는 섹션은 col_count=0 → 병합 없이 그대로)
                    if col_count > 0 and n < col_count / 2:
                        val = '<br>'.join(
                            esc(c.get('value', '')) for c in cells if c.get('value')
                        )
                        out.append('<tr>')
                        out.append(f'<td colspan="{col_count}" class="cell-wrap merged-desc">{val}</td>')
                        out.append('</tr>')

                    # ── 정상 행 ──────────────────────────────────────────────
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

        out.append('</div>')  # .table-wrap

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
def build_nav(manifest):
    lines = [
        '<nav class="sidebar" id="sidebar">',
        '  <div class="nav-section">',
        '    <div class="nav-section-title">대시보드</div>',
        '    <div class="nav-item active" onclick="showPage(\'home\')"><span class="icon">🏠</span>홈</div>',
        '  </div>',
        '  <div class="nav-section">',
        '    <div class="nav-section-title">운영 매뉴얼</div>',
    ]

    order = ['master','org','prog','sup','emb','port']
    ref   = ['qa']
    extra = ['orgv2','supplement']

    for key in order:
        if key not in manifest:
            continue
        info  = manifest[key]
        meta  = MODULE_META.get(key, {'icon':'📄','color':'#2C3E50'})
        icon  = meta['icon']
        label = info['label']
        lines.append(f'    <div class="nav-group">')
        lines.append(f'      <div class="nav-group-header" onclick="toggleGroup(this)"><span class="icon">{icon}</span>{esc(label)}<span class="arrow">▶</span></div>')
        lines.append(f'      <div class="nav-sub">')
        for idx, sheet_name in enumerate(info['sheets'], 1):
            pid = f'{key}-{idx}'
            lines.append(f'        <div class="nav-item" onclick="showPage(\'{pid}\')">{esc(sheet_name)}</div>')
        lines.append(f'      </div>')
        lines.append(f'    </div>')

    lines.append('  </div>')
    lines.append('  <div class="nav-section">')
    lines.append('    <div class="nav-section-title">참고자료</div>')

    for key in ref:
        if key not in manifest:
            continue
        info  = manifest[key]
        meta  = MODULE_META.get(key, {'icon':'📄','color':'#2C3E50'})
        icon  = meta['icon']
        label = info['label']
        lines.append(f'    <div class="nav-group">')
        lines.append(f'      <div class="nav-group-header" onclick="toggleGroup(this)"><span class="icon">{icon}</span>{esc(label)}<span class="arrow">▶</span></div>')
        lines.append(f'      <div class="nav-sub">')
        for idx, sheet_name in enumerate(info['sheets'], 1):
            pid = f'{key}-{idx}'
            lines.append(f'        <div class="nav-item" onclick="showPage(\'{pid}\')">{esc(sheet_name)}</div>')
        lines.append(f'      </div>')
        lines.append(f'    </div>')

    lines.append('  </div>')

    # ── 조직도·역량정의서 + 보강자료 ──
    lines.append('  <div class="nav-section">')
    lines.append('    <div class="nav-section-title">추가자료</div>')

    for key in extra:
        if key not in manifest:
            continue
        info  = manifest[key]
        meta  = MODULE_META.get(key, {'icon':'📄','color':'#2C3E50'})
        icon  = meta['icon']
        label = info['label']
        lines.append(f'    <div class="nav-group">')
        lines.append(f'      <div class="nav-group-header" onclick="toggleGroup(this)"><span class="icon">{icon}</span>{esc(label)}<span class="arrow">▶</span></div>')
        lines.append(f'      <div class="nav-sub">')
        for idx, sheet_name in enumerate(info['sheets'], 1):
            pid = f'{key}-{idx}'
            lines.append(f'        <div class="nav-item" onclick="showPage(\'{pid}\')">{esc(sheet_name)}</div>')
        lines.append(f'      </div>')
        lines.append(f'    </div>')

    lines.append('  </div>')

    # ── 아카이브 섹션 ──
    lines.append('  <div class="nav-section">')
    lines.append('    <div class="nav-section-title">📁 아카이브 (버전 기록)</div>')
    lines.append('    <div class="nav-item" onclick="window.open(\'archive/v1.0_2026-03-17.html\',\'_blank\')">v1.0 (2026-03-17)</div>')
    lines.append('  </div>')

    lines.append('</nav>')
    return '\n'.join(lines)

# ── 홈 대시보드 생성 ─────────────────────────────────────────────────
def build_dashboard(manifest):
    order = ['prog','master','org','sup','emb','port','qa','orgv2','supplement']
    cards = []
    for key in order:
        if key not in manifest:
            continue
        info  = manifest[key]
        meta  = MODULE_META.get(key, {'icon':'📄','color':'#555','desc':''})
        icon  = meta['icon']
        color = meta['color']
        label = info['label']
        desc  = meta['desc']
        first_pid = f'{key}-1'

        sheet_lis = '\n'.join(
            f'            <li onclick="showPage(\'{key}-{i}\');event.stopPropagation()">{esc(s)}</li>'
            for i, s in enumerate(info['sheets'], 1)
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

    return '\n'.join(cards)

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

            page_title, content_html = render_sheet(sheet, key)

            # 페이지 제목: JSON 첫 행 값 우선, 없으면 sheet_name
            display_title = page_title.strip() or sheet_name

            pages.append(f'''  <div class="page" id="page-{pid}">
    <div class="breadcrumb"><a href="#" onclick="showPage(\'home\')">홈</a> / {esc(label)} / {esc(sheet_name)}</div>
    <div class="page-header">
      <h1>{icon} {esc(display_title)}</h1>
    </div>
{content_html}
  </div>''')

    return '\n'.join(pages)

# ── CSS ──────────────────────────────────────────────────────────────
CSS = """/* ── 레이아웃 상수 ── */
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
.dashboard { display:grid; grid-template-columns:repeat(auto-fill, minmax(300px,1fr)); gap:20px; }
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

/* 섹션 타이틀 — 색상은 Python 인라인 스타일(MODULE_SECTION_COLORS)로 결정 */
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

/* ── thead: 라이트 = 연한 회색(#F5F6F8) + 진한 텍스트, Python 인라인 무시 ── */
.ops-table thead th {
  padding: 10px 12px;
  font-size: 12px; font-weight: 700;
  text-align: center; letter-spacing: .3px;
  white-space: nowrap;
  background: #F5F6F8 !important;
  color: #333333 !important;
  border-bottom: 2px solid #E0E4EA;
}

/* ── tbody: color-mix()로 원본 bg를 opacity 35% 수준으로 연하게 ── */
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

/* 병합 설명 행 (엑셀 merged cell 복원) */
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
/* 다크: 행 배경을 단색 기본으로 통일 (color-mix 파스텔 제거) */
body.dark .ops-table tbody tr {
  background: #1E1E30 !important;
}
body.dark .ops-table tbody td {
  color: #D8D8E8 !important;       /* 인라인 color 완전 무력화 */
  border-bottom-color: #2A2A40;
}
body.dark .ops-table tbody td:first-child { color: #88C8A8 !important; }
body.dark .ops-table tbody tr:hover { background: #282848 !important; }
body.dark .ops-table tbody tr:hover td { border-bottom-color: #303058; }

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
}
@media (min-width:769px) and (max-width:1024px) {
  :root { --sidebar-w:240px; }
  .dashboard { grid-template-columns:repeat(2,1fr); }
}"""

JS = """/* ── 다크모드 ── */
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
}"""

# ── 최종 HTML 조립 ────────────────────────────────────────────────────
def build_html(manifest, modules):
    nav       = build_nav(manifest)
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
    <span class="badge-ver">v2.1</span>
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
    <div class="dashboard">
{dashboard}
    </div>
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
    # 테스트 모드: python build.py prog  → prog.json만 처리
    test_key = sys.argv[1] if len(sys.argv) > 1 else None

    print('=== 크루즈 운영 데스크 빌더 ===')
    manifest, modules = load_all()

    if test_key:
        # 테스트: 지정 모듈만, 나머지는 placeholder
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
