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

# ── 모듈별 메타정보 ───────────────────────────────────────────────────
MODULE_META = {
    'master': {'icon':'📋','color':'#0D1B2A','desc':'D-90부터 D-Day까지 전체 준비사항'},
    'org':    {'icon':'🏢','color':'#2980B9','desc':'HQ 중심 선내운영 조직 체계'},
    'prog':   {'icon':'🎭','color':'#E67E22','desc':'공연·이벤트·환경재단행사 통합운영'},
    'sup':    {'icon':'🔧','color':'#1ABC9C','desc':'IT·총무·홍보 통합 후방지원'},
    'emb':    {'icon':'🚢','color':'#6C5CE7','desc':'2,400명 승선·하선 총괄 운영'},
    'port':   {'icon':'⚓','color':'#27AE60','desc':'하코다테·오타루 기항 운영'},
    'qa':     {'icon':'❓','color':'#34495E','desc':'롯데JTB 이문규 팀장 자문용'},
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

def cell_style(c, include_bg=True):
    parts = []
    if include_bg and c.get('bg'):
        parts.append(f"background:{c['bg']}")
    if c.get('fc'):
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

        # 섹션 제목
        if sec['title']:
            t  = sec['title']
            bg = t.get('bg', '#1B2A4A')
            fc = t.get('fc', '#FFFFFF')
            out.append(f'<div class="section-title" style="background:{bg};color:{fc}">{esc(t.get("value",""))}</div>')

        rows = sec['rows']
        if rows:
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
                    bg = cells[0].get('bg', '#1B2A4A')
                    out.append(f'<tr style="background:{bg}">')
                    for c in cells:
                        out.append(f'<th style="{cell_style(c)}">{esc(c.get("value",""))}</th>')
                    out.append('</tr>')
                else:  # 'd'
                    if in_head:
                        out.append('</thead>')
                        in_head = False
                    if not in_body:
                        out.append('<tbody>')
                        in_body = True
                    rb = cells[0].get('bg', '')
                    out.append(f'<tr style="background:{rb}">' if rb else '<tr>')
                    for c in cells:
                        out.append(f'<td style="{cell_style(c)}" class="cell-wrap">{esc(c.get("value",""))}</td>')
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
    lines.append('</nav>')
    return '\n'.join(lines)

# ── 홈 대시보드 생성 ─────────────────────────────────────────────────
def build_dashboard(manifest):
    order = ['prog','master','org','sup','emb','port','qa']
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

# ── CSS (index.html에서 가져옴, 수정 금지) ──────────────────────────
CSS = """:root {
  --navy: #0D1B2A;
  --dark: #1B2A4A;
  --steel: #2C3E50;
  --charcoal: #34495E;
  --blue: #2980B9;
  --teal: #1ABC9C;
  --green: #27AE60;
  --orange: #E67E22;
  --red: #E74C3C;
  --purple: #8E44AD;
  --indigo: #6C5CE7;
  --cyan: #00BCD4;
  --gold: #F39C12;
  --white: #FFFFFF;
  --light: #F8F9FA;
  --border: #E2E8F0;
  --text: #1a202c;
  --text-sub: #64748b;
  --sidebar-w: 280px;
  --header-h: 60px;
  --radius: 8px;
  --shadow: 0 1px 3px rgba(0,0,0,.08), 0 4px 12px rgba(0,0,0,.04);
  --shadow-lg: 0 4px 16px rgba(0,0,0,.12);
}

* { margin:0; padding:0; box-sizing:border-box; }
body { font-family:'Noto Sans KR',sans-serif; background:#f1f5f9; color:var(--text); line-height:1.6; }

/* ── Header ── */
.header {
  position:fixed; top:0; left:0; right:0; height:var(--header-h); z-index:100;
  background:var(--navy); color:white;
  display:flex; align-items:center; padding:0 20px;
  box-shadow:0 2px 8px rgba(0,0,0,.3);
}
.header-burger { display:none; cursor:pointer; padding:8px; margin-right:12px; border:none; background:none; color:white; font-size:20px; }
.header-logo { font-weight:900; font-size:18px; letter-spacing:-.5px; }
.header-logo span { color:var(--teal); }
.header-sub { margin-left:16px; font-size:12px; color:#94a3b8; font-weight:300; }
.header-badge { margin-left:auto; display:flex; gap:8px; align-items:center; }
.badge { padding:3px 10px; border-radius:12px; font-size:11px; font-weight:600; }
.badge-live { background:var(--red); color:white; }
.badge-ver { background:rgba(255,255,255,.15); color:#94a3b8; }

/* ── Sidebar ── */
.sidebar {
  position:fixed; top:var(--header-h); left:0; bottom:0; width:var(--sidebar-w);
  background:var(--dark); overflow-y:auto; z-index:90;
  transition:transform .3s ease;
}
.sidebar::-webkit-scrollbar { width:4px; }
.sidebar::-webkit-scrollbar-thumb { background:rgba(255,255,255,.15); border-radius:2px; }

.nav-section { padding:16px 0 4px; }
.nav-section-title { padding:0 16px; font-size:10px; text-transform:uppercase; letter-spacing:1.5px; color:#64748b; font-weight:600; margin-bottom:8px; }

.nav-item {
  display:flex; align-items:center; padding:10px 16px; cursor:pointer;
  color:#94a3b8; font-size:13px; font-weight:400; transition:all .15s;
  border-left:3px solid transparent;
}
.nav-item:hover { background:rgba(255,255,255,.05); color:#e2e8f0; }
.nav-item.active { background:rgba(26,188,156,.1); color:var(--teal); border-left-color:var(--teal); font-weight:600; }
.nav-item .icon { width:20px; margin-right:10px; text-align:center; font-size:14px; }

.nav-sub { padding-left:28px; }
.nav-sub .nav-item { font-size:12px; padding:7px 16px; }
.nav-sub .nav-item::before { content:''; display:inline-block; width:6px; height:6px; border-radius:50%; background:#475569; margin-right:10px; flex-shrink:0; }
.nav-sub .nav-item.active::before { background:var(--teal); }

.nav-group { border-bottom:1px solid rgba(255,255,255,.06); }
.nav-group-header {
  display:flex; align-items:center; padding:10px 16px; cursor:pointer;
  color:#cbd5e1; font-size:13px; font-weight:500; transition:all .15s;
}
.nav-group-header:hover { background:rgba(255,255,255,.05); }
.nav-group-header .icon { width:20px; margin-right:10px; text-align:center; }
.nav-group-header .arrow { margin-left:auto; font-size:10px; transition:transform .2s; }
.nav-group.open .arrow { transform:rotate(90deg); }
.nav-group .nav-sub { display:none; }
.nav-group.open .nav-sub { display:block; }

/* ── Main Content ── */
.main { margin-left:var(--sidebar-w); margin-top:var(--header-h); padding:24px; min-height:calc(100vh - var(--header-h)); }

/* Dashboard Home */
.dashboard { display:grid; grid-template-columns:repeat(auto-fill, minmax(320px,1fr)); gap:20px; }
.dash-card {
  background:white; border-radius:var(--radius); overflow:hidden;
  box-shadow:var(--shadow); transition:all .2s; cursor:pointer;
}
.dash-card:hover { box-shadow:var(--shadow-lg); transform:translateY(-2px); }
.dash-card-header { padding:16px 20px; color:white; }
.dash-card-header h3 { font-size:15px; font-weight:700; }
.dash-card-header p { font-size:11px; opacity:.8; margin-top:2px; }
.dash-card-body { padding:16px 20px; }
.dash-card-body .sheet-list { list-style:none; }
.dash-card-body .sheet-list li {
  padding:8px 0; border-bottom:1px solid var(--border); font-size:13px;
  display:flex; align-items:center; cursor:pointer; color:var(--steel);
}
.dash-card-body .sheet-list li:last-child { border:none; }
.dash-card-body .sheet-list li:hover { color:var(--blue); }
.dash-card-body .sheet-list li::before { content:'→'; margin-right:8px; color:#94a3b8; font-size:11px; }

/* Content Page */
.page { display:none; }
.page.active { display:block; }
.page-header { margin-bottom:24px; }
.page-header h1 { font-size:22px; font-weight:900; color:var(--navy); }
.page-header p { font-size:13px; color:var(--text-sub); margin-top:4px; }
.breadcrumb { font-size:12px; color:var(--text-sub); margin-bottom:8px; }
.breadcrumb a { color:var(--blue); text-decoration:none; }

/* Table Styles */
.table-wrap { overflow-x:auto; margin-bottom:32px; border-radius:var(--radius); box-shadow:var(--shadow); }
.section-title {
  padding:14px 20px; color:white; font-size:14px; font-weight:700;
  position:sticky; top:0; z-index:5;
}
.ops-table { width:100%; border-collapse:collapse; background:white; font-size:12px; }
.ops-table thead th {
  padding:10px 12px; font-weight:600; color:white; text-align:center;
  position:sticky; top:0; z-index:4; font-size:11px; letter-spacing:.3px;
}
.ops-table tbody td {
  padding:10px 12px; border-bottom:1px solid #eef2f7; vertical-align:top;
  line-height:1.65;
}
.ops-table tbody tr:hover { background:#f8fafc; }
.ops-table tbody td:first-child { font-weight:600; white-space:nowrap; }
.cell-wrap { white-space:pre-line; }

/* Mobile */
.overlay { display:none; position:fixed; inset:0; background:rgba(0,0,0,.5); z-index:85; }
@media (max-width:768px) {
  .sidebar { transform:translateX(-100%); }
  .sidebar.open { transform:translateX(0); }
  .overlay.open { display:block; }
  .header-burger { display:block; }
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

JS = """function showPage(id) {
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
  <div class="header-badge">
    <span class="badge badge-ver">v2.0</span>
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
