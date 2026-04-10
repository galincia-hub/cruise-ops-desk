#!/usr/bin/env python3
"""
build.py — 모두의 크루즈 운영 데스크 HTML 빌더 v3.1 (대안A 계층형)
Usage: python build.py
  data/*.json → index.html
"""

import json, os, sys, re

BASE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(BASE, 'data')
OUT  = os.path.join(BASE, 'index.html')
ARCHIVE_DIR = os.path.join(BASE, 'archive')
GITHUB_PAGES_BASE = 'https://galincia-hub.github.io/cruise-ops-desk'

# ── 어두운 배경색 목록 ──────────────────────────────────────────────
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

MODULE_SECTION_COLORS = {
    'master':'#4A5568','org':'#2B6CB0','prog':'#C05621','sup':'#2F855A',
    'emb':'#4C51BF','port':'#276749','qa':'#4A5568','orgv2':'#1565C0',
    'orgv3':'#283593','supplement':'#00695C','teamdocs':'#4527A0','matrix':'#BF360C',
}

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

# 팀 워크스페이스 정의
TEAM_WORKSPACE = [
    {
        'id': 'team-hq',
        'name': 'HQ 운영본부',
        'icon': '🏛️',
        'color': '#0D1B2A',
        'members': '모두투어 + CI 공동부서장',
        'mission': '전 팀 조율·관리 총괄 | VIP 의전 | 코스타 선사 최종 대응',
        'teamdocs_sheet': None,
        'matrix_teams': ['HQ', 'HQ(재무)', 'HQ→각 팀장', '전원'],
        'manual_keys': [],
        'sub_items': [
            ('조직도', 'org-chart'),
            ('업무 진행현황', 'progress'),
            ('기안·협조공문', 'docs'),
            ('핵심 일정', 'timeline'),
        ],
    },
    {
        'id': 'team-support',
        'name': '운영지원팀',
        'icon': '🏢',
        'color': '#2980B9',
        'members': '황지애(CI, 지원) · 양은희(모두투어) · 조아라(웅진) · 김지은(환경재단)',
        'mission': '고객응대 + 방송안내 + 정산관리 — 팀 과업, 유동배치',
        'teamdocs_sheet': '운영지원팀',
        'matrix_teams': ['운영지원팀', '운영지원팀+안내'],
        'manual_keys': [],
        'sub_items': [
            ('팀 개요·미션', 'overview'),
            ('업무 리스트', 'tasks'),
            ('SOP·시나리오', 'sop'),
        ],
    },
    {
        'id': 'team-event',
        'name': '공연/행사팀',
        'icon': '🎭',
        'color': '#E67E22',
        'members': '이상민 매니저(팀장) ★사전탑승 + 팀원4 + 환경재단15명',
        'mission': '공연·이벤트·환경재단행사 통합운영',
        'teamdocs_sheet': '공연행사팀',
        'matrix_teams': ['공연/행사팀', '기항지+공연', '승하선팀+공연팀'],
        'manual_keys': ['prog'],
        'sub_items': [
            ('팀 개요·미션', 'overview'),
            ('업무 리스트', 'tasks'),
            ('SOP·시나리오', 'sop'),
            ('매뉴얼 (초안)', 'manual'),
        ],
    },
    {
        'id': 'team-port',
        'name': '기항지운영팀',
        'icon': '⚓',
        'color': '#27AE60',
        'members': '이수일 팀장 ★듀얼 + 팀원1 + 타이요(6+가이드70)',
        'mission': '하코다테·오타루 기항지 투어 운영 + 승하선 겸임',
        'teamdocs_sheet': '기항지운영팀',
        'matrix_teams': ['기항지팀', '기항지+공연'],
        'manual_keys': ['port'],
        'sub_items': [
            ('팀 개요·미션', 'overview'),
            ('업무 리스트', 'tasks'),
            ('SOP·시나리오', 'sop'),
            ('매뉴얼 (초안)', 'manual'),
        ],
        'sub_team': {
            'id': 'team-embark',
            'name': '승하선팀',
            'icon': '🚢',
            'color': '#6C5CE7',
            'members': '김남민 매니저(미탑승/터미널전담) + 이수일★듀얼 + 알바3~4',
            'mission': '2,400명 승선·하선 터미널 총괄',
            'teamdocs_sheet': '승하선팀',
            'matrix_teams': ['승하선팀', 'IT홍보팀+승하선팀', '승하선팀+공연팀'],
            'manual_keys': ['emb'],
            'sub_items': [
                ('팀 개요', 'overview'),
                ('업무 리스트', 'tasks'),
                ('SOP·케이스스터디', 'sop'),
            ],
        },
    },
    {
        'id': 'team-fb',
        'name': '식음료파트',
        'icon': '🍽️',
        'color': '#E74C3C',
        'members': '문형식 매니저 ★사전탑승',
        'mission': '전체 밀스케줄 관리 + 한식 품질 모니터링',
        'teamdocs_sheet': '식음료파트',
        'matrix_teams': ['식음료'],
        'manual_keys': [],
        'sub_items': [
            ('팀 개요', 'overview'),
            ('업무 리스트', 'tasks'),
        ],
    },
    {
        'id': 'team-it',
        'name': 'IT홍보팀',
        'icon': '💻',
        'color': '#1ABC9C',
        'members': '최구철 매니저(팀장)',
        'mission': 'IT·통신 인프라 + 홍보·콘텐츠 + 총무 지원',
        'teamdocs_sheet': 'IT홍보팀',
        'matrix_teams': ['IT홍보팀', 'IT홍보팀+승하선팀', '승하선팀+IT홍보팀'],
        'manual_keys': ['sup'],
        'sub_items': [
            ('팀 개요·미션', 'overview'),
            ('업무 리스트', 'tasks'),
            ('SOP·시나리오', 'sop'),
            ('매뉴얼 (초안)', 'manual'),
        ],
    },
]

MASTER_EXCLUDE_SHEETS = ['인력운영안(원본)']

# 팀 ID → meeting_insights.json 키 매핑
TEAM_INSIGHT_MAP = {
    'team-hq': 'hq',
    'team-support': 'support',
    'team-event': 'program',
    'team-port': 'port',
    'team-embark': 'port',      # 승하선팀은 기항지 인사이트 공유
    'team-it': 'it',
}

# v3.1 업데이트로 변경된 팀 ID (NEW 뱃지 대상)
UPDATE_DATE = '2026-04-10'
UPDATED_TEAMS = {'team-hq', 'team-support', 'team-event', 'team-port', 'team-embark', 'team-it'}

# 마스터 업무 매트릭스 담당팀 명칭 통일 매핑 (원본 matrix.json 수정 없이 렌더링 시 적용)
TEAM_NORM = {
    'HQ(재무)':   'HQ',
    'HQ→각 팀장': 'HQ',
    '각 팀장':    'HQ',
    '기항지팀':   '기항지운영팀',
    '식음료':     '식음료파트',
}

# ── 유틸리티 ─────────────────────────────────────────────────────────
def esc(v):
    if v is None: return ''
    s = str(v).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
    return s.replace('\n','<br>')

def is_dark(bg):
    return bool(bg) and bg.upper() in DARK_BGS

def cell_style(c, include_bg=True, include_color=True):
    parts = []
    if include_bg and c.get('bg'): parts.append(f"background:{c['bg']}")
    if include_color and c.get('fc'): parts.append(f"color:{c['fc']}")
    if c.get('bold'): parts.append('font-weight:700')
    return ';'.join(parts)

def scan_archives():
    archives = []
    if not os.path.isdir(ARCHIVE_DIR): return archives
    for fname in os.listdir(ARCHIVE_DIR):
        m = re.match(r'v([\d.]+)_(\d{4}-\d{2}-\d{2})\.html$', fname)
        if m: archives.append((m.group(1), m.group(2), fname))
    archives.sort(key=lambda x: (x[1], x[0]), reverse=True)
    return archives

# ── 시트 렌더링 ──────────────────────────────────────────────────────
def render_competency_transpose(sheet, module_key):
    data = sheet.get('data', [])
    if not data: return '', ''
    page_title = data[0][0].get('value', '')
    positions = []
    cur = None
    for row in data[1:]:
        if not row: continue
        single = (len(row) == 1)
        if single and (all(is_dark(c.get('bg','')) for c in row) or is_dark(row[0].get('bg',''))):
            cur = {'name': row[0].get('value','').replace('\n',' '), 'rows': {}}
            positions.append(cur)
        elif cur is not None and len(row) >= 2:
            label = row[0].get('value','').replace('\n',' ').strip()
            cur['rows'][label] = row[1].get('value','')
    if not positions: return page_title, ''
    row_labels = list(positions[0]['rows'].keys())
    sec_color = MODULE_SECTION_COLORS.get(module_key, '#4A5568')
    out = ['<div class="table-wrap">',
           f'<div class="section-title" style="background:{sec_color}">포지션별 역량 비교표</div>',
           '<table class="ops-table">','<thead><tr>','<th style="min-width:140px">구분</th>']
    for l in row_labels: out.append(f'<th>{esc(l)}</th>')
    out.append('</tr></thead><tbody>')
    for pos in positions:
        name = esc(pos['name'].lstrip('▣').strip())
        out.append('<tr>')
        out.append(f'<td style="font-weight:700;white-space:nowrap;text-align:center" class="cell-wrap">{name}</td>')
        for l in row_labels:
            out.append(f'<td style="text-align:left" class="cell-wrap">{esc(pos["rows"].get(l,""))}</td>')
        out.append('</tr>')
    out.append('</tbody></table></div>')
    return page_title, '\n'.join(out)

def render_sheet(sheet, module_key):
    data = sheet.get('data', [])
    out, page_title, sections, cur = [], '', [], None
    for i, row in enumerate(data):
        if not row: continue
        if i == 0:
            page_title = row[0].get('value','')
            continue
        all_dark = all(is_dark(c.get('bg','')) for c in row)
        single = (len(row) == 1)
        if single and all_dark:
            cur = {'title': row[0], 'rows': []}
            sections.append(cur)
        else:
            if cur is None:
                cur = {'title': None, 'rows': []}
                sections.append(cur)
            cur['rows'].append(('h' if all_dark else 'd', row))

    for sec in sections:
        out.append('<div class="table-wrap">')
        if sec['title']:
            sc = MODULE_SECTION_COLORS.get(module_key, '#4A5568')
            out.append(f'<div class="section-title" style="background:{sc}">{esc(sec["title"].get("value",""))}</div>')
        rows = sec['rows']
        if rows:
            col_count = next((len(c) for t,c in rows if t=='h'), 0)
            LEFT_KW = ['세부','내용','체크리스트','항목','업무']
            left_cols = set()
            for t,cells in rows:
                if t == 'h':
                    for ci,c in enumerate(cells):
                        if any(k in str(c.get('value','')) for k in LEFT_KW): left_cols.add(ci)
            out.append('<table class="ops-table">')
            in_head = in_body = False
            for rtype, cells in rows:
                if rtype == 'h':
                    if in_body: out.append('</tbody>'); in_body = False
                    if not in_head: out.append('<thead>'); in_head = True
                    out.append('<tr>')
                    for c in cells:
                        out.append(f'<th style="{cell_style(c,include_bg=False)}">{esc(c.get("value",""))}</th>')
                    out.append('</tr>')
                else:
                    if in_head: out.append('</thead>'); in_head = False
                    if not in_body: out.append('<tbody>'); in_body = True
                    n = len(cells)
                    if col_count > 0 and n < col_count / 2:
                        val = '<br>'.join(esc(c.get('value','')) for c in cells if c.get('value'))
                        out.append(f'<tr><td colspan="{col_count}" class="cell-wrap merged-desc">{val}</td></tr>')
                    else:
                        rb = cells[0].get('bg','')
                        rs = f' style="--row-bg:{rb}"' if rb else ''
                        out.append(f'<tr{rs}>')
                        for idx,c in enumerate(cells):
                            cls = "cell-wrap text-left" if idx in left_cols else "cell-wrap"
                            val = c.get('value','')
                            # 매트릭스: 담당팀(col5) 명칭 통일
                            if module_key == 'matrix' and idx == 5:
                                val = TEAM_NORM.get(str(val).strip(), val)
                            out.append(f'<td style="{cell_style(c,include_bg=False,include_color=False)}" class="{cls}">{esc(val)}</td>')
                        out.append('</tr>')
            if in_head: out.append('</thead>')
            if in_body: out.append('</tbody>')
            out.append('</table>')
        out.append('</div>')
    return page_title, '\n'.join(out)

# ── 데이터 로드 ──────────────────────────────────────────────────────
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

# ── 미팅 인사이트 로드 ────────────────────────────────────────────────
def load_insights():
    path = os.path.join(DATA, 'meeting_insights.json')
    if not os.path.exists(path):
        return None
    with open(path, encoding='utf-8') as f:
        return json.load(f)

def build_insights_section(tid, insights):
    """팀 워크스페이스 하단에 타사 미팅 인사이트 섹션을 렌더링한다."""
    if not insights:
        return ''
    key = TEAM_INSIGHT_MAP.get(tid)
    if not key or key not in insights:
        return ''
    section = insights[key]
    items = section.get('items', [])
    if not items:
        return ''

    priority_cls = {'높음': 'high', '중간': 'mid', '낮음': 'low'}
    meta = insights.get('meta', {})
    source = esc(meta.get('source', ''))
    updated = esc(meta.get('updated', ''))

    o = []
    o.append(f'    <div id="{tid}-insights" class="insights-wrap insights-new">')
    o.append(f'      <div class="insights-toggle" onclick="this.parentElement.classList.toggle(\'open\')">')
    o.append(f'        <h3 class="section-heading" style="margin:0;border:none;cursor:pointer">')
    o.append(f'          📋 타사 미팅 인사이트 <span class="nav-new-badge" title="{UPDATE_DATE} 신규 추가">NEW</span> <span class="insights-arrow">▶</span></h3>')
    o.append(f'        <span class="insights-meta">{source} · 갱신 {updated}</span>')
    o.append(f'      </div>')
    o.append(f'      <div class="insights-body">')
    o.append(f'        <table class="ops-table insights-table"><thead><tr>')
    o.append(f'          <th style="width:70px">ID</th><th>과업명</th><th>내용</th>')
    o.append(f'          <th style="width:80px">시점</th><th style="width:70px">우선순위</th>')
    o.append(f'        </tr></thead><tbody>')
    for it in items:
        pcls = priority_cls.get(it.get('priority',''), '')
        o.append(f'        <tr>')
        o.append(f'          <td><code>{esc(it["id"])}</code></td>')
        o.append(f'          <td><strong>{esc(it["title"])}</strong></td>')
        o.append(f'          <td>{esc(it["content"])}</td>')
        o.append(f'          <td>{esc(it.get("phase",""))}</td>')
        o.append(f'          <td><span class="priority-badge {pcls}">{esc(it.get("priority",""))}</span></td>')
        o.append(f'        </tr>')
    o.append(f'        </tbody></table>')
    o.append(f'      </div>')
    o.append(f'    </div>')
    return '\n'.join(o)

# ── 매트릭스 통계 (진행현황 카운터) ──────────────────────────────────
def matrix_stats(modules):
    if 'matrix' not in modules: return 0,0,0,0
    sheet = modules['matrix']['sheets'][1]  # 매트릭스 v2
    total = done = progress = todo = 0
    for row in sheet.get('data',[])[2:]:
        if len(row) <= 1: continue
        total += 1
        st = row[7].get('value','').strip() if len(row) > 7 else ''
        if st == '완료': done += 1
        elif st == '진행중': progress += 1
        elif st == '미착수': todo += 1
    return total, done, progress, todo

# ── 사이드바 ─────────────────────────────────────────────────────────
def build_nav(manifest):
    L = ['<nav class="sidebar" id="sidebar">']

    # 전체 조망
    L.append('  <div class="nav-section">')
    L.append('    <div class="nav-section-title">전체 조망</div>')
    L.append('    <div class="nav-item active" onclick="showPage(\'home\')"><span class="icon">🏠</span>운영 대시보드</div>')
    L.append(f'    <div class="nav-item" onclick="showPage(\'master-timeline\')"><span class="icon">📋</span>마스터 타임라인</div>')
    L.append(f'    <div class="nav-item" onclick="showPage(\'matrix-2\')"><span class="icon">📐</span>마스터 매트릭스</div>')
    L.append('  </div>')

    # 팀별 워크스페이스
    L.append('  <div class="nav-section">')
    L.append('    <div class="nav-section-title">팀별 워크스페이스</div>')

    def _nav_team(tm, indent='    '):
        new_badge = ''
        if tm['id'] in UPDATED_TEAMS:
            new_badge = f'<span class="nav-new-badge" title="{UPDATE_DATE} 업데이트">NEW</span>'
        L.append(f'{indent}<div class="nav-group">')
        L.append(f'{indent}  <div class="nav-group-header" onclick="toggleGroup(this)"><span class="icon">{tm["icon"]}</span>{esc(tm["name"])}{new_badge}<span class="arrow">▶</span></div>')
        L.append(f'{indent}  <div class="nav-sub">')
        for label, anchor in tm.get('sub_items', []):
            L.append(f'{indent}    <div class="nav-item" onclick="showPageSection(\'{tm["id"]}\',\'{anchor}\')">{esc(label)}</div>')
        if 'sub_team' in tm:
            sub = tm['sub_team']
            _nav_team(sub, indent + '    ')
        L.append(f'{indent}  </div>')
        L.append(f'{indent}</div>')

    for tm in TEAM_WORKSPACE:
        _nav_team(tm)

    L.append('  </div>')

    # 참고자료실
    L.append('  <div class="nav-section">')
    L.append('    <div class="nav-section-title">참고자료실</div>')
    ref_items = [
        ('orgv3-2', '📝', '기안·협조공문'),
        ('orgv3-4', '📋', '온보드미팅 준비'),
        ('ref-supplement', '📎', '보강자료'),
        ('qa-1', '❓', '질의·자문 요청사항'),
        ('ref-manuals', '📚', '운영매뉴얼 (초안)'),
    ]
    for pid, icon, label in ref_items:
        if pid.startswith('ref-'):
            # 그룹 메뉴
            L.append(f'    <div class="nav-group">')
            L.append(f'      <div class="nav-group-header" onclick="toggleGroup(this)"><span class="icon">{icon}</span>{esc(label)}<span class="arrow">▶</span></div>')
            L.append(f'      <div class="nav-sub">')
            if pid == 'ref-supplement':
                for i, sn in enumerate(manifest.get('supplement',{}).get('sheets',[]),1):
                    L.append(f'        <div class="nav-item" onclick="showPage(\'supplement-{i}\')">{esc(sn)}</div>')
            elif pid == 'ref-manuals':
                for mk in ['master','org','prog','sup','emb','port']:
                    if mk not in manifest: continue
                    mi = manifest[mk]
                    micon = MODULE_META.get(mk,{}).get('icon','📄')
                    mlabel = 'IT홍보팀' if mk == 'sup' else mi['label']
                    L.append(f'        <div class="nav-group">')
                    L.append(f'          <div class="nav-group-header" onclick="toggleGroup(this)"><span class="icon">{micon}</span>{esc(mlabel)} (초안)<span class="arrow">▶</span></div>')
                    L.append(f'          <div class="nav-sub">')
                    for si, sn in enumerate(mi['sheets'],1):
                        if mk == 'master' and sn in MASTER_EXCLUDE_SHEETS: continue
                        L.append(f'            <div class="nav-item" onclick="showPage(\'{mk}-{si}\')">{esc(sn)}</div>')
                    L.append(f'          </div>')
                    L.append(f'        </div>')
            L.append(f'      </div>')
            L.append(f'    </div>')
        else:
            L.append(f'    <div class="nav-item" onclick="showPage(\'{pid}\')"><span class="icon">{icon}</span>{esc(label)}</div>')
    L.append('  </div>')

    # 아카이브
    archives = scan_archives()
    L.append('  <div class="nav-section">')
    L.append('    <div class="nav-section-title">📁 아카이브</div>')
    for ver, date, fname in archives:
        url = f'{GITHUB_PAGES_BASE}/archive/{fname}'
        L.append(f'    <div class="nav-item" onclick="window.open(\'{url}\',\'_blank\')">v{ver} ({date})</div>')
    L.append('  </div>')

    L.append('</nav>')
    return '\n'.join(L)

# ── 조직도 HTML ──────────────────────────────────────────────────────
def build_org_chart():
    return """
    <div class="org-chart-section">
      <h2 style="font-size:18px;font-weight:800;margin-bottom:16px;color:var(--text-primary)">📊 선내운영 조직도 v4</h2>
      <div class="org-chart">
        <div class="org-node org-hq" onclick="showPage('home')" style="cursor:pointer">
          <div class="org-node-title">HQ 운영본부 (모두투어 + CI)</div>
          <div class="org-node-role">공연/기항지/승하선/고객응대/정산/식음료/방송 등 전체 운영 과업을 공동 수행</div>
          <div class="org-node-note">VIP: 부서장 직속 관리</div>
        </div>
        <div class="org-branches">
          <div class="org-branch">
            <div class="org-node org-support" onclick="showPage('team-support')" style="cursor:pointer">
              <div class="org-node-title">운영지원팀</div>
              <div class="org-node-members">황지애(CI, 지원) · 양은희(모두투어) · 조아라(웅진) · 김지은(환경재단)</div>
              <div class="org-node-role">팀 과업, 유동배치</div>
            </div>
            <div class="org-node org-event" onclick="showPage('team-event')" style="cursor:pointer">
              <div class="org-node-title">공연/행사팀 <span class="org-tag tag-pre">★사전탑승</span></div>
              <div class="org-node-members">이상민 매니저(팀장) + 팀원4 + 환경재단15명 협업</div>
            </div>
            <div class="org-node org-port" onclick="showPage('team-port')" style="cursor:pointer">
              <div class="org-node-title">기항지운영팀 <span class="org-tag tag-dual">★듀얼</span></div>
              <div class="org-node-members">이수일 팀장 + 팀원1 + 타이요(6+가이드70)</div>
              <div class="org-sub-node">
                <div class="org-node org-embark" onclick="showPage('team-embark');event.stopPropagation()" style="cursor:pointer">
                  <div class="org-node-title">승하선팀 (기항지팀 산하)</div>
                  <div class="org-node-members">김남민 매니저 <span class="org-tag tag-off">미탑승/터미널</span> + 이수일<span class="org-tag tag-dual">★듀얼</span> + 알바3~4</div>
                </div>
              </div>
            </div>
          </div>
          <div class="org-branch">
            <div class="org-node org-fb" onclick="showPage('team-fb')" style="cursor:pointer">
              <div class="org-node-title">식음료파트 <span class="org-tag tag-pre">★사전탑승</span></div>
              <div class="org-node-members">문형식 매니저</div>
            </div>
            <div class="org-node org-it" onclick="showPage('team-it')" style="cursor:pointer">
              <div class="org-node-title">IT홍보팀</div>
              <div class="org-node-members">최구철 매니저(팀장)</div>
            </div>
            <div class="org-node org-alba">
              <div class="org-node-title">알바 12명</div>
              <div class="org-node-role">탄력 배치</div>
            </div>
            <div class="org-node org-temp">
              <div class="org-node-title">[임시] 부산지점 사전준비팀</div>
              <div class="org-node-sub">(탑승 전 임시조직)</div>
              <div class="org-node-members">이수일 / 김남민 / 정은혜</div>
            </div>
          </div>
        </div>
        <div class="org-legend">
          <span class="org-tag tag-pre">★사전탑승자</span>
          <span class="org-tag tag-dual">★듀얼임무자</span>
          <span class="org-tag tag-off">미탑승(터미널전담)</span>
        </div>
      </div>
    </div>
"""

# ── 대시보드 홈 ──────────────────────────────────────────────────────
def build_dashboard(manifest, modules):
    total, done, progress, todo = matrix_stats(modules)
    other = total - done - progress - todo

    counters = f"""
    <div class="stat-row">
      <div class="stat-card"><div class="stat-num">{total}</div><div class="stat-label">전체 업무</div></div>
      <div class="stat-card stat-done"><div class="stat-num">{done}</div><div class="stat-label">완료</div></div>
      <div class="stat-card stat-prog"><div class="stat-num">{progress}</div><div class="stat-label">진행중</div></div>
      <div class="stat-card stat-todo"><div class="stat-num">{todo}</div><div class="stat-label">미착수</div></div>
    </div>
"""

    milestones = """
    <h3 class="dash-section-title">🗓️ 핵심 일정 하이라이트</h3>
    <div class="milestone-row">
      <div class="ms-card ms-urgent"><div class="ms-date">4월 초</div><div class="ms-title">★★★ TF 기안 상신</div><div class="ms-desc">운영 TF 공식 편성 기안</div></div>
      <div class="ms-card"><div class="ms-date">D-60</div><div class="ms-title">공연 계약+대금</div><div class="ms-desc">아티스트 계약 완료 + 선급금</div></div>
      <div class="ms-card"><div class="ms-date">5/27</div><div class="ms-title">온보드미팅</div><div class="ms-desc">전 팀 합동 사전 미팅</div></div>
      <div class="ms-card"><div class="ms-date">6/13~19</div><div class="ms-title">사전탑승</div><div class="ms-desc">이상민·문형식 사전 승선</div></div>
      <div class="ms-card"><div class="ms-date">6/18</div><div class="ms-title">베이스캠프</div><div class="ms-desc">부산 터미널 사전 세팅</div></div>
      <div class="ms-card ms-dday"><div class="ms-date">6/19</div><div class="ms-title">D-Day 승선!</div><div class="ms-desc">2,400명 승선 개시</div></div>
    </div>
"""

    return counters + '\n' + milestones

# ── 팀 워크스페이스 페이지 ───────────────────────────────────────────
def build_team_pages(manifest, modules, insights=None):
    pages = []

    def _build_hq(tm):
        """HQ 운영본부 전용 페이지."""
        tid = tm['id']
        total, done, progress, todo = matrix_stats(modules)
        out = []
        out.append(f'  <div class="page" id="page-{tid}">')
        out.append(f'    <div class="breadcrumb"><a href="#" onclick="showPage(\'home\')">홈</a> / 팀별 워크스페이스 / {esc(tm["name"])}</div>')
        out.append(f'    <div class="team-header" style="border-left:5px solid {tm["color"]}">')
        out.append(f'      <h1>{tm["icon"]} {esc(tm["name"])}</h1>')
        out.append(f'      <div class="team-members">{esc(tm["members"])}</div>')
        out.append(f'      <div class="team-mission">{esc(tm["mission"])}</div>')
        out.append(f'    </div>')

        # ① 조직도
        out.append(f'    <div id="{tid}-org-chart">')
        out.append(build_org_chart())
        out.append(f'    </div>')

        # ② 업무 진행현황
        out.append(f'    <div id="{tid}-progress">')
        out.append(f'    <h3 class="section-heading">📊 업무 진행현황</h3>')
        out.append(f'    <div class="stat-row">')
        out.append(f'      <div class="stat-card"><div class="stat-num">{total}</div><div class="stat-label">전체 업무</div></div>')
        out.append(f'      <div class="stat-card stat-done"><div class="stat-num">{done}</div><div class="stat-label">완료</div></div>')
        out.append(f'      <div class="stat-card stat-prog"><div class="stat-num">{progress}</div><div class="stat-label">진행중</div></div>')
        out.append(f'      <div class="stat-card stat-todo"><div class="stat-num">{todo}</div><div class="stat-label">미착수</div></div>')
        out.append(f'    </div>')
        # master "업무분장표" 시트 (index 2)
        if 'master' in modules:
            msheets = modules['master'].get('sheets',[])
            if len(msheets) > 2:
                _, html = render_sheet(msheets[2], 'master')
                out.append(html)
        out.append(f'    </div>')

        # ③ 기안·협조공문
        out.append(f'    <div id="{tid}-docs">')
        out.append(f'    <h3 class="section-heading">📝 기안·협조공문</h3>')
        if 'orgv3' in modules:
            osheets = modules['orgv3'].get('sheets',[])
            # 협조요청 공문 (index 1), 기안 목록 (index 2)
            for si in [1, 2]:
                if si < len(osheets):
                    _, html = render_sheet(osheets[si], 'orgv3')
                    out.append(html)
        out.append(f'    </div>')

        # ④ 핵심 일정
        out.append(f'    <div id="{tid}-timeline">')
        out.append(f'    <h3 class="section-heading">🗓️ 핵심 일정 타임라인</h3>')
        # master "사전준비 마스터플랜" 시트 (index 0)
        if 'master' in modules:
            msheets = modules['master'].get('sheets',[])
            if len(msheets) > 0:
                _, html = render_sheet(msheets[0], 'master')
                out.append(html)
        out.append(f'    </div>')

        # ⑤ 타사 미팅 인사이트
        ins_html = build_insights_section(tid, insights)
        if ins_html:
            out.append(ins_html)

        out.append('  </div>')
        pages.append('\n'.join(out))

    def _build_one(tm):
        """일반 팀 페이지 — 섹션별 앵커 ID 포함."""
        if tm['id'] == 'team-hq':
            _build_hq(tm)
            return

        tid = tm['id']
        out = []
        out.append(f'  <div class="page" id="page-{tid}">')
        out.append(f'    <div class="breadcrumb"><a href="#" onclick="showPage(\'home\')">홈</a> / 팀별 워크스페이스 / {esc(tm["name"])}</div>')

        # 팀 개요 (overview 앵커)
        out.append(f'    <div id="{tid}-overview" class="team-header" style="border-left:5px solid {tm["color"]}">')
        out.append(f'      <h1>{tm["icon"]} {esc(tm["name"])}</h1>')
        out.append(f'      <div class="team-members">{esc(tm["members"])}</div>')
        out.append(f'      <div class="team-mission">{esc(tm["mission"])}</div>')
        out.append(f'    </div>')

        # 업무 리스트 (tasks 앵커) — teamdocs 시트
        out.append(f'    <div id="{tid}-tasks">')
        if 'teamdocs' in modules and tm.get('teamdocs_sheet'):
            tdsheets = modules['teamdocs'].get('sheets',[])
            td_names = manifest.get('teamdocs',{}).get('sheets',[])
            for si, sname in enumerate(td_names):
                if sname == tm['teamdocs_sheet'] and si < len(tdsheets):
                    _, html = render_sheet(tdsheets[si], 'teamdocs')
                    out.append(f'    <h3 class="section-heading">📑 업무분장서</h3>')
                    out.append(html)
        out.append(f'    </div>')

        # SOP·시나리오 (sop 앵커) — 매뉴얼 SOP 시트만
        out.append(f'    <div id="{tid}-sop">')
        # 관련 매뉴얼의 SOP/시나리오 시트만 추출
        for mk in tm.get('manual_keys',[]):
            if mk not in modules or mk not in manifest: continue
            mi = manifest[mk]
            msheets = modules[mk].get('sheets',[])
            micon = MODULE_META.get(mk,{}).get('icon','📄')
            sop_found = False
            for si, sheet in enumerate(msheets):
                sname = mi['sheets'][si] if si < len(mi['sheets']) else ''
                if any(kw in sname for kw in ['SOP','시나리오','케이스스터디','상황대응']):
                    if not sop_found:
                        out.append(f'    <h3 class="section-heading">{micon} SOP·시나리오</h3>')
                        sop_found = True
                    _, html = render_sheet(sheet, mk)
                    out.append(html)
        out.append(f'    </div>')

        # 매뉴얼 전체 (manual 앵커)
        out.append(f'    <div id="{tid}-manual">')
        for mk in tm.get('manual_keys',[]):
            if mk not in modules or mk not in manifest: continue
            mi = manifest[mk]
            msheets = modules[mk].get('sheets',[])
            micon = MODULE_META.get(mk,{}).get('icon','📄')
            mlabel = 'IT홍보팀' if mk == 'sup' else mi['label']
            out.append(f'    <h3 class="section-heading">{micon} {esc(mlabel)} (초안) 매뉴얼</h3>')
            for si, sheet in enumerate(msheets):
                sname = mi['sheets'][si] if si < len(mi['sheets']) else ''
                if mk == 'master' and sname in MASTER_EXCLUDE_SHEETS: continue
                _, html = render_sheet(sheet, mk)
                out.append(html)
        out.append(f'    </div>')

        # 타사 미팅 인사이트
        ins_html = build_insights_section(tid, insights)
        if ins_html:
            out.append(ins_html)

        out.append('  </div>')
        pages.append('\n'.join(out))

    for tm in TEAM_WORKSPACE:
        _build_one(tm)
        if 'sub_team' in tm:
            _build_one(tm['sub_team'])

    return '\n'.join(pages)

# ── 마스터 타임라인 페이지 ───────────────────────────────────────────
def build_master_timeline(manifest, modules):
    """master JSON + matrix "타사 스케줄" 시트를 타임라인 페이지로."""
    parts = []
    parts.append('  <div class="page" id="page-master-timeline">')
    parts.append('    <div class="breadcrumb"><a href="#" onclick="showPage(\'home\')">홈</a> / 전체 조망 / 마스터 타임라인</div>')
    parts.append('    <div class="page-header"><h1>📋 마스터 타임라인</h1></div>')

    # master JSON 시트들
    if 'master' in modules and 'master' in manifest:
        mi = manifest['master']
        for si, sheet in enumerate(modules['master'].get('sheets',[])):
            sname = mi['sheets'][si] if si < len(mi['sheets']) else sheet.get('name','')
            if sname in MASTER_EXCLUDE_SHEETS: continue
            _, html = render_sheet(sheet, 'master')
            parts.append(html)

    # matrix "타사 스케줄" 시트
    if 'matrix' in modules:
        msheets = modules['matrix'].get('sheets',[])
        if len(msheets) > 0:
            _, html = render_sheet(msheets[0], 'matrix')
            parts.append(f'    <h3 class="section-heading">📐 타사 스케줄·참관 계획</h3>')
            parts.append(html)

    parts.append('  </div>')
    return '\n'.join(parts)

# ── 참고자료 페이지들 ────────────────────────────────────────────────
def build_ref_pages(manifest, modules):
    pages = []
    order = list(manifest.keys())

    for key in order:
        if key not in modules: continue
        info = manifest[key]
        meta = MODULE_META.get(key, {'icon':'📄','color':'#2C3E50','desc':''})
        icon = meta['icon']
        sheets = modules[key].get('sheets',[])

        for idx, sheet in enumerate(sheets, 1):
            pid = f'{key}-{idx}'
            sname = info['sheets'][idx-1] if idx-1 < len(info['sheets']) else sheet.get('name','')
            label = info['label']

            if key == 'master' and sname in MASTER_EXCLUDE_SHEETS: continue

            is_competency = (pid == 'orgv2-2') or ('역량 정의서' in sname)
            if is_competency:
                pt, html = render_competency_transpose(sheet, key)
            else:
                pt, html = render_sheet(sheet, key)

            display_title = pt.strip() or sname
            extra = MATRIX_FILTER_HTML if pid == 'matrix-2' else ''

            pages.append(f'''  <div class="page" id="page-{pid}">
    <div class="breadcrumb"><a href="#" onclick="showPage('home')">홈</a> / {esc(label)} / {esc(sname)}</div>
    <div class="page-header"><h1>{icon} {esc(display_title)}</h1></div>
{extra}
{html}
  </div>''')

    return '\n'.join(pages)

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
CSS = r"""/* ── 디자인 시스템 v4.0 ── */
:root{
  --sidebar-w:240px;--header-h:56px;--tab-bar-h:52px;
  --radius:12px;--radius-sm:8px;--transition:0.2s ease;
  --bg-primary:#FAF8F5;--bg-secondary:#F2EFEB;--bg-sidebar:#FFFFFF;
  --bg-hover:#EDE9E3;--bg-active:#E8F5E9;
  --text-primary:#1A1A1A;--text-secondary:#6B6B6B;--text-muted:#9B9B9B;
  --accent:#2D6A4F;--accent-light:#52B788;--accent-bg:#D8F3DC;--accent-blue:#2563EB;
  --border:#E5E2DD;--border-light:#F0EDE8;
  --shadow-sm:0 1px 3px rgba(0,0,0,.04);
  --shadow-md:0 4px 12px rgba(0,0,0,.06);
  --shadow-lg:0 8px 24px rgba(0,0,0,.08)
}
[data-theme="dark"]{
  --bg-primary:#1A1F1E;--bg-secondary:#242928;--bg-sidebar:#1E2322;
  --bg-hover:#2D3330;--bg-active:#1E3028;
  --text-primary:#E8E6E3;--text-secondary:#9BA8A3;--text-muted:#6B7875;
  --accent:#52B788;--accent-light:#74C69D;--accent-bg:#1B3A2D;--accent-blue:#60A5FA;
  --border:#2D3330;--border-light:#252928;
  --shadow-sm:0 1px 3px rgba(0,0,0,.2);
  --shadow-md:0 4px 12px rgba(0,0,0,.3);
  --shadow-lg:0 8px 24px rgba(0,0,0,.4)
}

*,*::before,*::after{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Pretendard','Noto Sans KR',sans-serif;background:var(--bg-primary);color:var(--text-primary);line-height:1.65;transition:background var(--transition),color var(--transition)}

.header{position:fixed;top:0;left:0;right:0;height:var(--header-h);z-index:100;background:var(--bg-sidebar);border-bottom:1px solid var(--border);display:flex;align-items:center;padding:0 20px;gap:12px;transition:background var(--transition),border-color var(--transition)}
.header-burger{display:none;cursor:pointer;padding:6px 8px;border:none;background:none;color:var(--text-primary);font-size:18px;border-radius:6px}
.header-burger:hover{background:var(--bg-hover)}
.header-logo{font-weight:900;font-size:17px;letter-spacing:-.5px;color:var(--text-primary);white-space:nowrap}
.header-logo span{color:var(--accent)}
.header-sub{font-size:12px;color:var(--text-secondary);font-weight:300}
.header-right{margin-left:auto;display:flex;align-items:center;gap:10px}
.badge-ver{padding:3px 10px;border-radius:20px;font-size:11px;font-weight:600;background:var(--bg-active);color:var(--accent)}
.dark-toggle{display:flex;align-items:center;gap:6px;padding:5px 12px;border-radius:20px;border:1px solid var(--border);background:var(--bg-secondary);cursor:pointer;font-size:13px;font-weight:500;color:var(--text-secondary);transition:all var(--transition)}
.dark-toggle:hover{border-color:var(--accent);color:var(--accent)}

.sidebar{position:fixed;top:var(--header-h);left:0;bottom:0;width:var(--sidebar-w);background:var(--bg-sidebar);border-right:1px solid var(--border);overflow-y:auto;z-index:90;transition:transform .3s ease,background var(--transition),border-color var(--transition)}
.sidebar::-webkit-scrollbar{width:4px}
.sidebar::-webkit-scrollbar-thumb{background:var(--border);border-radius:2px}
.nav-section{padding:14px 0 4px}
.nav-section-title{padding:0 16px;font-size:10px;text-transform:uppercase;letter-spacing:1.5px;color:var(--text-muted);font-weight:600;margin-bottom:6px}
.nav-item{display:flex;align-items:center;padding:9px 16px;cursor:pointer;color:var(--text-secondary);font-size:13px;font-weight:400;border-left:3px solid transparent;transition:background .12s,color .12s,border-color .12s}
.nav-item:hover{background:var(--bg-hover);color:var(--text-primary)}
.nav-item.active{background:var(--bg-active);color:var(--accent);border-left-color:var(--accent);font-weight:600}
.nav-item .icon{width:22px;margin-right:8px;text-align:center;font-size:14px}
.nav-sub{padding-left:24px}
.nav-sub .nav-item{font-size:12px;padding:7px 16px}
.nav-sub .nav-item::before{content:'';display:inline-block;width:5px;height:5px;border-radius:50%;background:var(--text-muted);margin-right:10px;flex-shrink:0}
.nav-sub .nav-item.active::before{background:var(--accent)}
.nav-group{border-bottom:1px solid var(--border)}
.nav-group-header{display:flex;align-items:center;padding:10px 16px;cursor:pointer;color:var(--text-primary);font-size:13px;font-weight:600;background:var(--bg-secondary);margin:2px 8px;border-radius:6px;transition:background .12s}
.nav-group-header:hover{background:var(--bg-hover)}
.nav-group-header .icon{width:22px;margin-right:10px;text-align:center;font-size:14px}
.nav-group-header .arrow{margin-left:auto;font-size:10px;color:var(--text-muted);transition:transform .2s}
.nav-group.open .arrow{transform:rotate(90deg)}
.nav-group .nav-sub{display:none}
.nav-group.open .nav-sub{display:block}

.main{margin-left:var(--sidebar-w);margin-top:var(--header-h);padding:28px;min-height:calc(100vh - var(--header-h));transition:background var(--transition)}

/* ── 탭 바 ── */
.tab-bar-wrap{display:none;position:fixed;top:var(--header-h);left:var(--sidebar-w);right:0;z-index:89;background:var(--bg-sidebar);border-bottom:1px solid var(--border);padding:0 24px;overflow-x:auto;height:var(--tab-bar-h);align-items:center}
body.tabs-visible .tab-bar-wrap{display:flex}
body.tabs-visible .main{margin-top:calc(var(--header-h) + var(--tab-bar-h))}
.tab-item{flex-shrink:0;height:var(--tab-bar-h);display:inline-flex;align-items:center;justify-content:center;padding:0 20px;font-size:14px;font-weight:500;color:var(--text-secondary);border:none;background:none;cursor:pointer;border-bottom:3px solid transparent;white-space:normal;max-width:140px;text-align:center;line-height:1.3;transition:color .15s,border-color .15s}
.tab-item.active{color:var(--accent);border-bottom-color:var(--accent);font-weight:600}

/* 통계 카드 */
.stat-row{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:24px}
.stat-card{background:var(--bg-secondary);border:1px solid var(--border);border-radius:var(--radius-sm);padding:16px;text-align:center;box-shadow:var(--shadow-md)}
.stat-num{font-size:28px;font-weight:900;color:var(--text-primary)}
.stat-label{font-size:12px;color:var(--text-secondary);margin-top:4px}
.stat-done .stat-num{color:#2D6A4F}
.stat-prog .stat-num{color:#2563EB}
.stat-todo .stat-num{color:#F59E0B}

/* 마일스톤 */
.milestone-row{display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:12px;margin-bottom:28px}
.ms-card{background:var(--bg-secondary);border:1px solid var(--border);border-radius:var(--radius-sm);padding:14px;box-shadow:var(--shadow-md);transition:transform .15s}
.ms-card:hover{transform:translateY(-2px)}
.ms-date{font-size:11px;font-weight:700;color:var(--accent-blue)}
.ms-title{font-size:14px;font-weight:800;margin:4px 0 2px;color:var(--text-primary)}
.ms-desc{font-size:11px;color:var(--text-secondary)}
.ms-urgent{border-left:4px solid #EF4444}
.ms-urgent .ms-date{color:#EF4444}
.ms-dday{border-left:4px solid var(--accent);background:linear-gradient(135deg,var(--bg-secondary),var(--accent-bg))}

/* 팀 헤더 */
.team-header{background:var(--bg-secondary);border:1px solid var(--border);border-radius:var(--radius-sm);padding:20px;margin-bottom:24px;box-shadow:var(--shadow-md)}
.team-header h1{font-size:20px;font-weight:900}
.team-members{font-size:13px;color:var(--text-secondary);margin-top:6px}
.team-mission{font-size:12px;color:var(--accent);font-weight:600;margin-top:4px}
.section-heading{font-size:15px;font-weight:700;color:var(--text-primary);margin:24px 0 12px;padding-bottom:8px;border-bottom:2px solid var(--border)}

/* 대시보드 */
.dash-section-title{font-size:15px;font-weight:700;color:var(--text-primary);margin:28px 0 12px;padding-bottom:8px;border-bottom:2px solid var(--border)}
.dash-section-title:first-child{margin-top:0}
.dashboard{display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:20px;margin-bottom:8px}
.dash-card{background:var(--bg-secondary);border:1px solid var(--border);border-radius:var(--radius);overflow:hidden;box-shadow:var(--shadow-md);cursor:pointer;transition:box-shadow var(--transition),transform var(--transition)}
.dash-card:hover{box-shadow:var(--shadow-lg);transform:translateY(-3px)}
.dash-card-header{padding:18px 20px;color:white}
.dash-card-header h3{font-size:14px;font-weight:700}
.dash-card-header p{font-size:11px;opacity:.85;margin-top:3px}
.dash-card-body{padding:14px 20px}
.sheet-list{list-style:none}
.sheet-list li{padding:8px 0;border-bottom:1px solid var(--border);font-size:13px;display:flex;align-items:center;cursor:pointer;color:var(--text-secondary);transition:color .12s}
.sheet-list li:last-child{border:none}
.sheet-list li:hover{color:var(--accent-blue)}
.sheet-list li::before{content:'→';margin-right:8px;color:var(--text-muted);font-size:11px}

.page{display:none}
.page.active{display:block}
.page-header{margin-bottom:24px}
.page-header h1{font-size:21px;font-weight:900;color:var(--text-primary)}
.breadcrumb{font-size:12px;color:var(--text-secondary);margin-bottom:8px}
.breadcrumb a{color:var(--accent-blue);text-decoration:none}
.breadcrumb a:hover{text-decoration:underline}

.table-wrap{overflow-x:auto;margin-bottom:28px;border-radius:var(--radius-sm);border:1px solid var(--border);box-shadow:var(--shadow-md);transition:border-color var(--transition)}
.section-title{padding:11px 18px;font-size:13px;font-weight:700;color:#FFF;letter-spacing:.2px}
.ops-table{width:100%;border-collapse:collapse;background:var(--bg-primary);font-size:13px;transition:background var(--transition)}
.ops-table thead th{padding:10px 12px;font-size:12px;font-weight:700;text-align:center;letter-spacing:.3px;white-space:nowrap;background:var(--bg-secondary)!important;color:var(--text-primary)!important;border-bottom:2px solid var(--border)}
.ops-table tbody tr{background:color-mix(in srgb,var(--row-bg,transparent) 35%,var(--bg-primary));min-height:40px}
.ops-table tbody td{padding:10px 12px;font-size:13px;line-height:1.6;vertical-align:top;border-bottom:1px solid var(--border-light);color:var(--text-primary);text-align:center}
.ops-table tbody td.text-left{text-align:left}
.ops-table tbody td:first-child{font-weight:600;white-space:nowrap}
.ops-table tbody tr:last-child td{border-bottom:none}
.ops-table tbody tr:hover{background:var(--bg-hover)!important}
.ops-table tbody tr:hover td{border-bottom-color:var(--border)}
.cell-wrap{white-space:pre-line}
.ops-table tbody td.merged-desc{padding:14px 20px;font-size:13px;line-height:1.7;color:var(--text-secondary);background:var(--bg-secondary);border-bottom:1px solid var(--border)}

/* 조직도 */
.org-chart-section{margin-bottom:28px;padding:24px;background:var(--bg-secondary);border:1px solid var(--border);border-radius:var(--radius);box-shadow:var(--shadow-md)}
.org-chart{display:flex;flex-direction:column;align-items:center;gap:16px}
.org-node{border:2px solid var(--border);border-radius:10px;padding:12px 16px;background:var(--bg-secondary);min-width:220px;transition:all var(--transition)}
.org-node:hover{box-shadow:var(--shadow-lg);transform:translateY(-2px)}
.org-node-title{font-size:14px;font-weight:800;color:var(--text-primary)}
.org-node-sub{font-size:11px;color:var(--text-secondary);margin-top:2px}
.org-node-members{font-size:12px;color:var(--text-primary);margin-top:6px;line-height:1.6}
.org-node-role{font-size:11px;color:var(--text-secondary);margin-top:4px;font-style:italic}
.org-node-note{font-size:11px;color:var(--accent);margin-top:4px;font-weight:600}
.org-hq{border-color:#0D1B2A;background:linear-gradient(135deg,#0D1B2A,#1B2A4A)}
.org-hq .org-node-title,.org-hq .org-node-sub,.org-hq .org-node-role,.org-hq .org-node-note{color:white}
.org-branches{display:flex;gap:16px;flex-wrap:wrap;justify-content:center;width:100%}
.org-branch{display:flex;flex-direction:column;gap:10px;flex:1;min-width:260px;max-width:400px}
.org-support{border-left:4px solid #2980B9}
.org-fb{border-left:4px solid #E74C3C}
.org-it{border-left:4px solid #1ABC9C}
.org-alba{border-left:4px solid #95A5A6}
.org-event{border-left:4px solid #E67E22}
.org-port{border-left:4px solid #27AE60}
.org-embark{border-left:4px solid #6C5CE7;margin-left:20px}
.org-temp{border-left:4px solid #F39C12;border-style:dashed}
.org-sub-node{margin-top:8px}
.org-tag{display:inline-block;font-size:10px;padding:1px 6px;border-radius:10px;font-weight:600;margin-left:4px;vertical-align:middle}
.tag-pre{background:#FEF3C7;color:#92400E}
.tag-dual{background:#DBEAFE;color:#1E40AF}
.tag-off{background:var(--bg-secondary);color:var(--text-secondary)}
[data-theme="dark"] .tag-pre{background:#78350F;color:#FDE68A}
[data-theme="dark"] .tag-dual{background:#1E3A5F;color:#93C5FD}
.org-legend{display:flex;gap:12px;flex-wrap:wrap;justify-content:center;padding:12px;margin-top:8px;border-top:1px solid var(--border)}

/* 매트릭스 필터 */
.matrix-toolbar{margin-bottom:16px;padding:14px 16px;background:var(--bg-secondary);border:1px solid var(--border);border-radius:var(--radius-sm)}
.matrix-search-row{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.matrix-search{flex:1;min-width:200px;padding:8px 12px;border:1px solid var(--border);border-radius:6px;font-size:13px;background:var(--bg-primary);color:var(--text-primary);outline:none;transition:border-color .2s}
.matrix-search:focus{border-color:var(--accent-blue)}
.matrix-btn{padding:7px 14px;border:1px solid var(--border);border-radius:6px;background:var(--bg-secondary);color:var(--text-primary);font-size:12px;font-weight:600;cursor:pointer;transition:all .15s;white-space:nowrap}
.matrix-btn:hover{background:var(--bg-hover);border-color:var(--accent-blue)}
.matrix-btn-dl{background:var(--accent-blue);color:white;border-color:var(--accent-blue)}
.matrix-btn-dl:hover{opacity:.85}
.matrix-counter{font-size:12px;color:var(--text-secondary);font-weight:600;white-space:nowrap}
.col-filter-wrap{position:relative;display:inline-block}
.col-filter-btn{background:none;border:none;cursor:pointer;font-size:11px;color:var(--text-secondary);padding:0 2px;vertical-align:middle}
.col-filter-btn.active{color:var(--accent-blue)}
.col-filter-drop{display:none;position:fixed;z-index:9999;background:var(--bg-secondary);border:1px solid var(--border);border-radius:8px;box-shadow:var(--shadow-lg);padding:8px;min-width:240px;max-height:400px;overflow-y:auto}
.col-filter-drop.open{display:block}
.col-filter-drop label{display:flex;align-items:center;gap:8px;font-size:14px;padding:10px 14px;min-height:40px;line-height:1.4;cursor:pointer;color:var(--text-primary);border-radius:6px;white-space:nowrap}
.col-filter-drop label input[type=checkbox]{width:16px;height:16px;cursor:pointer;flex-shrink:0}
.col-filter-drop label:hover{background:var(--bg-hover)}
.col-filter-actions{display:flex;gap:6px;margin-bottom:8px;border-bottom:1px solid var(--border);padding-bottom:8px}
.col-filter-actions button{font-size:12px;padding:6px 12px;min-height:32px;border:1px solid var(--border);border-radius:6px;background:var(--bg-primary);color:var(--text-primary);cursor:pointer}
.col-filter-actions button:hover{background:var(--bg-hover)}
@media(max-width:768px){
  .col-filter-drop{min-width:260px;max-height:60vh}
  .col-filter-drop label{font-size:14px;min-height:44px;padding:12px 16px}
  .col-filter-drop label input[type=checkbox]{width:18px;height:18px}
}
.badge-cat{display:inline-block;font-size:10px;padding:2px 8px;border-radius:10px;font-weight:600;white-space:nowrap}
.badge-status{display:inline-block;font-size:10px;padding:2px 8px;border-radius:10px;font-weight:700;white-space:nowrap}
.badge-status-done{background:#D1FAE5;color:#065F46}
.badge-status-progress{background:#DBEAFE;color:#1E40AF}
.badge-status-todo{background:#FEF3C7;color:#92400E}
[data-theme="dark"] .badge-status-done{background:#064E3B;color:#6EE7B7}
[data-theme="dark"] .badge-status-progress{background:#1E3A5F;color:#93C5FD}
[data-theme="dark"] .badge-status-todo{background:#78350F;color:#FDE68A}

.overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.4);z-index:85}

@media(max-width:768px){
  .sidebar{transform:translateX(-100%);box-shadow:none;border-right:none}
  .sidebar.open{transform:translateX(0);box-shadow:4px 0 20px rgba(0,0,0,.15)}
  .overlay.open{display:block}
  .header-burger{display:flex}
  .main{margin-left:0;padding:16px}
  .header-sub{display:none}
  .dashboard{grid-template-columns:1fr}
  .stat-row{grid-template-columns:repeat(2,1fr)}
  .milestone-row{grid-template-columns:1fr}
  .ops-table{font-size:11px}
  .ops-table thead th,.ops-table tbody td{padding:8px 6px}
  .org-branches{flex-direction:column}
  .org-branch{max-width:100%}
  .tab-bar-wrap{left:0}
}
@media(min-width:769px) and (max-width:1024px){
  :root{--sidebar-w:220px}
  .dashboard{grid-template-columns:repeat(2,1fr)}
  .stat-row{grid-template-columns:repeat(4,1fr)}
}

/* ── NEW 뱃지 ── */
.nav-new-badge{display:inline-block;margin-left:6px;padding:1px 6px;font-size:9px;font-weight:800;line-height:1.5;background:#EF4444;color:#FFF;border-radius:8px;letter-spacing:.5px;vertical-align:middle;cursor:help;animation:newPulse 2s ease-in-out infinite}
@keyframes newPulse{0%,100%{box-shadow:0 0 0 0 rgba(239,68,68,.5)}50%{box-shadow:0 0 0 4px rgba(239,68,68,0)}}

/* ── 타사 미팅 인사이트 ── */
.insights-wrap{margin-top:28px;background:var(--bg-secondary);border:1px solid var(--border);border-radius:var(--radius-sm);overflow:hidden;box-shadow:var(--shadow-md)}
.insights-wrap.insights-new{border-left:5px solid #EF4444}
.insights-toggle{display:flex;align-items:center;justify-content:space-between;padding:14px 18px;cursor:pointer;user-select:none;background:var(--bg-secondary)}
.insights-toggle:hover{background:var(--bg-hover)}
.insights-arrow{font-size:11px;color:var(--text-secondary);transition:transform .2s;margin-left:8px}
.insights-wrap.open .insights-arrow{transform:rotate(90deg)}
.insights-meta{font-size:11px;color:var(--text-secondary);white-space:nowrap}
.insights-body{display:none;padding:0}
.insights-wrap.open .insights-body{display:block}
.insights-table td{vertical-align:top}
.priority-badge{display:inline-block;padding:2px 10px;border-radius:12px;font-size:11px;font-weight:700;line-height:1.6}
.priority-badge.high{background:#FEE2E2;color:#DC2626}
.priority-badge.mid{background:#FEF3C7;color:#D97706}
.priority-badge.low{background:#D1FAE5;color:#059669}
[data-theme="dark"] .priority-badge.high{background:#7F1D1D;color:#FCA5A5}
[data-theme="dark"] .priority-badge.mid{background:#78350F;color:#FCD34D}
[data-theme="dark"] .priority-badge.low{background:#064E3B;color:#6EE7B7}"""

JS = r"""(function(){var t=localStorage.getItem('theme')||localStorage.getItem('dark');if(t==='dark'||t==='1')document.documentElement.setAttribute('data-theme','dark')})();
function toggleDark(){var isDark=document.documentElement.getAttribute('data-theme')==='dark';var next=isDark?'light':'dark';document.documentElement.setAttribute('data-theme',next);localStorage.setItem('theme',next);document.getElementById('dark-icon').textContent=next==='dark'?'☀️':'🌙'}
window.addEventListener('DOMContentLoaded',function(){var isDark=document.documentElement.getAttribute('data-theme')==='dark';document.getElementById('dark-icon').textContent=isDark?'☀️':'🌙';if(document.getElementById('matrix-toolbar'))matrixInit()});

function showPage(id){
  document.querySelectorAll('.page').forEach(function(p){p.classList.remove('active')});
  document.querySelectorAll('.nav-item').forEach(function(n){n.classList.remove('active')});
  // 모든 펼침 메뉴 닫기
  document.querySelectorAll('.nav-group.open').forEach(function(g){g.classList.remove('open')});

  var page=document.getElementById('page-'+id);
  if(page)page.classList.add('active');
  else document.getElementById('page-home').classList.add('active');

  // 해당 메뉴 active 표시 + 부모 그룹 펼침
  document.querySelectorAll('.nav-item').forEach(function(n){
    var oc=n.getAttribute('onclick')||'';
    if(oc.indexOf("'"+id+"'")!==-1||oc.indexOf("'"+id+"',")!==-1){
      n.classList.add('active');
      var g=n.closest('.nav-group');
      while(g){g.classList.add('open');g=g.parentElement.closest('.nav-group')}
    }
  });

  document.getElementById('sidebar').classList.remove('open');
  document.getElementById('overlay').classList.remove('open');
  window.scrollTo(0,0);
  renderTabBar(id,null);
}

function showPageSection(pageId, anchor){
  showPage(pageId);
  renderTabBar(pageId,anchor);
  setTimeout(function(){
    var el=document.getElementById(pageId+'-'+anchor);
    if(el)el.scrollIntoView({behavior:'smooth',block:'start'});
  },50);
}

function toggleGroup(el){
  var grp=el.parentElement;
  var wasOpen=grp.classList.contains('open');
  // 같은 레벨의 다른 그룹 닫기
  var parent=grp.parentElement;
  if(parent){parent.querySelectorAll(':scope > .nav-group.open').forEach(function(g){g.classList.remove('open')})}
  if(!wasOpen)grp.classList.add('open');
}

function toggleSidebar(){document.getElementById('sidebar').classList.toggle('open');document.getElementById('overlay').classList.toggle('open')}
function openGroup(pid){showPage(pid);document.querySelectorAll('.nav-group').forEach(function(g){g.querySelectorAll('.nav-item').forEach(function(i){if(i.getAttribute('onclick')&&i.getAttribute('onclick').indexOf(pid)!==-1)g.classList.add('open')})})}

/* ══ 매트릭스 필터 ══ */
/* 필터 컬럼: 카테고리(2), 담당팀(5), 담당자(6) — 시기(1)·상태(7) 제거 */
var _mxTable=null,_mxFilters={},_mxFilterCols={},_mxFilterable=[2,5,6];
var STATUS_SET=new Set(['완료','진행중','미착수']);
var CAT_COLORS={'계약·행정':'#EBF5FB','인력·조직':'#E8F8F5','VIP·의전':'#FDEDEC','물품·물류':'#FEF9E7','통신·IT':'#EBF5FB','프로그램·공연':'#FDF2E9','식음료':'#E8F8F5','기항지·CIQ':'#EAFAF1','승하선':'#F4ECF7','콘텐츠·인쇄':'#FDEBD0','운영지원팀':'#D6EAF8','정산·사후':'#FADBD8','★공연매니지먼트':'#FAE5D3','★교육·훈련':'#F5EEF8','★채용(알바)':'#FEF9E7','★채용(통역)':'#FEF9E7'};
var CAT_COLORS_DARK={'계약·행정':'#1B4F72','인력·조직':'#0E6655','VIP·의전':'#78281F','물품·물류':'#7D6608','통신·IT':'#1A5276','프로그램·공연':'#784212','식음료':'#0B5345','기항지·CIQ':'#145A32','승하선':'#4A235A','콘텐츠·인쇄':'#784212','운영지원팀':'#1B4F72','정산·사후':'#78281F','★공연매니지먼트':'#6E2C00','★교육·훈련':'#4A235A','★채용(알바)':'#7D6608','★채용(통역)':'#7D6608'};

/* 컬럼 인덱스 ci에 대한 필터 비교값 반환.
   담당자(col6): 8+cols 이고 STATUS 아닌 경우만 실제 담당자, 나머지는 '(없음)'.
   기타: col이 없으면 null(해당 필터 skip). */
function _getFilterVal(ci,tds){
  if(ci===6){
    if(tds.length>=8&&!STATUS_SET.has(tds[6].textContent.trim()))return tds[6].textContent.trim();
    return '(없음)';
  }
  if(ci>=tds.length)return null;
  return tds[ci].textContent.trim();
}

/* page-matrix-2 안의 데이터 테이블(인덱스 1~N) tbody tr을 순회하는 헬퍼.
   인덱스 0은 thead 전용 빈 테이블이므로 건너뜀. */
function _mxDataRows(cb){
  var page=document.getElementById('page-matrix-2');
  if(!page)return;
  page.querySelectorAll('.ops-table').forEach(function(tbl,ti){
    if(ti===0)return;
    tbl.querySelectorAll('tbody tr').forEach(cb);
  });
}

function matrixInit(){
  var page=document.getElementById('page-matrix-2');
  if(!page)return;
  _mxTable=page.querySelector('.ops-table');  /* thead 버튼 삽입 전용 */
  if(!_mxTable)return;
  var ths=_mxTable.querySelectorAll('thead th');
  _mxFilterable.forEach(function(ci){
    if(ci>=ths.length)return;
    var th=ths[ci];
    _mxFilterCols[ci]=th.textContent.trim();
    var vals=new Set();
    _mxDataRows(function(tr){
      if(tr.querySelector('.merged-desc'))return;
      var tds2=tr.querySelectorAll('td');var v=_getFilterVal(ci,tds2);if(v)vals.add(v);
    });
    _mxFilters[ci]=new Set(vals);
    var wrap=document.createElement('span');wrap.className='col-filter-wrap';
    var btn=document.createElement('button');btn.className='col-filter-btn';btn.textContent=' ▼';btn.setAttribute('data-col',ci);
    var drop=document.createElement('div');drop.className='col-filter-drop';
    drop.addEventListener('click',function(e){e.stopPropagation()});
    btn.onclick=function(e){
      e.stopPropagation();
      document.querySelectorAll('.col-filter-drop.open').forEach(function(d){if(d!==drop)d.classList.remove('open')});
      if(!drop.classList.contains('open')){
        /* position:fixed 좌표를 버튼 위치 기준으로 계산 → overflow:auto 컨테이너 탈출 */
        var r=btn.getBoundingClientRect();
        drop.style.top=(r.bottom+4)+'px';
        var left=r.left;
        if(left+260>window.innerWidth)left=window.innerWidth-268;
        drop.style.left=Math.max(4,left)+'px';
      }
      drop.classList.toggle('open');
    };
    var actions=document.createElement('div');actions.className='col-filter-actions';
    var sa=document.createElement('button');sa.textContent='전체';sa.onclick=function(e){e.stopPropagation();toggleAllChecks(drop,true,ci)};
    var sn=document.createElement('button');sn.textContent='해제';sn.onclick=function(e){e.stopPropagation();toggleAllChecks(drop,false,ci)};
    actions.appendChild(sa);actions.appendChild(sn);drop.appendChild(actions);
    Array.from(vals).sort().forEach(function(v){var lbl=document.createElement('label');var cb=document.createElement('input');cb.type='checkbox';cb.checked=true;cb.value=v;cb.setAttribute('data-col',ci);cb.onchange=function(e){e.stopPropagation();onFilterChange(ci)};lbl.appendChild(cb);var sp=document.createElement('span');sp.textContent=v;lbl.appendChild(sp);drop.appendChild(lbl)});
    wrap.appendChild(btn);wrap.appendChild(drop);th.appendChild(wrap);
  });
  applyBadges();updateCounter();
  document.addEventListener('click',function(){document.querySelectorAll('.col-filter-drop.open').forEach(function(d){d.classList.remove('open')})});
}
function toggleAllChecks(drop,state,ci){drop.querySelectorAll('input[type=checkbox]').forEach(function(cb){cb.checked=state});onFilterChange(ci)}
function onFilterChange(ci){
  var checked=new Set();document.querySelectorAll('input[data-col="'+ci+'"]:checked').forEach(function(cb){checked.add(cb.value)});
  _mxFilters[ci]=checked;
  var all=new Set();document.querySelectorAll('input[data-col="'+ci+'"]').forEach(function(cb){all.add(cb.value)});
  var btn=document.querySelector('.col-filter-btn[data-col="'+ci+'"]');
  if(btn)btn.classList.toggle('active',checked.size<all.size);
  matrixFilter();
}
function matrixFilter(){
  if(!_mxTable)return;
  var qEl=document.getElementById('matrix-search');
  var q=(qEl&&qEl.value||'').toLowerCase().trim();
  _mxDataRows(function(tr){
    var tds=tr.querySelectorAll('td');
    if(tr.querySelector('.merged-desc')){tr.style.display='';return}
    var show=true;
    for(var ci in _mxFilters){ci=parseInt(ci);var ct=_getFilterVal(ci,tds);if(ct===null)continue;if(!_mxFilters[ci].has(ct)){show=false;break}}
    /* 검색: 전 컬럼 텍스트 대상 (컬럼 수 불일치 대응) */
    if(show&&q){var st='';tds.forEach(function(td){st+=' '+td.textContent});if(st.toLowerCase().indexOf(q)===-1)show=false}
    tr.style.display=show?'':'none';
  });
  updateCounter();
}
function updateCounter(){
  if(!_mxTable)return;var t=0,v=0;
  _mxDataRows(function(tr){if(tr.querySelector('.merged-desc'))return;t++;if(tr.style.display!=='none')v++});
  var el=document.getElementById('matrix-counter');if(el)el.textContent=v+' / '+t+'건';
}
function matrixResetAll(){
  document.getElementById('matrix-search').value='';
  document.querySelectorAll('#page-matrix-2 input[type=checkbox]').forEach(function(cb){cb.checked=true});
  document.querySelectorAll('.col-filter-btn').forEach(function(b){b.classList.remove('active')});
  for(var ci in _mxFilters){var all=new Set();document.querySelectorAll('input[data-col="'+ci+'"]').forEach(function(cb){all.add(cb.value)});_mxFilters[ci]=all}
  matrixFilter();
}
function applyBadges(){
  var isDark=document.documentElement.getAttribute('data-theme')==='dark';
  _mxDataRows(function(tr){
    var tds=tr.querySelectorAll('td');
    if(tds.length>2){var cat=tds[2].textContent.trim();var bg=isDark?(CAT_COLORS_DARK[cat]||'#2C3E50'):(CAT_COLORS[cat]||'#F2F3F4');var tc=isDark?'#E0E0E0':'#1A1A1A';if(cat)tds[2].innerHTML='<span class="badge-cat" style="background:'+bg+';color:'+tc+'">'+cat+'</span>'}
    if(tds.length>7){var st=tds[7].textContent.trim();var cls='';if(st==='완료')cls='badge-status-done';else if(st==='진행중')cls='badge-status-progress';else if(st==='미착수')cls='badge-status-todo';if(cls)tds[7].innerHTML='<span class="badge-status '+cls+'">'+st+'</span>'}
  });
}
function getVisibleRows(){
  if(!_mxTable)return[];var h=[];_mxTable.querySelectorAll('thead th').forEach(function(th){h.push(th.childNodes[0].textContent.trim())});var d=[h];
  _mxDataRows(function(tr){if(tr.style.display==='none'||tr.querySelector('.merged-desc'))return;var r=[];tr.querySelectorAll('td').forEach(function(td){r.push(td.textContent.trim())});d.push(r)});return d;
}
function matrixDownloadCSV(){var d=getVisibleRows();var csv=d.map(function(r){return r.map(function(c){return '"'+c.replace(/"/g,'""')+'"'}).join(',')}).join('\n');var blob=new Blob(['\uFEFF'+csv],{type:'text/csv;charset=utf-8'});var a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='업무매트릭스_'+new Date().toISOString().slice(0,10)+'.csv';a.click()}
function matrixDownloadXlsx(){if(typeof XLSX==='undefined'){var s=document.createElement('script');s.src='https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';s.onload=function(){doXlsxDownload()};document.head.appendChild(s)}else{doXlsxDownload()}}
function doXlsxDownload(){var d=getVisibleRows();var ws=XLSX.utils.aoa_to_sheet(d);var wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'업무매트릭스');XLSX.writeFile(wb,'업무매트릭스_'+new Date().toISOString().slice(0,10)+'.xlsx')}

/* ══ 탭 바 ══ */
var TAB_CONFIG={
  'team-hq':[
    {key:'org-chart',label:'조직도'},
    {key:'progress',label:'업무 진행현황'},
    {key:'docs',label:'기안·협조공문'},
    {key:'timeline',label:'핵심 일정'}
  ],
  'team-support':[
    {key:'overview',label:'팀 개요·미션'},
    {key:'tasks',label:'업무 리스트'},
    {key:'sop',label:'SOP·시나리오'}
  ],
  'team-event':[
    {key:'overview',label:'팀 개요·미션'},
    {key:'tasks',label:'업무 리스트'},
    {key:'sop',label:'SOP·시나리오'},
    {key:'manual',label:'매뉴얼 (초안)'}
  ],
  'team-port':[
    {key:'overview',label:'팀 개요·미션'},
    {key:'tasks',label:'업무 리스트'},
    {key:'sop',label:'SOP·시나리오'},
    {key:'manual',label:'매뉴얼 (초안)'}
  ],
  'team-embark':[
    {key:'overview',label:'팀 개요'},
    {key:'tasks',label:'업무 리스트'},
    {key:'sop',label:'SOP·케이스스터디'}
  ],
  'team-fb':[
    {key:'overview',label:'팀 개요'},
    {key:'tasks',label:'업무 리스트'}
  ],
  'team-it':[
    {key:'overview',label:'팀 개요·미션'},
    {key:'tasks',label:'업무 리스트'},
    {key:'sop',label:'SOP·시나리오'},
    {key:'manual',label:'매뉴얼 (초안)'}
  ]
};
function renderTabBar(pageId,activeAnchor){
  var wrap=document.getElementById('tab-bar');
  if(!wrap)return;
  var tabs=TAB_CONFIG[pageId];
  if(!tabs){wrap.innerHTML='';document.body.classList.remove('tabs-visible');return}
  document.body.classList.add('tabs-visible');
  wrap.innerHTML=tabs.map(function(t){
    var cls='tab-item'+(t.key===activeAnchor?' active':'');
    return '<button class="'+cls+'" onclick="showPageSection(\''+pageId+'\',\''+t.key+'\')">'+t.label+'</button>';
  }).join('');
}
"""

# ── 최종 HTML 조립 ────────────────────────────────────────────────────
def build_html(manifest, modules):
    nav = build_nav(manifest)
    org_chart = build_org_chart()
    dashboard = build_dashboard(manifest, modules)
    insights = load_insights()
    team_pages = build_team_pages(manifest, modules, insights)
    timeline = build_master_timeline(manifest, modules)
    ref_pages = build_ref_pages(manifest, modules)

    return f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>모두의 크루즈 운영 데스크</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://cdn.jsdelivr.net" crossorigin>
<link href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css" rel="stylesheet">
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;700;900&family=JetBrains+Mono:wght@400;600&display=swap" rel="stylesheet">
<style>
{CSS}
</style>
</head>
<body>
<header class="header">
  <button class="header-burger" onclick="toggleSidebar()">☰</button>
  <div class="header-logo">모두의 <span>크루즈</span> 운영 데스크</div>
  <div class="header-sub">2026 코스타 세레나 한일전세선 (6.19~6.25)</div>
  <div class="header-right">
    <span class="badge-ver" title="{UPDATE_DATE} 업데이트">v4.0</span>
    <button class="dark-toggle" id="dark-btn" onclick="toggleDark()" title="다크모드 전환"><span id="dark-icon">🌙</span> 다크</button>
  </div>
</header>

{nav}

<div class="tab-bar-wrap" id="tab-bar"></div>

<div class="overlay" id="overlay" onclick="toggleSidebar()"></div>

<main class="main">
  <div class="page active" id="page-home">
    <div class="page-header">
      <h1>🚢 모두의 크루즈 운영 데스크</h1>
      <p>2026 코스타 세레나 한일전세선 (6.19~6.25) — 총괄운영 대시보드</p>
    </div>
{dashboard}
{org_chart}
  </div>

{timeline}
{team_pages}
{ref_pages}
</main>

<script>
{JS}
</script>
</body>
</html>"""

# ── 메인 ─────────────────────────────────────────────────────────────
def main():
    test_key = sys.argv[1] if len(sys.argv) > 1 else None
    print('=== 크루즈 운영 데스크 빌더 v4.0 ===')
    manifest, modules = load_all()
    if test_key:
        print(f'  [TEST MODE] {test_key}만 렌더링')
        for k in list(modules.keys()):
            if k != test_key: modules.pop(k)
    html = build_html(manifest, modules)
    with open(OUT, 'w', encoding='utf-8') as f:
        f.write(html)
    size_kb = os.path.getsize(OUT) // 1024
    print(f'  → {OUT} ({size_kb} KB) 생성 완료')

if __name__ == '__main__':
    main()
