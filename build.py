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
    'orgv3':'#283593','supplement':'#00695C','teamdocs':'#4527A0','matrix':'#2D6A4F',
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
        'members': 'HQ (모두투어 + CI) 공동운영',
        'mission': '전 팀 조율·관리 총괄 | VIP 의전 | 코스타 선사 최종 대응 | 예산·일정·계약 통합관리',
        'teamdocs_sheet': None,
        'matrix_teams': ['HQ', 'HQ(재무)', 'HQ→각 팀장', '전원'],
        'manual_keys': [],
        'sub_items': [
            ('미션·역할', 'mission'),
            ('의사결정체계', 'decision'),
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
        ],
    },
    {
        'id': 'team-port',
        'name': '기항지운영팀 (VIP대응팀 겸임)',
        'icon': '⚓',
        'color': '#27AE60',
        'members': '이수일 팀장 ★듀얼 + 팀원1 + 타이요(6+가이드70)',
        'mission': '하코다테·오타루 기항지 투어 운영 + VIP 의전 대응 + 승하선 겸임',
        'teamdocs_sheet': '기항지운영팀',
        'matrix_teams': ['기항지팀', '기항지+공연'],
        'manual_keys': ['port'],
        'sub_items': [
            ('팀 개요·미션', 'overview'),
            ('업무 리스트', 'tasks'),
            ('SOP·시나리오', 'sop'),
            ('VIP·가수로지스틱', 'vip-logistics'),
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
UPDATE_DATE = '2026-04-12'
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
                        # 흰색 계열은 기본 배경(var(--bg-primary))으로 처리 — 다크모드 호환
                        if rb and rb.upper() in {'#FFFFFF', '#FFF', '#FFFFFF'.lower(), '#FFF'.lower()}:
                            rb = ''
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
    L.append(f'    <div class="nav-item" onclick="showPage(\'master-timeline\')"><span class="icon">📋</span>핵심요약</div>')
    L.append(f'    <div class="nav-item" onclick="showPage(\'matrix-2\')"><span class="icon">📐</span>마스터 매트릭스</div>')
    L.append(f'    <div class="nav-item" onclick="showPage(\'common-extra\')"><span class="icon">📌</span>공통 추가업무</div>')
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
      <h2 style="font-size:18px;font-weight:800;margin-bottom:4px;color:var(--text-primary)">📊 선내운영 조직도 v4</h2>
      <p style="font-size:12px;color:var(--text-muted);margin-bottom:16px">각 팀 카드를 클릭하면 주요 업무를 확인할 수 있습니다.</p>
      <div class="org-chart">
        <div class="org-node org-hq" onclick="toggleOrgDrop(event,'hq')" style="cursor:pointer">
          <div class="org-node-title">HQ 운영본부 (모두투어 + CI)</div>
          <div class="org-node-role">공연/기항지/승하선/고객응대/정산/식음료/방송 등 전체 운영 과업을 공동 수행</div>
          <div class="org-node-note">VIP: 부서장 직속 관리</div>
          <div class="org-drop" id="org-drop-hq">전 팀 조율·관리 / VIP 의전 / 코스타 선사 최종 대응 / 예산·일정·계약 통합관리</div>
        </div>
        <div class="org-branches">
          <div class="org-branch">
            <div class="org-node org-support" onclick="toggleOrgDrop(event,'support')" style="cursor:pointer">
              <div class="org-node-title">운영지원팀</div>
              <div class="org-node-members">황지애(CI, 지원) · 양은희(모두투어) · 조아라(웅진) · 김지은(환경재단)</div>
              <div class="org-node-role">팀 과업, 유동배치</div>
              <div class="org-drop" id="org-drop-support">안내데스크 13시간 운영(08~21시) / 고객응대·방송·정산 / 3사 캐빈 취합·선사 전달</div>
            </div>
            <div class="org-node org-event" onclick="toggleOrgDrop(event,'event')" style="cursor:pointer">
              <div class="org-node-title">공연/행사팀 <span class="org-tag tag-pre">★사전탑승</span></div>
              <div class="org-node-members">이상민 매니저(팀장) + 팀원4 + 환경재단15명 협업</div>
              <div class="org-drop" id="org-drop-event">자체공연 운영 / 환경재단 선상학교 지원 / 출항식·폐막행사 / 음향·장비 관리</div>
            </div>
            <div class="org-node org-port" onclick="toggleOrgDrop(event,'port')" style="cursor:pointer">
              <div class="org-node-title">기항지운영팀 <span class="org-tag tag-dual">★듀얼</span> <span class="org-tag" style="background:#8E44AD;color:#fff">VIP대응팀 겸임</span></div>
              <div class="org-node-members">이수일 팀장 + 팀원1 + 타이요(6+가이드70)</div>
              <div class="org-drop" id="org-drop-port">하코다테·오타루 투어 운영 / 타이요 76명 관리 / 갱웨이 관리 / 중간하선·승선 CIQ / 각사별 VIP 의전 프로토콜 관리</div>
              <div class="org-sub-node">
                <div class="org-node org-embark" onclick="toggleOrgDrop(event,'embark')" style="cursor:pointer">
                  <div class="org-node-title">승하선팀 (기항지팀 산하)</div>
                  <div class="org-node-members">김남민 매니저 <span class="org-tag tag-off">미탑승/터미널</span> + 이수일<span class="org-tag tag-dual">★듀얼</span> + 알바3~4</div>
                  <div class="org-drop" id="org-drop-embark">터미널 운영 / 수화물 프로세스 / 승선·하선 동선 관리</div>
                </div>
              </div>
            </div>
          </div>
          <div class="org-branch">
            <div class="org-node org-fb" onclick="toggleOrgDrop(event,'fb')" style="cursor:pointer">
              <div class="org-node-title">식음료파트 <span class="org-tag tag-pre">★사전탑승</span></div>
              <div class="org-node-members">문형식 매니저</div>
              <div class="org-drop" id="org-drop-fb">코스타 F&amp;B팀 조율 / 밀스케줄 관리 / 특별식·에코정책 반영</div>
            </div>
            <div class="org-node org-it" onclick="toggleOrgDrop(event,'it')" style="cursor:pointer">
              <div class="org-node-title">IT홍보팀</div>
              <div class="org-node-members">최구철 매니저(팀장)</div>
              <div class="org-drop" id="org-drop-it">무전기·Wi-Fi 통신장비 / 촬영·홍보 / 프로그램표·인쇄물 제작</div>
            </div>
            <div class="org-node org-alba" onclick="toggleOrgDrop(event,'alba')" style="cursor:pointer">
              <div class="org-node-title">알바 12명</div>
              <div class="org-node-role">탄력 배치</div>
              <div class="org-drop" id="org-drop-alba">탄력배치 / 각 팀 유동 지원</div>
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
    org_chart = build_org_chart()

    org_summary = """
    <h3 class="dash-section-title">🏛️ 조직 체계 요약</h3>
    <div class="table-wrap">
    <table class="ops-table">
    <thead><tr>
      <th style="color:#FFFFFF;font-weight:700">구분</th>
      <th style="color:#FFFFFF;font-weight:700">직책/구성</th>
      <th style="color:#FFFFFF;font-weight:700">인원</th>
      <th style="color:#FFFFFF;font-weight:700">주요 역할</th>
      <th style="color:#FFFFFF;font-weight:700">세부 담당 업무</th>
      <th style="color:#FFFFFF;font-weight:700">비고</th>
    </tr></thead>
    <tbody>
    <tr>
      <td style="font-weight:700" class="cell-wrap">운영부서장</td>
      <td class="cell-wrap">운영부서장 2명</td>
      <td class="cell-wrap">2</td>
      <td class="cell-wrap">선내 전체 운영 총괄 및 공동 지휘</td>
      <td class="cell-wrap text-left">- 전세선 전체 운영 및 최종 의사결정<br>- 선사·협력사 실무 조율<br>- 예산/일정/계약/보고 통합관리<br>- VIP 의전 총괄<br>- 코스타 Hotel Director·Cruise Director와 직접 소통</td>
      <td class="cell-wrap">HQ (모두투어 + CI)<br>공동운영</td>
    </tr>
    <tr>
      <td style="font-weight:700" class="cell-wrap">운영지원팀</td>
      <td class="cell-wrap">운영부서장 겸임 + 팀원3</td>
      <td class="cell-wrap">5</td>
      <td class="cell-wrap">안내데스크·고객응대·현장 컨트롤타워</td>
      <td class="cell-wrap text-left">- 안내데스크 운영 08~21시(13시간)<br>- 방송·공지 송출 관리<br>- VOC 수집 및 피드백<br>- 카카오톡 실시간 상황판 관리<br>- 3사 캐빈 취합 및 선사 전달<br>- 정산 관리</td>
      <td class="cell-wrap">운영부서장 겸임</td>
    </tr>
    <tr>
      <td style="font-weight:700" class="cell-wrap">기항지운영팀<br><small>(VIP대응팀 겸임)</small></td>
      <td class="cell-wrap">이수일 팀장 + 팀원1 + 타이요76</td>
      <td class="cell-wrap">78</td>
      <td class="cell-wrap">기항지투어 관리, 입출국 지원, VIP 의전</td>
      <td class="cell-wrap text-left">- 하코다테·오타루 투어 운영 총괄<br>- 타이요 76명 관리<br>- 갱웨이 관리<br>- 중간하선/승선 CIQ<br>- 각사별 VIP 의전 프로토콜 관리<br>- 자유여행자 귀선 통제</td>
      <td class="cell-wrap">타이요플랜 협력<br>VIP대응 겸임</td>
    </tr>
    <tr>
      <td style="font-weight:700" class="cell-wrap">공연/행사팀</td>
      <td class="cell-wrap">이상민 팀장(★사전탑승) + 팀원4 + 환경재단15명</td>
      <td class="cell-wrap">20</td>
      <td class="cell-wrap">공연·이벤트·환경재단행사 통합운영</td>
      <td class="cell-wrap text-left">- 자체공연 운영<br>- 환경재단 선상학교 지원<br>- 출항식·폐막행사 진행<br>- 음향·장비 관리<br>- GRM·코스타 테크니션 협업</td>
      <td class="cell-wrap">환경재단 15명 배치</td>
    </tr>
    <tr>
      <td style="font-weight:700" class="cell-wrap">승하선팀</td>
      <td class="cell-wrap">김남민 매니저(터미널전담) + 이수일★듀얼 + 알바3~4</td>
      <td class="cell-wrap">6</td>
      <td class="cell-wrap">터미널 운영, 수화물, 동선 관리</td>
      <td class="cell-wrap text-left">- 2,400명 승선·하선 터미널 총괄<br>- 수화물 프로세스 관리<br>- 승선·하선 동선 관리<br>- 타사 참관·답사 조율</td>
      <td class="cell-wrap">기항지팀 산하</td>
    </tr>
    <tr>
      <td style="font-weight:700" class="cell-wrap">식음료파트</td>
      <td class="cell-wrap">문형식 매니저(★사전탑승)</td>
      <td class="cell-wrap">1</td>
      <td class="cell-wrap">전체 밀스케줄 관리 + 한식 품질 모니터링</td>
      <td class="cell-wrap text-left">- 코스타 F&amp;B팀 조율<br>- 밀스케줄 관리<br>- 특별식·에코정책 반영</td>
      <td class="cell-wrap">★사전탑승</td>
    </tr>
    <tr>
      <td style="font-weight:700" class="cell-wrap">IT홍보팀</td>
      <td class="cell-wrap">최구철 매니저(팀장)</td>
      <td class="cell-wrap">1</td>
      <td class="cell-wrap">IT·통신 인프라 + 홍보·콘텐츠</td>
      <td class="cell-wrap text-left">- 무전기·Wi-Fi 통신장비 관리<br>- 촬영·홍보<br>- 프로그램표·인쇄물 제작</td>
      <td class="cell-wrap"></td>
    </tr>
    <tr>
      <td style="font-weight:700" class="cell-wrap">알바</td>
      <td class="cell-wrap">탄력배치</td>
      <td class="cell-wrap">12</td>
      <td class="cell-wrap">각 팀 유동 지원</td>
      <td class="cell-wrap text-left">- 탄력배치, 각 팀 유동 지원<br>- 안내데스크 보조<br>- 행사보조·물류지원</td>
      <td class="cell-wrap">12명</td>
    </tr>
    <tr>
      <td style="font-weight:700" class="cell-wrap">통역</td>
      <td class="cell-wrap">채용 후 코스타 인계</td>
      <td class="cell-wrap">18</td>
      <td class="cell-wrap">코스타에서 업무지시 및 배치</td>
      <td class="cell-wrap text-left">- 채용 후 승선일 코스타에 인계<br>- 코스타에서 업무지시 및 배치</td>
      <td class="cell-wrap">채용 후 승선일 코스타 인계</td>
    </tr>
    </tbody>
    </table>
    </div>
"""

    report_line = """
    <h3 class="dash-section-title">📡 보고라인 및 커뮤니케이션 체계</h3>
    <div class="table-wrap">
    <table class="ops-table">
    <thead><tr>
      <th style="color:#FFFFFF;font-weight:700">단계</th>
      <th style="color:#FFFFFF;font-weight:700">방식</th>
      <th style="color:#FFFFFF;font-weight:700">주요 내용</th>
      <th style="color:#FFFFFF;font-weight:700">시간</th>
      <th style="color:#FFFFFF;font-weight:700">참석자</th>
      <th style="color:#FFFFFF;font-weight:700">비고</th>
    </tr></thead>
    <tbody>
    <tr>
      <td style="font-weight:700" class="cell-wrap">보고라인</td>
      <td class="cell-wrap">단일체계</td>
      <td class="cell-wrap text-left">운영부서장(HQ) → 팀장 → 팀원/보조</td>
      <td class="cell-wrap">상시</td>
      <td class="cell-wrap">전 인력</td>
      <td class="cell-wrap"></td>
    </tr>
    <tr>
      <td style="font-weight:700" class="cell-wrap">1차 커뮤니케이션</td>
      <td class="cell-wrap">무전기 실시간 교신</td>
      <td class="cell-wrap text-left">현장 긴급 대응, 고객이슈, 이동통제</td>
      <td class="cell-wrap">상시</td>
      <td class="cell-wrap">팀장급 이상 + 키맨</td>
      <td class="cell-wrap">100대</td>
    </tr>
    <tr>
      <td style="font-weight:700" class="cell-wrap">2차 커뮤니케이션</td>
      <td class="cell-wrap">카카오톡 단체방</td>
      <td class="cell-wrap text-left">실시간 보고, 사진/문서 공유, 결재라인</td>
      <td class="cell-wrap">상시</td>
      <td class="cell-wrap">전 스태프</td>
      <td class="cell-wrap">팀별+전체방</td>
    </tr>
    <tr>
      <td style="font-weight:700" class="cell-wrap">3차 커뮤니케이션</td>
      <td class="cell-wrap">전체미팅</td>
      <td class="cell-wrap text-left">주요현안 브리핑, 개선사항 공유</td>
      <td class="cell-wrap">매일 22:00</td>
      <td class="cell-wrap">전 관계자 (필수근무자 제외)</td>
      <td class="cell-wrap">모두투어/CI/재단/웅진/타이요</td>
    </tr>
    <tr>
      <td style="font-weight:700" class="cell-wrap">데이터 공유</td>
      <td class="cell-wrap">Google Sheet</td>
      <td class="cell-wrap text-left">문서 실시간 연동, 자동 기록, 일일보고</td>
      <td class="cell-wrap">상시</td>
      <td class="cell-wrap">팀장급 이상</td>
      <td class="cell-wrap">자동화 시스템</td>
    </tr>
    </tbody>
    </table>
    </div>
"""

    milestones = """
    <h3 class="dash-section-title">🗓️ 주차별 핵심 일정</h3>
    <div class="milestone-row">
      <div class="ms-card ms-urgent">
        <div class="ms-date">D-90~D-60 (3~4월)</div>
        <div class="ms-title">사전 행정·계약</div>
        <div class="ms-desc">TF 기안 상신 / 3사 계약 관리 / 알바·통역 채용 공고 / 조직도 확정</div>
      </div>
      <div class="ms-card">
        <div class="ms-date">D-60~D-30 (4~5월)</div>
        <div class="ms-title">계약·SOP·기항지</div>
        <div class="ms-desc">공연 계약+대금 착수 / 팀별 SOP 초안 완성 / 타이요 미팅 / 기항지 인스펙션</div>
      </div>
      <div class="ms-card">
        <div class="ms-date">D-30~D-14 (5월초~중)</div>
        <div class="ms-title">온보드미팅·발주</div>
        <div class="ms-desc">온보드미팅(5/27~6/1) / 물품 제작 발주 / 알바 교육 / 통역 OT</div>
      </div>
      <div class="ms-card">
        <div class="ms-date">D-14~D-7 (6/5~12)</div>
        <div class="ms-title">최종 시뮬레이션</div>
        <div class="ms-desc">최종 시뮬레이션 / 비상대응 훈련 / 전체 2차 브리핑</div>
      </div>
      <div class="ms-card">
        <div class="ms-date">D-7~D-1 (6/12~18)</div>
        <div class="ms-title">사전탑승·베이스캠프</div>
        <div class="ms-desc">사전탑승자 승선(6/13) / 베이스캠프 세팅(6/18) / 터미널 최종 점검</div>
      </div>
      <div class="ms-card ms-dday">
        <div class="ms-date">승선일 (6/19)</div>
        <div class="ms-title">D-Day 2,400명 승선</div>
        <div class="ms-desc">2,400명 승선 개시 / 출항식(21:00)</div>
      </div>
      <div class="ms-card">
        <div class="ms-date">항차 Day1~2 (6/20~21)</div>
        <div class="ms-title">해상일·하코다테</div>
        <div class="ms-desc">해상일 프로그램 / 하코다테 기항(13:00~22:00)</div>
      </div>
      <div class="ms-card">
        <div class="ms-date">항차 Day3~4 (6/22~23)</div>
        <div class="ms-title">오타루 오버나잇</div>
        <div class="ms-desc">오타루 오버나잇 / 자유여행객 관리</div>
      </div>
      <div class="ms-card">
        <div class="ms-date">항차 Day5+하선 (6/24~25)</div>
        <div class="ms-title">폐막·입항·하선</div>
        <div class="ms-desc">폐막행사 / 부산 입항(15:00) / 하선·수화물</div>
      </div>
      <div class="ms-card">
        <div class="ms-date">사후 D+1~D+7 (6/26~7/2)</div>
        <div class="ms-title">정산·결과보고</div>
        <div class="ms-desc">3사 정산 / 결과보고서 / 사후 홍보</div>
      </div>
    </div>
"""

    return org_chart + org_summary + report_line

# ── 팀 워크스페이스 페이지 ───────────────────────────────────────────
def build_team_pages(manifest, modules, insights=None):
    pages = []
    insight_idx = build_all_insight_idx(insights)

    def _insight_block(tid, anchor):
        """해당 팀·앵커에 배분된 타사인사이트 블록 반환."""
        dist = TEAM_INSIGHT_DISTRIBUTION.get(tid, {})
        ids = dist.get(anchor, [])
        return build_insight_rows_html(ids, insight_idx)

    def _build_hq(tm):
        """HQ 운영본부 전용 페이지 — 미션·역할 / 의사결정체계."""
        tid = tm['id']
        out = []
        out.append(f'  <div class="page" id="page-{tid}">')
        out.append(f'    <div class="breadcrumb"><a href="#" onclick="showPage(\'home\')">홈</a> / 팀별 워크스페이스 / HQ 운영본부</div>')
        out.append(f'    <div class="team-header" style="border-left:5px solid {tm["color"]}">')
        out.append(f'      <h1>{tm["icon"]} HQ 운영본부</h1>')
        out.append(f'      <div class="team-members">{esc(tm["members"])}</div>')
        out.append(f'      <div class="team-mission">{esc(tm["mission"])}</div>')
        out.append(f'    </div>')

        # ① 미션·역할
        out.append(f'    <div id="{tid}-mission">')
        out.append(f'    <h3 class="section-heading">🎯 미션·역할</h3>')
        out.append(f'''    <div class="table-wrap">
    <table class="ops-table">
    <thead><tr>
      <th style="color:#FFFFFF;font-weight:700">역할</th>
      <th style="color:#FFFFFF;font-weight:700">구성</th>
      <th style="color:#FFFFFF;font-weight:700">핵심 책임</th>
      <th style="color:#FFFFFF;font-weight:700">세부 업무</th>
    </tr></thead>
    <tbody>
    <tr>
      <td style="font-weight:700" class="cell-wrap">운영부서장</td>
      <td class="cell-wrap">모두투어 + CI (공동운영)</td>
      <td class="cell-wrap">선내 전체 운영 총괄 및 공동 지휘</td>
      <td class="cell-wrap text-left">- 전세선 전체 운영 및 최종 의사결정<br>- 선사·협력사 실무 조율<br>- 예산/일정/계약/보고 통합관리<br>- VIP 의전 총괄<br>- 코스타 Hotel Director·Cruise Director와 직접 소통</td>
    </tr>
    <tr>
      <td style="font-weight:700" class="cell-wrap">운영지원팀 겸임</td>
      <td class="cell-wrap">부서장 직접 겸임</td>
      <td class="cell-wrap">안내데스크·고객응대 총괄 허브</td>
      <td class="cell-wrap text-left">- 안내데스크 운영 총괄 (08~21시)<br>- 방송·공지 송출 최종 검수<br>- VOC 수집 및 개선 지시<br>- 3사 캐빈 취합 → 선사 전달</td>
    </tr>
    <tr>
      <td style="font-weight:700" class="cell-wrap">VIP 의전</td>
      <td class="cell-wrap">부서장 직속 관리</td>
      <td class="cell-wrap">3사 VIP 일정 및 의전 프로토콜 관리</td>
      <td class="cell-wrap text-left">- 3사별 VIP 명단 관리<br>- 기항지 VIP 의전 조율 (기항지운영팀과 협업)<br>- VIP 객실 특별 세팅<br>- VIP 전용 일정표 배포</td>
    </tr>
    <tr>
      <td style="font-weight:700" class="cell-wrap">코스타 선사 대응</td>
      <td class="cell-wrap">HQ 단독 창구</td>
      <td class="cell-wrap">선사와의 공식 커뮤니케이션 유일 창구</td>
      <td class="cell-wrap text-left">- 코스타 Hotel Director 일일 브리핑 참석<br>- 운영 이슈 선사 보고 및 조율<br>- Emergency 채널 대응<br>- 운항 일정 변경 시 전 팀 즉시 전파</td>
    </tr>
    <tr>
      <td style="font-weight:700" class="cell-wrap">예산·계약 관리</td>
      <td class="cell-wrap">HQ 통합</td>
      <td class="cell-wrap">3사 공동 예산 및 계약 사항 통합 관리</td>
      <td class="cell-wrap text-left">- 현장 집행 예산 승인<br>- 협력사 지급 조율<br>- 계약 변경 사항 검토<br>- 정산 준비 (D+1~D+7)</td>
    </tr>
    </tbody>
    </table>
    </div>''')
        # 미션 인사이트 배분
        ins_mission = _insight_block(tid, 'mission')
        if ins_mission: out.append(ins_mission)
        out.append(f'    </div>')

        # ② 의사결정체계
        out.append(f'    <div id="{tid}-decision">')
        out.append(f'    <h3 class="section-heading">⚡ 의사결정체계</h3>')
        out.append(f'''    <div class="table-wrap">
    <table class="ops-table">
    <thead><tr>
      <th style="color:#FFFFFF;font-weight:700">상황</th>
      <th style="color:#FFFFFF;font-weight:700">결정권자</th>
      <th style="color:#FFFFFF;font-weight:700">프로세스</th>
      <th style="color:#FFFFFF;font-weight:700">수단</th>
      <th style="color:#FFFFFF;font-weight:700">비고</th>
    </tr></thead>
    <tbody>
    <tr>
      <td style="font-weight:700" class="cell-wrap">일반 현장운영</td>
      <td class="cell-wrap">해당 팀장</td>
      <td class="cell-wrap text-left">팀장 판단 → 즉시 실행 → HQ 보고</td>
      <td class="cell-wrap">무전기·카카오톡</td>
      <td class="cell-wrap">팀장 자율 처리</td>
    </tr>
    <tr>
      <td style="font-weight:700" class="cell-wrap">팀 간 조율 이슈</td>
      <td class="cell-wrap">HQ 운영부서장</td>
      <td class="cell-wrap text-left">팀장 보고 → HQ 검토 → 조율 결정 → 전파</td>
      <td class="cell-wrap">무전기·전체 카톡방</td>
      <td class="cell-wrap">30분 내 결정 원칙</td>
    </tr>
    <tr>
      <td style="font-weight:700" class="cell-wrap">예산 집행</td>
      <td class="cell-wrap">HQ 공동 합의</td>
      <td class="cell-wrap text-left">팀장 요청 → HQ 검토 → 공동 승인</td>
      <td class="cell-wrap">카카오톡·문서</td>
      <td class="cell-wrap">50만원 이상 사전 승인 필수</td>
    </tr>
    <tr>
      <td style="font-weight:700" class="cell-wrap">긴급 비상상황</td>
      <td class="cell-wrap">HQ 운영부서장</td>
      <td class="cell-wrap text-left">즉시 결정 → 팀장 전파 → 코스타 동시 보고</td>
      <td class="cell-wrap">무전기 비상채널</td>
      <td class="cell-wrap">코스타 Emergency 병행</td>
    </tr>
    <tr>
      <td style="font-weight:700" class="cell-wrap">선사 공식 소통</td>
      <td class="cell-wrap">HQ 운영부서장</td>
      <td class="cell-wrap text-left">HQ 단독 창구로 코스타에 직접 소통</td>
      <td class="cell-wrap">코스타 내선·영어</td>
      <td class="cell-wrap">팀장 직접 선사 소통 금지</td>
    </tr>
    <tr>
      <td style="font-weight:700" class="cell-wrap">VIP 이슈</td>
      <td class="cell-wrap">HQ 운영부서장</td>
      <td class="cell-wrap text-left">현장팀 보고 → HQ → 해당사 담당자 즉시 연락</td>
      <td class="cell-wrap">별도 VIP 채널</td>
      <td class="cell-wrap">3사 담당자 사전 배정</td>
    </tr>
    </tbody>
    </table>
    </div>''')
        # 의사결정 인사이트 배분
        ins_decision = _insight_block(tid, 'decision')
        if ins_decision: out.append(ins_decision)
        out.append(f'    </div>')

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
        # 업무분장 인사이트 배분
        ins_tasks = _insight_block(tid, 'tasks')
        if ins_tasks: out.append(ins_tasks)
        out.append(f'    </div>')

        # SOP·시나리오 (sop 앵커) — 매뉴얼 SOP 시트만
        out.append(f'    <div id="{tid}-sop">')
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
        # SOP 인사이트 배분
        ins_sop = _insight_block(tid, 'sop')
        if ins_sop: out.append(ins_sop)
        out.append(f'    </div>')

        # VIP·가수로지스틱 (vip-logistics 앵커) — team-port 전용
        if tid == 'team-port':
            out.append(f'    <div id="{tid}-vip-logistics">')
            out.append(f'    <h3 class="section-heading">🌟 VIP·가수로지스틱 (초대가수 승하선 + VIP 의전)</h3>')
            if 'master' in modules:
                msheets = modules['master'].get('sheets', [])
                if len(msheets) > 4:
                    _, html = render_sheet(msheets[4], 'master')
                    out.append(html)
            # VIP 인사이트 배분
            ins_vip = _insight_block(tid, 'vip-logistics')
            if ins_vip: out.append(ins_vip)
            out.append(f'    </div>')

        out.append('  </div>')
        pages.append('\n'.join(out))

    for tm in TEAM_WORKSPACE:
        _build_one(tm)
        if 'sub_team' in tm:
            _build_one(tm['sub_team'])

    return '\n'.join(pages)

# 마스터 타임라인에서 대시보드·공통 추가업무로 이동한 시트 목록
MASTER_TIMELINE_EXCLUDE_SHEETS = {'업무분장표', '선상운영 타임스케줄', '준비물 체크리스트', '특수업무·VIP의전'}

# ── 핵심준비사항 및 일정 (리드보드) 데이터 ───────────────────────────
MILESTONE_DATA = [
    {'id': 'ms-0', 'period': 'D-90~D-60 (3~4월)',     'title': '사전 행정·계약',      'detail': 'TF 기안 상신 / 3사 계약 관리 / 알바·통역 채용 공고 / 조직도 확정'},
    {'id': 'ms-1', 'period': 'D-60~D-30 (4~5월)',     'title': '계약·SOP·기항지',     'detail': '공연 계약+대금 착수 / 팀별 SOP 초안 완성 / 타이요 미팅 / 기항지 인스펙션'},
    {'id': 'ms-2', 'period': 'D-30~D-14 (5월초~중)',  'title': '온보드미팅·발주',      'detail': '온보드미팅(5/27~6/1) / 물품 제작 발주 / 알바 교육 / 통역 OT'},
    {'id': 'ms-3', 'period': 'D-14~D-7 (6/5~12)',     'title': '최종 시뮬레이션',      'detail': '최종 시뮬레이션 / 비상대응 훈련 / 전체 2차 브리핑'},
    {'id': 'ms-4', 'period': 'D-7~D-1 (6/12~18)',     'title': '사전탑승·베이스캠프',  'detail': '사전탑승자 승선(6/13) / 베이스캠프 세팅(6/18) / 터미널 최종 점검'},
    {'id': 'ms-5', 'period': '승선일 (6/19)',           'title': 'D-Day 2,400명 승선',  'detail': '2,400명 승선 개시 / 출항식(21:00)'},
    {'id': 'ms-6', 'period': '항차 Day1~2 (6/20~21)', 'title': '해상일·하코다테',      'detail': '해상일 프로그램 / 하코다테 기항(13:00~22:00)'},
    {'id': 'ms-7', 'period': '항차 Day3~4 (6/22~23)', 'title': '오타루 오버나잇',      'detail': '오타루 오버나잇 / 자유여행객 관리'},
    {'id': 'ms-8', 'period': '항차 Day5+하선 (6/24~25)', 'title': '폐막·입항·하선',   'detail': '폐막행사 / 부산 입항(15:00) / 하선·수화물'},
    {'id': 'ms-9', 'period': '사후 D+1~D+7 (6/26~7/2)', 'title': '정산·결과보고',    'detail': '3사 정산 / 결과보고서 / 사후 홍보'},
]

# 팀별 인사이트 배분 매핑: {tid: {section_anchor: [id_list]}}
TEAM_INSIGHT_DISTRIBUTION = {
    'team-hq':      {'mission': ['HQ-03'], 'decision': ['HQ-01', 'HQ-02']},
    'team-support': {'tasks': ['SUP-04'], 'sop': ['SUP-01', 'SUP-02', 'SUP-03']},
    'team-event':   {'tasks': ['PRG-02'], 'sop': ['PRG-01']},
    'team-port':    {
        'tasks':         ['HQ-05', 'PRT-05', 'PRT-06', 'PRT-07', 'PRT-08'],
        'sop':           ['PRT-01', 'PRT-02', 'PRT-03', 'PRT-04'],
        'vip-logistics': ['HQ-04'],
    },
    'team-it':      {'sop': ['IT-01']},
}

def build_all_insight_idx(insights):
    """insights JSON → {id: item} 딕셔너리."""
    idx = {}
    if not insights:
        return idx
    for team_key, team_data in insights.items():
        if team_key == 'meta':
            continue
        for item in team_data.get('items', []):
            idx[item['id']] = item
    return idx

def build_insight_rows_html(ids, insight_idx):
    """지정된 ID 목록 인사이트를 팀 업무에 통합된 행 테이블로 렌더링."""
    items = [insight_idx[i] for i in ids if i in insight_idx]
    if not items:
        return ''
    priority_cls = {'높음': 'high', '중간': 'mid', '낮음': 'low'}
    o = ['<div class="table-wrap">']
    o.append('<table class="ops-table"><thead><tr>')
    o.append('<th style="width:70px">ID</th>')
    o.append('<th>과업명 <span style="color:#DC3545;font-size:11px;font-weight:400">(타사인사이트)</span></th>')
    o.append('<th>내용</th><th style="width:80px">시점</th><th style="width:70px">우선순위</th>')
    o.append('</tr></thead><tbody>')
    for it in items:
        pcls = priority_cls.get(it.get('priority', ''), '')
        o.append('<tr>')
        o.append(f'<td><code style="color:#DC3545">{esc(it["id"])}</code></td>')
        o.append(f'<td style="color:#DC3545;font-weight:700">{esc(it["title"])}</td>')
        o.append(f'<td class="text-left">{esc(it["content"])}</td>')
        o.append(f'<td>{esc(it.get("phase",""))}</td>')
        o.append(f'<td><span class="priority-badge {pcls}">{esc(it.get("priority",""))}</span></td>')
        o.append('</tr>')
    o.append('</tbody></table></div>')
    return '\n'.join(o)

def build_milestone_table():
    """핵심준비사항 및 일정 — 리드보드 테이블 (진척 상태 드롭다운 + 프로그레스 바)."""
    o = []
    o.append('    <div class="ms-lead-board">')
    o.append('      <div class="ms-progress-wrap">')
    o.append('        <div class="ms-progress-bar">')
    o.append('          <div class="ms-bar-seg ms-bar-done" id="ms-bar-done"></div>')
    o.append('          <div class="ms-bar-seg ms-bar-prog" id="ms-bar-prog"></div>')
    o.append('          <div class="ms-bar-seg ms-bar-hold" id="ms-bar-hold"></div>')
    o.append('        </div>')
    o.append('        <div class="ms-progress-text" id="ms-progress-text">완료 0 / 진행중 0 / 진행전 10 / 보류 0</div>')
    o.append('      </div>')
    o.append('      <div class="table-wrap" style="margin-bottom:0;border-radius:0;border:none;box-shadow:none">')
    o.append('        <table class="ops-table ms-lead-table"><thead><tr>')
    o.append('          <th style="width:155px">시기</th>')
    o.append('          <th style="width:150px">핵심 항목</th>')
    o.append('          <th>세부 내용</th>')
    o.append('          <th style="width:125px">상태</th>')
    o.append('        </tr></thead><tbody>')
    for ms in MILESTONE_DATA:
        o.append(f'        <tr class="ms-row" data-ms-id="{ms["id"]}">')
        o.append(f'          <td class="ms-period-cell">{esc(ms["period"])}</td>')
        o.append(f'          <td style="font-weight:700">{esc(ms["title"])}</td>')
        o.append(f'          <td class="text-left">{esc(ms["detail"])}</td>')
        o.append(f'          <td>')
        o.append(f'            <select class="ms-status-sel" data-ms-id="{ms["id"]}" onchange="msStatusChange(this)">')
        o.append(f'              <option value="진행전">⚪ 진행전</option>')
        o.append(f'              <option value="진행중">🟡 진행중</option>')
        o.append(f'              <option value="완료">🟢 완료</option>')
        o.append(f'              <option value="보류">🔴 보류/폐기</option>')
        o.append(f'            </select>')
        o.append(f'          </td>')
        o.append(f'        </tr>')
    o.append('        </tbody></table>')
    o.append('      </div>')
    o.append('    </div>')
    return '\n'.join(o)

# ── 마스터 타임라인 페이지 ───────────────────────────────────────────
def build_master_timeline(manifest, modules):
    """핵심요약(리드보드) + master JSON 사전준비 마스터플랜 시트."""
    parts = []
    parts.append('  <div class="page" id="page-master-timeline">')
    parts.append('    <div class="breadcrumb"><a href="#" onclick="showPage(\'home\')">홈</a> / 전체 조망 / 핵심요약</div>')
    parts.append('    <div class="page-header"><h1>📋 핵심요약</h1><p style="font-size:13px;color:var(--text-secondary);margin-top:6px">전체 일정의 핵심준비사항을 누락 없이 확인하는 리드보드(Lead Board). 상태를 직접 관리하세요.</p></div>')

    # ① 핵심준비사항 및 일정 — 리드보드 (대시보드에서 이동)
    parts.append('    <h3 class="section-heading">🗓️ 핵심준비사항 및 일정</h3>')
    parts.append(build_milestone_table())

    # ② master JSON 시트들 (사전준비 마스터플랜만)
    parts.append('    <h3 class="section-heading" style="margin-top:36px">📋 사전준비 마스터플랜</h3>')
    if 'master' in modules and 'master' in manifest:
        mi = manifest['master']
        for si, sheet in enumerate(modules['master'].get('sheets',[])):
            sname = mi['sheets'][si] if si < len(mi['sheets']) else sheet.get('name','')
            if sname in MASTER_EXCLUDE_SHEETS: continue
            if sname in MASTER_TIMELINE_EXCLUDE_SHEETS: continue
            _, html = render_sheet(sheet, 'master')
            parts.append(html)

    parts.append('  </div>')
    return '\n'.join(parts)

# ── 공통 추가업무 페이지 ─────────────────────────────────────────────
def build_common_extra_page(manifest, modules):
    """타사 스케줄·참관 계획 / 기안·협조공문 / 외부협조 / 승인기안 등 통합."""
    parts = []
    parts.append('  <div class="page" id="page-common-extra">')
    parts.append('    <div class="breadcrumb"><a href="#" onclick="showPage(\'home\')">홈</a> / 전체 조망 / 공통 추가업무</div>')
    parts.append('    <div class="page-header"><h1>📌 공통 추가업무</h1></div>')

    # ① 타사 스케줄·참관 계획 + 참관·답사·사전탑승 일정 + 사전탑승자 보고체계
    #    (matrix sheet 0 — 이동: 마스터 타임라인 → 공통 추가업무)
    if 'matrix' in modules:
        msheets = modules['matrix'].get('sheets', [])
        if len(msheets) > 0:
            parts.append('    <h3 class="section-heading">📐 타사 스케줄·참관 계획 / 사전탑승자 보고체계</h3>')
            _, html = render_sheet(msheets[0], 'matrix')
            parts.append(html)

    # ② 기안·협조공문 (orgv3 sheet 1, 2 — 이동: HQ 운영본부 → 공통 추가업무)
    if 'orgv3' in modules:
        osheets = modules['orgv3'].get('sheets', [])
        parts.append('    <h3 class="section-heading">📝 기안·협조공문 / 외부 협조 요청 / 조직 승인 기안 / 별도 기안</h3>')
        for si in [1, 2]:
            if si < len(osheets):
                _, html = render_sheet(osheets[si], 'orgv3')
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
            # 준비물 체크리스트 페이지에 물품 목록 제목 추가 [9]
            extra_heading = ''
            if key == 'master' and sname == '준비물 체크리스트':
                extra_heading = '    <h3 class="section-heading">📦 물품 목록</h3>\n'

            pages.append(f'''  <div class="page" id="page-{pid}">
    <div class="breadcrumb"><a href="#" onclick="showPage('home')">홈</a> / {esc(label)} / {esc(sname)}</div>
    <div class="page-header"><h1>{icon} {esc(display_title)}</h1></div>
{extra}
{extra_heading}{html}
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
  --bg-primary:#FAF8F5;--bg-secondary:#F2EFEB;--bg-sidebar:#FAF8F5;
  --bg-hover:#EDE9E3;--bg-active:#E8F5E9;
  --text-primary:#1A1A1A;--text-secondary:#6B6B6B;--text-muted:#9B9B9B;
  --accent:#2D6A4F;--accent-light:#52B788;--accent-bg:#D8F3DC;--accent-blue:#2563EB;
  --border:#E5E2DD;--border-light:#F0EDE8;
  --shadow-sm:0 1px 3px rgba(0,0,0,.04);
  --shadow-md:0 4px 12px rgba(0,0,0,.06);
  --shadow-lg:0 8px 24px rgba(0,0,0,.08)
}
[data-theme="dark"]{
  --bg-primary:#1A1F1E;--bg-secondary:#242928;--bg-sidebar:#1A1F1E;
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
.theme-toggle{display:flex;align-items:center;background:var(--bg-secondary);border-radius:20px;padding:4px;border:1px solid var(--border)}
.theme-toggle button{padding:6px 12px;border-radius:16px;font-size:13px;font-weight:500;background:transparent;border:none;cursor:pointer;color:var(--text-secondary);transition:all var(--transition)}
.theme-toggle button.active{background:var(--accent);color:white}
.theme-toggle button:not(.active):hover{color:var(--text-primary)}

.sidebar{position:fixed;top:var(--header-h);left:0;bottom:0;width:var(--sidebar-w);background:var(--bg-sidebar);border-right:1px solid var(--border);overflow-y:auto;z-index:90;transition:transform .3s ease,background var(--transition),border-color var(--transition)}
.sidebar::-webkit-scrollbar{width:4px}
.sidebar::-webkit-scrollbar-thumb{background:var(--border);border-radius:2px}
.nav-section{padding:14px 0 4px}
.nav-section-title{padding:0 16px;font-size:10px;text-transform:uppercase;letter-spacing:1.5px;color:var(--text-muted);font-weight:600;margin-bottom:6px}
.nav-item{display:flex;align-items:center;padding:9px 16px;cursor:pointer;color:var(--text-secondary);font-size:13px;font-weight:400;border-left:3px solid transparent;transition:background .12s,color .12s,border-color .12s}
.nav-item:hover{background:var(--bg-hover);color:var(--text-primary);border-left-color:#CCCCCC}
.nav-item.active{background:var(--bg-active);color:var(--accent);border-left-color:var(--accent);font-weight:600}
.nav-item.active:hover{background:var(--bg-active);color:var(--accent);border-left-color:var(--accent)}
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

/* 섹션 카드 */
.section-card{background:var(--bg-secondary);border:1px solid var(--border);border-radius:var(--radius);padding:24px;margin-bottom:20px;box-shadow:var(--shadow-sm)}
.section-card h3{margin-bottom:16px;padding-bottom:12px;border-bottom:1px solid var(--border-light)}

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
.breadcrumb{font-size:13px;color:var(--text-muted);margin-bottom:16px}
.breadcrumb a{color:var(--accent);text-decoration:none}
.breadcrumb a:hover{text-decoration:underline}

.table-wrap{overflow-x:auto;margin-bottom:28px;border-radius:var(--radius-sm);border:1px solid var(--border);box-shadow:var(--shadow-md);transition:border-color var(--transition)}
.section-title{padding:11px 18px;font-size:13px;font-weight:700;color:#FFF;letter-spacing:.2px}
.ops-table{width:100%;border-collapse:collapse;background:var(--bg-primary);font-size:13px;transition:background var(--transition)}
.ops-table thead th{padding:12px 16px;font-size:13px;font-weight:600;text-align:center;letter-spacing:.03em;text-transform:uppercase;white-space:nowrap;background:var(--bg-secondary)!important;color:var(--text-secondary)!important;border-bottom:1px solid var(--border)}
.ops-table tbody tr{background:color-mix(in srgb,var(--row-bg,transparent) 35%,var(--bg-primary));min-height:40px}
.ops-table tbody td{padding:14px 16px;font-size:14px;line-height:1.6;vertical-align:top;border-bottom:1px solid var(--border-light);color:var(--text-primary);text-align:center}
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
.org-drop{display:none;margin-top:10px;padding:8px 12px;background:rgba(0,0,0,0.07);border-radius:8px;font-size:12px;line-height:1.7;color:var(--text-primary);text-align:left;border-left:3px solid var(--accent)}
[data-theme="dark"] .org-drop{background:rgba(255,255,255,0.08)}
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
.col-filter-drop{display:none;position:fixed;z-index:9999;background:var(--bg-sidebar);border:1px solid var(--border);border-radius:8px;box-shadow:var(--shadow-md);padding:8px;min-width:240px;max-height:400px;overflow-y:auto}
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
.badge-done,.badge-status-done{background:var(--accent-bg);color:var(--accent)}
.badge-progress,.badge-status-progress{background:#FFF3CD;color:#856404}
.badge-todo,.badge-status-todo{background:#F8D7DA;color:#721C24}
[data-theme="dark"] .badge-done,[data-theme="dark"] .badge-status-done{background:var(--accent-bg);color:var(--accent)}
[data-theme="dark"] .badge-progress,[data-theme="dark"] .badge-status-progress{background:#3D2B00;color:#FCD34D}
[data-theme="dark"] .badge-todo,[data-theme="dark"] .badge-status-todo{background:#3D0D14;color:#FCA5A5}

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
[data-theme="dark"] .priority-badge.low{background:#064E3B;color:#6EE7B7}

/* ── 핵심준비사항 리드보드 ── */
.ms-lead-board{background:var(--bg-secondary);border:1px solid var(--border);border-radius:var(--radius-sm);margin-bottom:28px;overflow:hidden;box-shadow:var(--shadow-md)}
.ms-progress-wrap{padding:14px 20px;border-bottom:1px solid var(--border);background:var(--bg-secondary)}
.ms-progress-bar{height:10px;background:var(--bg-primary);border-radius:5px;overflow:hidden;display:flex;margin-bottom:6px}
.ms-bar-seg{height:100%;transition:width .4s ease}
.ms-bar-done{background:#27AE60}
.ms-bar-prog{background:#F39C12}
.ms-bar-hold{background:#E74C3C}
.ms-progress-text{font-size:12px;font-weight:600;color:var(--text-secondary)}
.ms-lead-table .ms-period-cell{white-space:nowrap;font-size:12px;color:var(--accent-blue);font-weight:600}
.ms-status-sel{width:116px;padding:5px 8px;border:1px solid var(--border);border-radius:6px;background:var(--bg-primary);color:var(--text-primary);font-size:12px;cursor:pointer;outline:none}
.ms-status-sel:focus{border-color:var(--accent)}
.ms-row[data-ms-status="완료"]{background:rgba(39,174,96,0.10)!important}
.ms-row[data-ms-status="진행중"]{background:rgba(243,156,18,0.10)!important}
.ms-row[data-ms-status="보류"]{background:rgba(231,76,60,0.10)!important}

/* 아코디언 CSS 제거됨 */

/* ── 팀별 Excel 버튼 (탭 바) ── */
.tab-spacer{flex:1 1 auto}
.tab-excel-btn{flex-shrink:0;height:32px;padding:0 14px;margin:auto 0;border-radius:6px;font-size:13px;font-weight:600;background:var(--accent);color:#fff;border:none;cursor:pointer;white-space:nowrap;transition:opacity .15s}
.tab-excel-btn:hover{opacity:.82}"""

JS = r"""(function(){var t=localStorage.getItem('theme')||localStorage.getItem('dark');if(t==='dark'||t==='1')document.documentElement.setAttribute('data-theme','dark')})();
function _syncThemeButtons(){var isDark=document.documentElement.getAttribute('data-theme')==='dark';var bl=document.getElementById('btn-light');var bd=document.getElementById('btn-dark');if(bl)bl.classList.toggle('active',!isDark);if(bd)bd.classList.toggle('active',isDark)}
function setTheme(t){document.documentElement.setAttribute('data-theme',t);localStorage.setItem('theme',t);_syncThemeButtons()}
function toggleDark(){setTheme(document.documentElement.getAttribute('data-theme')==='dark'?'light':'dark')}
window.addEventListener('DOMContentLoaded',function(){_syncThemeButtons();if(document.getElementById('matrix-toolbar'))matrixInit()});

var _accInitialized={};
function showPage(id){
  document.querySelectorAll('.page').forEach(function(p){p.classList.remove('active')});
  document.querySelectorAll('.nav-item').forEach(function(n){n.classList.remove('active')});
  document.querySelectorAll('.nav-group.open').forEach(function(g){g.classList.remove('open')});

  var page=document.getElementById('page-'+id);
  if(page)page.classList.add('active');
  else document.getElementById('page-home').classList.add('active');

  document.querySelectorAll('.nav-item').forEach(function(n){
    var oc=n.getAttribute('onclick')||'';
    var hasId=(oc.indexOf("'"+id+"'")!==-1||oc.indexOf("'"+id+"',")!==-1);
    if(!hasId)return;
    var g=n.closest('.nav-group');
    while(g){g.classList.add('open');g=g.parentElement.closest('.nav-group')}
    if(oc.indexOf('showPageSection')===-1)n.classList.add('active');
  });

  document.getElementById('sidebar').classList.remove('open');
  document.getElementById('overlay').classList.remove('open');
  window.scrollTo(0,0);
  renderTabBar(id,null);

  // 아코디언 초기화 (페이지 첫 표시 시 1회)
  if(!_accInitialized[id]){
    _accInitialized[id]=true;
    var target=document.getElementById('page-'+id);
    /* 아코디언 제거 — 모든 섹션 항상 표시 */
    if(id==='master-timeline')msInit();
  }
}

function showPageSection(pageId, anchor){
  showPage(pageId);
  renderTabBar(pageId,anchor);
  // showPage가 active를 초기화했으므로 해당 앵커 서브아이템 하나만 active
  document.querySelectorAll('.nav-item').forEach(function(n){
    var oc=n.getAttribute('onclick')||'';
    if(oc.indexOf('showPageSection')!==-1&&
       oc.indexOf("'"+pageId+"'")!==-1&&
       oc.indexOf("'"+anchor+"'")!==-1){
      n.classList.add('active');
    }
  });
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

function toggleOrgDrop(e,id){
  e.stopPropagation();
  var drop=document.getElementById('org-drop-'+id);
  if(!drop)return;
  var wasOpen=drop.style.display==='block';
  document.querySelectorAll('.org-drop').forEach(function(d){d.style.display='none'});
  if(!wasOpen)drop.style.display='block';
}
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
    {key:'mission',label:'미션·역할'},
    {key:'decision',label:'의사결정체계'}
  ],
  'team-support':[
    {key:'overview',label:'팀 개요·미션'},
    {key:'tasks',label:'업무 리스트'},
    {key:'sop',label:'SOP·시나리오'}
  ],
  'team-event':[
    {key:'overview',label:'팀 개요·미션'},
    {key:'tasks',label:'업무 리스트'},
    {key:'sop',label:'SOP·시나리오'}
  ],
  'team-port':[
    {key:'overview',label:'팀 개요·미션'},
    {key:'tasks',label:'업무 리스트'},
    {key:'sop',label:'SOP·시나리오'},
    {key:'sub-embark',label:'↳ 승하선팀',goto:'team-embark'},
    {key:'vip-logistics',label:'VIP·가수로지스틱'}
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
    {key:'sop',label:'SOP·시나리오'}
  ]
};
var TEAM_EXCEL_NAMES={
  'team-hq':'HQ운영본부',
  'team-support':'운영지원팀',
  'team-event':'공연행사팀',
  'team-port':'기항지운영팀',
  'team-embark':'승하선팀',
  'team-fb':'식음료파트',
  'team-it':'IT홍보팀'
};
function renderTabBar(pageId,activeAnchor){
  var wrap=document.getElementById('tab-bar');
  if(!wrap)return;
  var tabs=TAB_CONFIG[pageId];
  if(!tabs){wrap.innerHTML='';document.body.classList.remove('tabs-visible');return}
  document.body.classList.add('tabs-visible');
  var tabHtml=tabs.map(function(t){
    var cls='tab-item'+(t.key===activeAnchor?' active':'');
    if(t.goto)return '<button class="'+cls+'" onclick="showPage(\''+t.goto+'\')">'+t.label+'</button>';
    return '<button class="'+cls+'" onclick="showPageSection(\''+pageId+'\',\''+t.key+'\')">'+t.label+'</button>';
  }).join('');
  if(TEAM_EXCEL_NAMES[pageId]){
    tabHtml+='<span class="tab-spacer"></span><button class="tab-excel-btn" onclick="teamDownloadXlsx(\''+pageId+'\',\''+TEAM_EXCEL_NAMES[pageId]+'\')">📥 Excel</button>';
  }
  wrap.innerHTML=tabHtml;
}

/* ══ 핵심준비사항 진척 관리 ══ */
var MS_KEY='cruise-ms-status';
function msLoad(){try{return JSON.parse(localStorage.getItem(MS_KEY)||'{}')}catch(e){return{}}}
function msStatusChange(sel){
  var id=sel.getAttribute('data-ms-id');
  var data=msLoad();data[id]=sel.value;
  localStorage.setItem(MS_KEY,JSON.stringify(data));
  var row=sel.closest('tr');if(row)row.setAttribute('data-ms-status',sel.value);
  msUpdateProgress();
}
function msInit(){
  var data=msLoad();
  document.querySelectorAll('.ms-status-sel').forEach(function(sel){
    var id=sel.getAttribute('data-ms-id');
    if(data[id]){sel.value=data[id];var row=sel.closest('tr');if(row)row.setAttribute('data-ms-status',data[id]);}
  });
  msUpdateProgress();
}
function msUpdateProgress(){
  var done=0,prog=0,todo=0,hold=0,total=0;
  document.querySelectorAll('.ms-status-sel').forEach(function(sel){
    total++;var v=sel.value;
    if(v==='완료')done++;else if(v==='진행중')prog++;else if(v==='진행전')todo++;else hold++;
  });
  var txt=document.getElementById('ms-progress-text');
  if(txt)txt.textContent='완료 '+done+' / 진행중 '+prog+' / 진행전 '+todo+' / 보류 '+hold;
  if(total>0){
    var d=document.getElementById('ms-bar-done');var p=document.getElementById('ms-bar-prog');var h=document.getElementById('ms-bar-hold');
    if(d)d.style.width=(done/total*100)+'%';
    if(p)p.style.width=(prog/total*100)+'%';
    if(h)h.style.width=(hold/total*100)+'%';
  }
}

/* 아코디언 제거됨 — 모든 섹션 항상 펼침 */

/* ══ 팀별 Excel 다운로드 ══ */
function teamDownloadXlsx(pageId,teamName){
  if(typeof XLSX==='undefined'){
    var s=document.createElement('script');
    s.src='https://cdn.sheetjs.com/xlsx-0.20.0/package/dist/xlsx.full.min.js';
    s.onload=function(){_doTeamXlsx(pageId,teamName);};
    document.head.appendChild(s);
  }else{_doTeamXlsx(pageId,teamName);}
}
function _doTeamXlsx(pageId,teamName){
  var wb=XLSX.utils.book_new();
  var page=document.getElementById('page-'+pageId);
  if(!page)return;
  /* 섹션 → 시트명 매핑 (없는 섹션은 자동 스킵) */
  var secs=[
    {id:pageId+'-mission',   label:'미션·역할'},
    {id:pageId+'-decision',  label:'의사결정체계'},
    {id:pageId+'-overview',  label:'팀 개요'},
    {id:pageId+'-tasks',     label:'업무분장서'},
    {id:pageId+'-sop',       label:'SOP·시나리오'},
    {id:pageId+'-vip-logistics',label:'VIP·로지스틱'}
  ];
  secs.forEach(function(sec){
    var el=document.getElementById(sec.id);
    if(!el)return;
    var tbls=el.querySelectorAll('table.ops-table');
    if(tbls.length===0)return;
    var rows=[];
    tbls.forEach(function(tbl,ti){
      if(ti>0)rows.push([]);  /* 테이블 간 빈 행 */
      tbl.querySelectorAll('tr').forEach(function(tr){
        var r=[];
        tr.querySelectorAll('th,td').forEach(function(cell){
          r.push(cell.textContent.replace(/\s+/g,' ').trim());
        });
        if(r.some(function(c){return c.length>0;}))rows.push(r);
      });
    });
    if(rows.length===0)return;
    var ws=XLSX.utils.aoa_to_sheet(rows);
    /* 컬럼 너비 자동 조정 */
    if(rows[0]){ws['!cols']=rows[0].map(function(v,ci){
      var max=8;rows.forEach(function(r){if(r[ci]&&r[ci].length>max)max=Math.min(r[ci].length,60);});return{wch:max};
    });}
    XLSX.utils.book_append_sheet(wb,ws,sec.label);
  });
  if(wb.SheetNames.length===0){alert('내보낼 데이터가 없습니다.');return;}
  var date=new Date().toISOString().slice(0,10);
  XLSX.writeFile(wb,'코스타세레나_'+teamName+'_'+date+'.xlsx');
}
"""

# ── 최종 HTML 조립 ────────────────────────────────────────────────────
def build_html(manifest, modules):
    nav = build_nav(manifest)
    dashboard = build_dashboard(manifest, modules)
    insights = load_insights()
    team_pages = build_team_pages(manifest, modules, insights)
    timeline = build_master_timeline(manifest, modules)
    common_extra = build_common_extra_page(manifest, modules)
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
    <span class="badge-ver" title="{UPDATE_DATE} 업데이트">v4.2</span>
    <div class="theme-toggle">
      <button id="btn-light" onclick="setTheme('light')">☀️ 라이트</button>
      <button id="btn-dark" onclick="setTheme('dark')">🌙 다크</button>
    </div>
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
  </div>

{timeline}
{common_extra}
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
    import shutil
    test_key = sys.argv[1] if len(sys.argv) > 1 else None
    print('=== 크루즈 운영 데스크 빌더 v4.2 ===')
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
    # archive 저장
    if not test_key:
        archive_path = os.path.join(ARCHIVE_DIR, 'v4.2_2026-04-12.html')
        shutil.copy2(OUT, archive_path)
        print(f'  → archive/{os.path.basename(archive_path)} 저장 완료')

if __name__ == '__main__':
    main()
