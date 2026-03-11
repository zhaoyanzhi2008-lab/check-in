"""
ICode 签到管理系统 v3.0
Design: Code Planet — deep space aesthetic, electric blue + neon orange
"""
from flask import Flask, render_template, request, jsonify, session, redirect, send_file
import sqlite3, hashlib, json, io, os
from datetime import datetime
from functools import wraps
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

app = Flask(__name__)
app.secret_key = 'icode-planet-v3-2025-secret'
DB = 'icode.db'

# ─── helpers ───────────────────────────────────────────────────────
def db():
    c = sqlite3.connect(DB)
    c.row_factory = sqlite3.Row
    return c

def sha(s):
    return hashlib.sha256(s.encode()).hexdigest()

def now_str():
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')

def init_db():
    conn = db(); c = conn.cursor()
    c.executescript('''
    CREATE TABLE IF NOT EXISTS admins (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        phone TEXT UNIQUE NOT NULL,
        password TEXT NOT NULL,
        is_main INTEGER DEFAULT 0,
        permissions TEXT DEFAULT '{}',
        created_at TEXT DEFAULT (datetime('now','localtime'))
    );
    CREATE TABLE IF NOT EXISTS competitions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        location TEXT DEFAULT '',
        start_time TEXT DEFAULT '',
        end_time TEXT DEFAULT '',
        description TEXT DEFAULT '',
        banner_text TEXT DEFAULT '欢迎参加ICode比赛',
        banner_color TEXT DEFAULT '#1a6fa8',
        banner_accent TEXT DEFAULT '#0099cc',
        groups TEXT DEFAULT '[]',
        display_fields TEXT DEFAULT '["name","school","group_name","session","seat_no","shirt_size"]',
        query_field TEXT DEFAULT 'player_no,account',
        query_hint TEXT DEFAULT '请输入报名编号或选手账号',
        is_active INTEGER DEFAULT 1,
        created_by INTEGER,
        created_at TEXT DEFAULT (datetime('now','localtime'))
    );
    CREATE TABLE IF NOT EXISTS players (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        competition_id INTEGER NOT NULL,
        player_no TEXT DEFAULT '',
        account TEXT DEFAULT '',
        name TEXT NOT NULL,
        school TEXT DEFAULT '',
        grade TEXT DEFAULT '',
        group_name TEXT DEFAULT '',
        comp_date TEXT DEFAULT '',
        session TEXT DEFAULT '',
        seat_no TEXT DEFAULT '',
        shirt_size TEXT DEFAULT '',
        checked_in INTEGER DEFAULT 0,
        checkin_time TEXT DEFAULT '',
        remark TEXT DEFAULT '',
        FOREIGN KEY (competition_id) REFERENCES competitions(id)
    );
    CREATE TABLE IF NOT EXISTS checkin_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        player_id INTEGER,
        competition_id INTEGER,
        operator TEXT DEFAULT '选手自助',
        checkin_time TEXT DEFAULT (datetime('now','localtime'))
    );
    ''')
    # default main admin
    if not c.execute("SELECT id FROM admins WHERE is_main=1").fetchone():
        c.execute("INSERT INTO admins(name,phone,password,is_main,permissions) VALUES(?,?,?,1,?)",
                  ('主管理员', 'admin', sha('admin123'), '{"all":true}'))
    # demo competition
    if not c.execute("SELECT id FROM competitions LIMIT 1").fetchone():
        c.execute("""INSERT INTO competitions(name,location,start_time,end_time,description,
                     banner_text,banner_color,banner_accent,groups,display_fields,query_field,query_hint,created_by)
                     VALUES(?,?,?,?,?,?,?,?,?,?,?,?,1)""",
                  ('2025 ICode全国青少年编程大赛', '北京·国家会议中心A厅',
                   '2025-06-15 08:30', '2025-06-16 17:00',
                   '📢 请凭报名编号完成现场签到。赛场请勿携带电子设备，请提前15分钟入场就座。祝各位选手发挥出色，取得佳绩！',
                   '欢迎参加 ICode 全国青少年编程大赛',
                   '#1a6fa8', '#0099cc',
                   '["初级组","中级组","高级组"]',
                   '["name","school","group_name","session","seat_no","shirt_size"]',
                   'player_no,account', '请输入报名编号或选手账号'))
        cid = c.lastrowid
        demo = [
            ('IC2025001','user001','张小明','北京海淀实验小学','六年级','初级组','2025-06-15','上午场','A-01','M',0),
            ('IC2025002','user002','李思远','清华大学附属小学','五年级','初级组','2025-06-15','上午场','A-02','S',0),
            ('IC2025003','user003','王浩然','北京第二实验小学','初一','中级组','2025-06-15','下午场','B-01','L',1),
            ('IC2025004','user004','赵雨欣','人民大学附属中学','初二','中级组','2025-06-15','下午场','B-02','M',0),
            ('IC2025005','user005','孙悦琪','北京师范大学附属中学','高一','高级组','2025-06-16','上午场','C-01','XL',0),
            ('IC2025006','user006','陈思涵','育才学校','高二','高级组','2025-06-16','上午场','C-02','S',0),
            ('IC2025007','user007','林志远','北大附中','高二','高级组','2025-06-16','下午场','C-03','M',0),
        ]
        nw = now_str()
        for d in demo:
            ct = nw if d[10] else ''
            c.execute("""INSERT INTO players(competition_id,player_no,account,name,school,grade,
                         group_name,comp_date,session,seat_no,shirt_size,checked_in,checkin_time)
                         VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                      (cid, *d[:10], d[10], ct))
    conn.commit(); conn.close()

def admin_required(f):
    @wraps(f)
    def w(*a, **k):
        if 'admin_id' not in session:
            return jsonify({'error': '未授权'}), 401
        return f(*a, **k)
    return w

def get_me():
    if 'admin_id' not in session: return None
    conn = db()
    a = conn.execute("SELECT * FROM admins WHERE id=?", (session['admin_id'],)).fetchone()
    conn.close()
    return a

def can(admin, perm):
    if not admin: return False
    if admin['is_main']: return True
    p = json.loads(admin['permissions'] or '{}')
    return p.get('all') or p.get(perm, False)

def admin_owns_comp(admin, comp_id):
    if admin['is_main']: return True
    conn = db()
    comp = conn.execute("SELECT created_by FROM competitions WHERE id=?", (comp_id,)).fetchone()
    conn.close()
    if not comp: return False
    return comp['created_by'] == admin['id']

# ═══════════════════════════════════════════════════════════════════
# PUBLIC PLAYER ROUTES
# ═══════════════════════════════════════════════════════════════════

@app.route('/')
def player_root():
    return render_template('player.html')

@app.route('/c/<int:cid>')
def player_comp(cid):
    return render_template('player.html', preset_cid=cid)

@app.route('/api/pub/competition/<int:cid>')
def pub_competition(cid):
    conn = db()
    c = conn.execute(
        "SELECT id,name,description,banner_text,banner_color,banner_accent,groups,"
        "display_fields,query_field,query_hint,location,start_time,end_time "
        "FROM competitions WHERE id=? AND is_active=1", (cid,)).fetchone()
    conn.close()
    if not c: return jsonify({'error': '该赛事不存在或已下线'}), 404
    return jsonify(dict(c))

@app.route('/api/pub/competitions')
def pub_competitions():
    conn = db()
    rows = conn.execute(
        "SELECT id,name,description,banner_text,banner_color,banner_accent,location,start_time "
        "FROM competitions WHERE is_active=1 ORDER BY id DESC").fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])

@app.route('/api/pub/query', methods=['POST'])
def pub_query():
    d = request.json or {}
    cid = d.get('competition_id')
    q = d.get('query', '').strip()
    if not q: return jsonify({'error': '请输入查询内容'}), 400
    if not cid: return jsonify({'error': '比赛参数缺失'}), 400
    conn = db()
    comp = conn.execute("SELECT display_fields,query_field FROM competitions WHERE id=?", (cid,)).fetchone()
    if not comp: conn.close(); return jsonify({'error': '赛事不存在'}), 404
    display_fields = json.loads(comp['display_fields'])
    query_fields = [f.strip() for f in (comp['query_field'] or 'player_no,account').split(',')]
    tokens = [t for t in q.split() if t]
    results = []
    seen = set()
    for t in tokens:
        conds = ' OR '.join([f"{f}=?" for f in query_fields])
        row = conn.execute(
            f"SELECT * FROM players WHERE competition_id=? AND ({conds})",
            (cid, *([t] * len(query_fields)))).fetchone()
        if row and row['id'] not in seen:
            results.append(dict(row))
            seen.add(row['id'])
    conn.close()
    if not results: return jsonify({'error': '未找到选手，请检查编号是否正确'}), 404
    return jsonify({'players': results, 'display_fields': display_fields})

@app.route('/api/pub/checkin', methods=['POST'])
def pub_checkin():
    d = request.json or {}
    ids = d.get('player_ids', [])
    if not ids: return jsonify({'error': '请选择要签到的选手'}), 400
    nw = now_str()
    conn = db()
    names = []
    for pid in ids:
        p = conn.execute("SELECT * FROM players WHERE id=?", (pid,)).fetchone()
        if p and not p['checked_in']:
            conn.execute("UPDATE players SET checked_in=1,checkin_time=? WHERE id=?", (nw, pid))
            conn.execute("INSERT INTO checkin_logs(player_id,competition_id,checkin_time) VALUES(?,?,?)",
                         (pid, p['competition_id'], nw))
            names.append(p['name'])
    conn.commit(); conn.close()
    return jsonify({'success': True, 'names': names})

# ═══════════════════════════════════════════════════════════════════
# ADMIN AUTH
# ═══════════════════════════════════════════════════════════════════

@app.route('/admin')
def admin_login_page():
    if 'admin_id' in session: return redirect('/admin/dashboard')
    return render_template('admin_login.html')

@app.route('/admin/dashboard')
def admin_dash():
    if 'admin_id' not in session: return redirect('/admin')
    return render_template('admin_dashboard.html')

@app.route('/api/admin/login', methods=['POST'])
def admin_login():
    d = request.json or {}
    conn = db()
    a = conn.execute("SELECT * FROM admins WHERE phone=?", (d.get('phone', ''),)).fetchone()
    conn.close()
    if not a or a['password'] != sha(d.get('password', '')):
        return jsonify({'error': '账号或密码错误'}), 401
    session['admin_id'] = a['id']
    session['admin_name'] = a['name']
    session['is_main'] = a['is_main']
    return jsonify({'success': True, 'name': a['name'], 'is_main': a['is_main']})

@app.route('/api/admin/logout', methods=['POST'])
def admin_logout():
    session.clear()
    return jsonify({'success': True})

@app.route('/api/admin/me')
@admin_required
def admin_me():
    a = get_me()
    return jsonify({'id': a['id'], 'name': a['name'], 'phone': a['phone'],
                    'is_main': a['is_main'],
                    'permissions': json.loads(a['permissions'] or '{}')})

@app.route('/api/admin/change-password', methods=['POST'])
@admin_required
def change_pwd():
    d = request.json or {}
    conn = db()
    a = conn.execute("SELECT * FROM admins WHERE id=?", (session['admin_id'],)).fetchone()
    if a['password'] != sha(d.get('old_password', '')):
        conn.close(); return jsonify({'error': '旧密码错误'}), 400
    conn.execute("UPDATE admins SET password=? WHERE id=?",
                 (sha(d['new_password']), session['admin_id']))
    conn.commit(); conn.close()
    return jsonify({'success': True})

# ═══════════════════════════════════════════════════════════════════
# COMPETITIONS
# ═══════════════════════════════════════════════════════════════════

@app.route('/api/competitions', methods=['GET'])
@admin_required
def list_competitions():
    a = get_me()
    conn = db()
    if a['is_main']:
        rows = conn.execute("""
            SELECT c.*,(SELECT COUNT(*) FROM players WHERE competition_id=c.id) pc,
            adm.name creator_name FROM competitions c
            LEFT JOIN admins adm ON c.created_by=adm.id ORDER BY c.id DESC""").fetchall()
    else:
        rows = conn.execute("""
            SELECT c.*,(SELECT COUNT(*) FROM players WHERE competition_id=c.id) pc,
            adm.name creator_name FROM competitions c
            LEFT JOIN admins adm ON c.created_by=adm.id
            WHERE c.created_by=? ORDER BY c.id DESC""", (a['id'],)).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])

@app.route('/api/competitions', methods=['POST'])
@admin_required
def create_competition():
    a = get_me()
    if not can(a, 'add_competition'): return jsonify({'error': '无权限'}), 403
    d = request.json or {}
    if not d.get('name'): return jsonify({'error': '请填写赛事名称'}), 400
    conn = db()
    conn.execute("""INSERT INTO competitions(name,location,start_time,end_time,description,
                    banner_text,banner_color,banner_accent,groups,display_fields,
                    query_field,query_hint,is_active,created_by) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                 (d['name'], d.get('location', ''), d.get('start_time', ''), d.get('end_time', ''),
                  d.get('description', ''),
                  d.get('banner_text', '欢迎参加ICode比赛'),
                  d.get('banner_color', '#1a6fa8'),
                  d.get('banner_accent', '#0099cc'),
                  json.dumps(d.get('groups', []), ensure_ascii=False),
                  json.dumps(d.get('display_fields', ['name','school','group_name','session','seat_no','shirt_size']), ensure_ascii=False),
                  d.get('query_field', 'player_no,account'),
                  d.get('query_hint', '请输入报名编号或选手账号'),
                  d.get('is_active', 1), session['admin_id']))
    conn.commit(); conn.close()
    return jsonify({'success': True})

@app.route('/api/competitions/<int:cid>', methods=['GET'])
@admin_required
def get_competition(cid):
    conn = db()
    r = conn.execute("SELECT * FROM competitions WHERE id=?", (cid,)).fetchone()
    conn.close()
    if not r: return jsonify({'error': '不存在'}), 404
    return jsonify(dict(r))

@app.route('/api/competitions/<int:cid>', methods=['PUT'])
@admin_required
def update_competition(cid):
    a = get_me()
    if not admin_owns_comp(a, cid): return jsonify({'error': '无权限'}), 403
    d = request.json or {}
    conn = db()
    conn.execute("""UPDATE competitions SET name=?,location=?,start_time=?,end_time=?,
                    description=?,banner_text=?,banner_color=?,banner_accent=?,groups=?,
                    display_fields=?,query_field=?,query_hint=?,is_active=? WHERE id=?""",
                 (d.get('name'), d.get('location', ''), d.get('start_time', ''), d.get('end_time', ''),
                  d.get('description', ''), d.get('banner_text', ''),
                  d.get('banner_color', '#1a6fa8'), d.get('banner_accent', '#0099cc'),
                  json.dumps(d.get('groups', []), ensure_ascii=False),
                  json.dumps(d.get('display_fields', []), ensure_ascii=False),
                  d.get('query_field', 'player_no,account'),
                  d.get('query_hint', ''), d.get('is_active', 1), cid))
    conn.commit(); conn.close()
    return jsonify({'success': True})

@app.route('/api/competitions/<int:cid>', methods=['DELETE'])
@admin_required
def delete_competition(cid):
    a = get_me()
    if not admin_owns_comp(a, cid): return jsonify({'error': '无权限'}), 403
    conn = db()
    conn.execute("DELETE FROM checkin_logs WHERE competition_id=?", (cid,))
    conn.execute("DELETE FROM players WHERE competition_id=?", (cid,))
    conn.execute("DELETE FROM competitions WHERE id=?", (cid,))
    conn.commit(); conn.close()
    return jsonify({'success': True})

# ═══════════════════════════════════════════════════════════════════
# PLAYERS
# ═══════════════════════════════════════════════════════════════════

@app.route('/api/players', methods=['GET'])
@admin_required
def list_players():
    cid = request.args.get('competition_id')
    if not cid: return jsonify([])
    a = get_me()
    if not admin_owns_comp(a, int(cid)): return jsonify({'error': '无权限'}), 403
    q = "SELECT * FROM players WHERE competition_id=?"
    params = [cid]
    for f, col in [('group', 'group_name'), ('date', 'comp_date'),
                   ('session', 'session'), ('shirt', 'shirt_size'),
                   ('school', 'school'), ('grade', 'grade')]:
        v = request.args.get(f, '')
        if v: q += f" AND {col}=?"; params.append(v)
    checked = request.args.get('checked', '')
    if checked != '': q += " AND checked_in=?"; params.append(int(checked))
    search = request.args.get('search', '').strip()
    if search:
        q += " AND (name LIKE ? OR player_no LIKE ? OR account LIKE ? OR school LIKE ?)"
        params.extend([f'%{search}%'] * 4)
    q += " ORDER BY id"
    conn = db()
    rows = conn.execute(q, params).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])

@app.route('/api/players', methods=['POST'])
@admin_required
def create_player():
    d = request.json or {}
    if not d.get('name'): return jsonify({'error': '姓名必填'}), 400
    conn = db()
    conn.execute("""INSERT INTO players(competition_id,player_no,account,name,school,grade,
                    group_name,comp_date,session,seat_no,shirt_size,remark) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
                 (d['competition_id'], d.get('player_no',''), d.get('account',''), d['name'],
                  d.get('school',''), d.get('grade',''), d.get('group_name',''),
                  d.get('comp_date',''), d.get('session',''), d.get('seat_no',''),
                  d.get('shirt_size',''), d.get('remark','')))
    conn.commit(); conn.close()
    return jsonify({'success': True})

@app.route('/api/players/<int:pid>', methods=['PUT'])
@admin_required
def update_player(pid):
    d = request.json or {}
    conn = db()
    conn.execute("""UPDATE players SET player_no=?,account=?,name=?,school=?,grade=?,
                    group_name=?,comp_date=?,session=?,seat_no=?,shirt_size=?,
                    checked_in=?,checkin_time=?,remark=? WHERE id=?""",
                 (d.get('player_no',''), d.get('account',''), d.get('name',''),
                  d.get('school',''), d.get('grade',''), d.get('group_name',''),
                  d.get('comp_date',''), d.get('session',''), d.get('seat_no',''),
                  d.get('shirt_size',''), d.get('checked_in',0),
                  d.get('checkin_time',''), d.get('remark',''), pid))
    conn.commit(); conn.close()
    return jsonify({'success': True})

@app.route('/api/players/<int:pid>', methods=['DELETE'])
@admin_required
def delete_player(pid):
    conn = db()
    conn.execute("DELETE FROM players WHERE id=?", (pid,))
    conn.commit(); conn.close()
    return jsonify({'success': True})

@app.route('/api/players/import', methods=['POST'])
@admin_required
def import_players():
    a = get_me()
    if not can(a, 'import_players'): return jsonify({'error': '无权限'}), 403
    cid = request.form.get('competition_id')
    upload = request.files.get('file')
    if not cid or not upload: return jsonify({'error': '参数缺失'}), 400

    conn = db()
    # 取赛事合法组别列表
    comp_row = conn.execute("SELECT groups FROM competitions WHERE id=?", (cid,)).fetchone()
    if not comp_row: conn.close(); return jsonify({'error': '赛事不存在'}), 404
    valid_groups = json.loads(comp_row['groups'] or '[]')   # e.g. ["初级组","中级组"]
    # 已存在的 player_no / account（非空）
    exist_nos  = set(r[0] for r in conn.execute(
        "SELECT player_no FROM players WHERE competition_id=? AND player_no!=''", (cid,)).fetchall())
    exist_accs = set(r[0] for r in conn.execute(
        "SELECT account  FROM players WHERE competition_id=? AND account!=''",   (cid,)).fetchall())

    wb = openpyxl.load_workbook(upload, data_only=True)
    ws = wb.active
    hdrs = [str(c.value).strip().rstrip('*').strip() if c.value else '' for c in ws[1]]
    fm = {'报名编号':'player_no','账号':'account','姓名':'name','学校':'school',
          '年级':'grade','组别':'group_name','比赛日期':'comp_date','场次':'session',
          '座位号':'seat_no','衣服尺码':'shirt_size','备注':'remark'}

    errors = []   # 收集所有错误行
    to_insert = []
    # 文件内去重集合（防止同一文件里重复）
    file_nos  = set()
    file_accs = set()

    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not any(row): continue
        pdata = {}
        for i, h in enumerate(hdrs):
            if h in fm and i < len(row):
                pdata[fm[h]] = str(row[i]).strip() if row[i] is not None else ''
        if not pdata.get('name'): continue

        row_errors = []
        pno  = pdata.get('player_no', '')
        pacc = pdata.get('account', '')
        grp  = pdata.get('group_name', '')

        # ── 功能1：重复检测 ──────────────────────────────────
        if pno:
            if pno in exist_nos or pno in file_nos:
                row_errors.append(f'报名编号"{pno}"重复')
            else:
                file_nos.add(pno)
        if pacc:
            if pacc in exist_accs or pacc in file_accs:
                row_errors.append(f'报名账号"{pacc}"重复')
            else:
                file_accs.add(pacc)

        # ── 功能2：组别校验 ──────────────────────────────────
        if valid_groups and grp and grp not in valid_groups:
            row_errors.append(f'组别"{grp}"不在赛事组别{valid_groups}中，请修改后重新导入')

        if row_errors:
            errors.append({'row': row_idx, 'name': pdata.get('name',''), 'errors': row_errors})
        else:
            to_insert.append(pdata)

    # 有任何错误 → 全部不导入，返回详细错误列表
    if errors:
        conn.close()
        msgs = []
        for e in errors:
            msgs.append(f"第{e['row']}行「{e['name']}」：{'；'.join(e['errors'])}")
        return jsonify({
            'success': False,
            'error': '导入失败，请修正以下问题后重新导入：\n' + '\n'.join(msgs),
            'error_rows': errors
        }), 400

    # 全部通过 → 批量插入
    cnt = 0
    for pdata in to_insert:
        conn.execute("""INSERT INTO players(competition_id,player_no,account,name,school,grade,
                        group_name,comp_date,session,seat_no,shirt_size,remark) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
                     (cid, pdata.get('player_no',''), pdata.get('account',''), pdata.get('name',''),
                      pdata.get('school',''), pdata.get('grade',''), pdata.get('group_name',''),
                      pdata.get('comp_date',''), pdata.get('session',''), pdata.get('seat_no',''),
                      pdata.get('shirt_size',''), pdata.get('remark','')))
        cnt += 1
    conn.commit(); conn.close()
    return jsonify({'success': True, 'count': cnt})

# ── 功能3：批量删除 ──────────────────────────────────────────────
@app.route('/api/players/batch_delete', methods=['POST'])
@admin_required
def batch_delete_players():
    d = request.json or {}
    ids = d.get('ids', [])
    if not ids: return jsonify({'error': '未选择选手'}), 400
    a = get_me()
    conn = db()
    # 鉴权：所有id都属于该管理员有权限的赛事
    for pid in ids:
        p = conn.execute("SELECT competition_id FROM players WHERE id=?", (pid,)).fetchone()
        if p and not admin_owns_comp(a, p['competition_id']):
            conn.close(); return jsonify({'error': '无权限'}), 403
    placeholders = ','.join('?' * len(ids))
    conn.execute(f"DELETE FROM players WHERE id IN ({placeholders})", ids)
    conn.commit(); conn.close()
    return jsonify({'success': True, 'count': len(ids)})

# ── 功能3：批量修改 ──────────────────────────────────────────────
@app.route('/api/players/batch_update', methods=['POST'])
@admin_required
def batch_update_players():
    d = request.json or {}
    ids    = d.get('ids', [])
    fields = d.get('fields', {})   # e.g. {"group_name":"高级组","session":"下午场"}
    if not ids:   return jsonify({'error': '未选择选手'}), 400
    if not fields: return jsonify({'error': '未指定修改字段'}), 400
    allowed = {'player_no','account','name','school','grade','group_name',
               'comp_date','session','seat_no','shirt_size','remark','checked_in','checkin_time'}
    safe = {k: v for k, v in fields.items() if k in allowed}
    if not safe: return jsonify({'error': '没有合法字段'}), 400
    a = get_me()
    conn = db()
    for pid in ids:
        p = conn.execute("SELECT competition_id FROM players WHERE id=?", (pid,)).fetchone()
        if p and not admin_owns_comp(a, p['competition_id']):
            conn.close(); return jsonify({'error': '无权限'}), 403
    set_clause = ', '.join(f"{k}=?" for k in safe)
    vals = list(safe.values())
    placeholders = ','.join('?' * len(ids))
    conn.execute(f"UPDATE players SET {set_clause} WHERE id IN ({placeholders})",
                 vals + ids)
    conn.commit(); conn.close()
    return jsonify({'success': True, 'count': len(ids)})

# ── 功能3：赛事批量删除 ──────────────────────────────────────────
@app.route('/api/competitions/batch_delete', methods=['POST'])
@admin_required
def batch_delete_competitions():
    a = get_me()
    d = request.json or {}
    ids = d.get('ids', [])
    if not ids: return jsonify({'error': '未选择赛事'}), 400
    conn = db()
    for cid in ids:
        if not admin_owns_comp(a, cid):
            conn.close(); return jsonify({'error': '无权限'}), 403
    placeholders = ','.join('?' * len(ids))
    conn.execute(f"DELETE FROM checkin_logs WHERE competition_id IN ({placeholders})", ids)
    conn.execute(f"DELETE FROM players WHERE competition_id IN ({placeholders})", ids)
    conn.execute(f"DELETE FROM competitions WHERE id IN ({placeholders})", ids)
    conn.commit(); conn.close()
    return jsonify({'success': True})

# ── 功能3：带筛选的选手导出 ──────────────────────────────────────
@app.route('/api/players/export/<int:cid>')
@admin_required
def export_players(cid):
    conn = db()
    # 支持筛选参数导出（功能3）
    q = "SELECT * FROM players WHERE competition_id=?"
    params = [cid]
    for f, col in [('group', 'group_name'), ('date', 'comp_date'),
                   ('session', 'session'), ('shirt', 'shirt_size'), ('school', 'school'), ('grade', 'grade')]:
        v = request.args.get(f, '')
        if v: q += f" AND {col}=?"; params.append(v)
    checked = request.args.get('checked', '')
    if checked != '': q += " AND checked_in=?"; params.append(int(checked))
    search = request.args.get('search', '').strip()
    if search:
        q += " AND (name LIKE ? OR player_no LIKE ? OR account LIKE ? OR school LIKE ?)"
        params.extend([f'%{search}%'] * 4)
    q += " ORDER BY id"
    players = conn.execute(q, params).fetchall()
    comp = conn.execute("SELECT name FROM competitions WHERE id=?", (cid,)).fetchone()
    conn.close()
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = '选手信息'
    hfill = PatternFill("solid", fgColor="050d1f")
    afill = PatternFill("solid", fgColor="E8F8FF")
    thin = Border(*[Side(style='thin', color='CBD5E1')] * 4)
    hdrs = ['ID','报名编号','账号','姓名','学校','年级','组别','比赛日期','场次','座位号','衣服尺码','是否签到','签到时间','备注']
    ws.append(hdrs)
    for cell in ws[1]:
        cell.fill = hfill
        cell.font = Font(color="00d4ff", bold=True, name='微软雅黑', size=11)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin
    ws.row_dimensions[1].height = 24
    for ri, p in enumerate(players):
        row = [p['id'], p['player_no'], p['account'], p['name'], p['school'],
               p['grade'], p['group_name'], p['comp_date'], p['session'],
               p['seat_no'], p['shirt_size'],
               '✅ 已签到' if p['checked_in'] else '⏳ 未签到',
               p['checkin_time'] or '', p['remark'] or '']
        ws.append(row)
        fill = afill if ri % 2 == 0 else PatternFill("solid", fgColor="FFFFFF")
        for cell in ws[ri + 2]:
            cell.fill = fill
            cell.font = Font(name='微软雅黑', size=10)
            cell.alignment = Alignment(horizontal='left', vertical='center')
            cell.border = thin
        ws.row_dimensions[ri + 2].height = 20
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = max(
            len(str(c.value or '')) * 2 + 4 for c in col)
    ws.freeze_panes = 'A2'
    out = io.BytesIO(); wb.save(out); out.seek(0)
    name = f"{comp['name']}_选手信息.xlsx" if comp else "选手信息.xlsx"
    return send_file(out, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name=name)

@app.route('/api/players/template')
@admin_required
def player_template():
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = '导入模板'
    hfill = PatternFill("solid", fgColor="E8F8FF")
    req_fill = PatternFill("solid", fgColor="FFE8E0")
    hdrs = ['报名编号','账号','姓名*','学校','年级','组别','比赛日期','场次','座位号','衣服尺码','备注']
    ws.append(hdrs)
    for i, cell in enumerate(ws[1]):
        cell.fill = req_fill if '*' in hdrs[i] else hfill
        cell.font = Font(bold=True, name='微软雅黑')
        cell.alignment = Alignment(horizontal='center')
    ws.append(['IC001','user001','张三','北京实验小学','六年级','初级组','2025-06-15','上午场','A-001','M',''])
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 16
    out = io.BytesIO(); wb.save(out); out.seek(0)
    return send_file(out, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name='选手导入模板.xlsx')

# ═══════════════════════════════════════════════════════════════════
# STATISTICS
# ═══════════════════════════════════════════════════════════════════

# ── 功能4：赛事地点列表（供统计筛选用）────────────────────────
@app.route('/api/competitions/locations')
@admin_required
def competition_locations():
    a = get_me()
    conn = db()
    if a['is_main']:
        rows = conn.execute("SELECT DISTINCT location FROM competitions WHERE location!='' ORDER BY location").fetchall()
    else:
        rows = conn.execute("SELECT DISTINCT location FROM competitions WHERE location!='' AND created_by=? ORDER BY location",
                            (a['id'],)).fetchall()
    conn.close()
    return jsonify([r['location'] for r in rows])

@app.route('/api/stats/<int:cid>')
@admin_required
def stats(cid):
    a = get_me()
    if not admin_owns_comp(a, cid): return jsonify({'error': '无权限'}), 403
    # 功能4：支持按地点筛选（location 筛选的是赛事列表层，这里支持跨赛事按地点汇总）
    location = request.args.get('location', '').strip()
    conn = db()

    if location:
        # 找出该地点下当前管理员有权限的所有赛事id
        if a['is_main']:
            cid_rows = conn.execute("SELECT id FROM competitions WHERE location=?", (location,)).fetchall()
        else:
            cid_rows = conn.execute("SELECT id FROM competitions WHERE location=? AND created_by=?",
                                    (location, a['id'])).fetchall()
        cids = [r['id'] for r in cid_rows]
        if not cids:
            conn.close()
            return jsonify({'total':0,'checked':0,'unchecked':0,'by_session':[],'by_group':[],
                            'by_shirt':[],'by_date':[],'recent':[],'location':location,'comp_names':[]})
        ph = ','.join('?'*len(cids))
        where = f"competition_id IN ({ph})"
        p = cids
        comp_names = [r['name'] for r in conn.execute(f"SELECT name FROM competitions WHERE id IN ({ph})", cids).fetchall()]
    else:
        where = "competition_id=?"
        p = [cid]
        comp_names = []

    total = conn.execute(f"SELECT COUNT(*) FROM players WHERE {where}", p).fetchone()[0]
    chk   = conn.execute(f"SELECT COUNT(*) FROM players WHERE {where} AND checked_in=1", p).fetchone()[0]
    by_session = conn.execute(f"""SELECT comp_date,session,COUNT(*) total,SUM(checked_in) checked
        FROM players WHERE {where} GROUP BY comp_date,session ORDER BY comp_date,session""", p).fetchall()
    by_group = conn.execute(f"""SELECT group_name,COUNT(*) total,SUM(checked_in) checked
        FROM players WHERE {where} GROUP BY group_name ORDER BY total DESC""", p).fetchall()
    by_shirt = conn.execute(f"""SELECT shirt_size,COUNT(*) total FROM players
        WHERE {where} GROUP BY shirt_size ORDER BY total DESC""", p).fetchall()
    by_date  = conn.execute(f"""SELECT comp_date,COUNT(*) total,SUM(checked_in) checked
        FROM players WHERE {where} GROUP BY comp_date ORDER BY comp_date""", p).fetchall()
    recent   = conn.execute(f"""SELECT pl.name,pl.group_name,pl.session,l.checkin_time
        FROM checkin_logs l JOIN players pl ON l.player_id=pl.id
        WHERE l.{where} ORDER BY l.checkin_time DESC LIMIT 10""", p).fetchall()
    conn.close()
    return jsonify({
        'total': total, 'checked': chk, 'unchecked': total - chk,
        'by_session': [dict(r) for r in by_session],
        'by_group':   [dict(r) for r in by_group],
        'by_shirt':   [dict(r) for r in by_shirt],
        'by_date':    [dict(r) for r in by_date],
        'recent':     [dict(r) for r in recent],
        'location':   location,
        'comp_names': comp_names,
    })

@app.route('/api/stats/export/<int:cid>')
@admin_required
def export_stats(cid):
    a = get_me()
    if not admin_owns_comp(a, cid): return jsonify({'error': '无权限'}), 403
    conn = db()
    comp = conn.execute("SELECT name FROM competitions WHERE id=?", (cid,)).fetchone()
    total = conn.execute("SELECT COUNT(*) FROM players WHERE competition_id=?", (cid,)).fetchone()[0]
    chk = conn.execute("SELECT COUNT(*) FROM players WHERE competition_id=? AND checked_in=1", (cid,)).fetchone()[0]
    by_session = conn.execute("""SELECT comp_date,session,COUNT(*) total,SUM(checked_in) checked
        FROM players WHERE competition_id=? GROUP BY comp_date,session""", (cid,)).fetchall()
    by_group = conn.execute("""SELECT group_name,COUNT(*) total,SUM(checked_in) checked
        FROM players WHERE competition_id=? GROUP BY group_name""", (cid,)).fetchall()
    by_shirt = conn.execute("""SELECT shirt_size,COUNT(*) total FROM players WHERE competition_id=?
        GROUP BY shirt_size""", (cid,)).fetchall()
    conn.close()
    wb = openpyxl.Workbook()
    hfill = PatternFill("solid", fgColor="050d1f")
    def add_sheet(title, headers, rows, ws=None):
        if ws is None: ws = wb.create_sheet(title)
        ws.append(headers)
        for cell in ws[1]:
            cell.fill = hfill
            cell.font = Font(color="00d4ff", bold=True, name='微软雅黑')
            cell.alignment = Alignment(horizontal='center')
        for row in rows: ws.append(row)
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 18
        return ws
    ws0 = wb.active; ws0.title = '总览'
    add_sheet('总览', ['指标','数值'],
              [['报名总人数',total],['已签到',chk],['未签到',total-chk],
               ['签到率',f"{round(chk/total*100,1) if total else 0}%"]], ws0)
    add_sheet('按场次', ['日期','场次','总人数','已签到','未签到'],
              [[r['comp_date'],r['session'],r['total'],r['checked'],r['total']-r['checked']]
               for r in by_session])
    add_sheet('按组别', ['组别','总人数','已签到','未签到'],
              [[r['group_name'],r['total'],r['checked'],r['total']-r['checked']]
               for r in by_group])
    add_sheet('衣服尺码', ['尺码','人数'],
              [[r['shirt_size'],r['total']] for r in by_shirt])
    out = io.BytesIO(); wb.save(out); out.seek(0)
    name = f"{comp['name']}_统计数据.xlsx" if comp else "统计数据.xlsx"
    return send_file(out, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name=name)

# ═══════════════════════════════════════════════════════════════════
# ADMINS
# ═══════════════════════════════════════════════════════════════════

@app.route('/api/admins', methods=['GET'])
@admin_required
def list_admins():
    a = get_me()
    if not a['is_main']: return jsonify({'error': '无权限'}), 403
    conn = db()
    rows = conn.execute("SELECT id,name,phone,is_main,permissions,created_at FROM admins ORDER BY id").fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])

@app.route('/api/admins', methods=['POST'])
@admin_required
def create_admin():
    a = get_me()
    if not a['is_main']: return jsonify({'error': '无权限'}), 403
    d = request.json or {}
    conn = db()
    try:
        conn.execute("INSERT INTO admins(name,phone,password,permissions) VALUES(?,?,?,?)",
                     (d['name'], d['phone'], sha(d['password']),
                      json.dumps(d.get('permissions', {}), ensure_ascii=False)))
        conn.commit()
    except Exception:
        conn.close(); return jsonify({'error': '手机号已存在'}), 400
    conn.close()
    return jsonify({'success': True})

@app.route('/api/admins/<int:aid>', methods=['PUT'])
@admin_required
def update_admin(aid):
    a = get_me(); d = request.json or {}
    if not a['is_main'] and a['id'] != aid: return jsonify({'error': '无权限'}), 403
    conn = db()
    if a['is_main']:
        conn.execute("UPDATE admins SET name=?,permissions=? WHERE id=?",
                     (d.get('name',''), json.dumps(d.get('permissions',{}), ensure_ascii=False), aid))
    if d.get('password'):
        conn.execute("UPDATE admins SET password=? WHERE id=?", (sha(d['password']), aid))
    conn.commit(); conn.close()
    return jsonify({'success': True})

@app.route('/api/admins/<int:aid>', methods=['DELETE'])
@admin_required
def delete_admin(aid):
    a = get_me()
    if not a['is_main']: return jsonify({'error': '无权限'}), 403
    conn = db()
    conn.execute("DELETE FROM admins WHERE id=? AND is_main=0", (aid,))
    conn.commit(); conn.close()
    return jsonify({'success': True})

if __name__ == '__main__':
    init_db()
    print("\n" + "="*55)
    print("  🚀 ICode 签到管理系统 v3.2  — 批量操作版")
    print("  选手签到: http://localhost:5000/")
    print("  管理后台: http://localhost:5000/admin")
    print("  账号: admin  密码: admin123")
    print("="*55 + "\n")
    app.run(debug=True, host='0.0.0.0', port=5000)
