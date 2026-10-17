import os, re, json, random, sqlite3, secrets
from datetime import datetime
from flask import Flask, request, jsonify, send_from_directory, session
from functools import wraps
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import load_workbook

BASE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE, '已导入考试系统')
DB_PATH = os.path.join(BASE, 'question_bank.db')
app = Flask(__name__, static_folder='static')
app.secret_key = os.environ.get('FLASK_SECRET_KEY') or secrets.token_hex(32)
app.config.update(MAX_CONTENT_LENGTH=16 * 1024 * 1024, SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SAMESITE='Lax', SESSION_COOKIE_SECURE=os.environ.get('COOKIE_SECURE') == '1')

SCHEMA = '''
CREATE TABLE IF NOT EXISTS users (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE NOT NULL, password_hash TEXT NOT NULL, role TEXT NOT NULL DEFAULT 'student', created_at TEXT);
CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS questions (
 id INTEGER PRIMARY KEY AUTOINCREMENT, type TEXT NOT NULL, chapter TEXT NOT NULL,
 score REAL DEFAULT 1, level TEXT, question TEXT NOT NULL, answer TEXT NOT NULL,
 options TEXT DEFAULT '{}', explanation TEXT DEFAULT '', source TEXT DEFAULT '',
 created_at TEXT, updated_at TEXT
);
CREATE TABLE IF NOT EXISTS favorites (user_id INTEGER NOT NULL, question_id INTEGER NOT NULL, created_at TEXT, PRIMARY KEY(user_id, question_id), FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE, FOREIGN KEY(question_id) REFERENCES questions(id) ON DELETE CASCADE);
CREATE TABLE IF NOT EXISTS wrong_questions (user_id INTEGER NOT NULL, question_id INTEGER NOT NULL, wrong_count INTEGER DEFAULT 1, last_answer TEXT, updated_at TEXT, PRIMARY KEY(user_id, question_id), FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE, FOREIGN KEY(question_id) REFERENCES questions(id) ON DELETE CASCADE);
CREATE TABLE IF NOT EXISTS records (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER NOT NULL, mode TEXT, total INTEGER, correct INTEGER, score REAL, question_ids TEXT, answers TEXT, created_at TEXT, FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE);
CREATE TABLE IF NOT EXISTS progress (
 user_id INTEGER NOT NULL, device_id TEXT NOT NULL, state TEXT NOT NULL, updated_at TEXT NOT NULL,
 PRIMARY KEY(user_id, device_id), FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE
);
'''

def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute('PRAGMA foreign_keys=ON')
    return conn

def now(): return datetime.now().isoformat(timespec='seconds')
def clean(v): return '' if v is None else str(v).strip()
def normalize_answer(v, qtype):
    s = clean(v).upper().replace('，', ',').replace('、', ',').replace(' ', '')
    if qtype == '判断':
        return {'T':'对','F':'错','TRUE':'对','FALSE':'错','Y':'对','N':'错'}.get(s, s)
    return ''.join(sorted(set(re.findall('[A-E]', s))))
def row_question(row, headers, qtype, source):
    vals = {headers[i]: row[i] if i < len(row) else None for i in range(len(headers))}
    question = clean(vals.get('question') or vals.get('题目') or vals.get('问题'))
    if not question: return None
    ans = normalize_answer(vals.get('answer') or vals.get('答案'), qtype)
    if not ans: return None
    options = {x: clean(vals.get(x.lower()) or vals.get(x)) for x in 'ABCDE' if clean(vals.get(x.lower()) or vals.get(x))}
    if qtype == '判断' and not options: options = {'T': '对', 'F': '错'}
    explanation = clean(vals.get('attachment') or vals.get('解析') or vals.get('explanation') or vals.get('analysis'))
    chapter = clean(vals.get('chapter') or vals.get('章节')) or clean(vals.get('章节名称')) or os.path.splitext(os.path.basename(source))[0]
    if chapter.lower() == 'undefined': chapter = os.path.splitext(os.path.basename(source))[0]
    return (qtype, chapter, clean(vals.get('score')) or '1', clean(vals.get('level')), question, ans, json.dumps(options, ensure_ascii=False), explanation, os.path.basename(source), now(), now())

def import_excel(clear=False):
    conn = db(); conn.executescript(SCHEMA)
    if clear: conn.execute('DELETE FROM questions')
    count, skipped, files = 0, 0, 0
    for path in sorted([os.path.join(DATA_DIR, x) for x in os.listdir(DATA_DIR) if x.lower().endswith('.xlsx')]):
        files += 1
        try: wb = load_workbook(path, read_only=True, data_only=True)
        except Exception: skipped += 1; continue
        for ws in wb.worksheets:
            qtype = {'choice':'单选','mchoice':'多选','judgement':'判断'}.get(ws.title.lower())
            if not qtype: continue
            rows = ws.iter_rows(values_only=True)
            try: headers = [clean(x).lower() for x in next(rows)]
            except StopIteration: continue
            for row in rows:
                item = row_question(row, headers, qtype, path)
                if item:
                    conn.execute('INSERT INTO questions(type,chapter,score,level,question,answer,options,explanation,source,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?)', item); count += 1
                else: skipped += 1
    conn.commit(); conn.close(); return {'files':files,'imported':count,'skipped':skipped}

def ensure_db():
    conn = db()
    conn.execute('BEGIN IMMEDIATE')
    try:
        for table, required in {
            'favorites': {'user_id'}, 'wrong_questions': {'user_id'},
            'records': {'user_id'}, 'progress': {'user_id'},
        }.items():
            exists = conn.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (table,)).fetchone()
            if exists:
                columns = {row['name'] for row in conn.execute(f'PRAGMA table_info({table})')}
                if not required.issubset(columns):
                    legacy = 'legacy_' + table
                    if conn.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (legacy,)).fetchone():
                        legacy += '_v1'
                    conn.execute(f'ALTER TABLE {table} RENAME TO {legacy}')
        for statement in SCHEMA.split(';'):
            if statement.strip(): conn.execute(statement)
        conn.commit()
    except Exception:
        conn.rollback(); conn.close(); raise
    if not conn.execute("SELECT 1 FROM users WHERE username='admin'").fetchone():
        conn.execute('INSERT INTO users(username,password_hash,role,created_at) VALUES(?,?,?,?)', ('admin', generate_password_hash(os.environ.get('ADMIN_PASSWORD', 'admin123')), 'admin', now()))
    conn.execute("UPDATE questions SET chapter=COALESCE(NULLIF(source,''),'未分类') WHERE chapter IS NULL OR TRIM(chapter)='' OR LOWER(chapter)='undefined'")
    conn.commit()
    n = conn.execute('SELECT COUNT(*) c FROM questions').fetchone()['c']
    missing = conn.execute("SELECT COUNT(*) c FROM questions WHERE options IS NULL OR options='{}' OR options=''").fetchone()['c']
    if n == 0:
        conn.close()
        import_excel(clear=False)
        conn = db()
    conn.execute("UPDATE questions SET options=? WHERE type='判断' AND (options IS NULL OR options='{}' OR options='')", (json.dumps({'T':'对','F':'错'}, ensure_ascii=False),))
    conn.commit(); conn.close()
ensure_db()

def qdict(row, include_answer=True):
    d = dict(row); d['options'] = json.loads(d.get('options') or '{}')
    d['is_favorite'] = bool(d.pop('favorite', 0)); d['is_wrong'] = bool(d.pop('wrong', 0))
    if not include_answer: d.pop('answer', None); d.pop('explanation', None)
    return d

def login_required(admin_only=False):
    def decorate(fn):
        @wraps(fn)
        def wrapped(*args, **kwargs):
            if not session.get('user_id') or (admin_only and session.get('role') != 'admin'):
                return jsonify({'error': '需要登录或管理员权限'}), 401
            return fn(*args, **kwargs)
        return wrapped
    return decorate

@app.route('/api/auth/me')
def auth_me(): return jsonify({'user': session.get('username'), 'role': session.get('role')})

@app.after_request
def no_cache_static(response):
    if request.path.startswith('/static/'):
        response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    return response

@app.route('/api/auth/login', methods=['POST'])
def auth_login():
    data=request.get_json(silent=True) or {}; c=db(); user=c.execute('SELECT * FROM users WHERE username=?',(clean(data.get('username')),)).fetchone(); c.close()
    if not user or not check_password_hash(user['password_hash'], clean(data.get('password'))): return jsonify({'error':'用户名或密码错误'}),401
    session.clear(); session['user_id']=user['id']; session['username']=user['username']; session['role']=user['role']
    return jsonify({'user':user['username'],'role':user['role']})

@app.route('/api/auth/change-password', methods=['POST'])
@login_required()
def change_password():
    data = request.get_json(silent=True) or {}
    current = clean(data.get('current_password'))
    new_password = clean(data.get('new_password'))
    c = db()
    user = c.execute('SELECT password_hash FROM users WHERE id=?', (session['user_id'],)).fetchone()
    if not user or not check_password_hash(user['password_hash'], current):
        c.close(); return jsonify({'error': '当前密码错误'}), 400
    if len(new_password) < 12:
        c.close(); return jsonify({'error': '新密码至少12位'}), 400
    c.execute('UPDATE users SET password_hash=? WHERE id=?', (generate_password_hash(new_password), session['user_id']))
    c.commit(); c.close(); session.clear()
    return jsonify({'ok': True, 'reauthenticate': True})

@app.route('/api/auth/register', methods=['POST'])
def auth_register():
    c=db(); admin_exists=c.execute("SELECT 1 FROM users WHERE role='admin'").fetchone(); c.close()
    if not admin_exists:
        return jsonify({'error':'管理员账户尚未初始化'}), 503
    data=request.get_json(silent=True) or {}; username=clean(data.get('username')); password=clean(data.get('password'))
    if not re.fullmatch(r'[A-Za-z0-9_]{3,32}',username) or len(password)<6: return jsonify({'error':'用户名需为3-32位字母、数字或下划线，密码至少6位'}),400
    c=db()
    try: c.execute('INSERT INTO users(username,password_hash,role,created_at) VALUES(?,?,?,?)',(username,generate_password_hash(password),'student',now())); c.commit()
    except sqlite3.IntegrityError: c.close(); return jsonify({'error':'用户名已存在'}),409
    c.close(); return jsonify({'ok':True})

@app.route('/api/auth/logout', methods=['POST'])
def auth_logout(): session.clear(); return jsonify({'ok':True})

@app.route('/')
def index(): return send_from_directory('static', 'index.html')
@app.route('/admin')
def admin(): return send_from_directory('static', 'admin.html')
@app.route('/result')
def result_page(): return send_from_directory('static', 'result.html')
def progress_key():
    return clean(request.headers.get('X-Device-ID'))

@app.route('/api/progress', methods=['GET', 'PUT', 'DELETE'])
@login_required()
def progress():
    device_id = progress_key()
    user_id = session['user_id']
    if not device_id:
        return jsonify({'error': 'X-Device-ID is required'}), 400
    c = db()
    if request.method == 'GET':
        row = c.execute('SELECT state FROM progress WHERE user_id=? AND device_id=?', (user_id, device_id)).fetchone()
        c.close()
        if not row:
            return jsonify({'state': None})
        try:
            return jsonify({'state': json.loads(row['state'])})
        except json.JSONDecodeError:
            return jsonify({'state': None})
    if request.method == 'DELETE':
        c.execute('DELETE FROM progress WHERE user_id=? AND device_id=?', (user_id, device_id))
    else:
        state = request.get_json(silent=True)
        if not isinstance(state, dict):
            c.close()
            return jsonify({'error': 'state must be a JSON object'}), 400
        c.execute('INSERT INTO progress(user_id,device_id,state,updated_at) VALUES(?,?,?,?) ON CONFLICT(user_id,device_id) DO UPDATE SET state=excluded.state,updated_at=excluded.updated_at', (user_id, device_id, json.dumps(state, ensure_ascii=False), now()))
    c.commit(); c.close()
    return jsonify({'ok': True})

@app.route('/api/stats')
@login_required(admin_only=True)
def stats():
    c=db(); total=c.execute('SELECT COUNT(*) n FROM questions').fetchone()['n']; types=c.execute('SELECT type,COUNT(*) n FROM questions GROUP BY type').fetchall(); chapters=c.execute('SELECT chapter,COUNT(*) n FROM questions GROUP BY chapter ORDER BY chapter').fetchall(); c.close(); return jsonify({'total':total,'types':{r['type']:r['n'] for r in types},'chapters':[dict(r) for r in chapters]})
@app.route('/api/chapters')
def chapters():
    c=db(); rows=c.execute("SELECT chapter,COUNT(*) count FROM questions WHERE chapter IS NOT NULL AND TRIM(chapter)<>'' AND LOWER(TRIM(chapter))<>'undefined' GROUP BY chapter ORDER BY chapter").fetchall(); c.close(); return jsonify([dict(r) for r in rows])
@app.route('/api/questions')
@login_required(admin_only=True)
def questions():
    c=db(); where=[]; args=[]; search=clean(request.args.get('search')); typ=clean(request.args.get('type')); types=request.args.getlist('types'); chapter=clean(request.args.get('chapter'))
    if search: where.append('(question LIKE ? OR chapter LIKE ?)'); args += [f'%{search}%',f'%{search}%']
    if types: where.append('type IN (%s)' % ','.join('?' * len(types))); args += types
    elif typ: where.append('type=?'); args.append(typ)
    if chapter: where.append('chapter=?'); args.append(chapter)
    sql='SELECT * FROM questions WHERE '+(' AND '.join(where) if where else '1=1')+' ORDER BY id DESC LIMIT 500'; rows=c.execute(sql,args).fetchall(); c.close(); return jsonify([qdict(r) for r in rows])
@app.route('/api/questions/<int:qid>', methods=['POST','PUT','DELETE'])
@login_required(admin_only=True)
def question_detail(qid):
    c=db()
    if request.method=='DELETE': c.execute('DELETE FROM questions WHERE id=?',(qid,)); c.commit(); c.close(); return jsonify({'ok':True})
    data=request.json or {}; typ=data.get('type','单选'); options=data.get('options',{})
    fields=(typ, data.get('chapter','未分类'), data.get('score',1), data.get('level',''), data.get('question',''), normalize_answer(data.get('answer',''),typ), json.dumps(options,ensure_ascii=False), data.get('explanation',''), data.get('source','手工录入'), now())
    if request.method=='POST': c.execute('INSERT INTO questions(type,chapter,score,level,question,answer,options,explanation,source,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?)',fields)
    else: c.execute('UPDATE questions SET type=?,chapter=?,score=?,level=?,question=?,answer=?,options=?,explanation=?,source=?,updated_at=? WHERE id=?',fields+(qid,))
    c.commit(); c.close(); return jsonify({'ok':True})
@app.route('/api/import', methods=['POST'])
@login_required(admin_only=True)
def do_import(): return jsonify(import_excel(bool((request.json or {}).get('clear'))))

@app.route('/api/import-mapped', methods=['POST'])
@login_required(admin_only=True)
def import_mapped():
    data = request.get_json(silent=True) or {}
    questions = data.get('questions')
    if not isinstance(questions, list):
        return jsonify({'error': 'questions must be a list'}), 400
    c = db(); imported = skipped = 0
    try:
        for item in questions:
            if not isinstance(item, dict):
                skipped += 1; continue
            qtype = item.get('type')
            question = clean(item.get('question'))
            answer = normalize_answer(item.get('answer'), qtype)
            if qtype not in ('单选', '多选', '判断') or not question or not answer:
                skipped += 1; continue
            options = item.get('options') if isinstance(item.get('options'), dict) else {}
            if qtype == '判断' and not options:
                options = {'T': '对', 'F': '错'}
            c.execute('INSERT INTO questions(type,chapter,score,level,question,answer,options,explanation,source,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?)', (qtype, clean(item.get('chapter')) or '未分类', 1, '', question, answer, json.dumps(options, ensure_ascii=False), clean(item.get('explanation')), '上传Excel（字段映射）', now(), now()))
            imported += 1
        c.commit()
    except Exception:
        c.rollback(); c.close(); raise
    c.close()
    return jsonify({'imported': imported, 'skipped': skipped})

@app.route('/api/upload-import', methods=['POST'])
@login_required(admin_only=True)
def upload_import():
    files=request.files.getlist('files'); chapter=clean(request.form.get('chapter')); count=skipped=0
    for file in files:
        if not file.filename or not file.filename.lower().endswith('.xlsx'): skipped+=1; continue
        try: wb=load_workbook(file,read_only=True,data_only=True)
        except Exception: skipped+=1; continue
        c=db()
        for ws in wb.worksheets:
            qtype={'choice':'单选','mchoice':'多选','judgement':'判断'}.get(ws.title.lower())
            if not qtype: continue
            rows=ws.iter_rows(values_only=True)
            try: headers=[clean(x).lower() for x in next(rows)]
            except StopIteration: continue
            for row in rows:
                item=row_question(row,headers,qtype,file.filename)
                if item:
                    if chapter and not (clean(row[headers.index('chapter')]) if 'chapter' in headers and headers.index('chapter')<len(row) else clean(row[headers.index('章节')]) if '章节' in headers and headers.index('章节')<len(row) else ''): item=(item[0],chapter)+item[2:]
                    c.execute('INSERT INTO questions(type,chapter,score,level,question,answer,options,explanation,source,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?)',item); count+=1
                else: skipped+=1
        c.commit(); c.close()
    return jsonify({'files':len(files),'imported':count,'skipped':skipped})
@app.route('/api/quiz', methods=['POST'])
@login_required()
def quiz():
    d=request.json or {}; types=d.get('types') or []
    counts=d.get('counts')
    base=[]; args=[]
    if counts is None:
        n=max(1,min(int(d.get('count',10)),100))
        if types: base.append('type IN (%s)'%','.join('?'*len(types))); args += types
        if d.get('chapter'): base.append('chapter=?'); args.append(d['chapter'])
        if d.get('wrong'): base.append('id IN (SELECT question_id FROM wrong_questions WHERE user_id=?)'); args.append(session['user_id'])
        if d.get('favorite'): base.append('id IN (SELECT question_id FROM favorites WHERE user_id=?)'); args.append(session['user_id'])
        c=db(); rows=c.execute('SELECT * FROM questions WHERE '+(' AND '.join(base) if base else '1=1')+' ORDER BY RANDOM() LIMIT ?',args+[n]).fetchall(); c.close(); return jsonify([qdict(r,False) for r in rows])
    requested = {t: max(0, min(int(counts.get(t, 0)), 100)) for t in ('单选','多选','判断')}
    if d.get('chapter'): base.append('chapter=?'); args.append(d['chapter'])
    if d.get('wrong'): base.append('id IN (SELECT question_id FROM wrong_questions WHERE user_id=?)'); args.append(session['user_id'])
    if d.get('favorite'): base.append('id IN (SELECT question_id FROM favorites WHERE user_id=?)'); args.append(session['user_id'])
    c=db(); rows=[]
    for typ, amount in requested.items():
        if amount <= 0: continue
        where=base + ['type=?']; query_args=args + [typ]
        rows.extend(c.execute('SELECT * FROM questions WHERE '+ ' AND '.join(where) +' ORDER BY RANDOM() LIMIT ?', query_args+[amount]).fetchall())
    c.close(); random.shuffle(rows); return jsonify([qdict(r,False) for r in rows])
@app.route('/api/submit', methods=['POST'])
@login_required()
def submit():
    d=request.json or {}; answers=d.get('answers',{}); ids=[int(x) for x in d.get('question_ids',[])]; c=db(); result=[]; correct=0
    for qid in ids:
        r=c.execute('SELECT * FROM questions WHERE id=?',(qid,)).fetchone()
        if not r: continue
        got=normalize_answer(answers.get(str(qid),answers.get(qid,'')),r['type']); ok=got==r['answer']; correct += ok
        if not ok: c.execute('INSERT INTO wrong_questions(user_id,question_id,wrong_count,last_answer,updated_at) VALUES(?,?,?,?,?) ON CONFLICT(user_id,question_id) DO UPDATE SET wrong_count=wrong_count+1,last_answer=excluded.last_answer,updated_at=excluded.updated_at',(session['user_id'],qid,1,got,now()))
        result.append({'id':qid,'correct':ok,'answer':r['answer'],'explanation':r['explanation'],'user_answer':got})
    score=round(sum(float(c.execute('SELECT score FROM questions WHERE id=?',(x,)).fetchone()['score'] or 1) for x in ids if any(z['id']==x and z['correct'] for z in result)),2)
    c.execute('INSERT INTO records(user_id,mode,total,correct,score,question_ids,answers,created_at) VALUES(?,?,?,?,?,?,?,?)',(session['user_id'],d.get('mode','练习'),len(ids),correct,score,json.dumps(ids),json.dumps(answers,ensure_ascii=False),now())); c.commit(); c.close(); return jsonify({'total':len(ids),'correct':correct,'score':score,'results':result})
@app.route('/api/favorites/<int:qid>', methods=['POST','DELETE'])
@login_required()
def favorite(qid):
    c=db(); user_id=session['user_id']; c.execute('DELETE FROM favorites WHERE user_id=? AND question_id=?',(user_id,qid));
    if request.method=='POST': c.execute('INSERT INTO favorites(user_id,question_id,created_at) VALUES(?,?,?)',(user_id,qid,now()))
    c.commit(); c.close(); return jsonify({'ok':True,'favorite':request.method=='POST'})
@app.route('/api/review/<kind>')
@login_required()
def review(kind):
    if kind not in ('wrong','favorite'): return jsonify([]),400
    table='wrong_questions' if kind=='wrong' else 'favorites'; order_col='updated_at' if kind=='wrong' else 'created_at'; c=db(); rows=c.execute(f'SELECT q.*,1 as {kind} FROM questions q JOIN {table} x ON x.question_id=q.id WHERE x.user_id=? ORDER BY x.{order_col} DESC',(session['user_id'],)).fetchall(); c.close(); return jsonify([qdict(r,False) for r in rows])
@app.route('/api/records')
@login_required(admin_only=True)
def records():
    c=db(); rows=c.execute('SELECT r.*,u.username FROM records r JOIN users u ON u.id=r.user_id ORDER BY r.id DESC LIMIT 50').fetchall(); c.close(); return jsonify([dict(r) for r in rows])
if __name__=='__main__': app.run(host=os.environ.get('HOST','127.0.0.1'),port=int(os.environ.get('PORT',5000)),debug=False)
