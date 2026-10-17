import os, re, json, random, sqlite3
from datetime import datetime
from flask import Flask, request, jsonify, send_from_directory
from openpyxl import load_workbook

BASE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE, '已导入考试系统')
DB_PATH = os.path.join(BASE, 'question_bank.db')
app = Flask(__name__, static_folder='static')

SCHEMA = '''
CREATE TABLE IF NOT EXISTS questions (
 id INTEGER PRIMARY KEY AUTOINCREMENT, type TEXT NOT NULL, chapter TEXT NOT NULL,
 score REAL DEFAULT 1, level TEXT, question TEXT NOT NULL, answer TEXT NOT NULL,
 options TEXT DEFAULT '{}', explanation TEXT DEFAULT '', source TEXT DEFAULT '',
 created_at TEXT, updated_at TEXT
);
CREATE TABLE IF NOT EXISTS favorites (question_id INTEGER PRIMARY KEY, created_at TEXT, FOREIGN KEY(question_id) REFERENCES questions(id) ON DELETE CASCADE);
CREATE TABLE IF NOT EXISTS wrong_questions (question_id INTEGER PRIMARY KEY, wrong_count INTEGER DEFAULT 1, last_answer TEXT, updated_at TEXT, FOREIGN KEY(question_id) REFERENCES questions(id) ON DELETE CASCADE);
CREATE TABLE IF NOT EXISTS records (id INTEGER PRIMARY KEY AUTOINCREMENT, mode TEXT, total INTEGER, correct INTEGER, score REAL, question_ids TEXT, answers TEXT, created_at TEXT);
CREATE TABLE IF NOT EXISTS progress (
 device_id TEXT NOT NULL, ip TEXT NOT NULL, state TEXT NOT NULL, updated_at TEXT NOT NULL,
 UNIQUE(device_id, ip)
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
    explanation = clean(vals.get('attachment') or vals.get('解析') or vals.get('explanation'))
    return (qtype, os.path.splitext(os.path.basename(source))[0], clean(vals.get('score')) or '1', clean(vals.get('level')), question, ans, json.dumps(options, ensure_ascii=False), explanation, os.path.basename(source), now(), now())

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
    conn = db(); conn.executescript(SCHEMA)
    n = conn.execute('SELECT COUNT(*) c FROM questions').fetchone()['c']
    missing = conn.execute("SELECT COUNT(*) c FROM questions WHERE options IS NULL OR options='{}' OR options=''").fetchone()['c']
    if n == 0 or missing:
        conn.close()
        import_excel(clear=True)
        conn = db()
    conn.execute("UPDATE questions SET options=? WHERE type='判断' AND (options IS NULL OR options='{}' OR options='')", (json.dumps({'T':'对','F':'错'}, ensure_ascii=False),))
    conn.commit(); conn.close()
ensure_db()

def qdict(row, include_answer=True):
    d = dict(row); d['options'] = json.loads(d.get('options') or '{}')
    d['is_favorite'] = bool(d.pop('favorite', 0)); d['is_wrong'] = bool(d.pop('wrong', 0))
    if not include_answer: d.pop('answer', None); d.pop('explanation', None)
    return d

@app.route('/')
def index(): return send_from_directory('static', 'index.html')
@app.route('/admin')
def admin(): return send_from_directory('static', 'admin.html')
@app.route('/result')
def result_page(): return send_from_directory('static', 'result.html')
def client_ip():
    return request.headers.get('X-Forwarded-For', request.remote_addr or '').split(',')[0].strip()

def progress_key():
    return clean(request.headers.get('X-Device-ID')), client_ip()

@app.route('/api/progress', methods=['GET', 'PUT', 'DELETE'])
def progress():
    device_id, ip = progress_key()
    if not device_id:
        return jsonify({'error': 'X-Device-ID is required'}), 400
    c = db()
    if request.method == 'GET':
        row = c.execute('SELECT state FROM progress WHERE device_id=? AND ip=?', (device_id, ip)).fetchone()
        c.close()
        if not row:
            return jsonify({'state': None})
        try:
            return jsonify({'state': json.loads(row['state'])})
        except json.JSONDecodeError:
            return jsonify({'state': None})
    if request.method == 'DELETE':
        c.execute('DELETE FROM progress WHERE device_id=? AND ip=?', (device_id, ip))
    else:
        state = request.get_json(silent=True)
        if not isinstance(state, dict):
            c.close()
            return jsonify({'error': 'state must be a JSON object'}), 400
        c.execute('INSERT INTO progress(device_id,ip,state,updated_at) VALUES(?,?,?,?) ON CONFLICT(device_id,ip) DO UPDATE SET state=excluded.state,updated_at=excluded.updated_at', (device_id, ip, json.dumps(state, ensure_ascii=False), now()))
    c.commit(); c.close()
    return jsonify({'ok': True})

@app.route('/api/stats')
def stats():
    c=db(); total=c.execute('SELECT COUNT(*) n FROM questions').fetchone()['n']; types=c.execute('SELECT type,COUNT(*) n FROM questions GROUP BY type').fetchall(); chapters=c.execute('SELECT chapter,COUNT(*) n FROM questions GROUP BY chapter ORDER BY chapter').fetchall(); c.close(); return jsonify({'total':total,'types':{r['type']:r['n'] for r in types},'chapters':[dict(r) for r in chapters]})
@app.route('/api/chapters')
def chapters(): return jsonify([dict(r) for r in db().execute('SELECT chapter,COUNT(*) count FROM questions GROUP BY chapter ORDER BY chapter').fetchall()])
@app.route('/api/questions')
def questions():
    c=db(); where=[]; args=[]; search=clean(request.args.get('search')); typ=clean(request.args.get('type')); types=request.args.getlist('types'); chapter=clean(request.args.get('chapter'))
    if search: where.append('(question LIKE ? OR chapter LIKE ?)'); args += [f'%{search}%',f'%{search}%']
    if types: where.append('type IN (%s)' % ','.join('?' * len(types))); args += types
    elif typ: where.append('type=?'); args.append(typ)
    if chapter: where.append('chapter=?'); args.append(chapter)
    sql='SELECT * FROM questions WHERE '+(' AND '.join(where) if where else '1=1')+' ORDER BY id DESC LIMIT 500'; rows=c.execute(sql,args).fetchall(); c.close(); return jsonify([qdict(r) for r in rows])
@app.route('/api/questions/<int:qid>', methods=['POST','PUT','DELETE'])
def question_detail(qid):
    c=db()
    if request.method=='DELETE': c.execute('DELETE FROM questions WHERE id=?',(qid,)); c.commit(); c.close(); return jsonify({'ok':True})
    data=request.json or {}; typ=data.get('type','单选'); options=data.get('options',{})
    fields=(typ, data.get('chapter','未分类'), data.get('score',1), data.get('level',''), data.get('question',''), normalize_answer(data.get('answer',''),typ), json.dumps(options,ensure_ascii=False), data.get('explanation',''), data.get('source','手工录入'), now())
    if request.method=='POST': c.execute('INSERT INTO questions(type,chapter,score,level,question,answer,options,explanation,source,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?)',fields)
    else: c.execute('UPDATE questions SET type=?,chapter=?,score=?,level=?,question=?,answer=?,options=?,explanation=?,source=?,updated_at=? WHERE id=?',fields+(qid,))
    c.commit(); c.close(); return jsonify({'ok':True})
@app.route('/api/import', methods=['POST'])
def do_import(): return jsonify(import_excel(bool((request.json or {}).get('clear'))))
@app.route('/api/quiz', methods=['POST'])
def quiz():
    d=request.json or {}; types=d.get('types') or []
    counts=d.get('counts')
    base=[]; args=[]
    if counts is None:
        n=max(1,min(int(d.get('count',10)),100))
        if types: base.append('type IN (%s)'%','.join('?'*len(types))); args += types
        if d.get('chapter'): base.append('chapter=?'); args.append(d['chapter'])
        if d.get('wrong'): base.append('id IN (SELECT question_id FROM wrong_questions)')
        if d.get('favorite'): base.append('id IN (SELECT question_id FROM favorites)')
        c=db(); rows=c.execute('SELECT * FROM questions WHERE '+(' AND '.join(base) if base else '1=1')+' ORDER BY RANDOM() LIMIT ?',args+[n]).fetchall(); c.close(); return jsonify([qdict(r,False) for r in rows])
    requested = {t: max(0, min(int(counts.get(t, 0)), 100)) for t in ('单选','多选','判断')}
    if d.get('chapter'): base.append('chapter=?'); args.append(d['chapter'])
    if d.get('wrong'): base.append('id IN (SELECT question_id FROM wrong_questions)')
    if d.get('favorite'): base.append('id IN (SELECT question_id FROM favorites)')
    c=db(); rows=[]
    for typ, amount in requested.items():
        if amount <= 0: continue
        where=base + ['type=?']; query_args=args + [typ]
        rows.extend(c.execute('SELECT * FROM questions WHERE '+ ' AND '.join(where) +' ORDER BY RANDOM() LIMIT ?', query_args+[amount]).fetchall())
    c.close(); random.shuffle(rows); return jsonify([qdict(r,False) for r in rows])
@app.route('/api/submit', methods=['POST'])
def submit():
    d=request.json or {}; answers=d.get('answers',{}); ids=[int(x) for x in d.get('question_ids',[])]; c=db(); result=[]; correct=0
    for qid in ids:
        r=c.execute('SELECT * FROM questions WHERE id=?',(qid,)).fetchone()
        if not r: continue
        got=normalize_answer(answers.get(str(qid),answers.get(qid,'')),r['type']); ok=got==r['answer']; correct += ok
        if not ok: c.execute('INSERT INTO wrong_questions(question_id,wrong_count,last_answer,updated_at) VALUES(?,?,?,?) ON CONFLICT(question_id) DO UPDATE SET wrong_count=wrong_count+1,last_answer=excluded.last_answer,updated_at=excluded.updated_at',(qid,1,got,now()))
        result.append({'id':qid,'correct':ok,'answer':r['answer'],'explanation':r['explanation'],'user_answer':got})
    score=round(sum(float(c.execute('SELECT score FROM questions WHERE id=?',(x,)).fetchone()['score'] or 1) for x in ids if any(z['id']==x and z['correct'] for z in result)),2)
    c.execute('INSERT INTO records(mode,total,correct,score,question_ids,answers,created_at) VALUES(?,?,?,?,?,?,?)',(d.get('mode','练习'),len(ids),correct,score,json.dumps(ids),json.dumps(answers,ensure_ascii=False),now())); c.commit(); c.close(); return jsonify({'total':len(ids),'correct':correct,'score':score,'results':result})
@app.route('/api/favorites/<int:qid>', methods=['POST','DELETE'])
def favorite(qid):
    c=db(); c.execute('DELETE FROM favorites WHERE question_id=?',(qid,));
    if request.method=='POST': c.execute('INSERT INTO favorites VALUES(?,?)',(qid,now()))
    c.commit(); c.close(); return jsonify({'ok':True,'favorite':request.method=='POST'})
@app.route('/api/review/<kind>')
def review(kind):
    if kind not in ('wrong','favorite'): return jsonify([]),400
    table='wrong_questions' if kind=='wrong' else 'favorites'; order_col='updated_at' if kind=='wrong' else 'created_at'; c=db(); rows=c.execute(f'SELECT q.*,1 as {kind} FROM questions q JOIN {table} x ON x.question_id=q.id ORDER BY x.{order_col} DESC').fetchall(); c.close(); return jsonify([qdict(r,False) for r in rows])
@app.route('/api/records')
def records(): return jsonify([dict(r) for r in db().execute('SELECT * FROM records ORDER BY id DESC LIMIT 50').fetchall()])
if __name__=='__main__': app.run(host='127.0.0.1',port=int(os.environ.get('PORT',5000)),debug=False)
