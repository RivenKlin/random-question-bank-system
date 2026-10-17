#!/bin/sh
set -e
cd /opt/random-question-bank
python3 -c 'import shutil, sqlite3; shutil.copy2("question_bank.db", "question_bank.db.pre-v1.0.1.bak") if __import__("os").path.exists("question_bank.db") else None'
python3 -c 'import zipfile; zipfile.ZipFile("deploy.zip").extractall(".")'
/opt/random-question-bank/.venv/bin/pip install -r requirements.txt
pkill -f 'python app.py' || true
nohup env PORT=8044 /opt/random-question-bank/.venv/bin/python app.py > /opt/random-question-bank/server.log 2>&1 </dev/null &
sleep 4
echo '---LOG---'
tail -40 /opt/random-question-bank/server.log
echo '---HTTP---'
python3 - <<'PY'
import urllib.request
for u in ('http://127.0.0.1:8044/','http://127.0.0.1:8044/api/chapters'):
    try:
        r=urllib.request.urlopen(u, timeout=5)
        print(u, r.status, r.read(300))
    except Exception as e:
        print(u, e)
PY